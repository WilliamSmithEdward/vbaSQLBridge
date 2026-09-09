"""Just enough of a SQL Server driver to log in and ask something.

This speaks the sequence the reference capture recorded: PRELOGIN, a TLS
handshake whose records travel inside TDS packets, LOGIN7 through that
tunnel carrying a token Windows itself produced, then the SSPI leg in the
clear. Nothing is mocked; the point is that a client's real sequence gets a
real login back.
"""
from __future__ import annotations

import datetime
import socket
import ssl
import struct
import time
import uuid

HEADER = struct.Struct(">BBHHBB")

PRELOGIN = 0x12
SQL_BATCH = 0x01
TABULAR_RESULT = 0x04
LOGIN7 = 0x10
SSPI = 0x11

TOKEN_COL_METADATA = 0x81
TOKEN_ERROR = 0xAA
TOKEN_INFO = 0xAB
TOKEN_LOGIN_ACK = 0xAD
TOKEN_ROW = 0xD1
TOKEN_ENV_CHANGE = 0xE3
TOKEN_SSPI = 0xED
TOKEN_DONE = 0xFD

# The six options a modern driver asks about, in the order the capture
# recorded: version 18.7.5, encryption off, no instance, a thread id, MARS
# off, and a trace id.
CLIENT_PRELOGIN_PAYLOAD = bytes.fromhex(
    "00001f000601002500010200260001030027000404002b000105002c0024ff"
    "1207000500000000483e0000004f9ccf1f076f8c4a8a2c701cf600230e1141"
    "e2d4671440499c86b81a5580359002000000"
)


def frame(packet_type: int, payload: bytes, packet_id: int = 1) -> bytes:
    return HEADER.pack(packet_type, 1, len(payload) + 8, 0, packet_id, 0) + payload


def connect(port: int, timeout: float = 20.0) -> socket.socket:
    """Keep trying until the server inside Excel is listening."""
    deadline = time.monotonic() + timeout
    while True:
        try:
            connection = socket.create_connection(("127.0.0.1", port), timeout=10)
            connection.settimeout(30)
            return connection
        except OSError:
            if time.monotonic() > deadline:
                raise
            time.sleep(0.2)


class TdsClient:
    def __init__(self, port: int) -> None:
        self.socket = connect(port)
        self.buffer = b""
        self.tls: ssl.SSLObject | None = None
        self.incoming = ssl.MemoryBIO()
        self.outgoing = ssl.MemoryBIO()
        self.encrypted = False

    def close(self) -> None:
        self.socket.close()

    # -- framing ---------------------------------------------------------
    def send(self, packet_type: int, payload: bytes, packet_id: int = 1) -> None:
        message = frame(packet_type, payload, packet_id)
        if self.encrypted:
            self.tls.write(message)
            message = self.outgoing.read()
        self.socket.sendall(message)

    def read_message(self) -> tuple[int, bytes]:
        payload, message_type = b"", None
        while True:
            while len(self.buffer) < 8:
                self._fill()
            kind, status, length = HEADER.unpack(self.buffer[:8])[:3]
            while len(self.buffer) < length:
                self._fill()
            if message_type is None:
                message_type = kind
            payload += self.buffer[8:length]
            self.buffer = self.buffer[length:]
            if status & 1:
                return message_type, payload

    def _fill(self) -> None:
        chunk = self.socket.recv(65536)
        if not chunk:
            raise ConnectionError("the server closed the connection")
        if self.encrypted:
            self.incoming.write(chunk)
            while True:
                try:
                    plain = self.tls.read()
                except ssl.SSLWantReadError:
                    break
                if not plain:
                    break
                self.buffer += plain
        else:
            self.buffer += chunk

    # -- the login sequence ----------------------------------------------
    def prelogin(self) -> tuple[int, bytes]:
        self.send(PRELOGIN, CLIENT_PRELOGIN_PAYLOAD)
        return self.read_message()

    def handshake(self) -> ssl.SSLObject:
        """A TLS handshake whose records travel inside TDS packets."""
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        self.tls = context.wrap_bio(self.incoming, self.outgoing,
                                    server_hostname="localhost")
        while True:
            try:
                self.tls.do_handshake()
            except ssl.SSLWantReadError:
                pass
            else:
                return self.tls
            records = self.outgoing.read()
            if records:
                self.socket.sendall(frame(PRELOGIN, records, packet_id=0))
            kind, payload = self.read_message()
            if kind != PRELOGIN:
                raise AssertionError(
                    f"expected framed TLS records, got packet type {kind}")
            self.incoming.write(payload)

    def login(self, port: int) -> list[tuple[int, bytes]]:
        """LOGIN7 and the SSPI legs, ending at the login response tokens."""
        import sspi
        import sspicon

        # The service principal name a client uses when it reached the server
        # by address. Nothing has registered one, so SPNEGO falls through to
        # NTLM, which is what the capture showed.
        initiator = sspi.ClientAuth(
            "Negotiate",
            targetspn=f"MSSQLSvc/127.0.0.1:{port}",
            scflags=sspicon.ISC_REQ_CONNECTION,
        )
        error, buffers = initiator.authorize(None)

        # The login goes through the tunnel; everything after it is cleartext
        # again, which is what "encryption off" means here.
        message = frame(LOGIN7, build_login7(buffers[0].Buffer))
        self.tls.write(message)
        self.socket.sendall(self.outgoing.read())

        while True:
            kind, payload = self.read_message()
            tokens = parse_tokens(payload)
            challenge = next((body for token, body in tokens
                              if token == TOKEN_SSPI), None)
            if challenge is None:
                return tokens
            error, buffers = initiator.authorize(challenge)
            # The client's half carries the blob raw: no token, no length.
            self.send(SSPI, buffers[0].Buffer)

    def query(self, sql: str) -> list[tuple[int, bytes]]:
        """A SQL batch, with the ALL_HEADERS block TDS 7.2 added."""
        headers = struct.pack("<IIHQI", 22, 18, 2, 0, 1)
        self.send(SQL_BATCH, headers + sql.encode("utf-16-le"))
        _, payload = self.read_message()
        return parse_tokens(payload)


def build_login7(sspi_blob: bytes, database: str = "master") -> bytes:
    """A LOGIN7 with the shape a Windows-authenticated client sends."""
    strings = [
        "WORKSTATION1",         # host name
        "",                     # user name, empty under Windows auth
        "",                     # password, likewise
        "vbaSQLBridge tests",   # application name
        "tcp:127.0.0.1",        # server name
        "",                     # extension
        "TdsClient",            # client interface name
        "",                     # language
        database,
    ]

    table, variable = b"", b""
    offset = 94
    for text in strings:
        encoded = text.encode("utf-16-le")
        table += struct.pack("<HH", offset if encoded else 0, len(text))
        variable += encoded
        offset += len(encoded)

    table += b"\x00" * 6                                  # client id
    table += struct.pack("<HH", offset, len(sspi_blob))
    variable += sspi_blob
    table += struct.pack("<HH", 0, 0)                     # AtchDBFile
    table += struct.pack("<HH", 0, 0)                     # ChangePassword
    table += struct.pack("<I", 0)                         # cbSSPILong

    fixed = struct.pack(
        "<5I4BiI",
        0x74000004,      # TDS 7.4
        4096,            # packet size
        0x07000000,      # client program version
        1234,            # client pid
        0,               # connection id
        0xE0, 0x03, 0x00, 0x00,
        0, 1033,
    )
    body = fixed + table + variable
    return struct.pack("<I", len(body) + 4) + body


# DONE and its two variants carry a fixed twelve-byte body rather than a
# counted one.
_FIXED_TOKENS = {TOKEN_DONE: 12, 0xFE: 12, 0xFF: 12}

# What a type declaration carries after its type byte.
_TYPE_EXTRA = {
    0x26: 1,    # INTN, one byte of width
    0x6D: 1,    # FLTN
    0x68: 1,    # BITN
    0x6F: 1,    # DATETIMN
    0x24: 1,    # GUIDN
    0x34: 0,    # INT2, fixed and never null
    0xE7: 7,    # NVARCHAR, two bytes of size and five of collation
}


def parse_tokens(payload: bytes) -> list[tuple[int, object]]:
    """Split a response into (token, body) pairs.

    COLMETADATA comes back as a list of column descriptions and ROW as a list
    of values, because neither can be skipped without being understood: their
    length is whatever the declaration says it is.
    """
    tokens: list[tuple[int, object]] = []
    columns: list[dict] = []
    cursor = 0
    while cursor < len(payload):
        token = payload[cursor]
        cursor += 1
        if token == TOKEN_COL_METADATA:
            columns, cursor = _read_columns(payload, cursor)
            tokens.append((token, columns))
        elif token == TOKEN_ROW:
            values, cursor = _read_row(payload, cursor, columns)
            tokens.append((token, values))
        elif token in _FIXED_TOKENS:
            size = _FIXED_TOKENS[token]
            tokens.append((token, payload[cursor:cursor + size]))
            cursor += size
        else:
            size = struct.unpack_from("<H", payload, cursor)[0]
            cursor += 2
            tokens.append((token, payload[cursor:cursor + size]))
            cursor += size
    return tokens


def _read_columns(payload: bytes, cursor: int) -> tuple[list[dict], int]:
    count = struct.unpack_from("<H", payload, cursor)[0]
    cursor += 2
    columns = []
    for _ in range(count):
        cursor += 4                                    # user type, TDS 7.2 on
        cursor += 2                                    # flags
        kind = payload[cursor]
        cursor += 1
        extra = _TYPE_EXTRA[kind]
        detail = payload[cursor:cursor + extra]
        cursor += extra
        size = detail[0] if extra == 1 else (
            struct.unpack_from("<H", detail, 0)[0] if kind == 0xE7 else 0)
        name_length = payload[cursor]
        cursor += 1
        name = payload[cursor:cursor + name_length * 2].decode("utf-16-le")
        cursor += name_length * 2
        columns.append({"name": name, "type": kind, "size": size})
    return columns, cursor


def _read_row(payload: bytes, cursor: int, columns: list[dict]) -> tuple[list, int]:
    values = []
    for column in columns:
        value, cursor = _read_value(payload, cursor, column)
        values.append(value)
    return values, cursor


def _read_value(payload: bytes, cursor: int, column: dict):
    kind = column["type"]

    if kind == 0x34:                                   # INT2, no length
        return struct.unpack_from("<h", payload, cursor)[0], cursor + 2

    if kind == 0xE7:                                   # NVARCHAR
        if column["size"] == 0xFFFF:                   # the MAX form
            total = struct.unpack_from("<Q", payload, cursor)[0]
            cursor += 8
            if total == 0xFFFFFFFFFFFFFFFF:
                return None, cursor
            if total == 0:
                return "", cursor
            text = ""
            while True:
                chunk = struct.unpack_from("<I", payload, cursor)[0]
                cursor += 4
                if chunk == 0:
                    return text, cursor
                text += payload[cursor:cursor + chunk].decode("utf-16-le")
                cursor += chunk
        length = struct.unpack_from("<H", payload, cursor)[0]
        cursor += 2
        if length == 0xFFFF:
            return None, cursor
        return payload[cursor:cursor + length].decode("utf-16-le"), cursor + length

    length = payload[cursor]
    cursor += 1
    if length == 0:
        return None, cursor
    raw = payload[cursor:cursor + length]
    cursor += length
    if kind == 0x26:                                   # INTN
        return int.from_bytes(raw, "little", signed=length > 1), cursor
    if kind == 0x6D:                                   # FLTN
        return struct.unpack("<d" if length == 8 else "<f", raw)[0], cursor
    if kind == 0x68:                                   # BITN
        return raw[0] != 0, cursor
    if kind == 0x6F:                                   # DATETIMN
        days, ticks = struct.unpack("<iI", raw)
        return (datetime.datetime(1900, 1, 1) + datetime.timedelta(days=days)
                + datetime.timedelta(seconds=ticks / 300)), cursor
    if kind == 0x24:                                   # GUIDN
        return uuid.UUID(bytes_le=raw), cursor
    return raw, cursor


def columns_of(tokens: list[tuple[int, object]]) -> list[dict]:
    for token, body in tokens:
        if token == TOKEN_COL_METADATA:
            return body
    return []


def rows_of(tokens: list[tuple[int, object]], columns=None) -> list[list]:
    return [body for token, body in tokens if token == TOKEN_ROW]


def message_text(body: bytes) -> str:
    """The text out of an INFO or ERROR token body."""
    characters = struct.unpack_from("<H", body, 6)[0]
    return body[8:8 + characters * 2].decode("utf-16-le")
