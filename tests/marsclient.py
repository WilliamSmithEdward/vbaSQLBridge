"""A TDS client that speaks the session layer MARS runs over.

`tdsclient` talks to the bridge with nothing under TDS. A client that asked
for MARS in PRELOGIN gets every message after the login wrapped in a
sixteen-byte session header, and that layer has rules of its own: one TDS
packet per header, a sequence number that counts packets, and a window that
says how far the other side may send before it waits to be allowed more.

The window is the part worth testing. A client grants four packets and the
answer to `SELECT * FROM anything` is a great deal more than four, so a
server that ignores the window is a server that cannot return a table. This
client grants exactly what a real one grants and acknowledges as it reads,
which is what makes an answer of any size go through.
"""
import select
import struct

import tdsclient

SMID = 0x53
SYN = 0x01
ACK = 0x02
FIN = 0x04
DATA = 0x08
HEADER = struct.Struct("<BBHIII")     # smid, flags, sid, length, seq, window
HEADER_SIZE = 16

# What a real driver grants, read off the wire from .NET's SqlClient.
WINDOW = 4


def wants_mars(payload: bytes = None) -> bytes:
    """The recorded PRELOGIN with its MARS byte turned on.

    The option table puts MARS one byte wide at offset 0x2b, and the capture
    the constant came from had it off.
    """
    out = bytearray(payload or tdsclient.CLIENT_PRELOGIN_PAYLOAD)
    out[0x2B] = 1
    return bytes(out)


class MarsClient(tdsclient.TdsClient):
    """A TdsClient that wraps what it sends and unwraps what it reads."""

    def __init__(self, port: int, session: int = 1) -> None:
        super().__init__(port)
        self.session = session
        self.out_seq = 0
        self.in_seq = 0
        self.window = WINDOW
        self.granted = WINDOW
        self.frames = []

    # -- the session layer -----------------------------------------------
    def smp(self, flags: int, payload: bytes = b"") -> None:
        if flags & DATA:
            self.out_seq += 1
        self.granted = self.in_seq + self.window
        message = HEADER.pack(SMID, flags, self.session,
                              len(payload) + HEADER_SIZE, self.out_seq,
                              self.granted) + payload
        if self.encrypted:
            self.tls.write(message)
            message = self.outgoing.read()
        self.socket.sendall(message)

    def open_session(self) -> None:
        self.smp(SYN)

    def batch(self, sql: str) -> None:
        self.smp(DATA, tdsclient.frame(tdsclient.SQL_BATCH,
                                       sql.encode("utf-16-le")))

    def read_answer(self) -> bytes:
        """Every frame of one answer, acknowledging as a real client does.

        The server may only send as far as the window reaches, so the window
        is moved on once per frame read. Without that the answer stops after
        four packets and the read waits for ever.
        """
        payload = b""
        while True:
            while len(self.buffer) < HEADER_SIZE:
                self._fill()
            smid, flags, sid, length, seq, window = HEADER.unpack(
                self.buffer[:HEADER_SIZE])
            if smid != SMID:
                raise AssertionError(
                    f"expected a session header, got 0x{smid:02x}")
            # A real driver hangs up on a server that sends past what it
            # was allowed, so this one refuses it rather than reading on.
            if seq > self.granted:
                raise AssertionError(
                    f"the server sent sequence {seq} where the window "
                    f"granted reached {self.granted}: a real client ends "
                    f"the session with a transport-level error here")
            while len(self.buffer) < length:
                self._fill()
            body = self.buffer[HEADER_SIZE:length]
            self.buffer = self.buffer[length:]
            self.frames.append((flags, sid, seq, len(body)))
            self.in_seq = seq
            # The eight bytes of TDS header on each packet are framing, the
            # same as they are without a session under them.
            payload += body[8:]

            # The window is only moved on once it has been used up, rather
            # than after every frame. Acknowledging each one as it arrives
            # would let a server that ignores the window keep just ahead of
            # the acknowledgements and never be caught at it.
            if seq >= self.granted:
                # Nothing may be waiting yet: the window is used up and
                # nothing has been sent to move it on. A server that sent
                # the rest of the answer anyway has already put it on the
                # wire, and that is what a real driver hangs up over.
                self._refuse_anything_early()
                self.smp(ACK)

            if len(body) >= 2 and body[1] & 1:
                return payload
            if not body:
                return payload

    def _refuse_anything_early(self) -> None:
        """Fail if the server has sent past the window already.

        Read as far as the window reaches and then look: a server that
        waits has sent nothing more, and one that does not has the rest of
        the answer sitting there to be read.
        """
        if len(self.buffer) >= HEADER_SIZE:
            raise AssertionError(
                "the server sent past the window: another session frame was "
                "waiting before anything acknowledged the last one")
        waiting = select.select([self.socket], [], [], 0.25)[0]
        if waiting:
            self._fill()
            if len(self.buffer) >= HEADER_SIZE:
                raise AssertionError(
                    "the server sent past the window: more arrived before "
                    "anything acknowledged what it had already sent")

    # -- the whole sequence ----------------------------------------------
    def start(self, port: int) -> None:
        self.send(tdsclient.PRELOGIN, wants_mars())
        self.read_message()
        self.handshake()
        self.login(port)
        self.open_session()

    def ask(self, sql: str) -> tuple:
        """One statement, answered: its columns and its rows."""
        self.batch(sql)
        payload = self.read_answer()
        tokens = tdsclient.parse_tokens(payload)
        columns = tdsclient.columns_of(tokens)
        return columns, tdsclient.rows_of(tokens, columns)
