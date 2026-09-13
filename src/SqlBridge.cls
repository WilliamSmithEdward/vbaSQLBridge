VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlBridge"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' vbaSQLBridge: answer SQL Server's wire protocol from Excel, so a client
' connects to a workbook and sees a database.
'
' Everything is Win32 through Declare statements: sockets from ws2_32, the
' login's TLS tunnel and Windows Authentication from secur32, and the
' certificate that tunnel presents from crypt32. Nothing here needs a
' reference, an installer, or a driver on the machine that connects.
'
' The wire format is taken from pySQLbridge, which measured it against SQL
' Server 2025 rather than reading it out of the specification. Byte layouts
' that carry a "measured" note came off that capture.

' MIT License
'
' Copyright (c) 2026 William Smith
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


' ============================================================================
' Windows
' ============================================================================

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal length As LongPtr _
)


' ============================================================================
' TDS packet framing
'
' Every TDS message travels as one or more packets behind an 8-byte header:
'
'     offset 0  type       1 byte
'     offset 1  status     1 byte, bit flags
'     offset 2  length     2 bytes, whole packet including this header
'     offset 4  spid       2 bytes
'     offset 6  packet id  1 byte
'     offset 7  window     1 byte
'
' Multi-byte fields in the header are big-endian, which is worth stating
' because nothing else in the protocol is.
' ============================================================================

Private Const TDS_HEADER_SIZE As Long = 8

' The length field is 16 bits, so a packet cannot exceed this whatever size
' the connection negotiates.
'
' Written in decimal deliberately. VBA types a hex literal by its width, so
' &HFFFF is an Integer holding -1 and &H8000 is -32768, and both keep that
' value when assigned to a Long. Every hex constant in this file is either
' four digits and positive, or eight digits where the sign bit is the bit
' pattern the API wants.
Private Const TDS_MAX_PACKET_SIZE As Long = 65535

' What a connection uses until ENVCHANGE negotiates something else.
Private Const TDS_DEFAULT_PACKET_SIZE As Long = 4096

' Packet types. Marked values were seen in the reference capture; the rest
' are protocol-defined and not yet exercised here.
Private Const TDS_SQL_BATCH As Long = &H1           ' observed: client query
Private Const TDS_RPC As Long = &H3
Private Const TDS_TABULAR_RESULT As Long = &H4      ' observed: every reply
Private Const TDS_ATTENTION As Long = &H6
Private Const TDS_BULK_LOAD As Long = &H7
Private Const TDS_TRANSACTION_MANAGER As Long = &HE

' The transaction-manager requests a linked server and a client's API send,
' and the ENVCHANGE types a real server answers them with, read off a capture
' of one. The request types below 5 and the 6 between are the distributed
' ones: asking where the coordinator is, joining, and promoting.
Private Const TM_BEGIN_XACT As Long = 5
Private Const TM_COMMIT_XACT As Long = 7
Private Const TM_ROLLBACK_XACT As Long = 8
Private Const TM_SAVE_XACT As Long = 9
Private Const ENV_BEGIN_TRANSACTION As Long = 8
Private Const ENV_COMMIT_TRANSACTION As Long = 9
Private Const ENV_ROLLBACK_TRANSACTION As Long = 10
Private Const TDS_LOGIN7 As Long = &H10
Private Const TDS_SSPI As Long = &H11               ' observed: NTLM authenticate
Private Const TDS_PRELOGIN As Long = &H12           ' observed: prelogin and TLS

' Status flags, the header's second byte.
Private Const TDS_STATUS_NORMAL As Long = &H0
Private Const TDS_STATUS_EOM As Long = &H1          ' observed on every packet
Private Const TDS_STATUS_IGNORE As Long = &H2
Private Const TDS_STATUS_RESET As Long = &H8
Private Const TDS_STATUS_RESET_SKIP_TRAN As Long = &H10

' TDS versions, as LOGIN7 and LOGINACK spell them.
Private Const TDS_VERSION_70 As Long = &H70000000
Private Const TDS_VERSION_71 As Long = &H71000001
Private Const TDS_VERSION_72 As Long = &H72090002
Private Const TDS_VERSION_73A As Long = &H730A0003
Private Const TDS_VERSION_73B As Long = &H730B0003
Private Const TDS_VERSION_74 As Long = &H74000004

' SQL Server numbers user sessions from 51 and keeps everything below that
' for its own background work. The capture showed 66.
Private Const TDS_FIRST_SESSION_ID As Long = 51


' ============================================================================
' PRELOGIN
'
' The payload is a table of five-byte entries followed by the blob they point
' into. Order and presence both matter: a server that answers with the
' options it feels like sending, in its own order, is not answering the way
' SQL Server does.
'
'     token   1 byte
'     offset  2 bytes, from the payload's first byte
'     length  2 bytes
'     ...
'     0xFF    terminator, with no offset or length of its own
' ============================================================================

Private Const PRELOGIN_VERSION As Long = &H0
Private Const PRELOGIN_ENCRYPTION As Long = &H1
Private Const PRELOGIN_INSTOPT As Long = &H2
Private Const PRELOGIN_THREADID As Long = &H3
Private Const PRELOGIN_MARS As Long = &H4

' The session multiplex layer MARS runs over. Once a client and this server
' have agreed to MARS in PRELOGIN, every message after the login response is
' wrapped in a sixteen-byte header of its own, and one connection carries
' several sessions that each look like a connection to whatever is asking.
'
' Read off the wire from the linked-server provider rather than from the
' specification: SYN, then DATA carrying an ordinary TDS batch.
Private Const SMP_SMID As Long = &H53
Private Const SMP_SYN As Long = &H1
Private Const SMP_ACK As Long = &H2
Private Const SMP_FIN As Long = &H4
Private Const SMP_DATA As Long = &H8
Private Const SMP_HEADER_SIZE As Long = 16
Private Const SMP_WINDOW As Long = 4
Private Const PRELOGIN_TRACEID As Long = &H5
Private Const PRELOGIN_FEDAUTHREQUIRED As Long = &H6
Private Const PRELOGIN_NONCEOPT As Long = &H7
Private Const PRELOGIN_TERMINATOR As Long = &HFF

' The ENCRYPTION option's single byte. OFF is what both sides sent in the
' reference capture, and it does not mean no encryption: the login packet
' still goes through a TLS handshake, after which the connection reverts to
' cleartext.
Private Const ENCRYPT_OFF As Long = &H0
Private Const ENCRYPT_ON As Long = &H1
Private Const ENCRYPT_NOT_SUPPORTED As Long = &H2
Private Const ENCRYPT_REQUIRED As Long = &H3

' What this server reports as itself, taken from the capture so a client
' sees the version it would see from SQL Server 2025.
Private Const SERVER_VERSION_MAJOR As Long = 17
Private Const SERVER_VERSION_MINOR As Long = 0
Private Const SERVER_VERSION_BUILD As Long = 1000


' ============================================================================
' Errors
' ============================================================================

' Raised for anything the peer sent that cannot be read as TDS.
Private Const ERR_PROTOCOL As Long = vbObjectError + 4001



' ============================================================================
' Sockets
'
' ws2_32 through Declare statements, non-blocking throughout. Blocking accept
' and recv would stop Excel dead, and a server that freezes its host until a
' client turns up is not one anybody can leave running in a workbook.
'
' Every imported name is aliased. bind, listen and send are ordinary enough
' words that VBA code around them reads badly.
' ============================================================================

Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" ( _
    ByVal wVersionRequested As Integer, _
    lpWSAData As Any _
) As Long

Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long

Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long

Private Declare PtrSafe Function sock_socket Lib "ws2_32.dll" Alias "socket" ( _
    ByVal af As Long, _
    ByVal socketType As Long, _
    ByVal protocol As Long _
) As LongPtr

Private Declare PtrSafe Function sock_bind Lib "ws2_32.dll" Alias "bind" ( _
    ByVal s As LongPtr, _
    addr As Any, _
    ByVal namelen As Long _
) As Long

Private Declare PtrSafe Function sock_listen Lib "ws2_32.dll" Alias "listen" ( _
    ByVal s As LongPtr, _
    ByVal backlog As Long _
) As Long

Private Declare PtrSafe Function sock_accept Lib "ws2_32.dll" Alias "accept" ( _
    ByVal s As LongPtr, _
    addr As Any, _
    addrlen As Any _
) As LongPtr

Private Declare PtrSafe Function sock_recv Lib "ws2_32.dll" Alias "recv" ( _
    ByVal s As LongPtr, _
    buf As Any, _
    ByVal length As Long, _
    ByVal flags As Long _
) As Long

Private Declare PtrSafe Function sock_send Lib "ws2_32.dll" Alias "send" ( _
    ByVal s As LongPtr, _
    buf As Any, _
    ByVal length As Long, _
    ByVal flags As Long _
) As Long

Private Declare PtrSafe Function sock_closesocket Lib "ws2_32.dll" Alias "closesocket" ( _
    ByVal s As LongPtr _
) As Long

Private Declare PtrSafe Function sock_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" ( _
    ByVal s As LongPtr, _
    ByVal cmd As Long, _
    argp As Long _
) As Long

Private Declare PtrSafe Function sock_setsockopt Lib "ws2_32.dll" Alias "setsockopt" ( _
    ByVal s As LongPtr, _
    ByVal level As Long, _
    ByVal optname As Long, _
    optval As Any, _
    ByVal optlen As Long _
) As Long

Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong

Private Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal pointer As LongPtr _
) As Long

Private Declare PtrSafe Sub SleepMs Lib "kernel32" Alias "Sleep" ( _
    ByVal milliseconds As Long _
)

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_REUSEADDR As Long = 4
Private Const SOCKET_ERROR As Long = -1
Private Const INVALID_SOCKET As LongPtr = -1

' FIONBIO, the ioctl that turns blocking off. Eight hex digits, so VBA types
' it as a Long and keeps the sign bit the API expects.
Private Const FIONBIO As Long = &H8004667E

Private Const WSAEWOULDBLOCK As Long = 10035

' The key an IN lookup remembers a NULL candidate under. Neither of the
' prefixes ValueKey writes, so nothing can collide with it.
Private Const NULL_CANDIDATE As String = "!null"

' Winsock's own numbers start here. Anything below it did not come from
' winsock, whatever the thread's last error happens to say.
Private Const WSABASEERR As Long = 10000
Private Const WSAEADDRINUSE As Long = 10048

' What a receive reports beyond a byte count.
Private Const RECEIVE_NOTHING As Long = 0
Private Const RECEIVE_CLOSED As Long = -1

' Raised when Windows refuses a socket, a credential or a security context.
Private Const ERR_TRANSPORT As Long = vbObjectError + 4002


' ============================================================================
' The certificate the login tunnel presents
'
' SQL Server generates a self-signed certificate at startup when none is
' configured and its clients accept it, because with encryption negotiated
' off the tunnel covers only the login packet and drivers do not validate a
' chain for that. This does the same thing for the same reason.
'
' The key lives in a named CryptoAPI container for as long as the server
' runs and is deleted on shutdown. Schannel finds it through the key
' provider information stored in the certificate, so the container name is
' load-bearing rather than decorative.
' ============================================================================

Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32.dll" ( _
    phProv As LongPtr, _
    ByVal pszContainer As LongPtr, _
    ByVal pszProvider As LongPtr, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long _
) As Long

Private Declare PtrSafe Function CryptGenKey Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, _
    ByVal algId As Long, _
    ByVal dwFlags As Long, _
    phKey As LongPtr _
) As Long

Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr _
) As Long

Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, _
    ByVal dwFlags As Long _
) As Long

Private Declare PtrSafe Function CertStrToNameW Lib "crypt32.dll" ( _
    ByVal dwCertEncodingType As Long, _
    ByVal pszX500 As LongPtr, _
    ByVal dwStrType As Long, _
    ByVal pvReserved As LongPtr, _
    pbEncoded As Any, _
    pcbEncoded As Long, _
    ByVal ppszError As LongPtr _
) As Long

Private Declare PtrSafe Function CertCreateSelfSignCertificate Lib "crypt32.dll" ( _
    ByVal hCryptProv As LongPtr, _
    pSubjectIssuerBlob As Any, _
    ByVal dwFlags As Long, _
    pKeyProvInfo As Any, _
    pSignatureAlgorithm As Any, _
    pStartTime As Any, _
    pEndTime As Any, _
    ByVal pExtensions As LongPtr _
) As LongPtr

Private Declare PtrSafe Function CertFreeCertificateContext Lib "crypt32.dll" ( _
    ByVal pCertContext As LongPtr _
) As Long

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" ( _
    lpSystemTime As SYSTEMTIME _
)

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type CRYPTOAPI_BLOB
    cbData As Long
    pbData As LongPtr
End Type

Private Type CRYPT_KEY_PROV_INFO
    pwszContainerName As LongPtr
    pwszProvName As LongPtr
    dwProvType As Long
    dwFlags As Long
    cProvParam As Long
    rgProvParam As LongPtr
    dwKeySpec As Long
End Type

Private Type CRYPT_ALGORITHM_IDENTIFIER
    pszObjId As LongPtr
    Parameters As CRYPTOAPI_BLOB
End Type

' The provider that carries SHA-2. PROV_RSA_FULL cannot sign a certificate
' with SHA-256, and a certificate signed with SHA-1 is one a hardened client
' will refuse.
Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const CRYPT_DELETEKEYSET As Long = 16
Private Const AT_KEYEXCHANGE As Long = 1
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const X509_ASN_ENCODING As Long = 1
Private Const CERT_X500_NAME_STR As Long = 3
Private Const RSA_KEY_BITS As Long = 2048
Private Const CERTIFICATE_DAYS As Long = 365
Private Const SHA256_RSA_OID As String = "1.2.840.113549.1.1.11"


' ============================================================================
' SQL Server logins, which the bridge checks itself
'
' Windows authentication leaves the decision to Windows. A login with a
' name and a password leaves it here, so the rules are narrow: nothing is
' admitted that the host did not name, a password is kept only as a salted
' PBKDF2 hash and never written down, and a name that is not there costs as
' much to refuse as a password that is wrong. CNG does the hashing, which
' VBA would take seconds over.
' ============================================================================

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    phAlgorithm As LongPtr, _
    ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, _
    ByVal dwFlags As Long _
) As Long

Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal dwFlags As Long _
) As Long

Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt.dll" ( _
    ByVal hPrf As LongPtr, _
    ByVal pbPassword As LongPtr, _
    ByVal cbPassword As Long, _
    ByVal pbSalt As LongPtr, _
    ByVal cbSalt As Long, _
    ByVal cIterations As LongLong, _
    ByVal pbDerivedKey As LongPtr, _
    ByVal cbDerivedKey As Long, _
    ByVal dwFlags As Long _
) As Long

Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal pbBuffer As LongPtr, _
    ByVal cbBuffer As Long, _
    ByVal dwFlags As Long _
) As Long

Private Const BCRYPT_ALG_HANDLE_HMAC_FLAG As Long = 8
Private Const BCRYPT_USE_SYSTEM_PREFERRED_RNG As Long = 2

' How a password is kept, which is the form pySQLbridge keeps one in, so
' either reads the other's. The count travels inside each hash, so raising
' it later leaves the hashes already written working.
Private Const PASSWORD_ALGORITHM As String = "pbkdf2_sha256"
Private Const PASSWORD_ITERATIONS As Long = 210000
Private Const PASSWORD_SALT_BYTES As Long = 16
Private Const PASSWORD_DIGEST_BYTES As Long = 32

' What LOGIN7's password field is masked with, a constant [MS-TDS]
' publishes.
Private Const PASSWORD_MASK As Long = &HA5

' What a real server answers a login it will not admit, measured against
' one: 18456 at severity 14 and state 1, then a DONE, then it hangs up. The
' same for a name it does not have and a password that is wrong, so the
' answer never says which.
Private Const LOGIN_FAILED As Long = 18456
Private Const LOGIN_FAILED_SEVERITY As Long = 14

Private Const BASE64_ALPHABET As String = _
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"


' ============================================================================
' TLS and Windows Authentication, both through SSPI
'
' The two use the same buffer machinery and differ only in which package is
' asked for: Schannel for the tunnel that carries the login, Negotiate for
' the login itself. Neither ever sees a password. Windows validates the
' credential against the local account database or the domain, which is the
' point: a login server that implements its own credential check is a login
' server that gets it wrong.
' ============================================================================

Private Declare PtrSafe Function AcquireCredentialsHandleW Lib "secur32.dll" ( _
    ByVal pszPrincipal As LongPtr, _
    ByVal pszPackage As LongPtr, _
    ByVal fCredentialUse As Long, _
    ByVal pvLogonId As LongPtr, _
    pAuthData As Any, _
    ByVal pGetKeyFn As LongPtr, _
    ByVal pvGetKeyArgument As LongPtr, _
    phCredential As SecHandle, _
    ptsExpiry As SECURITY_INTEGER _
) As Long

Private Declare PtrSafe Function AcceptSecurityContext Lib "secur32.dll" ( _
    phCredential As SecHandle, _
    ByVal phContext As LongPtr, _
    pInput As SecBufferDesc, _
    ByVal fContextReq As Long, _
    ByVal targetDataRep As Long, _
    phNewContext As SecHandle, _
    pOutput As SecBufferDesc, _
    pfContextAttr As Long, _
    ptsTimeStamp As SECURITY_INTEGER _
) As Long

Private Declare PtrSafe Function DeleteSecurityContext Lib "secur32.dll" ( _
    phContext As SecHandle _
) As Long

Private Declare PtrSafe Function FreeCredentialsHandle Lib "secur32.dll" ( _
    phCredential As SecHandle _
) As Long

Private Declare PtrSafe Function FreeContextBuffer Lib "secur32.dll" ( _
    ByVal pvContextBuffer As LongPtr _
) As Long

Private Declare PtrSafe Function QueryContextAttributesW Lib "secur32.dll" ( _
    phContext As SecHandle, _
    ByVal ulAttribute As Long, _
    pBuffer As Any _
) As Long

Private Declare PtrSafe Function EncryptMessage Lib "secur32.dll" ( _
    phContext As SecHandle, _
    ByVal fQOP As Long, _
    pMessage As SecBufferDesc, _
    ByVal messageSeqNo As Long _
) As Long

Private Declare PtrSafe Function DecryptMessage Lib "secur32.dll" ( _
    phContext As SecHandle, _
    pMessage As SecBufferDesc, _
    ByVal messageSeqNo As Long, _
    pfQOP As Long _
) As Long

Private Type SecHandle
    dwLower As LongPtr
    dwUpper As LongPtr
End Type

Private Type SECURITY_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type SecBuffer
    cbBuffer As Long
    BufferType As Long
    pvBuffer As LongPtr
End Type

Private Type SecBufferDesc
    ulVersion As Long
    cBuffers As Long
    pBuffers As LongPtr
End Type

Private Type SCHANNEL_CRED
    dwVersion As Long
    cCreds As Long
    paCred As LongPtr
    hRootStore As LongPtr
    cMappers As Long
    aphMappers As LongPtr
    cSupportedAlgs As Long
    palgSupportedAlgs As LongPtr
    grbitEnabledProtocols As Long
    dwMinimumCipherStrength As Long
    dwMaximumCipherStrength As Long
    dwSessionLifespan As Long
    dwFlags As Long
    dwCredFormat As Long
End Type

Private Type SecPkgContext_StreamSizes
    cbHeader As Long
    cbTrailer As Long
    cbMaximumMessage As Long
    cBuffers As Long
    cbBlockSize As Long
End Type

Private Const UNISP_NAME As String = "Microsoft Unified Security Protocol Provider"
Private Const NEGOTIATE_NAME As String = "Negotiate"

Private Const SECPKG_CRED_INBOUND As Long = 1
Private Const SCHANNEL_CRED_VERSION As Long = 4

' TLS 1.2, server side, and only that.
'
' The reference server negotiated exactly this: its ServerHello carried
' version 0x0303 with no supported_versions extension, and the client flight
' held the ChangeCipherSpec that 1.3 does away with. Pinning also avoids a
' problem TDS predates. TLS 1.3 sends NewSessionTicket after the handshake
' reports itself complete, by which point the TDS framing has already
' stopped, and there is no correct way to carry those: framed is wrong
' because the tunnel is up, bare is wrong because the client is not reading
' bare records yet.
Private Const SP_PROT_TLS1_2_SERVER As Long = &H400

Private Const SECBUFFER_EMPTY As Long = 0
Private Const SECBUFFER_DATA As Long = 1
Private Const SECBUFFER_TOKEN As Long = 2
Private Const SECBUFFER_EXTRA As Long = 5
Private Const SECBUFFER_STREAM_TRAILER As Long = 6
Private Const SECBUFFER_STREAM_HEADER As Long = 7
Private Const SECBUFFER_VERSION As Long = 0

' What Schannel's inbound side accepts, and no more.
'
' Measured, because the published sample for a TLS server names six flags.
' Fed a real ClientHello, this credential answers SEC_E_UNSUPPORTED_FUNCTION
' to REPLAY_DETECT, SEQUENCE_DETECT, CONFIDENTIALITY and EXTENDED_ERROR in
' any combination, and SEC_E_INSUFFICIENT_MEMORY without ALLOCATE_MEMORY,
' which is a null output buffer rather than a flag problem. Those four
' belong to the initiating side.
Private Const ASC_REQ_ALLOCATE_MEMORY As Long = &H100
Private Const ASC_REQ_STREAM As Long = &H10000
Private Const ASC_REQ_CONNECTION As Long = &H800

Private Const SECPKG_ATTR_STREAM_SIZES As Long = 4
Private Const SECPKG_ATTR_NAMES As Long = 1

Private Const SEC_E_OK As Long = 0
Private Const SEC_I_CONTINUE_NEEDED As Long = &H90312
Private Const SEC_I_COMPLETE_NEEDED As Long = &H90313
Private Const SEC_I_COMPLETE_AND_CONTINUE As Long = &H90314
Private Const SEC_E_INCOMPLETE_MESSAGE As Long = &H80090318


' ============================================================================
' What one object is
'
' One class serves two roles, because VBA offers no other way to ship a
' library as a file somebody imports. A server owns the listening socket,
' the certificate and the credentials; a connection owns one socket and the
' state machine driving it, and reaches back to its server for the rest.
' ============================================================================

Private Const ROLE_NONE As Long = 0
Private Const ROLE_SERVER As Long = 1
Private Const ROLE_CONNECTION As Long = 2

' The login sequence, following the measured handshake.
'
'   EXPECT_PRELOGIN  a TDS message arrives, a PRELOGIN response goes back
'   TLS_HANDSHAKE    TLS records arrive framed in TDS, replies go back framed
'   EXPECT_LOGIN     TLS records arrive bare, LOGIN7 comes out of them
'   EXPECT_SSPI      the client answers the challenge, in the clear
'   READY            logged in; batches are answered from here on
Private Const STATE_EXPECT_PRELOGIN As Long = 0
Private Const STATE_TLS_HANDSHAKE As Long = 1
Private Const STATE_EXPECT_LOGIN As Long = 2
Private Const STATE_EXPECT_SSPI As Long = 3
Private Const STATE_READY As Long = 4
Private Const STATE_CLOSED As Long = 5
Private Const STATE_FAILED As Long = 6
' A login refused: the refusal goes out, then the connection closes.
Private Const STATE_REFUSED As Long = 7

Private mRole As Long

' Server
Private mPort As Long
Private mAddress As String
Private mServerName As String
Private mDatabase As String
Private mEncryptionPolicy As Long
Private mListener As LongPtr
Private mConnections As Collection
Private mWinsockStarted As Boolean
Private mCertificate As LongPtr
Private mProvider As LongPtr
Private mKey As LongPtr
Private mContainer As String
Private mTlsCredentials As SecHandle
Private mHasTlsCredentials As Boolean
Private mLoginCredentials As SecHandle
Private mHasLoginCredentials As Boolean
Private mNextSpid As Long
Private mLog As Collection
Private mLogLimit As Long

' Connection
Private mServer As SqlBridge
Private mSocket As LongPtr
Private mState As Long
Private mIn() As Byte
Private mInUsed As Long
Private mPlain() As Byte
Private mPlainUsed As Long
Private mOut() As Byte
Private mOutUsed As Long
Private mTlsContext As SecHandle
Private mHasTlsContext As Boolean
Private mTlsEstablished As Boolean
Private mStreamSizes As SecPkgContext_StreamSizes
Private mLoginContext As SecHandle
Private mHasLoginContext As Boolean
Private mSessionEncrypted As Boolean
Private mSpid As Long
Private mTdsVersion As Long
Private mClientPrelogin As Collection

' MARS, and the session whatever is being answered belongs to.
Private mMars As Boolean
Private mSmpSid As Long
Private mSmpOutSeq As Long
Private mSmpInSeq As Long

' SMP is flow controlled, and one connection carries several sessions at
' once, which is the whole point of MARS: the Object Explorer opens a
' session per thing it is drawing and keeps them going together. Each has a
' sequence number of its own, a window of its own, and frames of its own
' waiting to go, so all three are held per session rather than once. Shared,
' a question put on one session while another was still answering was
' answered on the wrong one.
Private mSmpOutSeqs As Collection
Private mSmpWindows As Collection
Private mSmpQueues As Collection
Private mSmpOrder As Collection

' And a buffer each on the way in, with the whole messages they made waiting
' to be answered, each remembering the session it arrived on.
Private mSmpInBufs As Collection
Private mSmpReady As Collection
Private mSmpIn() As Byte
Private mSmpInUsed As Long

' The transaction the transaction manager last opened on this connection,
' which is what the descriptor in its ENVCHANGE counts.
Private mTransactionDescriptor As Long

' The statements prepared on this connection, by handle, and the handle the
' call being answered was prepared under or asks to run.
Private mPrepared As Collection
Private mLastHandle As Long
Private mCallHandle As Long
Private mUserName As String
Private mAppName As String
Private mHostName As String


' ============================================================================
' Tokens in the server-to-client stream
'
' A TABULAR_RESULT payload is a sequence of tokens, each opening with a
' one-byte type. Lengths inside tokens are little-endian, unlike the packet
' header's big-endian length, and two string conventions appear that are not
' interchangeable:
'
'     B_VARCHAR   one length byte, counting characters, then UTF-16LE
'     US_VARCHAR  two length bytes, counting characters, then UTF-16LE
'
' The collation ENVCHANGE breaks the pattern by carrying five raw bytes
' behind a length that counts bytes rather than characters, because a
' collation is not text.
' ============================================================================

Private Const TOKEN_COL_METADATA As Long = &H81
Private Const TOKEN_ERROR As Long = &HAA
Private Const TOKEN_INFO As Long = &HAB
Private Const TOKEN_LOGIN_ACK As Long = &HAD
Private Const TOKEN_FEATURE_EXT_ACK As Long = &HAE
Private Const TOKEN_ROW As Long = &HD1
Private Const TOKEN_ENV_CHANGE As Long = &HE3
Private Const TOKEN_SSPI As Long = &HED
Private Const TOKEN_DONE As Long = &HFD

Private Const ENV_DATABASE As Long = 1
Private Const ENV_LANGUAGE As Long = 2
Private Const ENV_PACKET_SIZE As Long = 4
Private Const ENV_SQL_COLLATION As Long = 7

Private Const DONE_FINAL As Long = &H0
Private Const DONE_MORE As Long = &H1
Private Const DONE_ERROR As Long = &H2
Private Const DONE_COUNT As Long = &H10
' What acknowledges a cancellation. A client that sent one waits for it and
' will not reuse the connection until it arrives.
Private Const DONE_ATTENTION As Long = &H20

' The reference server's own name for itself. Note the two trailing nulls:
' the B_VARCHAR counts 22 characters, not the 20 the text has, and trimming
' them produces a token two bytes short of the real one.
Private Const SERVER_PROGRAM_NAME As String = "Microsoft SQL Server" & _
    vbNullChar & vbNullChar

' The collation the reference server announced. Five opaque bytes; nothing
' here interprets them, and sending what the real server sends is the point.
Private Const DEFAULT_COLLATION As String = "0904d00034"

Private Const DEFAULT_LANGUAGE As String = "us_english"
' What the one database this serves is called.
'
' Not master. A client files a database away by its name: SSMS puts master
' under System Databases, where somebody looking for their worksheet will
' not think to open it, and the Databases node reads as empty. A name of its
' own puts it where a user database goes.
Private Const DEFAULT_DATABASE As String = "vbaSQLBridge"

' How many separate stretches of rows go to Excel in one delete. Measured:
' the cost of a delete grows faster than the number of areas in it.
Private Const DELETE_AREAS As Long = 100

' Where a SQL Server keeps its own files. Nothing here has any, and a client
' that asks is told what it would be told by a server that did.
Private Const DEFAULT_DATA_PATH As String = _
    "C:\Program Files\Microsoft SQL Server\MSSQL\DATA\"
Private Const INSTALL_PATH As String = _
    "C:\Program Files\Microsoft SQL Server\MSSQL"

' The line number in INFO and ERROR grew from two bytes to four in TDS 7.2,
' the same version that widened the row count in DONE.
Private Const WIDENED_IN As Long = TDS_VERSION_72

' Where SQL Server's user-defined error numbers begin. Anything this server
' refuses is its own complaint rather than one of the real server's.
Private Const ERR_NUMBER_BASE As Long = 50000


' ============================================================================
' LOGIN7
'
' Everything in it is little-endian, which is worth stating because the
' packet header wrapping it is big-endian. The offset and length table holds
' nine string pairs, a six-byte client id, then pairs for SSPI, AtchDBFile
' and ChangePassword, and finally cbSSPILong: 58 bytes from offset 36, so
' the variable data starts at 94.
' ============================================================================

Private Const LOGIN7_FIXED_SIZE As Long = 36
Private Const LOGIN7_VARIABLE_START As Long = 94

' cbSSPI is 16 bits. When a blob will not fit, the client writes this
' sentinel and puts the real length in the 32-bit cbSSPILong instead.
Private Const SSPI_LENGTH_ESCAPE As Long = 65535

' The reference server sends its handshake packets with packet id 0, where
' PRELOGIN itself used 1. Clients are not known to inspect the field, and
' this project's whole approach is to send what the real server sends.
Private Const HANDSHAKE_PACKET_ID As Long = 0

' The words that begin a statement when they follow a finished one.
Private Const STATEMENT_STARTS As String = _
    "|SELECT|INSERT|UPDATE|DELETE|CREATE|DROP|ALTER|DECLARE|SET|EXEC|" & _
    "EXECUTE|IF|BEGIN|WITH|USE|PRINT|RETURN|TRUNCATE|GRANT|REVOKE|" & _
    "WAITFOR|MERGE|WHILE|COMMIT|ROLLBACK|SAVE|"

' How many times a WHILE may go round before it is called a runaway. A real
' server has no such limit; a workbook does, because the loop runs on the
' timer that keeps Excel answering and nobody can reach a stuck one.
Private Const WHILE_ROUNDS As Long = 100000

' What one SELECT is joined to another by, so the second belongs to the same
' statement rather than beginning one.
Private Const JOINS_TWO_SELECTS As String = "|UNION|EXCEPT|INTERSECT|"

' A statement no SELECT can be part of, so one after it begins another.
' INSERT INTO t SELECT is a single statement, and so is a branch of an IF.
' Nor can a statement that begins, ends or marks a transaction: BEGIN TRAN
' with a SELECT on the line after it was one statement, and the SELECT was
' lost. A BEGIN that opens a block never gets this far.
Private Const HOLDS_NO_SELECT As String = _
    "|DECLARE|SET|EXEC|EXECUTE|DROP|CREATE|BEGIN|COMMIT|ROLLBACK|SAVE|"


' Connection state that belongs to the login rather than the transport.
Private mHandshake() As Byte
Private mHandshakeUsed As Long
Private mClientTdsVersion As Long
Private mLastQuery As String
Private mStopRequested As Boolean




' ============================================================================
' Result sets
'
' Two tokens carry a SELECT's answer. COLMETADATA declares the shape once,
' then one ROW follows per row, and the values inside a row are encoded
' according to that declaration rather than describing themselves. The
' coupling is why the column types live with the encoders.
'
' Nullability is not a column property in the encoding. The nullable forms
' carry a length prefix and spend it to say NULL, which is why INTN and FLTN
' appear here rather than the fixed-width INT4 and FLT8.
' ============================================================================

Private Const TYPE_GUIDN As Long = &H24
Private Const TYPE_INTN As Long = &H26
Private Const TYPE_INT2 As Long = &H34
Private Const TYPE_BITN As Long = &H68
Private Const TYPE_DECIMALN As Long = &H6A
Private Const TYPE_FLTN As Long = &H6D
Private Const TYPE_DATETIMN As Long = &H6F
Private Const TYPE_BIGVARBINARY As Long = &HA5
Private Const TYPE_BIGBINARY As Long = &HAD
Private Const TYPE_NVARCHAR As Long = &HE7

' A declared size of 0xFFFF is not a size, it is the MAX form: the value
' arrives as an eight-byte total length followed by length-prefixed chunks
' and a zero-length terminator.
Private Const NVARCHAR_MAX_MARKER As Long = 65535

' The size beyond which a sized nvarchar cannot be declared, so MAX is the
' only way to carry the value.
Private Const NVARCHAR_SIZED_LIMIT As Long = 4000

' Bit 0 marks a column nullable and bit 5 marks it computed, which is what
' the reference server sent for the CAST expressions it was asked for.
Private Const SELECT_COLUMN_FLAGS As Long = &H21

' What SQL Server sends on the columns of its own catalog rowsets: bit 3 of
' the updateable field rather than the computed bit, since those come from a
' procedure rather than an expression.
Private Const CATALOG_COLUMN_FLAGS As Long = &H9

' What a real server flags a column read straight from a table with:
' nullable, and updatable as far as it knows (bits 2 and 3 holding 2), with
' no computed bit. Measured on a cursor a linked server opened over a table.
Private Const TABLE_COLUMN_FLAGS As Long = &H9

' ODBC type codes. A driver maps a column into its own type system from
' these, so an int reported as SQL_WVARCHAR arrives in Excel as text.
Private Const SQL_INTEGER As Long = 4
Private Const SQL_FLOAT As Long = 6
Private Const SQL_BIT As Long = -7
Private Const SQL_BIGINT As Long = -5
Private Const SQL_WVARCHAR As Long = -9
Private Const SQL_WLONGVARCHAR As Long = -10
Private Const SQL_GUID As Long = -11
Private Const SQL_TYPE_TIMESTAMP As Long = 93

' OLE DB's own codes, which the rowset family reports instead. Not a
' different spelling of the same numbers: nvarchar is -9 to ODBC and 130 to
' OLE DB, and -9 is not a DBTYPE at all.
Private Const DBTYPE_I2 As Long = 2
Private Const DBTYPE_I4 As Long = 3
Private Const DBTYPE_R8 As Long = 5
Private Const DBTYPE_BOOL As Long = 11
Private Const DBTYPE_I8 As Long = 20
Private Const DBTYPE_GUID As Long = 72
Private Const DBTYPE_WSTR As Long = 130
Private Const DBTYPE_DBTIMESTAMP As Long = 135

' What OLE DB says about a column in COLUMN_FLAGS.
Private Const DBCOLUMNFLAGS_ISFIXEDLENGTH As Long = &H10
Private Const DBCOLUMNFLAGS_ISNULLABLE As Long = &H20
Private Const DBCOLUMNFLAGS_MAYBENULL As Long = &H40
Private Const DBCOLUMNFLAGS_WRITE As Long = &H4

' The reference server reported this in DONE's current-command field after a
' SELECT, where the login response used zero.
Private Const SELECT_COMMAND As Long = &HC1

' What the reference server put in DONE's current-command field after each
' kind of write, read off the wire between sqlcmd and a real server on this
' machine. A client shown a count under the SELECT command prints nothing at
' all: it has a SELECT that answered with no rows, not a write.
Private Const INSERT_COMMAND As Long = &HC3
Private Const DELETE_COMMAND As Long = &HC4
Private Const UPDATE_COMMAND As Long = &HC5

' What a real server closes a MERGE with, read off the wire: its own
' command, and every row it inserted, updated or deleted in the count.
Private Const MERGE_COMMAND As Long = &H117

' How a real server closes the answer to a procedure called by RPC, read off
' the wire: DONEINPROC for its rows, RETURNSTATUS with the procedure's return
' value, then DONEPROC under the RPC command. A plain DONE is how a batch
' ends, not a call.
Private Const TOKEN_DONEINPROC As Long = &HFF
Private Const TOKEN_DONEPROC As Long = &HFE
Private Const TOKEN_RETURNSTATUS As Long = &H79
Private Const RPC_COMMAND As Long = &HE0

' An output parameter's value, which a call hands back between its return
' status and its DONEPROC.
Private Const TOKEN_RETURNVALUE As Long = &HAC

' What a cursor's columns are read from: the tables by name, then which
' table each column came from and whether it is hidden.
Private Const TOKEN_TABNAME As Long = &HA4
Private Const TOKEN_COLINFO As Long = &HA5

' A column is one of these: name, type, size, and the two the encoder needs
' to spell it out. Held as a Variant array rather than a type, because a
' class module cannot expose one and a Collection cannot hold one.
Private Const COLUMN_NAME As Long = 0
Private Const COLUMN_TYPE As Long = 1
Private Const COLUMN_SIZE As Long = 2
Private Const COLUMN_FLAGS As Long = 3
' Which source the column came from, so dtb.name and avar.name are two
' columns rather than one asked for twice.
Private Const COLUMN_QUALIFIER As Long = 4

' Only a decimal has one, and only a decimal is asked for it.
Private Const COLUMN_SCALE As Long = 5

' What a FROM clause holds: the first source, and how each one after it is
' put beside what came before.
Private Const FROM_BASE As Long = 0
Private Const FROM_VALUES As Long = 1
Private Const FROM_INNER As Long = 2
Private Const FROM_LEFT As Long = 3
Private Const FROM_CROSS As Long = 4
Private Const FROM_APPLY As Long = 5
Private Const FROM_OUTER_APPLY As Long = 6
Private Const FROM_RIGHT As Long = 7
Private Const FROM_FULL As Long = 8

' Errors a query raises, and the number the client is told. SQL Server's
' user-defined range begins at 50000; 208 is "invalid object name", which
' every client already knows how to present.
Private Const ERR_QUERY As Long = vbObjectError + 4003
Private Const SQL_ERR_INVALID_OBJECT As Long = 208

Private mTables As Collection
Private mQueryErrorNumber As Long

' Whether a statement that fails with an error a real server carries a batch
' on past is kept among the answers, which is so everywhere but inside a
' TRY; and the error a CATCH was sent by, for ERROR_NUMBER() to read.
Private mCatching As Boolean
Private mCaughtNumber As Long
Private mCaughtMessage As String
Private mParameters As Collection

' What one connection's batch left behind: the variables it declared
' and the temp tables it made. Owned by the connection and handed to
' the server for the length of a call, because the catalog is shared
' and a session is not.
Private mVariables As Collection
Private mTempTables As Collection

' What the connection said about itself at login: who it is, what it is
' called, where it came from. A client asks for all three back.
Private mSession As Collection


' ============================================================================
' SQL
'
' A parsed statement is a tree of Variant arrays: kind first, then whatever
' that kind needs. VBA has no discriminated union and a class module cannot
' expose a type, so the array is the node.
' ============================================================================

Private Const T_NAME As Long = 1
Private Const T_NUMBER As Long = 2
Private Const T_STRING As Long = 3
Private Const T_OPERATOR As Long = 4

Private Const N_LITERAL As Long = 1
Private Const N_COLUMN As Long = 2
Private Const N_STAR As Long = 3
Private Const N_BINARY As Long = 4
Private Const N_UNARY As Long = 5
Private Const N_CALL As Long = 6
Private Const N_SUBQUERY As Long = 7
Private Const N_EXISTS As Long = 8
Private Const N_WINDOW As Long = 9

' A column reference with its position already worked out, which is what a
' name becomes before a loop runs over rows. Never parsed and never stored:
' Prepared makes one, the loop uses it, and it goes with the loop.
Private Const N_AT As Long = 10

' The copy of a table an UPDATE or DELETE with a FROM joins, and the column
' on it that carries each row's place in the table.
Private Const PLACE_TABLE As String = "$target"
Private Const PLACE_COLUMN As String = "$place"

' The copy of a MERGE's source, and the column that says a joined row has
' a source row in it.
Private Const SOURCE_TABLE As String = "$source"
Private Const SOURCED_COLUMN As String = "$sourced"

' The two ends a window frame can have that are not a number of rows.
Private Const WINDOW_UNBOUNDED_PRECEDING As Long = -2147483647
Private Const WINDOW_UNBOUNDED_FOLLOWING As Long = 2147483647

' The parser's position, held here rather than threaded through every rule.
' Parsing one statement is not reentrant, and one connection parses one
' statement at a time.
Private mTokens As Collection
Private mCursor As Long

' The rows a subquery is running inside, innermost last. A condition in a
' subquery may name a column of the read that is running it, and SSMS writes
' several that do; without this the name is simply not there.
Private mOuterValues As Collection
Private mOuterColumns As Collection

' One per outer frame: what each name that frame was asked for resolved to,
' kept for as long as the frame is, because the row cannot change while it
' is the outer row.
Private mOuterCaches As Collection

' What the reads inside this statement answered, for the ones whose answer
' cannot change from row to row, and how many times a read has looked at the
' row running it.
Private mSubqueryAnswers As Collection
Private mSubqueryLookups As Collection
Private mExistsLookups As Collection
Private mExistsLookupsBuilt As Long
Private mJoinLookupsBuilt As Long
Private mNoJoinShortcut As Boolean
Private mReadOnly As Boolean

' The SQL Server logins a client may use, each a name and its password's
' hash, and a hash of a password nobody knows, which a name that is not
' there is checked against.
Private mLogins As Collection
Private mDecoyHash As String

' The views a workbook defines, kept among the tables with the statement
' each stands for. Their columns are whatever their statements answer with,
' worked out once a statement, because a client listing the catalog asks
' about every one of them at once; mViewsOpen is how deep a view reading
' another view has gone.
Private mViewColumns As Collection
Private mViewsOpen As Collection

' Set while a view is being run for the shape of its answer rather than the
' answer. Every source underneath hands back its columns and no rows, so
' nothing is joined, filtered or grouped on the way to the column list.
Private mColumnsOnly As Boolean

' The system rowsets as they were last built. sys.all_columns is a row per
' column of every table served, and building it walks them all, types each
' one and copies the lot into an array; the Object Explorer's Tables node
' is one batch that names sys.tables ten times and the columns seven, and
' each statement of it was building them again. They are worked out from
' the tables and nothing else, so what invalidates them is a table added,
' removed or changed, which is where they are dropped.
Private mSystemRowsets As Collection

' Which connection's open transaction holds each table it has written, by
' table. A connection's own transaction is kept with its variables; this is
' the part every connection has to see.
Private mHolds As Collection

' Where a connection's transaction is kept among its variables: names that
' begin with two at signs, which a statement's own variables cannot, and
' which a read of any @@ name takes to the server's variables instead.
Private Const TRAN_COUNT As String = "@@__trancount"
Private Const TRAN_NAME As String = "@@__tranname"
Private Const TRAN_SNAPSHOTS As String = "@@__snapshots"
Private Const TRAN_SAVEPOINTS As String = "@@__savepoints"
Private Const TRAN_HOLDS As String = "@@__holds"

' Where a connection keeps the number of its last statement's error, for
' @@ERROR to read, beside its transaction.
Private Const LAST_ERROR As String = "@@__error"

' An answer that is a statement's error rather than rows or a count, kept
' among the others when the batch carries on past it.
Private Const ANSWER_FAILED As Long = -2

' How many views deep a read may go before it is refused, which is the limit
' a real server puts on one thing reading another.
Private Const VIEW_NESTING_LIMIT As Long = 32

' What @@ROWCOUNT answers: what the statement before this one came to.
Private mRowCount As Long

' The output parameters of the call being answered, when it has any.
Private mCallOutputs As Collection

' The OUTPUT parameters the RPC being answered declared, kept by the
' connection until the statement has run and they can be handed back.
Private mOutputs As Collection

' The server-side cursors open, by handle, each the table it was opened over,
' and the last handle given out. Handles are the server's rather than a
' connection's, so two connections never hold the same one.
Private mCursors As Collection
Private mLastCursor As Long
' The common table expressions the statement being answered began with,
' each run once and kept as its columns and rows.
Private mCommonTables As Collection
Private mLoadedSources As Collection
Private mNoExistsShortcut As Boolean
Private mOuterReads As Long

' The statement being answered.
Private mSelectItems As Collection
Private mFromItems As Collection
Private mFromName As String
Private mWhereNode As Variant
Private mGroupBy As Collection
Private mHavingNode As Variant
Private mOrderBy As Collection
Private mTopCount As Long
Private mTopPercent As Double

' OFFSET and FETCH, which is how a client asks for one page of an answer.
Private mOffsetCount As Long
Private mFetchCount As Long
Private mIsDistinct As Boolean


' ============================================================================
' Bytes
'
' A byte array plus a used count, passed as a pair, is the buffer type here.
' VBA has no growable byte buffer and ReDim Preserve copies, so appends grow
' capacity by doubling and callers carry the length themselves.
' ============================================================================

' How many bytes an array holds, answering zero for one that was never
' dimensioned. VBA raises on UBound of an unallocated array rather than
' reporting -1, so the trap is the test.
Public Function ByteCount(ByRef data() As Byte) As Long
    On Error Resume Next
    ByteCount = UBound(data) - LBound(data) + 1
    If Err.Number <> 0 Then
        Err.Clear
        ByteCount = 0
    End If
    On Error GoTo 0
End Function

' Room for at least wanted bytes, keeping what is already there.
'
' UBound is asked directly rather than through ByteCount, because this runs
' once per field of every row: a million-value answer asked an error-trapped
' function a million times to be told there was already room.
Public Sub EnsureCapacity(ByRef data() As Byte, ByVal wanted As Long)
    Dim capacity As Long

    On Error GoTo Unallocated
    capacity = UBound(data) + 1
    On Error GoTo 0
    If capacity >= wanted Then Exit Sub

    Do While capacity < wanted
        capacity = capacity * 2
    Loop
    ReDim Preserve data(0 To capacity - 1)
    Exit Sub

Unallocated:
    On Error GoTo 0
    capacity = 256
    Do While capacity < wanted
        capacity = capacity * 2
    Loop
    ReDim data(0 To capacity - 1)
End Sub

' Append count bytes of source to target, advancing used.
Public Sub AppendBytes(ByRef target() As Byte, ByRef used As Long, _
                       ByRef source() As Byte, Optional ByVal count As Long = -1)
    If count < 0 Then count = ByteCount(source)
    If count = 0 Then Exit Sub

    EnsureCapacity target, used + count
    CopyMemory target(used), source(0), CLngPtr(count)
    used = used + count
End Sub

' A string's own bytes, straight into the buffer.
'
' A VBA string is already UTF-16, so encoding one is a copy and nothing
' else. Going through a Byte array to do it allocates an array, copies into
' it, and then asks its length through an error-trapped UBound twice: once
' at the call and once inside AppendBytes. Over a column of twenty thousand
' strings that is forty thousand error handlers set up and torn down for a
' length the string already knows.
Public Sub AppendText(ByRef target() As Byte, ByRef used As Long, _
                      ByRef text As String)
    Dim count As Long

    count = LenB(text)
    If count = 0 Then Exit Sub

    EnsureCapacity target, used + count
    CopyMemory target(used), ByVal StrPtr(text), CLngPtr(count)
    used = used + count
End Sub

Public Sub AppendByte(ByRef target() As Byte, ByRef used As Long, _
                      ByVal value As Long)
    EnsureCapacity target, used + 1
    target(used) = value And &HFF
    used = used + 1
End Sub

' Two bytes, most significant first. The TDS header is the only place this
' order appears, so it is named for the header rather than for the protocol.
Public Sub AppendUInt16BE(ByRef target() As Byte, ByRef used As Long, _
                          ByVal value As Long)
    EnsureCapacity target, used + 2
    target(used) = (value \ &H100) And &HFF
    target(used + 1) = value And &HFF
    used = used + 2
End Sub

Public Sub AppendUInt16LE(ByRef target() As Byte, ByRef used As Long, _
                          ByVal value As Long)
    EnsureCapacity target, used + 2
    target(used) = value And &HFF
    target(used + 1) = (value \ &H100) And &HFF
    used = used + 2
End Sub

' Four bytes, least significant first, which is every multi-byte field in
' the protocol except the packet header.
Public Sub AppendUInt32LE(ByRef target() As Byte, ByRef used As Long, _
                          ByVal value As Long)
    EnsureCapacity target, used + 4
    target(used) = value And &HFF
    target(used + 1) = (value \ &H100) And &HFF
    target(used + 2) = (value \ &H10000) And &HFF
    target(used + 3) = (value \ &H1000000) And &HFF
    used = used + 4
End Sub

' The bytes of a string as UCS-2, which is how TDS writes every string.
Public Function Utf16Bytes(ByVal text As String) As Byte()
    Dim out() As Byte

    If LenB(text) = 0 Then
        Utf16Bytes = EmptyBytes()
        Exit Function
    End If
    out = text
    Utf16Bytes = out
End Function

' The first count bytes, as an array of exactly that length.
Public Function LeftBytes(ByRef data() As Byte, ByVal count As Long) As Byte()
    Dim out() As Byte

    If count <= 0 Then
        LeftBytes = EmptyBytes()
        Exit Function
    End If
    ReDim out(0 To count - 1)
    CopyMemory out(0), data(0), CLngPtr(count)
    LeftBytes = out
End Function

' count bytes starting at start.
Public Function MidBytes(ByRef data() As Byte, ByVal start As Long, _
                         ByVal count As Long) As Byte()
    Dim out() As Byte

    If count <= 0 Then
        MidBytes = EmptyBytes()
        Exit Function
    End If
    ReDim out(0 To count - 1)
    CopyMemory out(0), data(start), CLngPtr(count)
    MidBytes = out
End Function

' A byte array holding nothing.
'
' The body is deliberately empty. ReDim cannot express a zero-length array,
' so the only empty a VBA Byte() has is an unallocated one, and an
' array-valued function that assigns nothing returns exactly that. ByteCount
' reads it as zero.
Public Function EmptyBytes() As Byte()
End Function

Public Function ToHex(ByRef data() As Byte) As String
    Dim out As String
    Dim i As Long
    Dim count As Long

    count = ByteCount(data)
    If count = 0 Then Exit Function

    out = Space$(count * 2)
    For i = 0 To count - 1
        Mid$(out, i * 2 + 1, 2) = Right$("0" & Hex$(data(i)), 2)
    Next i
    ToHex = LCase$(out)
End Function

' The double eight bytes hold, in the order the wire uses.
Public Function DoubleFromBytes(ByRef data() As Byte, _
                                Optional ByVal at As Long = 0) As Double
    Dim value As Double

    If ByteCount(data) < at + 8 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.DoubleFromBytes", _
                  "a double needs eight bytes"
    End If
    CopyMemory value, data(at), 8^
    DoubleFromBytes = value
End Function

Public Function FromHex(ByVal text As String) As Byte()
    Dim out() As Byte
    Dim i As Long

    If Len(text) = 0 Then
        FromHex = EmptyBytes()
        Exit Function
    End If
    If Len(text) Mod 2 <> 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.FromHex", _
                  "a hex string needs an even number of digits"
    End If

    ReDim out(0 To Len(text) \ 2 - 1)
    For i = 0 To UBound(out)
        out(i) = CByte("&H" & Mid$(text, i * 2 + 1, 2))
    Next i
    FromHex = out
End Function


' ============================================================================
' Packets
' ============================================================================

' Frame one payload as a single packet. The payload must fit; BuildMessage
' handles anything that might not.
Public Function BuildPacket(ByVal packetType As Long, ByRef payload() As Byte, _
                            Optional ByVal status As Long = TDS_STATUS_EOM, _
                            Optional ByVal spid As Long = 0, _
                            Optional ByVal packetId As Long = 1, _
                            Optional ByVal window As Long = 0) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim length As Long

    length = TDS_HEADER_SIZE + ByteCount(payload)
    If length > TDS_MAX_PACKET_SIZE Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.BuildPacket", _
                  "a payload of " & ByteCount(payload) & " bytes exceeds the " & _
                  "maximum packet size; use BuildMessage to split it"
    End If

    EnsureCapacity out, length
    AppendByte out, used, packetType
    AppendByte out, used, status
    AppendUInt16BE out, used, length
    AppendUInt16BE out, used, spid
    AppendByte out, used, packetId
    AppendByte out, used, window
    AppendBytes out, used, payload

    BuildPacket = LeftBytes(out, used)
End Function

' Split a payload across as many packets as it needs, returning them joined.
' Only the last carries END_OF_MESSAGE. The packet id counts up and wraps at
' 256, which the protocol specifies rather than leaving to an overflow. It
' starts where the caller says because the reference server numbers PRELOGIN
' from 1 and its TLS handshake packets from 0.
'
' An empty payload still produces one packet: a message has to be terminated
' before the peer will act on it, and a message of no packets never is.
Public Function BuildMessage(ByVal packetType As Long, ByRef payload() As Byte, _
                             Optional ByVal packetSize As Long = TDS_DEFAULT_PACKET_SIZE, _
                             Optional ByVal spid As Long = 0, _
                             Optional ByVal packetIdStart As Long = 1) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim body As Long
    Dim total As Long
    Dim offset As Long
    Dim take As Long
    Dim index As Long
    Dim chunk() As Byte
    Dim packet() As Byte
    Dim isLast As Boolean

    If packetSize <= TDS_HEADER_SIZE Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.BuildMessage", _
                  "packetSize must exceed the " & TDS_HEADER_SIZE & "-byte header"
    End If
    If packetSize > TDS_MAX_PACKET_SIZE Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.BuildMessage", _
                  "packetSize cannot exceed " & TDS_MAX_PACKET_SIZE
    End If

    body = packetSize - TDS_HEADER_SIZE
    total = ByteCount(payload)

    Do
        take = total - offset
        If take > body Then take = body
        isLast = (offset + take >= total)

        chunk = MidBytes(payload, offset, take)
        packet = BuildPacket(packetType, chunk, _
                             IIf(isLast, TDS_STATUS_EOM, TDS_STATUS_NORMAL), _
                             spid, (packetIdStart + index) Mod 256)
        AppendBytes out, used, packet

        offset = offset + take
        index = index + 1
    Loop While offset < total

    BuildMessage = LeftBytes(out, used)
End Function

Public Function PacketType(ByRef data() As Byte, Optional ByVal at As Long = 0) As Long
    PacketType = data(at)
End Function

Public Function PacketLength(ByRef data() As Byte, Optional ByVal at As Long = 0) As Long
    PacketLength = CLng(data(at + 2)) * &H100 + CLng(data(at + 3))
End Function

Public Function PacketSpid(ByRef data() As Byte, Optional ByVal at As Long = 0) As Long
    PacketSpid = CLng(data(at + 4)) * &H100 + CLng(data(at + 5))
End Function

Public Function PacketIsEndOfMessage(ByRef data() As Byte, _
                                     Optional ByVal at As Long = 0) As Boolean
    PacketIsEndOfMessage = ((data(at + 1) And TDS_STATUS_EOM) <> 0)
End Function

' Join the packets of one message. Answers the payload through the arguments
' and reports whether a whole message was there: a buffer holding only part
' of one is the normal case on a socket, not an error.
'
' consumed is how many bytes of the buffer the message used, so a caller can
' drop exactly that and keep whatever followed it in the same read.
Public Function Reassemble(ByRef stream() As Byte, ByVal available As Long, _
                           ByRef messageType As Long, ByRef payload() As Byte, _
                           ByRef consumed As Long) As Boolean
    Dim offset As Long
    Dim length As Long
    Dim out() As Byte
    Dim used As Long
    Dim chunk() As Byte

    messageType = 0
    consumed = 0

    Do While offset + TDS_HEADER_SIZE <= available
        length = PacketLength(stream, offset)
        If length < TDS_HEADER_SIZE Then
            Err.Raise ERR_PROTOCOL, "SqlBridge.Reassemble", _
                      "the TDS length field says " & length & ", which is " & _
                      "smaller than the " & TDS_HEADER_SIZE & "-byte header"
        End If
        If offset + length > available Then Exit Do

        If messageType = 0 Then messageType = PacketType(stream, offset)
        chunk = MidBytes(stream, offset + TDS_HEADER_SIZE, length - TDS_HEADER_SIZE)
        AppendBytes out, used, chunk

        If PacketIsEndOfMessage(stream, offset) Then
            consumed = offset + length
            payload = LeftBytes(out, used)
            Reassemble = True
            Exit Function
        End If
        offset = offset + length
    Loop

    Reassemble = False
End Function


' ============================================================================
' PRELOGIN
'
' Parsed messages are held as a Collection of two-element arrays, token and
' data, in the order they arrived. A Dictionary would lose that order, and
' the order is part of what makes the answer look like SQL Server's.
' ============================================================================

Public Function ParsePrelogin(ByRef payload() As Byte) As Collection
    Dim options As Collection
    Dim cursor As Long
    Dim available As Long
    Dim token As Long
    Dim offset As Long
    Dim length As Long
    Dim entry(0 To 1) As Variant

    Set options = New Collection
    available = ByteCount(payload)

    Do
        If cursor >= available Then
            Err.Raise ERR_PROTOCOL, "SqlBridge.ParsePrelogin", _
                      "the PRELOGIN option table has no terminator"
        End If
        If payload(cursor) = PRELOGIN_TERMINATOR Then Exit Do
        If cursor + 5 > available Then
            Err.Raise ERR_PROTOCOL, "SqlBridge.ParsePrelogin", _
                      "the PRELOGIN option entry at " & cursor & _
                      " runs past the payload"
        End If

        token = payload(cursor)
        offset = CLng(payload(cursor + 1)) * &H100 + CLng(payload(cursor + 2))
        length = CLng(payload(cursor + 3)) * &H100 + CLng(payload(cursor + 4))
        If offset + length > available Then
            Err.Raise ERR_PROTOCOL, "SqlBridge.ParsePrelogin", _
                      "PRELOGIN option 0x" & Hex$(token) & " points to bytes " & _
                      offset & ".." & (offset + length) & ", past the " & _
                      available & "-byte payload"
        End If

        entry(0) = token
        entry(1) = MidBytes(payload, offset, length)
        options.Add entry
        cursor = cursor + 5
    Loop

    Set ParsePrelogin = options
End Function

' Serialise options back to a payload. Offsets are computed from the entry
' count, so a parsed message rebuilds to the same bytes as long as the
' original packed its data contiguously in table order. The reference server
' does.
Public Function BuildPrelogin(ByVal options As Collection) As Byte()
    Dim table() As Byte
    Dim tableUsed As Long
    Dim blob() As Byte
    Dim blobUsed As Long
    Dim tableSize As Long
    Dim entry As Variant
    Dim data() As Byte
    Dim out() As Byte
    Dim used As Long

    tableSize = options.count * 5 + 1

    For Each entry In options
        data = entry(1)
        AppendByte table, tableUsed, entry(0)
        AppendUInt16BE table, tableUsed, tableSize + blobUsed
        AppendUInt16BE table, tableUsed, ByteCount(data)
        AppendBytes blob, blobUsed, data
    Next entry
    AppendByte table, tableUsed, PRELOGIN_TERMINATOR

    AppendBytes out, used, table, tableUsed
    AppendBytes out, used, blob, blobUsed
    BuildPrelogin = LeftBytes(out, used)
End Function

' The data an option carries, or an empty array when it is absent.
Public Function PreloginOption(ByVal options As Collection, _
                               ByVal token As Long) As Byte()
    Dim entry As Variant

    For Each entry In options
        If entry(0) = token Then
            PreloginOption = entry(1)
            Exit Function
        End If
    Next entry
    PreloginOption = EmptyBytes()
End Function

Public Function PreloginHas(ByVal options As Collection, ByVal token As Long) As Boolean
    Dim entry As Variant

    For Each entry In options
        If entry(0) = token Then
            PreloginHas = True
            Exit Function
        End If
    Next entry
End Function

' The ENCRYPTION byte a client sent, or -1 if it sent no such option.
Public Function PreloginEncryption(ByVal options As Collection) As Long
    Dim data() As Byte

    data = PreloginOption(options, PRELOGIN_ENCRYPTION)
    If ByteCount(data) = 0 Then
        PreloginEncryption = -1
    Else
        PreloginEncryption = data(0)
    End If
End Function

' The VERSION option's six bytes: major, minor, build big-endian, subbuild.
Public Function PreloginVersionBytes(ByVal major As Long, ByVal minor As Long, _
                                     ByVal build As Long, _
                                     Optional ByVal subbuild As Long = 0) As Byte()
    Dim out() As Byte
    Dim used As Long

    AppendByte out, used, major
    AppendByte out, used, minor
    AppendUInt16BE out, used, build
    AppendUInt16BE out, used, subbuild
    PreloginVersionBytes = LeftBytes(out, used)
End Function

' The response, mirroring the options the client asked about.
'
' A real server answers what it was sent and nothing else. Measured through a
' proxy against SQL Server 2025: a modern driver is answered with six
' options, and the legacy "SQL Server" ODBC driver, which asks about four,
' is answered with four. Sending it the extra two makes it read the response
' as a version older than SQL Server 6.5 and hang up.
'
' With nothing to mirror, all six go out, which is the shape the capture
' recorded and what every modern driver asks for.
Public Function PreloginResponse(Optional ByVal asked As Collection = Nothing, _
                                 Optional ByVal encryption As Long = ENCRYPT_OFF, _
                                 Optional ByVal mars As Boolean = False) As Collection
    Dim answers As Collection
    Dim out As Collection
    Dim entry As Variant

    Set answers = New Collection
    answers.Add PreloginEntry(PRELOGIN_VERSION, _
        PreloginVersionBytes(SERVER_VERSION_MAJOR, SERVER_VERSION_MINOR, _
                             SERVER_VERSION_BUILD))
    answers.Add PreloginEntry(PRELOGIN_ENCRYPTION, ByteOf(encryption))
    answers.Add PreloginEntry(PRELOGIN_INSTOPT, ByteOf(0))
    answers.Add PreloginEntry(PRELOGIN_THREADID, EmptyBytes())
    answers.Add PreloginEntry(PRELOGIN_MARS, ByteOf(IIf(mars, 1, 0)))
    answers.Add PreloginEntry(PRELOGIN_TRACEID, EmptyBytes())

    If asked Is Nothing Then
        Set PreloginResponse = answers
        Exit Function
    End If

    Set out = New Collection
    For Each entry In answers
        ' VERSION and ENCRYPTION always go: a client cannot proceed without
        ' either, whether or not it thought to ask.
        If entry(0) = PRELOGIN_VERSION Or entry(0) = PRELOGIN_ENCRYPTION _
           Or PreloginHas(asked, entry(0)) Then
            out.Add entry
        End If
    Next entry
    Set PreloginResponse = out
End Function

' What this server answers a client's ENCRYPTION option with.
'
' A client that asked for ON is saying it will not read cleartext afterwards,
' so agreeing to OFF strands it: it waits for encrypted bytes that never come
' and times out after the login. The reference server's OFF is the answer to
' an OFF request rather than a fixed policy.
Public Function NegotiateEncryption(ByVal requested As Long, _
                                    Optional ByVal policy As Long = ENCRYPT_OFF) As Long
    If policy = ENCRYPT_REQUIRED Then
        NegotiateEncryption = ENCRYPT_ON
    ElseIf requested = ENCRYPT_ON Or requested = ENCRYPT_REQUIRED Then
        NegotiateEncryption = ENCRYPT_ON
    Else
        NegotiateEncryption = ENCRYPT_OFF
    End If
End Function

Public Function PreloginEntry(ByVal token As Long, ByRef data() As Byte) As Variant
    Dim entry(0 To 1) As Variant

    entry(0) = token
    entry(1) = data
    PreloginEntry = entry
End Function

Private Function ByteOf(ByVal value As Long) As Byte()
    Dim out(0 To 0) As Byte

    out(0) = value And &HFF
    ByteOf = out
End Function


' ============================================================================
' Sockets
' ============================================================================

' A listening socket on one address and port, ready to accept without
' blocking.
Private Function OpenListener(ByVal port As Long, ByVal address As String) As LongPtr
    Dim listener As LongPtr
    Dim addr(0 To 15) As Byte
    Dim mode As Long
    Dim reuse As Long
    Dim octets() As String
    Dim code As Long
    Dim i As Long

    listener = sock_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If listener = INVALID_SOCKET Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.OpenListener", _
                  "could not create a socket: winsock error " & WSAGetLastError()
    End If

    ' Without this a restart inside the operating system's TIME_WAIT window
    ' is refused, which is exactly when a server is most likely to restart.
    reuse = 1
    sock_setsockopt listener, SOL_SOCKET, SO_REUSEADDR, reuse, 4

    ' sockaddr_in: family, port big-endian, four address bytes, eight unused.
    addr(0) = AF_INET
    addr(2) = (port \ 256) And &HFF
    addr(3) = port And &HFF
    octets = Split(address, ".")
    For i = 0 To 3
        addr(4 + i) = CLng(octets(i)) And &HFF
    Next i

    If sock_bind(listener, addr(0), 16) = SOCKET_ERROR Then
        code = WSAGetLastError()
        sock_closesocket listener
        If code = WSAEADDRINUSE Then
            Err.Raise ERR_TRANSPORT, "SqlBridge.OpenListener", _
                      "port " & port & " is already in use; another server, or " & _
                      "SQL Server itself, is listening there"
        End If
        Err.Raise ERR_TRANSPORT, "SqlBridge.OpenListener", _
                  "could not bind to " & address & ":" & port & _
                  ": winsock error " & code
    End If

    If sock_listen(listener, 8) = SOCKET_ERROR Then
        code = WSAGetLastError()
        sock_closesocket listener
        Err.Raise ERR_TRANSPORT, "SqlBridge.OpenListener", _
                  "could not listen on " & address & ":" & port & _
                  ": winsock error " & code
    End If

    mode = 1
    If sock_ioctlsocket(listener, FIONBIO, mode) = SOCKET_ERROR Then
        code = WSAGetLastError()
        sock_closesocket listener
        Err.Raise ERR_TRANSPORT, "SqlBridge.OpenListener", _
                  "could not make the listener non-blocking: winsock error " & code
    End If

    OpenListener = listener
End Function

' What winsock last reported on this thread, or "nothing waiting".
'
' `WSAGetLastError` reads the thread's last-error value, and anything else
' that runs on that thread can overwrite it first. Under a harness that
' polls Excel over COM while a macro is running, the accept in the idle loop
' came back with 5: a Windows access-denied left there by something else,
' where winsock's own codes all start at 10004. Reading a code winsock
' cannot have produced as "nothing waiting" costs a spin round the loop; the
' alternative cost a raise that took the whole server down, and, in the two
' places that read a send or a receive, a live connection dropped.
Private Function LastSocketError() As Long
    Dim code As Long

    code = WSAGetLastError()
    If code < WSABASEERR Then code = WSAEWOULDBLOCK
    LastSocketError = code
End Function

' The next waiting connection, or zero when none is waiting.
Private Function AcceptOne(ByVal listener As LongPtr) As LongPtr
    Dim client As LongPtr
    Dim mode As Long
    Dim code As Long

    client = sock_accept(listener, ByVal 0&, ByVal 0&)
    If client = INVALID_SOCKET Then
        code = LastSocketError()
        If code <> WSAEWOULDBLOCK Then
            Err.Raise ERR_TRANSPORT, "SqlBridge.AcceptOne", _
                      "accept failed: winsock error " & code
        End If
        Exit Function
    End If

    ' A socket accepted from a non-blocking listener inherits nothing, so it
    ' has to be told separately.
    mode = 1
    sock_ioctlsocket client, FIONBIO, mode
    AcceptOne = client
End Function

' Read whatever has arrived into target, growing it, and answer how many
' bytes that was. Zero means nothing was waiting; RECEIVE_CLOSED means the
' peer has gone.
Private Function ReceiveInto(ByVal s As LongPtr, ByRef target() As Byte, _
                             ByRef used As Long) As Long
    Dim chunk(0 To 16383) As Byte
    Dim got As Long

    got = sock_recv(s, chunk(0), 16384, 0)
    If got > 0 Then
        EnsureCapacity target, used + got
        CopyMemory target(used), chunk(0), CLngPtr(got)
        used = used + got
        ReceiveInto = got
        Exit Function
    End If

    If got = 0 Then
        ReceiveInto = RECEIVE_CLOSED
        Exit Function
    End If

    If LastSocketError() = WSAEWOULDBLOCK Then
        ReceiveInto = RECEIVE_NOTHING
    Else
        ' A reset is how a client that gave up looks from here, so it counts
        ' as a closed connection rather than a fault worth raising.
        ReceiveInto = RECEIVE_CLOSED
    End If
End Function

' Write every byte, or report that the peer is gone. A non-blocking socket
' can accept part of a write, so this loops rather than trusting one call.
Private Function SendAll(ByVal s As LongPtr, ByRef data() As Byte, _
                         ByVal count As Long) As Boolean
    Dim sent As Long
    Dim wrote As Long
    Dim spins As Long

    If count <= 0 Then
        SendAll = True
        Exit Function
    End If

    Do While sent < count
        wrote = sock_send(s, data(sent), count - sent, 0)
        If wrote > 0 Then
            sent = sent + wrote
            spins = 0
        Else
            If LastSocketError() <> WSAEWOULDBLOCK Then Exit Function
            ' The send buffer is full. Giving the message pump a turn is the
            ' only way it drains, and a client that never reads must not hold
            ' this loop forever.
            spins = spins + 1
            If spins > 10000 Then Exit Function
            DoEvents
        End If
    Loop
    SendAll = True
End Function


' ============================================================================
' The certificate
' ============================================================================

Private Sub CreateCertificate()
    Dim subjectBlobBytes() As Byte
    Dim subjectSize As Long
    Dim subject As CRYPTOAPI_BLOB
    Dim provInfo As CRYPT_KEY_PROV_INFO
    Dim signature As CRYPT_ALGORITHM_IDENTIFIER
    Dim objectId() As Byte
    Dim subjectName As String
    Dim notBefore As SYSTEMTIME
    Dim notAfter As SYSTEMTIME
    Dim issuedAt As Date

    mContainer = "vbaSQLBridge-" & Hex$(GetTickCount64())

    If CryptAcquireContextW(mProvider, StrPtr(mContainer), 0, PROV_RSA_AES, _
                            CRYPT_NEWKEYSET) = 0 Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.CreateCertificate", _
                  "could not create a key container: " & WindowsCode(Err.LastDllError)
    End If

    If CryptGenKey(mProvider, AT_KEYEXCHANGE, _
                   (RSA_KEY_BITS * &H10000) Or CRYPT_EXPORTABLE, mKey) = 0 Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.CreateCertificate", _
                  "could not generate a " & RSA_KEY_BITS & "-bit key: " & _
                  WindowsCode(Err.LastDllError)
    End If

    subjectName = "CN=" & mServerName
    If CertStrToNameW(X509_ASN_ENCODING, StrPtr(subjectName), CERT_X500_NAME_STR, _
                      0, ByVal 0&, subjectSize, 0) = 0 Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.CreateCertificate", _
                  "could not encode the subject name: " & WindowsCode(Err.LastDllError)
    End If
    ReDim subjectBlobBytes(0 To subjectSize - 1)
    CertStrToNameW X509_ASN_ENCODING, StrPtr(subjectName), CERT_X500_NAME_STR, _
                   0, subjectBlobBytes(0), subjectSize, 0
    subject.cbData = subjectSize
    subject.pbData = VarPtr(subjectBlobBytes(0))

    provInfo.pwszContainerName = StrPtr(mContainer)
    provInfo.pwszProvName = 0                 ' the default AES provider
    provInfo.dwProvType = PROV_RSA_AES
    provInfo.dwKeySpec = AT_KEYEXCHANGE

    objectId = AnsiBytes(SHA256_RSA_OID)
    signature.pszObjId = VarPtr(objectId(0))

    ' Backdated a day so a client whose clock runs slightly behind this one
    ' does not reject a certificate generated moments ago.
    issuedAt = UtcNow()
    notBefore = ToSystemTime(issuedAt - 1)
    notAfter = ToSystemTime(issuedAt + CERTIFICATE_DAYS)

    mCertificate = CertCreateSelfSignCertificate(mProvider, subject, 0, provInfo, _
                                                 signature, notBefore, notAfter, 0)
    If mCertificate = 0 Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.CreateCertificate", _
                  "could not create the certificate: " & WindowsCode(Err.LastDllError)
    End If
End Sub

Private Sub ReleaseCertificate()
    Dim scrap As LongPtr

    If mCertificate <> 0 Then
        CertFreeCertificateContext mCertificate
        mCertificate = 0
    End If
    If mKey <> 0 Then
        CryptDestroyKey mKey
        mKey = 0
    End If
    If mProvider <> 0 Then
        CryptReleaseContext mProvider, 0
        mProvider = 0
    End If
    If LenB(mContainer) > 0 Then
        ' The key outlives the process otherwise: a container is a file in
        ' the user's profile rather than something the handle owned.
        CryptAcquireContextW scrap, StrPtr(mContainer), 0, PROV_RSA_AES, _
                             CRYPT_DELETEKEYSET
        mContainer = vbNullString
    End If
End Sub

Private Function UtcNow() As Date
    Dim stamp As SYSTEMTIME

    GetSystemTime stamp
    UtcNow = DateSerial(stamp.wYear, stamp.wMonth, stamp.wDay) + _
             TimeSerial(stamp.wHour, stamp.wMinute, stamp.wSecond)
End Function

Private Function ToSystemTime(ByVal value As Date) As SYSTEMTIME
    Dim out As SYSTEMTIME

    out.wYear = Year(value)
    out.wMonth = Month(value)
    out.wDay = Day(value)
    out.wHour = Hour(value)
    out.wMinute = Minute(value)
    out.wSecond = Second(value)
    ToSystemTime = out
End Function

' A null-terminated ANSI string as bytes, for the API fields that take LPSTR.
Private Function AnsiBytes(ByVal text As String) As Byte()
    Dim out() As Byte

    out = StrConv(text & vbNullChar, vbFromUnicode)
    AnsiBytes = out
End Function

Private Function StringFromPointer(ByVal pointer As LongPtr) As String
    Dim length As Long
    Dim out As String

    If pointer = 0 Then Exit Function
    length = lstrlenW(pointer)
    If length <= 0 Then Exit Function

    out = Space$(length)
    CopyMemory ByVal StrPtr(out), ByVal pointer, CLngPtr(length * 2)
    StringFromPointer = out
End Function

Private Function WindowsCode(ByVal code As Long) As String
    WindowsCode = "0x" & Right$("00000000" & Hex$(code), 8)
End Function


' ============================================================================
' Credentials and security contexts
' ============================================================================

' The Schannel credential a server presents. Held on the server rather than
' per connection: acquiring one costs a key container lookup, and every
' connection presents the same certificate.
Private Sub AcquireTlsCredentials()
    Dim cred As SCHANNEL_CRED
    Dim certificate As LongPtr
    Dim expiry As SECURITY_INTEGER
    Dim status As Long

    certificate = mCertificate
    cred.dwVersion = SCHANNEL_CRED_VERSION
    cred.cCreds = 1
    cred.paCred = VarPtr(certificate)
    cred.grbitEnabledProtocols = SP_PROT_TLS1_2_SERVER

    status = AcquireCredentialsHandleW(0, StrPtr(UNISP_NAME), SECPKG_CRED_INBOUND, _
                                       0, cred, 0, 0, mTlsCredentials, expiry)
    If status <> SEC_E_OK Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.AcquireTlsCredentials", _
                  "Schannel refused the certificate: " & WindowsCode(status)
    End If
    mHasTlsCredentials = True
End Sub

Private Sub AcquireLoginCredentials()
    Dim expiry As SECURITY_INTEGER
    Dim status As Long

    status = AcquireCredentialsHandleW(0, StrPtr(NEGOTIATE_NAME), _
                                       SECPKG_CRED_INBOUND, 0, ByVal 0&, 0, 0, _
                                       mLoginCredentials, expiry)
    If status <> SEC_E_OK Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.AcquireLoginCredentials", _
                  "could not acquire a Negotiate credential: " & WindowsCode(status)
    End If
    mHasLoginCredentials = True
End Sub

Private Sub ReleaseCredentials()
    If mHasTlsCredentials Then
        FreeCredentialsHandle mTlsCredentials
        mHasTlsCredentials = False
    End If
    If mHasLoginCredentials Then
        FreeCredentialsHandle mLoginCredentials
        mHasLoginCredentials = False
    End If
End Sub


' ============================================================================
' The TLS tunnel, one connection's worth
'
' Handshake records arrive as the payload of TDS packets and application
' records arrive bare, so the engine is driven by buffers rather than by a
' socket. That is also what makes the whole login path testable with no
' socket at all.
' ============================================================================

' Feed handshake bytes and collect the reply.
'
' data is consumed: whatever the engine did not use is moved to the front
' and count is set to how much that was, because the same read can hold the
' end of the handshake and the start of what follows it.
Private Function TlsHandshakeStep(ByRef data() As Byte, ByRef count As Long, _
                                  ByRef reply() As Byte, _
                                  ByRef replyLen As Long) As Long
    Dim inSec(0 To 1) As SecBuffer
    Dim outSec(0 To 0) As SecBuffer
    Dim inDesc As SecBufferDesc
    Dim outDesc As SecBufferDesc
    Dim expiry As SECURITY_INTEGER
    Dim attrs As Long
    Dim status As Long
    Dim leftover As Long
    Dim credential As SecHandle

    replyLen = 0

    inSec(0).cbBuffer = count
    inSec(0).BufferType = SECBUFFER_TOKEN
    inSec(0).pvBuffer = VarPtr(data(0))
    inSec(1).cbBuffer = 0
    inSec(1).BufferType = SECBUFFER_EMPTY
    inSec(1).pvBuffer = 0
    inDesc.ulVersion = SECBUFFER_VERSION
    inDesc.cBuffers = 2
    inDesc.pBuffers = VarPtr(inSec(0))

    outSec(0).cbBuffer = 0
    outSec(0).BufferType = SECBUFFER_TOKEN
    outSec(0).pvBuffer = 0
    outDesc.ulVersion = SECBUFFER_VERSION
    outDesc.cBuffers = 1
    outDesc.pBuffers = VarPtr(outSec(0))

    credential = ServerCredential(True)
    If mHasTlsContext Then
        status = AcceptSecurityContext(credential, VarPtr(mTlsContext), inDesc, _
                                       ASC_REQ_ALLOCATE_MEMORY Or ASC_REQ_STREAM, _
                                       0, mTlsContext, outDesc, attrs, expiry)
    Else
        status = AcceptSecurityContext(credential, 0, inDesc, _
                                       ASC_REQ_ALLOCATE_MEMORY Or ASC_REQ_STREAM, _
                                       0, mTlsContext, outDesc, attrs, expiry)
    End If

    ' More of this flight is still on the wire. Nothing was consumed.
    If status = SEC_E_INCOMPLETE_MESSAGE Then
        TlsHandshakeStep = status
        Exit Function
    End If

    If status = SEC_E_OK Or status = SEC_I_CONTINUE_NEEDED _
       Or status = SEC_I_COMPLETE_NEEDED Or status = SEC_I_COMPLETE_AND_CONTINUE Then
        mHasTlsContext = True
        If outSec(0).cbBuffer > 0 And outSec(0).pvBuffer <> 0 Then
            replyLen = outSec(0).cbBuffer
            EnsureCapacity reply, replyLen
            CopyMemory reply(0), ByVal outSec(0).pvBuffer, CLngPtr(replyLen)
            FreeContextBuffer outSec(0).pvBuffer
        End If

        If inSec(1).BufferType = SECBUFFER_EXTRA Then
            leftover = inSec(1).cbBuffer
            CopyMemory data(0), data(count - leftover), CLngPtr(leftover)
            count = leftover
        Else
            count = 0
        End If

        If status = SEC_E_OK Then
            mTlsEstablished = True
            If QueryContextAttributesW(mTlsContext, SECPKG_ATTR_STREAM_SIZES, _
                                       mStreamSizes) <> SEC_E_OK Then
                Err.Raise ERR_TRANSPORT, "SqlBridge.TlsHandshakeStep", _
                          "the tunnel came up but would not report its record sizes"
            End If
        End If
        TlsHandshakeStep = status
        Exit Function
    End If

    Err.Raise ERR_TRANSPORT, "SqlBridge.TlsHandshakeStep", _
              "the TLS handshake failed: " & WindowsCode(status)
End Function

' Decrypt whole application records out of the front of source, appending
' the plaintext to target. Leftover ciphertext is moved to the front of
' source and sourceUsed set to what is left, so a partial record waits for
' the rest rather than being lost.
Private Sub TlsUnwrap(ByRef source() As Byte, ByRef sourceUsed As Long, _
                      ByRef target() As Byte, ByRef targetUsed As Long)
    Dim sec(0 To 3) As SecBuffer
    Dim desc As SecBufferDesc
    Dim status As Long
    Dim qop As Long
    Dim i As Long
    Dim leftover As Long
    Dim plain() As Byte

    Do While sourceUsed > 0
        sec(0).cbBuffer = sourceUsed
        sec(0).BufferType = SECBUFFER_DATA
        sec(0).pvBuffer = VarPtr(source(0))
        For i = 1 To 3
            sec(i).cbBuffer = 0
            sec(i).BufferType = SECBUFFER_EMPTY
            sec(i).pvBuffer = 0
        Next i
        desc.ulVersion = SECBUFFER_VERSION
        desc.cBuffers = 4
        desc.pBuffers = VarPtr(sec(0))

        status = DecryptMessage(mTlsContext, desc, 0, qop)
        If status = SEC_E_INCOMPLETE_MESSAGE Then Exit Sub
        If status <> SEC_E_OK Then
            Err.Raise ERR_TRANSPORT, "SqlBridge.TlsUnwrap", _
                      "could not decrypt what the client sent: " & WindowsCode(status)
        End If

        leftover = 0
        For i = 0 To 3
            If sec(i).BufferType = SECBUFFER_DATA And sec(i).cbBuffer > 0 Then
                EnsureCapacity plain, sec(i).cbBuffer
                CopyMemory plain(0), ByVal sec(i).pvBuffer, CLngPtr(sec(i).cbBuffer)
                AppendBytes target, targetUsed, plain, sec(i).cbBuffer
            ElseIf sec(i).BufferType = SECBUFFER_EXTRA And sec(i).cbBuffer > 0 Then
                leftover = sec(i).cbBuffer
            End If
        Next i

        If leftover > 0 Then
            CopyMemory source(0), source(sourceUsed - leftover), CLngPtr(leftover)
            sourceUsed = leftover
        Else
            sourceUsed = 0
        End If
    Loop
End Sub

' Encrypt plaintext into application records ready for the wire. Anything
' longer than one record's worth is split, because Schannel encrypts up to
' cbMaximumMessage at a time and a result set is routinely larger.
Private Function TlsWrap(ByRef plaintext() As Byte, ByVal count As Long) As Byte()
    Dim out() As Byte
    Dim outUsed As Long
    Dim record() As Byte
    Dim sec(0 To 3) As SecBuffer
    Dim desc As SecBufferDesc
    Dim status As Long
    Dim offset As Long
    Dim take As Long
    Dim total As Long

    Do While offset < count
        take = count - offset
        If take > mStreamSizes.cbMaximumMessage Then take = mStreamSizes.cbMaximumMessage

        ReDim record(0 To mStreamSizes.cbHeader + take + mStreamSizes.cbTrailer - 1)
        CopyMemory record(mStreamSizes.cbHeader), plaintext(offset), CLngPtr(take)

        sec(0).cbBuffer = mStreamSizes.cbHeader
        sec(0).BufferType = SECBUFFER_STREAM_HEADER
        sec(0).pvBuffer = VarPtr(record(0))
        sec(1).cbBuffer = take
        sec(1).BufferType = SECBUFFER_DATA
        sec(1).pvBuffer = VarPtr(record(mStreamSizes.cbHeader))
        sec(2).cbBuffer = mStreamSizes.cbTrailer
        sec(2).BufferType = SECBUFFER_STREAM_TRAILER
        sec(2).pvBuffer = VarPtr(record(mStreamSizes.cbHeader + take))
        sec(3).cbBuffer = 0
        sec(3).BufferType = SECBUFFER_EMPTY
        sec(3).pvBuffer = 0
        desc.ulVersion = SECBUFFER_VERSION
        desc.cBuffers = 4
        desc.pBuffers = VarPtr(sec(0))

        status = EncryptMessage(mTlsContext, 0, desc, 0)
        If status <> SEC_E_OK Then
            Err.Raise ERR_TRANSPORT, "SqlBridge.TlsWrap", _
                      "could not encrypt the reply: " & WindowsCode(status)
        End If

        total = sec(0).cbBuffer + sec(1).cbBuffer + sec(2).cbBuffer
        AppendBytes out, outUsed, record, total
        offset = offset + take
    Loop

    TlsWrap = LeftBytes(out, outUsed)
End Function


' ============================================================================
' Windows Authentication
' ============================================================================

' One leg of the SSPI exchange. Answers the status, fills reply with the
' token to send back, and sets complete when Windows has accepted the
' client.
Private Function LoginStep(ByRef token() As Byte, ByVal count As Long, _
                           ByRef reply() As Byte, ByRef replyLen As Long, _
                           ByRef complete As Boolean) As Long
    Dim inSec(0 To 0) As SecBuffer
    Dim outSec(0 To 0) As SecBuffer
    Dim inDesc As SecBufferDesc
    Dim outDesc As SecBufferDesc
    Dim expiry As SECURITY_INTEGER
    Dim attrs As Long
    Dim status As Long
    Dim credential As SecHandle

    replyLen = 0
    complete = False

    inSec(0).cbBuffer = count
    inSec(0).BufferType = SECBUFFER_TOKEN
    inSec(0).pvBuffer = VarPtr(token(0))
    inDesc.ulVersion = SECBUFFER_VERSION
    inDesc.cBuffers = 1
    inDesc.pBuffers = VarPtr(inSec(0))

    outSec(0).cbBuffer = 0
    outSec(0).BufferType = SECBUFFER_TOKEN
    outSec(0).pvBuffer = 0
    outDesc.ulVersion = SECBUFFER_VERSION
    outDesc.cBuffers = 1
    outDesc.pBuffers = VarPtr(outSec(0))

    credential = ServerCredential(False)
    If mHasLoginContext Then
        status = AcceptSecurityContext(credential, VarPtr(mLoginContext), inDesc, _
                                       ASC_REQ_ALLOCATE_MEMORY Or ASC_REQ_CONNECTION, _
                                       0, mLoginContext, outDesc, attrs, expiry)
    Else
        status = AcceptSecurityContext(credential, 0, inDesc, _
                                       ASC_REQ_ALLOCATE_MEMORY Or ASC_REQ_CONNECTION, _
                                       0, mLoginContext, outDesc, attrs, expiry)
    End If

    If status <> SEC_E_OK And status <> SEC_I_CONTINUE_NEEDED _
       And status <> SEC_I_COMPLETE_NEEDED And status <> SEC_I_COMPLETE_AND_CONTINUE Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.LoginStep", _
                  "Windows refused the login: " & WindowsCode(status)
    End If

    mHasLoginContext = True
    If outSec(0).cbBuffer > 0 And outSec(0).pvBuffer <> 0 Then
        replyLen = outSec(0).cbBuffer
        EnsureCapacity reply, replyLen
        CopyMemory reply(0), ByVal outSec(0).pvBuffer, CLngPtr(replyLen)
        FreeContextBuffer outSec(0).pvBuffer
    End If

    complete = (status = SEC_E_OK)
    If complete Then mUserName = ContextUserName()
    LoginStep = status
End Function

' Who Windows says authenticated, as DOMAIN\user.
Private Function ContextUserName() As String
    Dim namePointer As LongPtr

    If QueryContextAttributesW(mLoginContext, SECPKG_ATTR_NAMES, namePointer) _
       <> SEC_E_OK Then
        Exit Function
    End If
    ContextUserName = StringFromPointer(namePointer)
    FreeContextBuffer namePointer
End Function

Private Sub ReleaseContexts()
    If mHasTlsContext Then
        DeleteSecurityContext mTlsContext
        mHasTlsContext = False
        mTlsEstablished = False
    End If
    If mHasLoginContext Then
        DeleteSecurityContext mLoginContext
        mHasLoginContext = False
    End If
End Sub


' A credential handle belongs to the server, and a Friend member cannot hand
' over a private type, so a connection reads back the two halves and rebuilds
' the handle it passes to SSPI.
Private Function ServerCredential(ByVal forTls As Boolean) As SecHandle
    Dim handle As SecHandle

    If forTls Then
        handle.dwLower = mServer.TlsCredentialLower
        handle.dwUpper = mServer.TlsCredentialUpper
    Else
        handle.dwLower = mServer.LoginCredentialLower
        handle.dwUpper = mServer.LoginCredentialUpper
    End If
    ServerCredential = handle
End Function

Friend Property Get TlsCredentialLower() As LongPtr
    TlsCredentialLower = mTlsCredentials.dwLower
End Property

Friend Property Get TlsCredentialUpper() As LongPtr
    TlsCredentialUpper = mTlsCredentials.dwUpper
End Property

Friend Property Get LoginCredentialLower() As LongPtr
    LoginCredentialLower = mLoginCredentials.dwLower
End Property

Friend Property Get LoginCredentialUpper() As LongPtr
    LoginCredentialUpper = mLoginCredentials.dwUpper
End Property


' ============================================================================
' Tokens
' ============================================================================

Private Function BVarchar(ByVal text As String) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim encoded() As Byte

    AppendByte out, used, Len(text)
    encoded = Utf16Bytes(text)
    AppendBytes out, used, encoded
    BVarchar = LeftBytes(out, used)
End Function

Private Function UsVarchar(ByVal text As String) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim encoded() As Byte

    AppendUInt16LE out, used, Len(text)
    encoded = Utf16Bytes(text)
    AppendBytes out, used, encoded
    UsVarchar = LeftBytes(out, used)
End Function

' A token byte, a two-byte body length, and the body. The tokens that carry
' no length field build their own bytes instead.
Private Function Token(ByVal kind As Long, ByRef body() As Byte) As Byte()
    Dim out() As Byte
    Dim used As Long

    If ByteCount(body) > TDS_MAX_PACKET_SIZE Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.Token", _
                  "a token body of " & ByteCount(body) & " bytes exceeds the " & _
                  "two-byte token length field"
    End If
    AppendByte out, used, kind
    AppendUInt16LE out, used, ByteCount(body)
    AppendBytes out, used, body
    Token = LeftBytes(out, used)
End Function

' Frame an SSPI blob for the server-to-client stream.
'
' The client's half is not symmetric: it arrives in a packet typed SSPI
' carrying the blob raw, with no token byte and no length. Only this
' direction is framed, because only this direction shares a stream with
' other tokens.
Private Function TokenSspi(ByRef blob() As Byte, ByVal count As Long) As Byte()
    Dim body() As Byte

    body = LeftBytes(blob, count)
    TokenSspi = Token(TOKEN_SSPI, body)
End Function

Private Function TokenEnvChange(ByVal kind As Long, ByVal newValue As String, _
                                Optional ByVal oldValue As String = "") As Byte()
    Dim body() As Byte
    Dim used As Long
    Dim part() As Byte

    AppendByte body, used, kind
    part = BVarchar(newValue)
    AppendBytes body, used, part
    part = BVarchar(oldValue)
    AppendBytes body, used, part
    TokenEnvChange = Token(TOKEN_ENV_CHANGE, LeftBytes(body, used))
End Function

' The collation ENVCHANGE, whose value is opaque bytes counted in bytes.
Private Function TokenEnvChangeBytes(ByVal kind As Long, _
                                     ByRef newValue() As Byte) As Byte()
    Dim body() As Byte
    Dim used As Long

    AppendByte body, used, kind
    AppendByte body, used, ByteCount(newValue)
    AppendBytes body, used, newValue
    AppendByte body, used, 0                  ' no previous value
    TokenEnvChangeBytes = Token(TOKEN_ENV_CHANGE, LeftBytes(body, used))
End Function

' A transaction ENVCHANGE that ends one: no new value, and the descriptor of
' the transaction that ended as the old one.
Private Function TokenEnvChangeEnded(ByVal kind As Long, _
                                     ByRef oldValue() As Byte) As Byte()
    Dim body() As Byte
    Dim used As Long

    AppendByte body, used, kind
    AppendByte body, used, 0                  ' no new value
    AppendByte body, used, ByteCount(oldValue)
    AppendBytes body, used, oldValue
    TokenEnvChangeEnded = Token(TOKEN_ENV_CHANGE, LeftBytes(body, used))
End Function

' The body INFO and ERROR share.
'
' The line number is four bytes from TDS 7.2 and two before it. Sending four
' to a 7.1 client puts two bytes it does not expect between this token and
' the next, and everything after is read at the wrong offset.
Private Function MessageBody(ByVal number As Long, ByVal message As String, _
                             ByVal state As Long, ByVal severity As Long, _
                             ByVal server As String, ByVal procedure As String, _
                             ByVal line As Long, ByVal tdsVersion As Long) As Byte()
    Dim body() As Byte
    Dim used As Long
    Dim part() As Byte

    AppendUInt32LE body, used, number
    AppendByte body, used, state
    AppendByte body, used, severity
    part = UsVarchar(message)
    AppendBytes body, used, part
    part = BVarchar(server)
    AppendBytes body, used, part
    part = BVarchar(procedure)
    AppendBytes body, used, part
    If UnsignedLess(tdsVersion, WIDENED_IN) Then
        AppendUInt16LE body, used, line
    Else
        AppendUInt32LE body, used, line
    End If
    MessageBody = LeftBytes(body, used)
End Function

' An informational message the client will surface to its user. The
' reference server sends two during login, reporting the database and
' language it selected, and clients display them: a bare LOGINACK is a
' quieter login than a real one.
Private Function TokenInfo(ByVal number As Long, ByVal message As String, _
                           ByVal state As Long, ByVal server As String, _
                           ByVal tdsVersion As Long) As Byte()
    TokenInfo = Token(TOKEN_INFO, _
        MessageBody(number, message, state, 0, server, "", 1, tdsVersion))
End Function

' Report a failure the client should surface as an error. Same body as INFO;
' the token byte is what makes a client raise rather than log. Severity 16 is
' the conventional level for an error the caller caused and can correct.
Private Function TokenError(ByVal number As Long, ByVal message As String, _
                            ByVal server As String, ByVal severity As Long, _
                            ByVal tdsVersion As Long) As Byte()
    TokenError = Token(TOKEN_ERROR, _
        MessageBody(number, message, 1, severity, server, "", 1, tdsVersion))
End Function

' Tell the client the login succeeded, and what it is talking to.
Private Function TokenLoginAck(ByVal tdsVersion As Long) As Byte()
    Dim body() As Byte
    Dim used As Long
    Dim part() As Byte

    AppendByte body, used, 1                  ' interface: SQL_TSQL
    ' Big-endian here, where LOGIN7 wrote the same value little-endian. A
    ' server that echoed back the bytes it parsed would send 04 00 00 74.
    AppendByte body, used, (tdsVersion \ &H1000000) And &HFF
    AppendByte body, used, (tdsVersion \ &H10000) And &HFF
    AppendByte body, used, (tdsVersion \ &H100) And &HFF
    AppendByte body, used, tdsVersion And &HFF
    part = BVarchar(SERVER_PROGRAM_NAME)
    AppendBytes body, used, part
    AppendByte body, used, SERVER_VERSION_MAJOR
    AppendByte body, used, SERVER_VERSION_MINOR
    AppendByte body, used, (SERVER_VERSION_BUILD \ &H100) And &HFF
    AppendByte body, used, SERVER_VERSION_BUILD And &HFF
    TokenLoginAck = Token(TOKEN_LOGIN_ACK, LeftBytes(body, used))
End Function

' Close off a response. The row count is 64-bit from TDS 7.2 onward and
' 32-bit before it, so a client negotiated down to 7.1 would otherwise read
' four bytes of count and then four bytes of the next token as more count.
Private Function TokenDone(ByVal status As Long, ByVal rowCount As Long, _
                           ByVal tdsVersion As Long) As Byte()
    Dim out() As Byte
    Dim used As Long

    AppendByte out, used, TOKEN_DONE
    AppendUInt16LE out, used, status
    AppendUInt16LE out, used, 0               ' current command
    AppendUInt32LE out, used, rowCount
    If Not UnsignedLess(tdsVersion, WIDENED_IN) Then
        AppendUInt32LE out, used, 0
    End If
    TokenDone = LeftBytes(out, used)
End Function

' An ERROR token and the DONE that closes the failed batch. A client that
' receives the error without a DONE keeps waiting, because nothing has told
' it the batch finished.
Private Function ErrorResponse(ByVal number As Long, ByVal message As String, _
                               ByVal tdsVersion As Long) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim part() As Byte

    part = TokenError(number, message, ServerName, 16, tdsVersion)
    AppendBytes out, used, part
    part = TokenDone(DONE_ERROR, 0, tdsVersion)
    AppendBytes out, used, part
    ErrorResponse = LeftBytes(out, used)
End Function

' The whole token stream a client gets when its login succeeds, ordered as
' the reference server ordered it.
'
' FEATUREEXTACK is deliberately absent. The real server sent one because the
' client asked for session recovery, data classification and UTF-8 support,
' none of which this implements, and acknowledging a feature that does not
' work would be worse than staying quiet about it.
Private Function LoginResponse(ByVal tdsVersion As Long) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim part() As Byte
    Dim collation() As Byte
    Dim database As String
    Dim server As String

    database = ServedDatabase
    server = ServerName

    part = TokenEnvChange(ENV_DATABASE, database, database)
    AppendBytes out, used, part
    part = TokenInfo(5701, "Changed database context to '" & database & "'.", _
                     2, server, tdsVersion)
    AppendBytes out, used, part
    collation = FromHex(DEFAULT_COLLATION)
    part = TokenEnvChangeBytes(ENV_SQL_COLLATION, collation)
    AppendBytes out, used, part
    part = TokenEnvChange(ENV_LANGUAGE, DEFAULT_LANGUAGE)
    AppendBytes out, used, part
    part = TokenInfo(5703, "Changed language setting to " & DEFAULT_LANGUAGE & ".", _
                     1, server, tdsVersion)
    AppendBytes out, used, part
    part = TokenLoginAck(tdsVersion)
    AppendBytes out, used, part
    part = TokenEnvChange(ENV_PACKET_SIZE, CStr(TDS_DEFAULT_PACKET_SIZE), _
                          CStr(TDS_DEFAULT_PACKET_SIZE))
    AppendBytes out, used, part
    part = TokenDone(DONE_FINAL, 0, tdsVersion)
    AppendBytes out, used, part

    LoginResponse = LeftBytes(out, used)
End Function

' TDS versions have the top bit set, so a signed comparison puts 7.4 below
' 7.0. Every version test goes through here.
Private Function UnsignedLess(ByVal left As Long, ByVal right As Long) As Boolean
    Dim a As LongLong
    Dim b As LongLong

    a = left
    b = right
    If a < 0 Then a = a + &H100000000^
    If b < 0 Then b = b + &H100000000^
    UnsignedLess = (a < b)
End Function

' The version to answer a login with: never newer than the client asked for,
' and never older than the oldest this speaks.
'
' This is what makes an old client work at all. The legacy "SQL Server" ODBC
' driver asks for 7.1, and told 7.4 it decides it is talking to something
' from before 6.5 and hangs up. An unrecognised value is treated as the
' newest, because a client naming a version from the future expects to be
' negotiated down rather than refused.
Private Function NegotiateTdsVersion(ByVal requested As Long) As Long
    Dim known As Variant
    Dim best As Long
    Dim i As Long

    known = Array(TDS_VERSION_70, TDS_VERSION_71, TDS_VERSION_72, _
                  TDS_VERSION_73A, TDS_VERSION_73B, TDS_VERSION_74)

    For i = LBound(known) To UBound(known)
        If known(i) = requested Then
            NegotiateTdsVersion = requested
            Exit Function
        End If
    Next i

    best = TDS_VERSION_70
    For i = LBound(known) To UBound(known)
        If UnsignedLess(known(i), requested) Then best = known(i)
    Next i
    NegotiateTdsVersion = best
End Function


' ============================================================================
' LOGIN7
' ============================================================================

' Read what the login says about itself into this connection's fields, and
' answer the SSPI blob it carried.
'
' Offsets come from the client, so every one is checked rather than trusted.
Private Function ParseLogin7(ByRef payload() As Byte, ByRef user As String, _
                             ByRef masked() As Byte) As Byte()
    Dim available As Long
    Dim declared As Long
    Dim cursor As Long
    Dim sspiOffset As Long
    Dim sspiLength As Long
    Dim sspiLong As Long

    available = ByteCount(payload)
    If available < LOGIN7_VARIABLE_START Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.ParseLogin7", _
                  "LOGIN7 needs at least " & LOGIN7_VARIABLE_START & _
                  " bytes, got " & available
    End If

    declared = ReadUInt32LE(payload, 0)
    If declared <> available Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.ParseLogin7", _
                  "LOGIN7 says it is " & declared & " bytes but the message is " & _
                  available
    End If

    mClientTdsVersion = ReadUInt32LE(payload, 4)

    cursor = LOGIN7_FIXED_SIZE
    mHostName = ReadLoginString(payload, cursor)
    user = ReadLoginString(payload, cursor)      ' empty under SSPI
    masked = ReadLoginBytes(payload, cursor)     ' the password, masked
    mAppName = ReadLoginString(payload, cursor)
    ReadLoginString payload, cursor              ' server name
    ReadLoginString payload, cursor              ' extension
    ReadLoginString payload, cursor              ' client interface name
    ReadLoginString payload, cursor              ' language
    mDatabase = ReadLoginString(payload, cursor)

    cursor = cursor + 6                          ' the client's MAC address
    sspiOffset = ReadUInt16LE(payload, cursor)
    sspiLength = ReadUInt16LE(payload, cursor + 2)
    cursor = cursor + 4
    ReadLoginString payload, cursor              ' AtchDBFile
    ReadLoginString payload, cursor              ' ChangePassword
    sspiLong = ReadUInt32LE(payload, cursor)

    If sspiLength = SSPI_LENGTH_ESCAPE Then sspiLength = sspiLong
    If sspiLength = 0 Then
        ParseLogin7 = EmptyBytes()
        Exit Function
    End If
    If sspiOffset + sspiLength > available Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.ParseLogin7", _
                  "LOGIN7 points its SSPI blob at bytes " & sspiOffset & ".." & _
                  (sspiOffset + sspiLength) & ", past the " & available & _
                  "-byte message"
    End If
    ParseLogin7 = MidBytes(payload, sspiOffset, sspiLength)
End Function

' One (offset, character count) pair, advancing the cursor past it.
Private Function ReadLoginString(ByRef payload() As Byte, _
                                 ByRef cursor As Long) As String
    Dim raw() As Byte

    raw = ReadLoginBytes(payload, cursor)
    If ByteCount(raw) > 0 Then ReadLoginString = raw
End Function

' The same pair's bytes as they arrived. The password is masked, and read
' as text before it is unmasked it would only be rubbish.
Private Function ReadLoginBytes(ByRef payload() As Byte, _
                                ByRef cursor As Long) As Byte()
    Dim offset As Long
    Dim characters As Long

    offset = ReadUInt16LE(payload, cursor)
    characters = ReadUInt16LE(payload, cursor + 2)
    cursor = cursor + 4
    If characters = 0 Then
        ReadLoginBytes = EmptyBytes()
        Exit Function
    End If

    If offset + characters * 2 > ByteCount(payload) Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.ReadLoginString", _
                  "LOGIN7 points a string at bytes " & offset & ".." & _
                  (offset + characters * 2) & ", past the " & _
                  ByteCount(payload) & "-byte message"
    End If
    ReadLoginBytes = MidBytes(payload, offset, characters * 2)
End Function

Private Function ReadUInt16LE(ByRef data() As Byte, ByVal at As Long) As Long
    ReadUInt16LE = CLng(data(at)) + CLng(data(at + 1)) * &H100
End Function

Private Function ReadUInt32LE(ByRef data() As Byte, ByVal at As Long) As Long
    Dim value As LongLong

    value = CLngLng(data(at)) + CLngLng(data(at + 1)) * &H100^ + _
            CLngLng(data(at + 2)) * &H10000^ + CLngLng(data(at + 3)) * &H1000000^
    If value > 2147483647^ Then value = value - &H100000000^
    ReadUInt32LE = CLng(value)
End Function


' ============================================================================
' The login state machine
'
' Bytes in, bytes out. The connection owns no socket while it is stepping,
' which is what makes the whole login path testable without one.
' ============================================================================

' Take bytes that arrived and run the machine as far as they allow.
Friend Sub Feed(ByRef data() As Byte, ByVal count As Long)
    If count > 0 Then AppendBytes mIn, mInUsed, data, count
    Do While StepOnce()
    Loop
End Sub

' Everything the machine has produced, cleared as it is taken.
Friend Function TakeOutput(ByRef target() As Byte) As Long
    ' An answer held back by the session window goes out a few packets at a
    ' time, so every turn of the pump releases whatever is allowed now.
    FlushSmp
    target = LeftBytes(mOut, mOutUsed)
    TakeOutput = mOutUsed
    mOutUsed = 0
End Function

Private Function StepOnce() As Boolean
    Select Case mState
        Case STATE_EXPECT_PRELOGIN
            StepOnce = StepPrelogin()
        Case STATE_TLS_HANDSHAKE
            StepOnce = StepHandshake()
        Case STATE_EXPECT_LOGIN
            StepOnce = StepLogin()
        Case STATE_EXPECT_SSPI
            StepOnce = StepSspi()
        Case STATE_READY
            StepOnce = StepReady()
    End Select
End Function

Private Function StepPrelogin() As Boolean
    Dim messageType As Long
    Dim payload() As Byte
    Dim consumed As Long
    Dim agreed As Long
    Dim response() As Byte
    Dim packet() As Byte

    If Not Reassemble(mIn, mInUsed, messageType, payload, consumed) Then Exit Function
    DropFront mIn, mInUsed, consumed

    If messageType <> TDS_PRELOGIN Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.StepPrelogin", _
                  "expected PRELOGIN to open the connection, got packet type 0x" & _
                  Hex$(messageType)
    End If

    Set mClientPrelogin = ParsePrelogin(payload)
    agreed = NegotiateEncryption(PreloginEncryption(mClientPrelogin), _
                                 mServer.EncryptionPolicy)
    mSessionEncrypted = (agreed = ENCRYPT_ON)

    ' MARS is answered yes only where it was asked for. A client that never
    ' asked would be sent a session header it has no reason to expect.
    mMars = PreloginWantsMars(mClientPrelogin)

    response = BuildPrelogin(PreloginResponse(mClientPrelogin, agreed, mMars))
    ' The reference server answers PRELOGIN with a TABULAR_RESULT packet
    ' rather than echoing the PRELOGIN type back.
    packet = BuildPacket(TDS_TABULAR_RESULT, response)
    AppendBytes mOut, mOutUsed, packet

    mServer.Note "prelogin from " & PreloginClientVersion(mClientPrelogin) & _
                 ", encryption " & agreed & IIf(mMars, ", mars", "")
    mState = STATE_TLS_HANDSHAKE
    StepPrelogin = True
End Function

' The handshake records arrive as the payload of TDS packets, so they are
' unframed here and reframed on the way back. Schannel keeps no input of its
' own, so a flight split across packets is accumulated before each step.
Private Function StepHandshake() As Boolean
    Dim messageType As Long
    Dim payload() As Byte
    Dim consumed As Long
    Dim reply() As Byte
    Dim replyLen As Long
    Dim framed() As Byte
    Dim status As Long
    Dim leftover() As Byte

    If Not Reassemble(mIn, mInUsed, messageType, payload, consumed) Then Exit Function
    DropFront mIn, mInUsed, consumed
    AppendBytes mHandshake, mHandshakeUsed, payload

    Do While mHandshakeUsed > 0
        status = TlsHandshakeStep(mHandshake, mHandshakeUsed, reply, replyLen)
        If replyLen > 0 Then
            framed = BuildMessage(TDS_PRELOGIN, LeftBytes(reply, replyLen), _
                                  TDS_DEFAULT_PACKET_SIZE, 0, HANDSHAKE_PACKET_ID)
            AppendBytes mOut, mOutUsed, framed
        End If
        If status = SEC_E_INCOMPLETE_MESSAGE Then Exit Do
        If mTlsEstablished Then
            ' Anything still here is application data, which is not framed.
            ' One read can hold the end of the handshake and the start of
            ' what follows it.
            If mHandshakeUsed > 0 Then
                leftover = LeftBytes(mHandshake, mHandshakeUsed)
                PrependBytes mIn, mInUsed, leftover, mHandshakeUsed
                mHandshakeUsed = 0
            End If
            mServer.Note "tls tunnel up"
            mState = STATE_EXPECT_LOGIN
            Exit Do
        End If
        If replyLen = 0 And status = SEC_I_CONTINUE_NEEDED Then Exit Do
    Loop

    StepHandshake = True
End Function

Private Function StepLogin() As Boolean
    Dim messageType As Long
    Dim payload() As Byte
    Dim blob() As Byte
    Dim user As String
    Dim masked() As Byte
    Dim reply() As Byte
    Dim replyLen As Long
    Dim complete As Boolean

    If Not TakeMessage(messageType, payload) Then Exit Function

    If messageType <> TDS_LOGIN7 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.StepLogin", _
                  "expected LOGIN7 through the tunnel, got packet type 0x" & _
                  Hex$(messageType)
    End If

    blob = ParseLogin7(payload, user, masked)
    mTdsVersion = NegotiateTdsVersion(mClientTdsVersion)

    If ByteCount(blob) = 0 Then
        ' A name and a password rather than a Windows token. There is no
        ' exchange to run, since the credential arrived whole, so it is
        ' decided and answered here. Raised as not implemented, it closed
        ' the socket, and a client read a server that was down where it
        ' should have read a login that had failed.
        AnswerSqlLogin user, masked
        StepLogin = True
        Exit Function
    End If

    ' The tunnel has done its job. Everything from here is cleartext unless
    ' the client asked for an encrypted session, which is what the reference
    ' server does once the login is through.
    LoginStep blob, ByteCount(blob), reply, replyLen, complete
    If replyLen > 0 And Not complete Then
        SendPayload TokenSspi(reply, replyLen), 0
    End If

    If complete Then
        FinishLogin
    Else
        mState = STATE_EXPECT_SSPI
    End If
    StepLogin = True
End Function

Private Function StepSspi() As Boolean
    Dim messageType As Long
    Dim payload() As Byte
    Dim reply() As Byte
    Dim replyLen As Long
    Dim complete As Boolean

    If Not TakeMessage(messageType, payload) Then Exit Function

    If messageType <> TDS_SSPI Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.StepSspi", _
                  "expected an SSPI message, got packet type 0x" & Hex$(messageType)
    End If

    ' The client's half carries the blob raw: no token byte, no length.
    LoginStep payload, ByteCount(payload), reply, replyLen, complete

    ' The token that comes back on the completing step is not sent. SSPI
    ' produces one and SQL Server does not pass it on: measured through a
    ' proxy, a real server answers the last client token with the login
    ' response alone, and sending it as well leaves the legacy ODBC driver
    ' trying to continue a handshake that is already finished.
    If replyLen > 0 And Not complete Then
        SendPayload TokenSspi(reply, replyLen), 0
    End If

    If complete Then FinishLogin
    StepSspi = True
End Function

' Admit or refuse a name and a password, and say which. The refusal is sent
' rather than raised, and the connection closed after it: a client told
' 18456 knows its login failed, where one that finds its socket closed
' reads a server that is down. It is the same refusal for a name that is
' not there and a password that is wrong. The password is unmasked here and
' nowhere else, is not logged, and is not in the refusal, which names only
' the login.
Private Sub AnswerSqlLogin(ByVal user As String, ByRef masked() As Byte)
    Dim admitted As Boolean
    Dim out() As Byte
    Dim used As Long
    Dim part() As Byte

    On Error Resume Next
    admitted = mServer.AdmitsLogin(user, UnmaskPassword(masked))
    If Err.Number <> 0 Then
        Note "could not check the password offered for '" & user & "': " & _
             Err.Description
        admitted = False
    End If
    Err.Clear
    On Error GoTo 0

    If admitted Then
        mUserName = user
        FinishLogin
        Exit Sub
    End If

    If mServer.HasLogins Then
        Note "refused a login for '" & user & "'"
    Else
        Note "refused a login for '" & user & "': no logins are configured, " & _
             "so only Windows authentication is admitted"
    End If
    part = TokenError(LOGIN_FAILED, "Login failed for user '" & user & "'.", _
                      ServerName, LOGIN_FAILED_SEVERITY, mTdsVersion)
    AppendBytes out, used, part
    part = TokenDone(DONE_ERROR, 0, mTdsVersion)
    AppendBytes out, used, part
    SendPayload LeftBytes(out, used)
    mState = STATE_REFUSED
End Sub

' The password a client put in LOGIN7, unmasked. The client swaps the two
' halves of each byte and XORs it with 0xA5, so undoing that XORs first and
' swaps after; done in the client's order it gives plausible rubbish rather
' than failing. A published constant is not encryption: what keeps the
' password private is the TLS tunnel the login travels through.
Friend Function UnmaskPassword(ByRef masked() As Byte) As String
    Dim plain() As Byte
    Dim count As Long
    Dim index As Long
    Dim value As Long

    count = ByteCount(masked)
    If count = 0 Then Exit Function
    ReDim plain(0 To count - 1)
    For index = 0 To count - 1
        value = masked(index) Xor PASSWORD_MASK
        plain(index) = ((value And &HF) * &H10) Or (value \ &H10)
    Next index
    UnmaskPassword = plain
End Function

' Tell the client it is in. Until this goes out the client has authenticated
' but heard nothing back, so it sits waiting on the reply that decides
' whether it connected.
Private Sub FinishLogin()
    mSpid = mServer.NextSessionId()

    Set mSession = New Collection
    mSession.Add mUserName, "login"
    mSession.Add mAppName, "app"
    mSession.Add mHostName, "host"

    SendPayload LoginResponse(mTdsVersion)
    mState = STATE_READY
    mServer.Note "logged in as " & mUserName & " (app '" & mAppName & _
                 "', database '" & mDatabase & "', spid " & mSpid & ")"
End Sub

Private Function StepReady() As Boolean
    Dim messageType As Long
    Dim payload() As Byte
    Dim response() As Byte
    Dim parameters As Collection
    Dim parameterNames As String
    Dim procedureName As String

    If Not TakeMessage(messageType, payload) Then Exit Function

    Select Case messageType
        Case TDS_SSPI
            ' A trailing authentication token, arriving after the login was
            ' already answered. SPNEGO lets the client send one last leg once
            ' the server has accepted, and Windows has: SSPI returned
            ' SEC_E_OK, which is what made the login complete. The legacy
            ' ODBC driver always sends it, and refusing it there ends the
            ' connection with a protocol error rather than a query.
            StepReady = True
            Exit Function

        Case TDS_ATTENTION
            ' A cancellation. Requests here are answered before the next is
            ' read, so by the time one arrives there is nothing left to stop
            ' and the acknowledgement is the whole of the work.
            SendPayload TokenDone(DONE_FINAL Or DONE_ATTENTION, 0, mTdsVersion)
            StepReady = True
            Exit Function

        Case TDS_SQL_BATCH
            mLastQuery = ParseSqlBatch(payload)
            Set mOutputs = Nothing

        Case TDS_RPC
            ' Clients send anything parameterised, and every catalog query
            ' .NET issues, as a call to sp_executesql rather than as a batch.
            Set parameters = New Collection
            mLastQuery = ParseRpc(payload, parameters, parameterNames, _
                                  procedureName)

        Case TDS_TRANSACTION_MANAGER
            ' A linked server opens a transaction around each read and rolls
            ' it back after, as requests of their own rather than statements,
            ' and .NET's BeginTransaction, Save, Rollback and Commit arrive
            ' the same way. Taken for an unknown packet, the connection was
            ' lost here.
            SendPayload TransactionManagerAnswer(payload)
            StepReady = True
            Exit Function

        Case Else
            Err.Raise ERR_PROTOCOL, "SqlBridge.StepReady", _
                      "expected a SQL batch or an RPC, got packet type 0x" & _
                      Hex$(messageType)
    End Select

    ' The names and types, never the values: a parameter is somebody's data
    ' and a log is not the place for it.
    If Len(parameterNames) > 0 Then
        mServer.Note "query: " & mLastQuery & "  [" & parameterNames & "]"
    Else
        mServer.Note "query: " & mLastQuery
    End If

    ' A failed query is a normal answer rather than a broken connection: the
    ' client reports it and stays connected to ask something else.
    On Error GoTo QueryFailed
    Select Case procedureName
        Case ""
            If HasOutputs() Then
                ' OUTPUT parameters, handed back after the rows the way a
                ' real server hands them back from a call.
                response = mServer.AnswerWithOutputs(mLastQuery, _
                                                     mTdsVersion, parameters, _
                                                     mOutputs, mVariables, _
                                                     mTempTables, mSession)
            Else
                response = mServer.AnswerQuery(mLastQuery, mTdsVersion, _
                                               parameters, mVariables, _
                                               mTempTables, mSession)
            End If
        Case "sp_prepexec", "sp_execute"
            ' A prepared statement run, and answered as a call: its rows,
            ' then for sp_prepexec the handle it was prepared under.
            response = mServer.AnswerPrepared(mLastQuery, mCallHandle, _
                                              procedureName = "sp_prepexec", _
                                              mTdsVersion, parameters, _
                                              mVariables, mTempTables, _
                                              mSession)
        Case "sp_prepare"
            ' Prepared and not run: the columns it would answer with, which
            ' is what a linked server asks before it runs an OPENQUERY.
            response = mServer.AnswerDescribed(mLastQuery, mCallHandle, _
                                               mTdsVersion, parameters, _
                                               mVariables, mTempTables, _
                                               mSession)
        Case Else
            ' A client calling something by name is asking about the
            ' catalog, which is not a question the protocol layer can
            ' answer.
            response = mServer.AnswerProcedure(procedureName, parameters, _
                                               mTdsVersion, True, mVariables, _
                                               mTempTables, mSession)
    End Select
    On Error GoTo 0
    SendPayload response
    StepReady = True
    Exit Function

QueryFailed:
    mServer.Note "refused: " & Err.Description
    SendPayload ErrorResponse(mServer.LastQueryErrorNumber, Err.Description, _
                              mTdsVersion)
    StepReady = True
End Function

' What a transaction-manager request is answered with.
'
' The ENVCHANGE a real server sends, carrying an eight-byte descriptor the
' client quotes back in every request's headers until the transaction ends.
' The transaction itself is the one the statements use, so a write a linked
' server makes inside it is put back when it rolls back.
Private Function TransactionManagerAnswer(ByRef payload() As Byte) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim at As Long
    Dim request As Long
    Dim name As String
    Dim chained As Boolean
    Dim chainedName As String
    Dim descriptor() As Byte
    Dim part() As Byte
    Dim before As Long
    Dim after As Long

    at = ReadUInt32LE(payload, 0)            ' past ALL_HEADERS
    request = ReadUInt16LE(payload, at)
    at = at + 2

    On Error GoTo Refused
    Select Case request
        Case TM_BEGIN_XACT
            ' The isolation level asked for, which changes nothing here, then
            ' the name, when a begin carries either.
            If at <= UBound(payload) Then
                at = at + 1
                name = RequestName(payload, at)
            End If
            mServer.TransactionRequested request, name, mVariables, _
                                         mTempTables, before, after
            AppendBegun out, used

        Case TM_COMMIT_XACT, TM_ROLLBACK_XACT
            ' The name, then whether to begin another straight after, with
            ' that one's isolation level and name: a driver committing by
            ' hand ends one transaction and opens the next in one request.
            name = RequestName(payload, at)
            If at <= UBound(payload) Then chained = (payload(at) <> 0)
            at = at + 1
            If chained Then
                at = at + 1
                chainedName = RequestName(payload, at)
            End If
            mServer.TransactionRequested request, name, mVariables, _
                                         mTempTables, before, after

            ' Back to a savepoint, or out of one nested level: the
            ' transaction is still open, so nothing says it ended.
            If after = 0 Then
                descriptor = DescriptorBytes(mTransactionDescriptor)
                If request = TM_COMMIT_XACT Then
                    part = TokenEnvChangeEnded(ENV_COMMIT_TRANSACTION, descriptor)
                Else
                    part = TokenEnvChangeEnded(ENV_ROLLBACK_TRANSACTION, _
                                               descriptor)
                End If
                AppendBytes out, used, part
            End If
            If chained Then
                mServer.TransactionRequested TM_BEGIN_XACT, chainedName, _
                                             mVariables, mTempTables, _
                                             before, after
                AppendBegun out, used
            End If

        Case TM_SAVE_XACT
            name = RequestName(payload, at)
            mServer.TransactionRequested request, name, mVariables, _
                                         mTempTables, before, after

        Case Else
            ' Asking where the coordinator is, or to join or promote a
            ' distributed transaction. There is no coordinator here to join,
            ' and saying so keeps the connection, where answering with
            ' nothing let the client believe it had enlisted.
            TransactionManagerAnswer = ErrorResponse(ERR_NUMBER_BASE, _
                "vbaSQLBridge does not take part in distributed " & _
                "transactions.", mTdsVersion)
            Exit Function
    End Select
    On Error GoTo 0

    part = TokenDone(DONE_FINAL, 0, mTdsVersion)
    AppendBytes out, used, part
    TransactionManagerAnswer = LeftBytes(out, used)
    Exit Function

Refused:
    ' A savepoint marked with nothing open, or a rollback to one that is not
    ' there: the server's own number and words, the ones the statements give.
    mServer.Note "refused: " & Err.Description
    TransactionManagerAnswer = ErrorResponse(mServer.LastQueryErrorNumber, _
                                             Err.Description, mTdsVersion)
End Function

' The ENVCHANGE a begin is answered with: a new descriptor, which the client
' quotes back in the headers of everything it sends inside the transaction.
Private Sub AppendBegun(ByRef out() As Byte, ByRef used As Long)
    Dim descriptor() As Byte
    Dim part() As Byte

    mTransactionDescriptor = mTransactionDescriptor + 1
    descriptor = DescriptorBytes(mTransactionDescriptor)
    part = TokenEnvChangeBytes(ENV_BEGIN_TRANSACTION, descriptor)
    AppendBytes out, used, part
End Sub

' A name a transaction-manager request carries: a length in bytes, then the
' name in UTF-16. Empty where the request named nothing, which is what
' BeginTransaction() and Rollback() with no argument send. Moves at past it.
Private Function RequestName(ByRef payload() As Byte, ByRef at As Long) As String
    Dim length As Long
    Dim index As Long
    Dim name As String

    If at > UBound(payload) Then Exit Function
    length = payload(at)
    at = at + 1
    For index = at To at + length - 2 Step 2
        If index + 1 > UBound(payload) Then Exit For
        name = name & ChrW$(CLng(payload(index)) + CLng(payload(index + 1)) * &H100&)
    Next index
    at = at + length
    RequestName = name
End Function

' A transaction's descriptor: its number, in eight bytes.
Private Function DescriptorBytes(ByVal number As Long) As Byte()
    Dim out() As Byte
    Dim used As Long

    AppendUInt32LE out, used, number
    AppendUInt32LE out, used, 0
    DescriptorBytes = LeftBytes(out, used)
End Function

' The query text out of a SQL batch.
'
' From TDS 7.2 a request carries an ALL_HEADERS block in front of it, whose
' first four bytes are the total length of the block. Before that version
' the payload is the query text and nothing else.
Private Function ParseSqlBatch(ByRef payload() As Byte) As String
    Dim start As Long
    Dim headerLength As Long
    Dim raw() As Byte

    If ByteCount(payload) = 0 Then Exit Function

    If Not UnsignedLess(mTdsVersion, TDS_VERSION_72) And ByteCount(payload) >= 4 Then
        headerLength = ReadUInt32LE(payload, 0)
        If headerLength > 0 And headerLength <= ByteCount(payload) Then
            start = headerLength
        End If
    End If

    raw = MidBytes(payload, start, ByteCount(payload) - start)
    ParseSqlBatch = raw
End Function


' ============================================================================
' Buffers the state machine moves bytes through
' ============================================================================

' Take whatever whole SMP frames are in front, and keep what they carry.
'
' A SYN opens a session and is answered with one; a FIN closes it and is
' answered with one; DATA carries the TDS message the session is asking.
' Anything else is flow control, which is read for its sequence number and
' otherwise left alone: this answers one request at a time, so a window
' never closes on it.
Private Sub UnwrapSmp(ByRef source() As Byte, ByRef sourceUsed As Long)
    Dim at As Long
    Dim flags As Long
    Dim sid As Long
    Dim length As Long
    Dim window As Long
    Dim inner() As Byte

    Do While at + SMP_HEADER_SIZE <= sourceUsed
        If source(at) <> SMP_SMID Then
            Err.Raise ERR_PROTOCOL, "SqlBridge.UnwrapSmp", _
                      "expected a session header, got 0x" & Hex$(source(at))
        End If

        length = ReadUInt32LE(source, at + 4)
        If length < SMP_HEADER_SIZE Then
            Err.Raise ERR_PROTOCOL, "SqlBridge.UnwrapSmp", _
                      "a session header says its length is " & length
        End If
        If at + length > sourceUsed Then Exit Do

        flags = source(at + 1)
        sid = ReadUInt16LE(source, at + 2)
        mSmpInSeq = ReadUInt32LE(source, at + 8)

        ' The window is how far that session may be sent before it waits,
        ' and it only ever moves forward: an ACK carrying a larger one is
        ' the client saying it has read what was sent and will take more.
        window = ReadUInt32LE(source, at + 12)
        If window > SessionWindow(sid) Then SetSessionWindow sid, window

        If (flags And SMP_SYN) <> 0 Then
            ' Opening a session is not answered. A real server's side of
            ' this exchange, read off the wire, holds no SYN at all: it
            ' says nothing until it has something to say, and a SYN going
            ' back is what made the provider hang up.
            '
            ' Only this session is reset. Resetting whatever was shared
            ' threw away another session's answer half way through it.
            mSmpSid = sid
            SetSessionOutSeq sid, 0
            SetSessionWindow sid, window
            ResetSessionQueue sid
        ElseIf (flags And SMP_FIN) <> 0 Then
            ' Closing one is not answered either, for want of any evidence
            ' that it should be.
            mSmpInUsed = 0
        ElseIf (flags And SMP_DATA) <> 0 Then
            ' A buffer each. A client sends a question on one session while
            ' another is still answering, and both can arrive in the same
            ' read: sharing one buffer made the second look like it had
            ' interrupted the first, and the connection was dropped.
            inner = MidBytes(source, at + SMP_HEADER_SIZE, _
                             length - SMP_HEADER_SIZE)
            TakeSessionData sid, inner, length - SMP_HEADER_SIZE
        End If

        at = at + length
    Loop

    If at > 0 Then DropFront source, sourceUsed, at

    ' An ACK may have moved the window on, which is what lets the rest of an
    ' answer already worked out go out.
    FlushSmp
End Sub

' One SMP frame: the header, then whatever it carries.
Private Function SmpFrame(ByVal flags As Long, ByVal sid As Long, _
                          ByVal seq As Long, ByRef payload() As Byte, _
                          ByVal count As Long) As Byte()
    Dim out() As Byte
    Dim used As Long

    AppendByte out, used, SMP_SMID
    AppendByte out, used, flags
    AppendUInt16LE out, used, sid
    AppendUInt32LE out, used, SMP_HEADER_SIZE + count
    AppendUInt32LE out, used, seq
    AppendUInt32LE out, used, mSmpInSeq + SMP_WINDOW
    If count > 0 Then AppendBytes out, used, payload, count
    SmpFrame = LeftBytes(out, used)
End Function

' Every TDS packet of a message wrapped in a session header of its own, and
' queued rather than sent.
'
' SMP carries one TDS packet at a time: the length in its header covers the
' header and the single packet after it, and the sequence number counts
' packets. It is also flow controlled. Every header carries a window, which
' is the highest sequence number the other side may send before it waits to
' be told it may send more, and a client that grants four and is sent
' twenty-eight ends the session with a transport-level error rather than
' reading them.
'
' An answer of a few dozen rows fits inside the window, which is why this
' went unnoticed: a linked server reads the catalog a row or two at a time.
' Anything larger failed, and the Object Explorer opens every table over
' MARS, so a table of any size could not be opened at all.
Private Sub QueueSmp(ByRef message() As Byte, ByVal count As Long)
    Dim at As Long
    Dim length As Long
    Dim packet() As Byte
    Dim waiting(0 To 1) As Variant
    Dim queue As Collection
    Dim seq As Long

    Set queue = SessionQueue(mSmpSid)
    seq = SessionOutSeq(mSmpSid)
    Do While at < count
        length = PacketLength(message, at)
        If length <= 0 Or at + length > count Then length = count - at
        packet = MidBytes(message, at, length)
        seq = seq + 1
        waiting(0) = seq
        waiting(1) = SmpFrame(SMP_DATA, mSmpSid, seq, packet, length)
        queue.Add waiting
        at = at + length
    Loop
    SetSessionOutSeq mSmpSid, seq
End Sub

' Send what each session's window allows, and no more.
Private Sub FlushSmp()
    Dim sid As Variant
    Dim queue As Collection
    Dim waiting As Variant
    Dim frame() As Byte
    Dim window As Long

    If mSmpQueues Is Nothing Then Exit Sub
    For Each sid In mSmpSids()
        Set queue = SessionQueue(CLng(sid))
        window = SessionWindow(CLng(sid))
        Do While queue.count > 0
            waiting = queue(1)
            If CLng(waiting(0)) > window Then Exit Do
            frame = waiting(1)
            If mSessionEncrypted And mTlsEstablished Then
                frame = TlsWrap(frame, ByteCount(frame))
            End If
            AppendBytes mOut, mOutUsed, frame
            queue.Remove 1
        Loop
    Next sid
End Sub

' ---------------------------------------------------------------------
' What each session knows about itself
' ---------------------------------------------------------------------

' The sessions that have anything queued, in the order they were opened.
Private Function mSmpSids() As Collection
    Dim out As Collection
    Dim index As Long

    Set out = New Collection
    If mSmpQueues Is Nothing Then
        Set mSmpSids = out
        Exit Function
    End If
    For index = 1 To mSmpQueues.count
        out.Add mSmpOrder(index)
    Next index
    Set mSmpSids = out
End Function

Private Function SessionQueue(ByVal sid As Long) As Collection
    Dim held As Collection

    If mSmpQueues Is Nothing Then
        Set mSmpQueues = New Collection
        Set mSmpOrder = New Collection
    End If
    On Error Resume Next
    Set held = mSmpQueues(SessionKey(sid))
    On Error GoTo 0
    If held Is Nothing Then
        Set held = New Collection
        mSmpQueues.Add held, SessionKey(sid)
        mSmpOrder.Add sid
    End If
    Set SessionQueue = held
End Function

Private Function SessionOutSeq(ByVal sid As Long) As Long
    On Error Resume Next
    If Not mSmpOutSeqs Is Nothing Then
        SessionOutSeq = mSmpOutSeqs(SessionKey(sid))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SetSessionOutSeq(ByVal sid As Long, ByVal value As Long)
    If mSmpOutSeqs Is Nothing Then Set mSmpOutSeqs = New Collection
    On Error Resume Next
    mSmpOutSeqs.Remove SessionKey(sid)
    Err.Clear
    On Error GoTo 0
    mSmpOutSeqs.Add value, SessionKey(sid)
End Sub

Private Function SessionWindow(ByVal sid As Long) As Long
    On Error Resume Next
    If Not mSmpWindows Is Nothing Then
        SessionWindow = mSmpWindows(SessionKey(sid))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SetSessionWindow(ByVal sid As Long, ByVal value As Long)
    If mSmpWindows Is Nothing Then Set mSmpWindows = New Collection
    On Error Resume Next
    mSmpWindows.Remove SessionKey(sid)
    Err.Clear
    On Error GoTo 0
    mSmpWindows.Add value, SessionKey(sid)
End Sub

Private Function SessionKey(ByVal sid As Long) As String
    SessionKey = "s" & sid
End Function

' What arrived for one session, kept until it makes a whole message.
'
' Each session reassembles on its own, and a message that is complete goes
' on a queue with the session it belongs to, so the answer can be sent back
' on that session however many others are running.
Private Sub TakeSessionData(ByVal sid As Long, ByRef data() As Byte, _
                            ByVal count As Long)
    Dim held As Variant
    Dim buffer() As Byte
    Dim used As Long
    Dim messageType As Long
    Dim payload() As Byte
    Dim consumed As Long
    Dim ready(0 To 2) As Variant

    held = SessionInBuffer(sid)
    If IsArray(held) Then
        buffer = held(0)
        used = held(1)
    Else
        buffer = EmptyBytes()
    End If
    AppendBytes buffer, used, data, count

    Do While Reassemble(buffer, used, messageType, payload, consumed)
        DropFront buffer, used, consumed
        If mSmpReady Is Nothing Then Set mSmpReady = New Collection
        ready(0) = sid
        ready(1) = messageType
        ready(2) = payload
        mSmpReady.Add ready
    Loop
    SetSessionInBuffer sid, buffer, used
End Sub

Private Function SessionInBuffer(ByVal sid As Long) As Variant
    On Error Resume Next
    If Not mSmpInBufs Is Nothing Then
        SessionInBuffer = mSmpInBufs(SessionKey(sid))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SetSessionInBuffer(ByVal sid As Long, ByRef buffer() As Byte, _
                               ByVal used As Long)
    Dim held(0 To 1) As Variant

    If mSmpInBufs Is Nothing Then Set mSmpInBufs = New Collection
    held(0) = LeftBytes(buffer, used)
    held(1) = used
    On Error Resume Next
    mSmpInBufs.Remove SessionKey(sid)
    Err.Clear
    On Error GoTo 0
    mSmpInBufs.Add held, SessionKey(sid)
End Sub

' Throw away what one session had waiting, leaving every other session's
' alone. A session reopened under the same id starts again from nothing.
Private Sub ResetSessionQueue(ByVal sid As Long)
    Dim queue As Collection

    Set queue = SessionQueue(sid)
    Do While queue.count > 0
        queue.Remove 1
    Loop
End Sub

' The next whole message, decrypted first when the session is encrypted.
' Before the tunnel exists, and after it has served its purpose in a
' login-only session, the bytes on the wire are already plaintext.
Private Function TakeMessage(ByRef messageType As Long, _
                             ByRef payload() As Byte) As Boolean
    Dim consumed As Long
    Dim ready As Variant

    If mMars And mState = STATE_READY Then
        ' Past the login, a session that agreed to MARS wraps everything.
        If mSessionEncrypted And mTlsEstablished Then
            If mInUsed > 0 Then TlsUnwrap mIn, mInUsed, mPlain, mPlainUsed
            UnwrapSmp mPlain, mPlainUsed
        Else
            UnwrapSmp mIn, mInUsed
        End If
        If mSmpReady Is Nothing Then Exit Function
        If mSmpReady.count = 0 Then Exit Function

        ' Whichever session has a whole message, answered on that session:
        ' what is being replied to decides where the reply goes.
        ready = mSmpReady(1)
        mSmpReady.Remove 1
        mSmpSid = CLng(ready(0))
        messageType = CLng(ready(1))
        payload = ready(2)
        TakeMessage = True
        Exit Function
    End If

    If mSessionEncrypted And mTlsEstablished Then
        If mInUsed > 0 Then TlsUnwrap mIn, mInUsed, mPlain, mPlainUsed
        If Not Reassemble(mPlain, mPlainUsed, messageType, payload, consumed) Then
            Exit Function
        End If
        DropFront mPlain, mPlainUsed, consumed
    ElseIf mState = STATE_EXPECT_LOGIN Then
        ' The login itself always comes through the tunnel, whatever the
        ' session agreed to afterwards.
        If mInUsed > 0 Then TlsUnwrap mIn, mInUsed, mPlain, mPlainUsed
        If Not Reassemble(mPlain, mPlainUsed, messageType, payload, consumed) Then
            Exit Function
        End If
        DropFront mPlain, mPlainUsed, consumed
    Else
        If Not Reassemble(mIn, mInUsed, messageType, payload, consumed) Then
            Exit Function
        End If
        DropFront mIn, mInUsed, consumed
    End If
    TakeMessage = True
End Function

' Queue a response, encrypting it when the session is encrypted.
'
' Stamped with the session id from the login response onward. SQL Server
' sends zero through the handshake and the real id afterwards, and a client
' reading zero on an established session has nothing to identify it by.
Private Sub SendPayload(ByRef payload() As Byte, _
                        Optional ByVal packetId As Long = 1)
    Dim message() As Byte

    message = BuildMessage(TDS_TABULAR_RESULT, payload, TDS_DEFAULT_PACKET_SIZE, _
                           mSpid, packetId)
    If mMars And mState = STATE_READY Then
        ' Queued, and let out as the client says it may be. Each frame is
        ' encrypted as it goes rather than the answer as a whole, because
        ' the answer now leaves in pieces.
        QueueSmp message, ByteCount(message)
        FlushSmp
        Exit Sub
    End If
    If mSessionEncrypted And mTlsEstablished Then
        message = TlsWrap(message, ByteCount(message))
    End If
    AppendBytes mOut, mOutUsed, message
End Sub

Private Sub DropFront(ByRef buffer() As Byte, ByRef used As Long, _
                      ByVal count As Long)
    If count >= used Then
        used = 0
    ElseIf count > 0 Then
        CopyMemory buffer(0), buffer(count), CLngPtr(used - count)
        used = used - count
    End If
End Sub

Private Sub PrependBytes(ByRef target() As Byte, ByRef used As Long, _
                         ByRef source() As Byte, ByVal count As Long)
    If count <= 0 Then Exit Sub

    EnsureCapacity target, used + count
    If used > 0 Then
        CopyMemory target(count), target(0), CLngPtr(used)
    End If
    CopyMemory target(0), source(0), CLngPtr(count)
    used = used + count
End Sub

' Whether the client asked for MARS, which the linked-server provider does
' and will not connect without.
Private Function PreloginWantsMars(ByVal options As Collection) As Boolean
    Dim entry As Variant
    Dim data() As Byte

    If options Is Nothing Then Exit Function
    For Each entry In options
        If entry(0) = PRELOGIN_MARS Then
            data = entry(1)
            If ByteCount(data) > 0 Then PreloginWantsMars = (data(0) = 1)
            Exit Function
        End If
    Next entry
End Function

Private Function PreloginClientVersion(ByVal options As Collection) As String
    Dim data() As Byte

    data = PreloginOption(options, PRELOGIN_VERSION)
    If ByteCount(data) < 6 Then
        PreloginClientVersion = "an unnamed client"
        Exit Function
    End If
    PreloginClientVersion = data(0) & "." & data(1) & "." & _
                            (CLng(data(2)) * 256 + CLng(data(3)))
End Function


' ============================================================================
' The server
' ============================================================================

' Start answering as SQL Server on a port.
'
' Binds to the loopback address by default. A bridge reachable from the
' network is a database anyone on it can read, and that is a decision for
' whoever runs it rather than a default.
Public Function Serve(Optional ByVal port As Long = 1433, _
                      Optional ByVal address As String = "127.0.0.1") As SqlBridge
    Dim server As SqlBridge

    Set server = New SqlBridge
    server.ConfigureServer port, address
    Set Serve = server
End Function

' A server that answers queries without listening for anyone.
'
' The protocol needs a socket and a certificate; the catalog does not, and
' keeping them apart is what lets the SQL surface be exercised with no
' network in the way, in a test or from a worksheet.
Public Function Offline() As SqlBridge
    Dim server As SqlBridge

    Set server = New SqlBridge
    server.ConfigureCatalog
    Set Offline = server
End Function

Friend Sub ConfigureCatalog()
    mRole = ROLE_SERVER
    mServerName = Environ$("COMPUTERNAME")
    If LenB(mServerName) = 0 Then mServerName = "localhost"
    mDatabase = DEFAULT_DATABASE
    mEncryptionPolicy = ENCRYPT_OFF
    mNextSpid = TDS_FIRST_SESSION_ID
    mLogLimit = 500
    Set mConnections = New Collection
    Set mLog = New Collection
    Set mTables = New Collection
End Sub

Friend Sub ConfigureServer(ByVal port As Long, ByVal address As String)
    Dim winsockData(0 To 511) As Byte

    ConfigureCatalog
    mPort = port
    mAddress = address

    If WSAStartup(&H202, winsockData(0)) <> 0 Then
        Err.Raise ERR_TRANSPORT, "SqlBridge.ConfigureServer", _
                  "winsock would not start"
    End If
    mWinsockStarted = True

    CreateCertificate
    AcquireTlsCredentials
    AcquireLoginCredentials

    mListener = OpenListener(port, address)
    Note "listening on " & address & ":" & port & " as " & mServerName
End Sub

' Take one turn: accept whatever is waiting, then give every connection a
' chance to read, step and write. Never blocks.
Public Sub Poll()
    Dim client As LongPtr
    Dim connection As SqlBridge
    Dim finished As Collection
    Dim index As Long

    If mRole <> ROLE_SERVER Or mListener = 0 Then Exit Sub

    Do
        client = AcceptOne(mListener)
        If client = 0 Then Exit Do
        Set connection = CreateConnection(client)
        mConnections.Add connection
        Note "connection accepted, " & mConnections.count & " open"
    Loop

    Set finished = New Collection
    For index = 1 To mConnections.count
        Set connection = mConnections(index)
        connection.Pump
        If connection.HasClosed Then finished.Add index
    Next index

    For index = finished.count To 1 Step -1
        Set connection = mConnections(finished(index))
        connection.Detach
        mConnections.Remove finished(index)
    Next index
End Sub

' Poll until the time runs out, giving Excel its message pump between turns.
' This is the shape a test or a console script wants; a workbook that should
' stay usable arms the timer in SqlBridgeHost instead.
Public Sub RunFor(ByVal milliseconds As Long)
    Dim deadline As LongLong

    mStopRequested = False
    deadline = GetTickCount64() + milliseconds
    Do While GetTickCount64() < deadline And Not mStopRequested
        Poll
        DoEvents
        SleepMs 1
    Loop
End Sub

' Ask a RunFor loop to finish early. Safe to call from a worksheet button.
Public Sub RequestStop()
    mStopRequested = True
End Sub

' Close every connection, give the port back, and delete the key container
' the certificate lived in.
Public Sub Shutdown()
    Dim connection As SqlBridge
    Dim index As Long

    If mRole <> ROLE_SERVER Then Exit Sub

    If Not mConnections Is Nothing Then
        For index = mConnections.count To 1 Step -1
            Set connection = mConnections(index)
            connection.CloseConnection
            connection.Detach
            mConnections.Remove index
        Next index
    End If

    If mListener <> 0 Then
        sock_closesocket mListener
        mListener = 0
    End If

    ReleaseCredentials
    ReleaseCertificate

    If mWinsockStarted Then
        WSACleanup
        mWinsockStarted = False
    End If
    Note "stopped"
End Sub

Public Property Get Port() As Long
    Port = mPort
End Property

Public Property Get IsListening() As Boolean
    IsListening = (mListener <> 0)
End Property

Public Property Get ConnectionCount() As Long
    If mConnections Is Nothing Then Exit Property
    ConnectionCount = mConnections.count
End Property

' Whether the whole session is encrypted rather than only the login. A client
' that asks for encryption gets it either way; this makes the server insist.
Public Property Get RequireEncryption() As Boolean
    RequireEncryption = (mEncryptionPolicy = ENCRYPT_REQUIRED)
End Property

Public Property Let RequireEncryption(ByVal value As Boolean)
    If value Then
        mEncryptionPolicy = ENCRYPT_REQUIRED
    Else
        mEncryptionPolicy = ENCRYPT_OFF
    End If
End Property

' What the one database this serves is called. A client shows this in its
' tree and puts it in front of every table name.
Public Property Get Database() As String
    Database = mDatabase
End Property

Public Property Let Database(ByVal value As String)
    If Len(Trim$(value)) > 0 Then mDatabase = Trim$(value)
End Property

' What the server calls itself in the messages a client displays.
Public Property Get Name() As String
    Name = mServerName
End Property

Public Property Let Name(ByVal value As String)
    mServerName = value
End Property

' What happened, newest last. A server with no console needs somewhere to
' say what it did, and this is it.
Public Property Get LogText() As String
    Dim lines() As String
    Dim index As Long

    If mLog Is Nothing Then Exit Property
    If mLog.count = 0 Then Exit Property

    ReDim lines(0 To mLog.count - 1)
    For index = 1 To mLog.count
        lines(index - 1) = mLog(index)
    Next index
    LogText = Join(lines, vbLf)
End Property

Public Sub ClearLog()
    If mLog Is Nothing Then Exit Sub
    Set mLog = New Collection
End Sub

' One connection object, whether or not it has a socket. A connection with
' none is exactly what the state machine needs, so the login path can be
' driven from byte arrays with no network in the way.
Friend Function CreateConnection(ByVal socket As LongPtr) As SqlBridge
    Dim connection As SqlBridge

    Set connection = New SqlBridge
    connection.AttachTo Me, socket
    Set CreateConnection = connection
End Function

Friend Sub AttachTo(ByVal server As SqlBridge, ByVal socket As LongPtr)
    mRole = ROLE_CONNECTION
    Set mServer = server
    mSocket = socket
    mState = STATE_EXPECT_PRELOGIN
    mTdsVersion = TDS_VERSION_74

    ' A session's own state, which the server borrows for the length of a
    ' call. Temp tables and variables belong to the connection that made
    ' them, and the catalog behind them does not.
    Set mVariables = New Collection
    Set mTempTables = New Collection
End Sub

Friend Sub Detach()
    ' A connection holds its server and the server holds its connections, so
    ' nothing is released until one of them lets go.
    Set mServer = Nothing
End Sub

Friend Sub Note(ByVal message As String)
    If mRole = ROLE_CONNECTION Then
        If Not mServer Is Nothing Then mServer.Note message
        Exit Sub
    End If
    If mLog Is Nothing Then Exit Sub
    mLog.Add Format$(Now, "hh:nn:ss") & " " & message
    Do While mLog.count > mLogLimit
        mLog.Remove 1
    Loop
End Sub

Friend Property Get EncryptionPolicy() As Long
    EncryptionPolicy = mEncryptionPolicy
End Property

Friend Property Get ConfiguredServerName() As String
    ConfiguredServerName = mServerName
End Property

Friend Property Get ConfiguredDatabase() As String
    ConfiguredDatabase = mDatabase
End Property

' Session ids run from 51; everything below that belongs to the real
' server's own background work, and a client shown one of those knows it is
' not talking to a session. Two bytes on the wire, so it wraps.
Friend Function NextSessionId() As Long
    NextSessionId = mNextSpid
    mNextSpid = mNextSpid + 1
    If mNextSpid > 65535 Then mNextSpid = TDS_FIRST_SESSION_ID
End Function


' ============================================================================
' The connection
' ============================================================================

' Read whatever arrived, step the machine, write whatever it produced.
Friend Sub Pump()
    Dim got As Long
    Dim reads As Long
    Dim outgoing() As Byte
    Dim count As Long

    If mState = STATE_CLOSED Or mState = STATE_FAILED Then Exit Sub
    If mSocket = 0 Then Exit Sub

    Do
        got = ReceiveInto(mSocket, mIn, mInUsed)
        If got = RECEIVE_CLOSED Then
            Note "client disconnected"
            CloseConnection
            Exit Sub
        End If
        If got = RECEIVE_NOTHING Then Exit Do
        reads = reads + 1
    Loop While reads < 64

    On Error GoTo Failed
    Do While StepOnce()
    Loop
    On Error GoTo 0

    count = TakeOutput(outgoing)
    If count > 0 Then
        If Not SendAll(mSocket, outgoing, count) Then
            Note "could not write to the client"
            CloseConnection
        End If
    End If

    ' A refused login has had its refusal sent, and the connection goes
    ' after it, the way a real server hangs up on a login it will not have.
    If mState = STATE_REFUSED Then CloseConnection
    Exit Sub

Failed:
    Note "connection failed: " & Err.Description
    mState = STATE_FAILED
    CloseConnection
End Sub

Friend Sub CloseConnection()
    ' A transaction the client left open goes when it does, as on a real
    ' server: the rows it changed are put back and the tables it held let
    ' go. Left open, they stayed held by a connection that was not there.
    If mRole = ROLE_CONNECTION And Not mServer Is Nothing Then
        mServer.EndAbandoned mVariables, mTempTables, mSession
    End If
    ReleaseContexts
    If mSocket <> 0 Then
        sock_closesocket mSocket
        mSocket = 0
    End If
    If mState <> STATE_FAILED Then mState = STATE_CLOSED
End Sub

Friend Property Get HasClosed() As Boolean
    HasClosed = (mState = STATE_CLOSED Or mState = STATE_FAILED)
End Property

' Who Windows says is on the other end, once it has said so.
Public Property Get UserName() As String
    UserName = mUserName
End Property

Public Property Get ApplicationName() As String
    ApplicationName = mAppName
End Property

Public Property Get IsLoggedIn() As Boolean
    IsLoggedIn = (mState = STATE_READY)
End Property

Public Property Get SessionId() As Long
    SessionId = mSpid
End Property

' The server's name and database, wherever this object is asked from.
Private Property Get ServerName() As String
    If mRole = ROLE_CONNECTION Then
        If Not mServer Is Nothing Then ServerName = mServer.ConfiguredServerName
    Else
        ServerName = mServerName
    End If
End Property

Private Property Get ServedDatabase() As String
    If mRole = ROLE_CONNECTION Then
        If LenB(mDatabase) > 0 Then
            ' One database is served whatever a client names, so this is
            ' what it is told it reached. A client told master and then
            ' shown no master in the list shows no databases at all.
            ServedDatabase = mServer.ConfiguredDatabase
        Else
            ServedDatabase = DEFAULT_DATABASE
        End If
    Else
        ServedDatabase = mDatabase
    End If
End Property

Private Sub Class_Terminate()
    If mRole = ROLE_SERVER Then
        Shutdown
    ElseIf mRole = ROLE_CONNECTION Then
        CloseConnection
    End If
End Sub


' ============================================================================
' Result sets
' ============================================================================

' One column of an answer: what it is called, how it is typed, and how wide.
Friend Function ColumnOf(ByVal name As String, ByVal typeCode As Long, _
                         Optional ByVal size As Long = 0, _
                         Optional ByVal flags As Long = -1, _
                         Optional ByVal qualifier As String = "", _
                         Optional ByVal places As Long = 0) As Variant
    Dim out(0 To 5) As Variant

    If Len(name) > 255 Then
        ' One byte holds the name length. A select list written without
        ' commas parses as one column named after the rest of the statement,
        ' and reaches here as a name nothing can send.
        Err.Raise ERR_PROTOCOL, "SqlBridge.ColumnOf", _
                  "a column name of " & Len(name) & " characters cannot be " & _
                  "sent; the length field holds 255"
    End If
    out(COLUMN_NAME) = name
    out(COLUMN_TYPE) = typeCode
    out(COLUMN_SIZE) = size
    If flags < 0 Then
        out(COLUMN_FLAGS) = SELECT_COLUMN_FLAGS
    Else
        out(COLUMN_FLAGS) = flags
    End If
    out(COLUMN_QUALIFIER) = qualifier
    out(COLUMN_SCALE) = places
    ColumnOf = out
End Function

' How a column declares itself: the type byte and whatever follows it.
Private Function TypeInfo(ByVal typeCode As Long, ByVal size As Long, _
                          Optional ByVal places As Long = 0) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim collation() As Byte

    Select Case typeCode
        Case TYPE_DECIMALN
            ' The size a decimal declares is its precision, and how many
            ' bytes the value takes follows from that.
            AppendByte out, used, TYPE_DECIMALN
            AppendByte out, used, DecimalWidth(size)
            AppendByte out, used, size
            AppendByte out, used, places

        Case TYPE_INTN
            AppendByte out, used, TYPE_INTN
            AppendByte out, used, size
        Case TYPE_FLTN
            AppendByte out, used, TYPE_FLTN
            AppendByte out, used, size
        Case TYPE_BITN
            AppendByte out, used, TYPE_BITN
            AppendByte out, used, 1
        Case TYPE_DATETIMN
            AppendByte out, used, TYPE_DATETIMN
            AppendByte out, used, 8
        Case TYPE_GUIDN
            AppendByte out, used, TYPE_GUIDN
            AppendByte out, used, 16
        Case TYPE_INT2
            AppendByte out, used, TYPE_INT2
        Case TYPE_BIGBINARY, TYPE_BIGVARBINARY
            AppendByte out, used, typeCode
            AppendUInt16LE out, used, size

        Case TYPE_NVARCHAR
            AppendByte out, used, TYPE_NVARCHAR
            If size <= 0 Or size > NVARCHAR_SIZED_LIMIT Then
                AppendUInt16LE out, used, NVARCHAR_MAX_MARKER
            Else
                AppendUInt16LE out, used, size * 2
            End If
            collation = FromHex(DEFAULT_COLLATION)
            AppendBytes out, used, collation
        Case Else
            Err.Raise ERR_PROTOCOL, "SqlBridge.TypeInfo", _
                      "no encoding for column type 0x" & Hex$(typeCode)
    End Select
    TypeInfo = LeftBytes(out, used)
End Function

' One value, encoded against the column that declared it.
' One value, written where the answer is being built.
'
' Written there rather than returned: a value of its own is an array to
' allocate, fill, trim to length and copy in, four times over per field, and
' an answer of a million fields did all of that a million times.
Private Sub AppendEncodedValue(ByRef out() As Byte, ByRef used As Long, _
                               ByVal typeCode As Long, ByVal size As Long, _
                               ByVal value As Variant, _
                               Optional ByVal places As Long = 0)
    Dim text As String
    Dim encoded() As Byte
    Dim absent As Boolean

    absent = IsNull(value) Or IsEmpty(value)

    Select Case typeCode
        Case TYPE_INTN
            If absent Then
                ' A zero length is how a nullable type says NULL.
                AppendByte out, used, 0
            Else
                AppendByte out, used, size
                AppendSignedInteger out, used, CLngLng(value), size
            End If

        Case TYPE_FLTN
            If absent Then
                AppendByte out, used, 0
            Else
                AppendByte out, used, 8
                AppendDouble out, used, CDbl(value)
            End If

        Case TYPE_DECIMALN
            If absent Then
                AppendByte out, used, 0
            Else
                AppendDecimal out, used, value, size, places
            End If

        Case TYPE_BITN
            If absent Then
                AppendByte out, used, 0
            Else
                AppendByte out, used, 1
                AppendByte out, used, IIf(CBool(value), 1, 0)
            End If

        Case TYPE_DATETIMN
            If absent Then
                AppendByte out, used, 0
            Else
                AppendByte out, used, 8
                AppendDateTime out, used, CDate(value)
            End If

        Case TYPE_INT2
            ' The fixed form: two bytes, no length prefix, never NULL.
            If absent Then
                AppendUInt16LE out, used, 0
            Else
                AppendUInt16LE out, used, CLng(value) And &HFFFF&
            End If

        Case TYPE_GUIDN
            If absent Then
                AppendByte out, used, 0
            Else
                AppendByte out, used, 16
                AppendBytes out, used, GuidBytes(CStr(value))
            End If

        Case TYPE_BIGBINARY, TYPE_BIGVARBINARY
            ' A two-byte length, and all ones for null rather than a zero
            ' length, which would be a legitimate empty value.
            If absent Then
                AppendByte out, used, &HFF
                AppendByte out, used, &HFF
            Else
                encoded = FromHex(CStr(value))
                AppendUInt16LE out, used, ByteCount(encoded)
                AppendBytes out, used, encoded
            End If

        Case TYPE_NVARCHAR
            If size <= 0 Or size > NVARCHAR_SIZED_LIMIT Then
                AppendMaxText out, used, value, absent
                Exit Sub
            End If
            If absent Then
                ' NVARCHAR spends its whole two-byte length on the null
                ' marker, where the one-byte types use zero. Sending zero
                ' here would be a legitimate empty string instead.
                AppendByte out, used, &HFF
                AppendByte out, used, &HFF
            Else
                text = CStr(value)
                AppendUInt16LE out, used, LenB(text)
                AppendText out, used, text
            End If

        Case Else
            Err.Raise ERR_PROTOCOL, "SqlBridge.AppendEncodedValue", _
                      "no encoding for column type 0x" & Hex$(typeCode)
    End Select
End Sub

' The chunked form: a total length, then length-prefixed chunks, then a
' zero-length terminator.
Private Sub AppendMaxText(ByRef out() As Byte, ByRef used As Long, _
                          ByVal value As Variant, ByVal absent As Boolean)
    Dim text As String
    Dim i As Long
    Dim length As Long

    If absent Then
        ' A MAX column says NULL with an all-ones total length rather than
        ' the 0xffff a sized one uses.
        For i = 1 To 8
            AppendByte out, used, &HFF
        Next i
        Exit Sub
    End If

    text = CStr(value)
    length = LenB(text)
    If length = 0 Then
        ' An empty MAX value is its own marker. A zero total length followed
        ' by a zero chunk would be read as one chunk of nothing.
        For i = 1 To 8
            AppendByte out, used, 0
        Next i
        Exit Sub
    End If

    AppendUInt32LE out, used, length
    AppendUInt32LE out, used, 0
    AppendUInt32LE out, used, length
    AppendText out, used, text
    AppendUInt32LE out, used, 0
End Sub

' How many bytes a decimal of this precision occupies, sign byte included.
' The four sizes are what the protocol allows and nothing between them.
Private Function DecimalWidth(ByVal digits As Long) As Long
    If digits <= 9 Then
        DecimalWidth = 5
    ElseIf digits <= 19 Then
        DecimalWidth = 9
    ElseIf digits <= 28 Then
        DecimalWidth = 13
    Else
        DecimalWidth = 17
    End If
End Function

' A decimal on the wire: its own length, then a sign byte, then the value
' scaled to a whole number, least significant byte first.
'
' The arithmetic runs in VBA's Decimal, which carries 28 digits where a
' Double carries 15. A numeric(38) beyond that cannot be represented at all
' and is sent as far as it goes rather than not sent.
Private Sub AppendDecimal(ByRef target() As Byte, ByRef used As Long, _
                          ByVal value As Variant, ByVal digits As Long, _
                          ByVal places As Long)
    Dim width As Long
    Dim scaled As Variant
    Dim index As Long
    Dim negative As Boolean

    width = DecimalWidth(digits)
    AppendByte target, used, width

    On Error Resume Next
    scaled = CDec(value)
    If Err.Number <> 0 Then
        Err.Clear
        scaled = CDec(0)
    End If
    For index = 1 To places
        scaled = scaled * CDec(10)
    Next index
    scaled = Int(scaled)
    negative = (scaled < CDec(0))
    If negative Then scaled = -scaled
    On Error GoTo 0

    AppendByte target, used, IIf(negative, 0, 1)
    For index = 1 To width - 1
        AppendByte target, used, CLng(scaled - Int(scaled / CDec(256)) * CDec(256))
        scaled = Int(scaled / CDec(256))
    Next index
End Sub

Private Sub AppendSignedInteger(ByRef target() As Byte, ByRef used As Long, _
                                ByVal value As LongLong, ByVal width As Long)
    Dim i As Long
    Dim remaining As LongLong

    ' A number too big for the column it was declared in cannot go down
    ' it. Sending the low bytes and saying nothing turned 100000 * 100000
    ' into 1410065408, which is a wrong answer with nothing to show for it.
    If Not FitsInteger(value, width) Then
        RaiseQueryError 8115, _
            "Arithmetic overflow error converting expression to data " & _
            "type int."
    End If

    ' Room for the whole number at once. Asking per byte is four checks
    ' where one will do, and every integer of every row pays for them.
    EnsureCapacity target, used + width
    remaining = value
    For i = 1 To width
        target(used) = CLng(remaining And 255^)
        used = used + 1
        remaining = remaining \ 256^
        ' Arithmetic shift: a negative number keeps its sign as it narrows,
        ' which integer division alone does not do for the last byte.
        If value < 0 And remaining = 0 Then remaining = -1
    Next i
End Sub

Private Sub AppendDouble(ByRef target() As Byte, ByRef used As Long, _
                         ByVal value As Double)
    EnsureCapacity target, used + 8
    CopyMemory target(used), value, 8^
    used = used + 8
End Sub

' DATETIMN, the eight-byte form: four bytes of days since 1900-01-01, then
' four of three-hundredths of a second since midnight.
Private Sub AppendDateTime(ByRef target() As Byte, ByRef used As Long, _
                           ByVal value As Date)
    Dim days As Long
    Dim ticks As Long
    Dim seconds As Double

    days = CLng(Int(CDbl(value))) - 2
    seconds = (CDbl(value) - Int(CDbl(value))) * 86400#
    ticks = CLng(seconds * 300#)
    EnsureCapacity target, used + 8
    AppendUInt32LE target, used, days
    AppendUInt32LE target, used, ticks
End Sub

' Declare the shape of the rows that follow.
'
' The user type is four bytes from TDS 7.2 and two before it. Two extra bytes
' per column is enough to make the whole declaration unreadable, which a
' client reports as a protocol error in the stream rather than as anything
' naming a column.
Private Function ColMetadata(ByVal columns As Collection, _
                             ByVal tdsVersion As Long) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim column As Variant
    Dim info() As Byte
    Dim name() As Byte
    Dim flags As Long

    AppendByte out, used, TOKEN_COL_METADATA
    AppendUInt16LE out, used, columns.count

    For Each column In columns
        If UnsignedLess(tdsVersion, TDS_VERSION_72) Then
            AppendUInt16LE out, used, 0
        Else
            AppendUInt32LE out, used, 0
        End If
        ' A fixed-width column cannot say NULL, so the flag is cleared here
        ' rather than trusted: a client reading a nullable declaration on a
        ' fixed type looks for an indicator byte the row does not carry.
        flags = CLng(column(COLUMN_FLAGS))
        If column(COLUMN_TYPE) = TYPE_INT2 Then flags = flags And Not 1
        AppendUInt16LE out, used, flags
        info = TypeInfo(CLng(column(COLUMN_TYPE)), CLng(column(COLUMN_SIZE)), _
                        CLng(NullToZero(column(COLUMN_SCALE))))
        AppendBytes out, used, info
        AppendByte out, used, Len(CStr(column(COLUMN_NAME)))
        name = Utf16Bytes(CStr(column(COLUMN_NAME)))
        AppendBytes out, used, name
    Next column

    ColMetadata = LeftBytes(out, used)
End Function

' A whole answer: the declaration, the rows, and a DONE carrying the count.
'
' No columns means no result set, and the answer is a bare DONE. That is not
' the same as a query returning no rows, which still declares its shape:
' statements like SET produce nothing, and a client sent COLMETADATA for one
' of those reports an invalid cursor state when it later tries to read the
' results it was actually waiting for.
Friend Function ResultSet(ByVal columns As Collection, ByVal rows As Collection, _
                          ByVal tdsVersion As Long, _
                          Optional ByVal more As Boolean = False) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim part() As Byte
    Dim row As Variant
    Dim values As Variant
    Dim index As Long
    Dim column As Variant
    Dim width As Long
    Dim kinds() As Long
    Dim sizes() As Long
    Dim scales() As Long
    Dim first As Long

    If columns Is Nothing Then
        ResultSet = TokenDone(IIf(more, DONE_MORE, DONE_FINAL), 0, tdsVersion)
        Exit Function
    End If
    If columns.count = 0 Then
        ResultSet = TokenDone(IIf(more, DONE_MORE, DONE_FINAL), 0, tdsVersion)
        Exit Function
    End If

    part = ColMetadata(columns, tdsVersion)
    AppendBytes out, used, part

    ' What each column is, read once. Asking the Collection again for every
    ' value of every row is a walk down the list and three conversions, a
    ' million times over on an answer worth measuring.
    width = columns.count
    ReDim kinds(1 To width)
    ReDim sizes(1 To width)
    ReDim scales(1 To width)
    index = 1
    For Each column In columns
        kinds(index) = CLng(column(COLUMN_TYPE))
        sizes(index) = CLng(column(COLUMN_SIZE))
        scales(index) = CLng(NullToZero(column(COLUMN_SCALE)))
        index = index + 1
    Next column

    If Not rows Is Nothing Then
        For Each row In rows
            values = row
            If UBound(values) - LBound(values) + 1 <> width Then
                Err.Raise ERR_PROTOCOL, "SqlBridge.ResultSet", _
                          "a row carries " & (UBound(values) - LBound(values) + 1) & _
                          " values where " & width & " columns were declared"
            End If
            AppendByte out, used, TOKEN_ROW
            first = LBound(values)
            For index = 1 To width
                AppendEncodedValue out, used, kinds(index), sizes(index), _
                                   values(first + index - 1), scales(index)
            Next index
        Next row
    End If

    part = TokenDoneCounted(DONE_COUNT Or IIf(more, DONE_MORE, DONE_FINAL), _
                            RowCountOf(rows), tdsVersion)
    AppendBytes out, used, part
    ResultSet = LeftBytes(out, used)
End Function

Private Function RowCountOf(ByVal rows As Collection) As Long
    If rows Is Nothing Then Exit Function
    RowCountOf = rows.count
End Function

' DONE with the command field the reference server used after a SELECT,
' where the login response sent zero.
Private Function TokenDoneCounted(ByVal status As Long, ByVal rowCount As Long, _
                                  ByVal tdsVersion As Long, _
                                  Optional ByVal command As Long = SELECT_COMMAND) _
                                  As Byte()
    Dim out() As Byte
    Dim used As Long

    AppendByte out, used, TOKEN_DONE
    AppendUInt16LE out, used, status
    AppendUInt16LE out, used, command
    AppendUInt32LE out, used, rowCount
    If Not UnsignedLess(tdsVersion, WIDENED_IN) Then
        AppendUInt32LE out, used, 0
    End If
    TokenDoneCounted = LeftBytes(out, used)
End Function

Private Sub RaiseQueryError(ByVal number As Long, ByVal message As String)
    mQueryErrorNumber = number
    Err.Raise ERR_QUERY, "SqlBridge", message
End Sub

' What a failed statement does to the batch around it, measured on a real
' server one number at a time, as SELECT 'BEFORE'; <the statement>; SELECT
' 'AFTER'. After these the rest of the batch still runs, and the error
' arrives among its answers rather than in place of them.
Private Function GoesOnAfter(ByVal number As Long) As Boolean
    Select Case number
        Case 220, 244, 248, 517, 535, 1222, 2812, 3701, 3902, 3903, 6401, _
             8115, 8134, 8179, 9828
            GoesOnAfter = True
    End Select
End Function

' After these the batch ends, but what ran before it is kept: BEFORE, then
' the error. Every other error is answered by itself, the way an error a
' real server finds before it runs anything is: a syntax error, or a column
' that is not there, answers neither.
Private Function EndsTheBatchAfter(ByVal number As Long) As Boolean
    Select Case number
        Case 208, 241, 245, 281, 530, 628, 3623, 8114, 8169, 8672
            EndsTheBatchAfter = True
    End Select
End Function

' One statement of a batch, kept going past an error a real server carries
' on past: measured, the statements after a divide by zero or a COMMIT with
' nothing open still run, and their results reach the client behind the
' error. Every other error goes on up and ends the batch.
Private Sub RunCaught(ByVal written As String, ByVal answers As Collection)
    Dim raised As Long
    Dim failedIn As String
    Dim message As String

    On Error GoTo Failed
    RunStatement written, answers
    Exit Sub

Failed:
    raised = Err.Number
    failedIn = Err.Source
    message = Err.Description
    ' Only ever reached carrying on, so carrying on is put back, whatever
    ' the statement that failed had set it to on the way.
    mCatching = True
    If raised = ERR_QUERY Then KeepErrorNumber mQueryErrorNumber
    If raised <> ERR_QUERY Then Err.Raise raised, failedIn, message
    If Not GoesOnAfter(mQueryErrorNumber) Then Err.Raise raised, failedIn, message

    answers.Add FailedAnswer(mQueryErrorNumber, message)
    ' Measured: a statement that failed leaves the count at nought.
    mRowCount = 0
End Sub

' One statement, carried past its error where a real server would carry on,
' or not, inside a TRY.
Private Sub RunOne(ByVal written As String, ByVal answers As Collection)
    If mCatching Then
        RunCaught written, answers
    Else
        RunStatement written, answers
    End If
End Sub

' The number @@ERROR reads: the last statement's error on this connection,
' or nought where it finished.
Private Function ErrorNumber() As Long
    If mVariables Is Nothing Then Exit Function
    On Error Resume Next
    ErrorNumber = CLng(mVariables(LAST_ERROR))
    On Error GoTo 0
End Function

Private Sub KeepErrorNumber(ByVal number As Long)
    If mVariables Is Nothing Then Exit Sub
    KeepTranValue LAST_ERROR, number
End Sub

' A statement that finishes puts @@ERROR back to nought, measured against a
' real server, except where it only holds other statements: a DECLARE with
' no value, an IF, a WHILE, an EXEC of text and a plain BEGIN ... END leave
' it to what they ran, so IF 1 = 1 COMMIT still reads 3902 afterwards. A TRY
' clears it once its CATCH has run, which is what writing one is for.
Private Sub ForgetErrorAfter(ByVal head As String, ByVal written As String)
    Dim ends As String
    Dim second As String

    Select Case head
        Case "IF", "WHILE"
            Exit Sub
        Case "DECLARE"
            If InStr(written, "=") = 0 Then Exit Sub
        Case "EXEC", "EXECUTE"
            If ExecRunsText(written) Then Exit Sub
        Case "BEGIN"
            WordAt written, SkipSpaceAt(written, 1), ends
            second = UCase$(WordAt(written, SkipSpaceAt(written, CLng(ends)), _
                                   ends))
            Select Case second
                Case "TRY", "TRAN", "TRANSACTION", "DISTRIBUTED"
                Case Else
                    Exit Sub
            End Select
    End Select
    KeepErrorNumber 0
End Sub

' Whether an EXEC runs text rather than a procedure: EXEC ('...'), or
' sp_executesql named outright.
Private Function ExecRunsText(ByVal written As String) As Boolean
    Dim ends As String
    Dim rest As String

    WordAt written, 1, ends
    rest = TrimWhitespace(Mid$(written, CLng(ends)))
    If Left$(rest, 1) = "(" Then
        ExecRunsText = True
    ElseIf BareProcedureName(ProcedureFromBatch(written)) = "sp_executesql" Then
        ExecRunsText = True
    End If
End Function

Friend Property Get LastQueryErrorNumber() As Long
    If mQueryErrorNumber = 0 Then
        LastQueryErrorNumber = ERR_NUMBER_BASE
    Else
        LastQueryErrorNumber = mQueryErrorNumber
    End If
End Property


' ============================================================================
' Tables
'
' A table is a name, a two-dimensional Variant array whose first row is the
' header, and the columns inferred from it. Excel hands ranges over in
' exactly that shape, so a worksheet becomes a table with no copying.
' ============================================================================

' Serve a worksheet, a range, a table, or an array of values under a name.
'
' The first row is the header. Types are inferred from every value in a
' column rather than per row, because COLMETADATA declares a column once and
' every row is encoded against that declaration.
'
' A worksheet or a range is read again for every statement, so a client sees
' the sheet as it is rather than as it was when the server started. A
' worksheet grows and shrinks with what is typed on it; a range is the
' rectangle it names and nothing outside it. An array is a copy and stays
' one: there is nothing left to read it from.
Public Sub AddTable(ByVal name As String, ByVal source As Variant, _
                    Optional ByVal hasHeader As Variant)
    Dim data As Variant
    Dim raw As Variant
    Dim entry(0 To 7) As Variant
    Dim key As String
    Dim wide As Long
    Dim rowMap() As Long
    Dim columnMap() As Long

    If mRole <> ROLE_SERVER Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.AddTable", _
                  "only a server serves tables"
    End If
    If Len(name) = 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.AddTable", "a table needs a name"
    End If

    ' An argument nobody passed is Missing rather than Empty, and Missing
    ' put into an array is an error value rather than either. Settled here,
    ' so everything after it has one thing to test.
    If IsMissing(hasHeader) Then hasHeader = Empty

    raw = TableValues(source, name)
    data = CleanTable(raw, hasHeader, wide, rowMap, columnMap)

    entry(0) = name
    entry(1) = data
    If wide = 0 Then
        Set entry(2) = New Collection
    Else
        Set entry(2) = InferColumns(data)
    End If

    ' The range itself, not only what it held a moment ago. A workbook is
    ' edited while it is being served, and a client that asks again is
    ' asking about now: keeping the values and dropping the range is how a
    ' server answers with what the sheet used to say.
    If IsObject(source) Then Set entry(3) = source Else entry(3) = Empty
    entry(4) = hasHeader
    entry(5) = raw
    entry(6) = rowMap
    entry(7) = columnMap

    key = LCase$(name)
    On Error Resume Next
    mTables.Remove key
    On Error GoTo 0
    mTables.Add entry, key
    ForgetViewColumns
    Note "serving table " & name & " (" & (UBound(data, 1) - LBound(data, 1)) & _
         " rows, " & (UBound(data, 2) - LBound(data, 2) + 1) & " columns)"
End Sub

' Take the served ranges as they are now.
'
' A table given as an array has nothing to read again and keeps the values
' it was handed. A range whose workbook has closed keeps the last values it
' had, because answering with what was there is better than refusing to
' answer at all.
Private Sub RefreshTables()
    Dim names As Collection
    Dim named As Variant
    Dim entry As Variant
    Dim raw As Variant
    Dim key As String
    Dim wide As Long
    Dim rowMap() As Long
    Dim columnMap() As Long
    Dim wasColumns As Collection
    Dim reshaped As Boolean

    If mTables Is Nothing Then Exit Sub

    ' The names first. Replacing an entry means removing it and adding it
    ' back, and a Collection being walked does not survive that.
    Set names = New Collection
    For Each entry In mTables
        names.Add CStr(entry(0))
    Next entry

    For Each named In names
        key = LCase$(CStr(named))
        entry = mTables(key)
        If UBound(entry) >= 3 Then
            If IsObject(entry(3)) Then
                On Error Resume Next
                raw = TableValues(entry(3), CStr(entry(0)))
                If Err.Number = 0 Then
                    ' The rectangle as it came, compared with the rectangle
                    ' that came last time. Comparing what was made of it
                    ' instead would differ every time and read the sheet
                    ' again for nothing.
                    If Changed(entry(5), raw) Then
                        ' The rowsets carry row counts as well as shapes, so
                        ' a sheet that merely gained a row invalidates them
                        ' even where its columns are the same.
                        ForgetSystemRowsets
                        Set wasColumns = entry(2)
                        entry(5) = raw
                        entry(1) = CleanTable(raw, entry(4), wide, _
                                              rowMap, columnMap)
                        entry(6) = rowMap
                        entry(7) = columnMap
                        If wide = 0 Then
                            Set entry(2) = New Collection
                        Else
                            Set entry(2) = InferColumns(entry(1))
                        End If
                        If SameColumns(wasColumns, entry(2)) = False Then
                            reshaped = True
                        End If
                        mTables.Remove key
                        mTables.Add entry, key
                    End If
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next named

    ' A view's columns are worked out by running it, which for a view over a
    ' join of fifty thousand rows is seconds. They only change when a sheet
    ' underneath changes shape, so typing a value into a cell no longer
    ' throws them away: a client reading the catalog asks about every view
    ' several times over, and it used to run all of them every time.
    If reshaped Then ForgetViewColumns
End Sub

' Whether two column lists say the same thing. A row typed into a sheet
' leaves the columns alone; a column added, renamed or retyped does not.
Private Function SameColumns(ByVal before As Collection, _
                             ByVal after As Collection) As Boolean
    Dim index As Long

    If before Is Nothing Or after Is Nothing Then Exit Function
    If before.count <> after.count Then Exit Function
    For index = 1 To before.count
        If StrComp(CStr(before(index)(COLUMN_NAME)), _
                   CStr(after(index)(COLUMN_NAME)), vbBinaryCompare) <> 0 Then
            Exit Function
        End If
        If before(index)(COLUMN_TYPE) <> after(index)(COLUMN_TYPE) Then
            Exit Function
        End If
        If before(index)(COLUMN_SIZE) <> after(index)(COLUMN_SIZE) Then
            Exit Function
        End If
    Next index
    SameColumns = True
End Function

' Whether a range has anything new in it. Comparing is far cheaper than
' typing every column again, and a sheet nobody has touched is the usual
' case.
Private Function Changed(ByVal before As Variant, _
                         ByVal after As Variant) As Boolean
    Dim row As Long
    Dim column As Long
    Dim earlier As Variant
    Dim later As Variant

    If Not IsArray(before) Then
        Changed = True
        Exit Function
    End If

    On Error GoTo Differs
    If UBound(before, 1) <> UBound(after, 1) Then GoTo Differs
    If UBound(before, 2) <> UBound(after, 2) Then GoTo Differs
    If LBound(before, 1) <> LBound(after, 1) Then GoTo Differs
    If LBound(before, 2) <> LBound(after, 2) Then GoTo Differs

    ' Compared as they are rather than as text. Every cell of the sheet
    ' goes through here before every statement, and turning each one into a
    ' string to compare it was the price of asking whether anything had
    ' changed at all.
    For row = LBound(after, 1) To UBound(after, 1)
        For column = LBound(after, 2) To UBound(after, 2)
            earlier = before(row, column)
            later = after(row, column)
            If IsEmpty(earlier) Or IsNull(earlier) Then
                If Not (IsEmpty(later) Or IsNull(later)) Then GoTo Differs
            ElseIf IsEmpty(later) Or IsNull(later) Then
                GoTo Differs
            ElseIf VarType(earlier) <> VarType(later) Then
                GoTo Differs
            ElseIf earlier <> later Then
                GoTo Differs
            End If
        Next column
    Next row
    On Error GoTo 0
    Exit Function

Differs:
    Changed = True
End Function

' How many correlated reads have been turned into a lookup. A shortcut that
' never fired would agree with the slow way on everything and prove nothing,
' so which shapes take it is counted rather than assumed.
Public Property Get ExistsLookupsBuilt() As Long
    ExistsLookupsBuilt = mExistsLookupsBuilt
End Property

' How many joins have been answered by looking rows up rather than by
' trying every pair.
Public Property Get JoinLookupsBuilt() As Long
    JoinLookupsBuilt = mJoinLookupsBuilt
End Property

Public Property Let JoinShortcut(ByVal wanted As Boolean)
    mNoJoinShortcut = Not wanted
End Property

Public Property Get JoinShortcut() As Boolean
    JoinShortcut = Not mNoJoinShortcut
End Property

' Refuse to write. Off by default, because a workbook served to a client
' that can log in is a workbook that client can be given a reason to change;
' on, this is the read-only bridge it used to be, and says so in the catalog
' so a client shows the database that way.
Public Property Let ReadOnly(ByVal wanted As Boolean)
    mReadOnly = wanted
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property

' Turn the EXISTS lookup off, so a test can put the same question twice
' and compare. Off is the slow way and always correct; on is the shortcut.
Public Property Let ExistsShortcut(ByVal wanted As Boolean)
    mNoExistsShortcut = Not wanted
End Property

Public Property Get ExistsShortcut() As Boolean
    ExistsShortcut = Not mNoExistsShortcut
End Property

Public Sub RemoveTable(ByVal name As String)
    On Error Resume Next
    mTables.Remove LCase$(name)
    On Error GoTo 0
    ForgetViewColumns
End Sub

Public Property Get TableNames() As String
    Dim entry As Variant
    Dim names() As String
    Dim index As Long

    If mTables Is Nothing Then Exit Property
    If mTables.count = 0 Then Exit Property

    ReDim names(0 To mTables.count - 1)
    For Each entry In mTables
        names(index) = entry(0)
        index = index + 1
    Next entry
    TableNames = Join(names, ", ")
End Property

' ----------------------------------------------------------------------------
' Views
' ----------------------------------------------------------------------------

' Serve a statement under a name, the way a view serves one. A client reads
' it as it reads a table and finds it in the catalog as a view, and what it
' answers is worked out when it is read, so a view over a sheet follows the
' sheet. Views are kept among the tables, which is how everything that
' lists what is served finds them.
Public Sub AddView(ByVal name As String, ByVal sql As String)
    Dim entry(0 To 8) As Variant
    Dim bare As String

    bare = Trim$(Replace(Replace(name, "[", ""), "]", ""))
    If Len(bare) = 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.AddView", "a view needs a name"
    End If
    If Len(TrimWhitespace(sql)) = 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.AddView", _
                  "the view '" & bare & "' has no statement"
    End If
    If Not OnlyReads(sql) Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.AddView", _
                  "the view '" & bare & "' is not a read: a view holds a " & _
                  "SELECT, and nothing that changes the workbook"
    End If

    entry(0) = bare
    entry(1) = Empty
    Set entry(2) = New Collection
    entry(3) = Empty
    entry(4) = Empty
    entry(5) = Empty
    entry(6) = Empty
    entry(7) = Empty
    entry(8) = sql
    StoreTable bare, entry
    ForgetViewColumns
End Sub

Public Sub RemoveView(ByVal name As String)
    Dim entry As Variant

    If mTables Is Nothing Then Exit Sub
    On Error Resume Next
    entry = mTables(LCase$(Trim$(Replace(Replace(name, "[", ""), "]", ""))))
    On Error GoTo 0
    If Not IsViewEntry(entry) Then Exit Sub
    RemoveTable name
    ForgetViewColumns
End Sub

Public Sub ClearViews()
    Dim entry As Variant
    Dim names As Collection
    Dim name As Variant

    If mTables Is Nothing Then Exit Sub

    ' The names first: removing an entry from a Collection being walked
    ' does not survive the walk.
    Set names = New Collection
    For Each entry In mTables
        If IsViewEntry(entry) Then names.Add CStr(entry(0))
    Next entry
    For Each name In names
        RemoveTable CStr(name)
    Next name
    ForgetViewColumns
End Sub

' The views being served, never the statements behind them.
Public Property Get ViewNames() As String
    Dim entry As Variant
    Dim names As Collection
    Dim out() As String
    Dim index As Long

    If mTables Is Nothing Then Exit Property
    Set names = New Collection
    For Each entry In mTables
        If IsViewEntry(entry) Then names.Add CStr(entry(0))
    Next entry
    If names.count = 0 Then Exit Property

    ReDim out(0 To names.count - 1)
    For index = 1 To names.count
        out(index - 1) = names(index)
    Next index
    ViewNames = Join(out, ", ")
End Property

' The statement a view stands for, or nothing where the entry is a table.
Private Function ViewSql(ByVal entry As Variant) As String
    If Not IsArray(entry) Then Exit Function
    If UBound(entry) < 8 Then Exit Function
    If VarType(entry(8)) <> vbString Then Exit Function
    ViewSql = CStr(entry(8))
End Function

Private Function IsViewEntry(ByVal entry As Variant) As Boolean
    IsViewEntry = (Len(ViewSql(entry)) > 0)
End Function

' What a view answers now. Run where it is read rather than kept, so a view
' over a sheet follows the sheet, and counted on the way in, so a view that
' reads itself is refused rather than read until Excel gives up.
Private Sub ReadView(ByVal entry As Variant, ByRef outColumns As Collection, _
                     ByRef outRows As Collection)
    Dim kept As Variant
    Dim name As String
    Dim raised As Long
    Dim failedIn As String
    Dim message As String

    name = LCase$(CStr(entry(0)))
    If mViewsOpen Is Nothing Then Set mViewsOpen = New Collection
    If HasKey(mViewsOpen, name) Or mViewsOpen.count >= VIEW_NESTING_LIMIT Then
        RaiseQueryError 217, "Maximum stored procedure, function, trigger, " & _
            "or view nesting level exceeded (limit " & _
            VIEW_NESTING_LIMIT & ")."
    End If
    mViewsOpen.Add name, name

    kept = SaveParseState()
    On Error GoTo Restore
    Set mTokens = Tokenize(ViewSql(entry))
    mCursor = 1
    ParseSelect
    ExecuteWholeSelect outColumns, outRows
    RestoreParseState kept
    mViewsOpen.Remove name
    Exit Sub

Restore:
    raised = Err.Number
    failedIn = Err.Source
    message = Err.Description
    RestoreParseState kept
    On Error Resume Next
    mViewsOpen.Remove name
    On Error GoTo 0
    Err.Raise raised, failedIn, message
End Sub

' The columns an entry answers with: a table's are known, and a view's are
' whatever its statement produces. Worked out once a statement and kept,
' because a client reading the catalog asks about every view at once.
Private Function EntryColumns(ByVal entry As Variant) As Collection
    Dim columns As Collection
    Dim rows As Collection
    Dim key As String
    Dim wasColumnsOnly As Boolean

    If Not IsViewEntry(entry) Then
        Set EntryColumns = entry(2)
        Exit Function
    End If

    key = LCase$(CStr(entry(0)))
    If mViewColumns Is Nothing Then Set mViewColumns = New Collection
    On Error Resume Next
    Set columns = mViewColumns(key)
    On Error GoTo 0
    If Not columns Is Nothing Then
        Set EntryColumns = columns
        Exit Function
    End If

    ' A view whose statement no longer reads, because what it named was
    ' renamed or taken away, is listed with no columns rather than ending
    ' the catalog query that asked about every view there is.
    '
    ' Run for its shape rather than its answer. The sources underneath hand
    ' back their columns and no rows, so the statement is walked, the joins
    ' and the grouping decide what the columns are, and none of it touches a
    ' row. A column whose type its expression does not state is declared the
    ' way a view that answered with no rows has always declared one.
    wasColumnsOnly = mColumnsOnly
    mColumnsOnly = True
    On Error Resume Next
    ReadView entry, columns, rows
    On Error GoTo 0
    mColumnsOnly = wasColumnsOnly
    If columns Is Nothing Then Set columns = New Collection
    mViewColumns.Add columns, key
    Set EntryColumns = columns
End Function

Private Sub ForgetViewColumns()
    Set mViewColumns = Nothing
    ForgetSystemRowsets
End Sub

' Whether an object belongs in the list a client asked for: sys.tables holds
' the tables, sys.views the views, and sys.objects both.
Private Function ObjectBelongs(ByVal which As String, _
                               ByVal entry As Variant) As Boolean
    Select Case which
        Case "tables"
            ObjectBelongs = Not IsViewEntry(entry)
        Case "views", "all_views"
            ObjectBelongs = IsViewEntry(entry)
        Case Else
            ObjectBelongs = True
    End Select
End Function

' ----------------------------------------------------------------------------
' SQL Server logins
' ----------------------------------------------------------------------------

' Let a client log in with a name and a password, a SQL Server login rather
' than Windows authentication. Nothing is admitted that way until one is
' added: with none, a client asking for SQL authentication is refused as a
' real server refuses a login it does not have. The password is kept only
' as a salted hash. A name matches in any case, a password only exactly.
Public Sub AddLogin(ByVal name As String, ByVal password As String)
    If Len(password) = 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.AddLogin", "a login needs a password"
    End If
    KeepLogin name, HashPassword(password)
End Sub

' The same with the hash HashPassword writes in place of the password, so
' what sits in a workbook or a module is something nobody can log in with.
' pySQLbridge writes the same form, and either reads the other's.
Public Sub AddLoginHash(ByVal name As String, ByVal passwordHash As String)
    Dim iterations As Long
    Dim salt() As Byte
    Dim digest() As Byte

    If Not ReadHash(passwordHash, iterations, salt, digest) Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.AddLoginHash", _
                  "a password hash should look like """ & PASSWORD_ALGORITHM & _
                  "$<iterations>$<salt>$<hash>"", the form HashPassword writes"
    End If
    KeepLogin name, passwordHash
End Sub

Public Sub RemoveLogin(ByVal name As String)
    If mLogins Is Nothing Then Exit Sub
    On Error Resume Next
    mLogins.Remove LCase$(Trim$(name))
    On Error GoTo 0
End Sub

Public Sub ClearLogins()
    Set mLogins = Nothing
End Sub

' The names that may log in, never what they log in with.
Public Property Get LoginNames() As String
    Dim entry As Variant
    Dim names() As String
    Dim index As Long

    If mLogins Is Nothing Then Exit Property
    If mLogins.count = 0 Then Exit Property

    ReDim names(0 To mLogins.count - 1)
    For Each entry In mLogins
        names(index) = entry(0)
        index = index + 1
    Next entry
    LoginNames = Join(names, ", ")
End Property

' A password as it can be kept: PBKDF2 over HMAC-SHA256 with a salt of its
' own, written algorithm$iterations$salt$hash. Two logins with the same
' password get different hashes, so a copy of the list does not say which
' they were.
Public Function HashPassword(ByVal password As String) As String
    Dim salt() As Byte

    If Len(password) = 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.HashPassword", _
                  "a password to hash cannot be empty"
    End If
    salt = RandomBytes(PASSWORD_SALT_BYTES)
    HashPassword = PasswordHashOf(password, salt, PASSWORD_ITERATIONS)
End Function

' Whether a name and a password may log in. A name that is not there is
' checked against a decoy all the same, so the time taken does not say
' whether the name exists and is worth guessing a password for.
Friend Function AdmitsLogin(ByVal name As String, _
                            ByVal password As String) As Boolean
    Dim entry As Variant
    Dim known As Boolean

    If Not mLogins Is Nothing Then
        On Error Resume Next
        entry = mLogins(LCase$(Trim$(name)))
        known = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
    End If
    If known Then
        AdmitsLogin = MatchesHash(CStr(entry(1)), password)
    Else
        MatchesHash DecoyHash(), password
    End If
End Function

Friend Property Get HasLogins() As Boolean
    If mLogins Is Nothing Then Exit Property
    HasLogins = (mLogins.count > 0)
End Property

Private Sub KeepLogin(ByVal name As String, ByVal passwordHash As String)
    name = Trim$(name)
    If Len(name) = 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.AddLogin", "a login needs a name"
    End If
    If mLogins Is Nothing Then Set mLogins = New Collection
    RemoveLogin name
    mLogins.Add Array(name, passwordHash), LCase$(name)
End Sub

' Whether a password gives the digest a stored hash holds, compared in the
' same time whichever byte differs.
Private Function MatchesHash(ByVal stored As String, _
                             ByVal offered As String) As Boolean
    Dim iterations As Long
    Dim salt() As Byte
    Dim expected() As Byte
    Dim computed() As Byte
    Dim plain() As Byte
    Dim index As Long
    Dim difference As Long

    If Not ReadHash(stored, iterations, salt, expected) Then Exit Function
    plain = Utf8Bytes(offered)
    computed = Pbkdf2Sha256(plain, salt, iterations, ByteCount(expected))
    For index = 0 To ByteCount(expected) - 1
        difference = difference Or (computed(index) Xor expected(index))
    Next index
    MatchesHash = (difference = 0)
End Function

' The parts of a stored hash, or False where it is not one this can read.
Private Function ReadHash(ByVal stored As String, ByRef iterations As Long, _
                          ByRef salt() As Byte, ByRef digest() As Byte) As Boolean
    Dim parts() As String

    parts = Split(stored, "$")
    If UBound(parts) <> 3 Then Exit Function
    If parts(0) <> PASSWORD_ALGORITHM Then Exit Function
    If Len(parts(1)) = 0 Or Len(parts(1)) > 9 Then Exit Function
    If parts(1) Like "*[!0-9]*" Then Exit Function
    iterations = CLng(parts(1))
    If iterations < 1 Then Exit Function
    If Not Base64Decode(parts(2), salt) Then Exit Function
    If Not Base64Decode(parts(3), digest) Then Exit Function
    If ByteCount(salt) = 0 Or ByteCount(digest) < 16 Then Exit Function
    ReadHash = True
End Function

Private Function PasswordHashOf(ByVal password As String, ByRef salt() As Byte, _
                                ByVal iterations As Long) As String
    Dim plain() As Byte
    Dim digest() As Byte

    plain = Utf8Bytes(password)
    digest = Pbkdf2Sha256(plain, salt, iterations, PASSWORD_DIGEST_BYTES)
    PasswordHashOf = PASSWORD_ALGORITHM & "$" & iterations & "$" & _
                     Base64Encode(salt) & "$" & Base64Encode(digest)
End Function

' A hash of a password nothing knows, made once and kept, for a name that is
' not there to be checked against.
Private Function DecoyHash() As String
    Dim secret() As Byte
    Dim salt() As Byte

    If Len(mDecoyHash) = 0 Then
        secret = RandomBytes(24)
        salt = RandomBytes(PASSWORD_SALT_BYTES)
        mDecoyHash = PasswordHashOf(Base64Encode(secret), salt, _
                                    PASSWORD_ITERATIONS)
    End If
    DecoyHash = mDecoyHash
End Function

' PBKDF2 over HMAC-SHA256, through CNG.
Private Function Pbkdf2Sha256(ByRef password() As Byte, ByRef salt() As Byte, _
                              ByVal iterations As Long, _
                              ByVal length As Long) As Byte()
    Dim provider As LongPtr
    Dim derived() As Byte
    Dim passwordAt As LongPtr
    Dim status As Long
    Dim algorithm As String

    algorithm = "SHA256"
    status = BCryptOpenAlgorithmProvider(provider, StrPtr(algorithm), 0, _
                                         BCRYPT_ALG_HANDLE_HMAC_FLAG)
    If status <> 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.Pbkdf2Sha256", _
                  "Windows would not open HMAC-SHA256 (status 0x" & _
                  Hex$(status) & ")"
    End If

    ReDim derived(0 To length - 1)
    If ByteCount(password) > 0 Then passwordAt = VarPtr(password(0))
    status = BCryptDeriveKeyPBKDF2(provider, passwordAt, ByteCount(password), _
                                   VarPtr(salt(0)), ByteCount(salt), _
                                   CLngLng(iterations), VarPtr(derived(0)), _
                                   length, 0)
    BCryptCloseAlgorithmProvider provider, 0
    If status <> 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.Pbkdf2Sha256", _
                  "Windows could not derive the key (status 0x" & _
                  Hex$(status) & ")"
    End If
    Pbkdf2Sha256 = derived
End Function

Private Function RandomBytes(ByVal count As Long) As Byte()
    Dim out() As Byte
    Dim status As Long

    ReDim out(0 To count - 1)
    status = BCryptGenRandom(0, VarPtr(out(0)), count, _
                             BCRYPT_USE_SYSTEM_PREFERRED_RNG)
    If status <> 0 Then
        Err.Raise ERR_PROTOCOL, "SqlBridge.RandomBytes", _
                  "Windows would not give random bytes (status 0x" & _
                  Hex$(status) & ")"
    End If
    RandomBytes = out
End Function

' Text as UTF-8, which is what the hash is taken over, so a password with an
' accent in it hashes here as it does in Python.
Private Function Utf8Bytes(ByVal text As String) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim index As Long
    Dim code As Long
    Dim low As Long

    index = 1
    Do While index <= Len(text)
        code = AscW(Mid$(text, index, 1)) And &HFFFF&
        If code >= &HD800& And code <= &HDBFF& And index < Len(text) Then
            low = AscW(Mid$(text, index + 1, 1)) And &HFFFF&
            If low >= &HDC00& And low <= &HDFFF& Then
                code = &H10000 + (code - &HD800&) * &H400& + (low - &HDC00&)
                index = index + 1
            End If
        End If
        If code < &H80& Then
            AppendByte out, used, code
        ElseIf code < &H800& Then
            AppendByte out, used, &HC0& Or (code \ &H40&)
            AppendByte out, used, &H80& Or (code And &H3F&)
        ElseIf code < &H10000 Then
            AppendByte out, used, &HE0& Or (code \ &H1000&)
            AppendByte out, used, &H80& Or ((code \ &H40&) And &H3F&)
            AppendByte out, used, &H80& Or (code And &H3F&)
        Else
            AppendByte out, used, &HF0& Or (code \ &H40000)
            AppendByte out, used, &H80& Or ((code \ &H1000&) And &H3F&)
            AppendByte out, used, &H80& Or ((code \ &H40&) And &H3F&)
            AppendByte out, used, &H80& Or (code And &H3F&)
        End If
        index = index + 1
    Loop
    Utf8Bytes = LeftBytes(out, used)
End Function

Private Function Base64Encode(ByRef data() As Byte) As String
    Dim count As Long
    Dim index As Long
    Dim chunk As Long
    Dim out As String

    count = ByteCount(data)
    For index = 0 To count - 1 Step 3
        chunk = CLng(data(index)) * &H10000
        If index + 1 < count Then chunk = chunk + CLng(data(index + 1)) * &H100&
        If index + 2 < count Then chunk = chunk + CLng(data(index + 2))
        out = out & Base64Char(chunk \ &H40000) & _
              Base64Char((chunk \ &H1000&) And &H3F&)
        If index + 1 < count Then
            out = out & Base64Char((chunk \ &H40&) And &H3F&)
        Else
            out = out & "="
        End If
        If index + 2 < count Then
            out = out & Base64Char(chunk And &H3F&)
        Else
            out = out & "="
        End If
    Next index
    Base64Encode = out
End Function

Private Function Base64Char(ByVal sextet As Long) As String
    Base64Char = Mid$(BASE64_ALPHABET, sextet + 1, 1)
End Function

' Base64 back to bytes, or False where it is not base64.
Private Function Base64Decode(ByVal text As String, ByRef out() As Byte) As Boolean
    Dim buffer() As Byte
    Dim used As Long
    Dim index As Long
    Dim value As Long
    Dim bits As Long
    Dim accumulated As Long
    Dim padding As Long
    Dim ch As String

    If Len(text) = 0 Or Len(text) Mod 4 <> 0 Then Exit Function
    For index = 1 To Len(text)
        ch = Mid$(text, index, 1)
        If ch = "=" Then
            If index < Len(text) - 1 Then Exit Function
            padding = padding + 1
        ElseIf padding > 0 Then
            Exit Function
        Else
            value = InStr(1, BASE64_ALPHABET, ch, vbBinaryCompare) - 1
            If value < 0 Then Exit Function
            accumulated = accumulated * 64 + value
            bits = bits + 6
            If bits >= 8 Then
                bits = bits - 8
                AppendByte buffer, used, accumulated \ CLng(2 ^ bits)
                accumulated = accumulated Mod CLng(2 ^ bits)
            End If
        End If
    Next index
    out = LeftBytes(buffer, used)
    Base64Decode = True
End Function

' An Excel table's header and rows.
'
' A totals row is Excel's arithmetic over the rows rather than one of them,
' so it is left out: read as part of the table, it was served as a record
' called Total. A table with its header row turned off still has column
' names, and they are its header: read as its range, its first record was
' taken for the names and lost as a row.
Private Function ListObjectValues(ByVal listed As Object) As Variant
    Dim body As Variant
    Dim out() As Variant
    Dim height As Long
    Dim width As Long
    Dim row As Long
    Dim column As Long

    If listed.ShowHeaders And Not listed.ShowTotals Then
        ListObjectValues = OneCellSafe(listed.Range)
        Exit Function
    End If

    width = listed.ListColumns.count
    If Not listed.DataBodyRange Is Nothing Then
        body = OneCellSafe(listed.DataBodyRange)
        height = UBound(body, 1)
    End If
    ReDim out(1 To height + 1, 1 To width)
    For column = 1 To width
        out(1, column) = listed.ListColumns(column).name
    Next column
    For row = 1 To height
        For column = 1 To width
            out(row + 1, column) = body(row, column)
        Next column
    Next row
    ListObjectValues = out
End Function

' What a caller handed over, as a two-dimensional array with a header row.
Private Function TableValues(ByVal source As Variant, _
                             ByVal name As String) As Variant
    Dim data As Variant

    Select Case TypeName(source)
        Case "Range"
            If source.Cells.count = 1 Then
                Err.Raise ERR_PROTOCOL, "SqlBridge.AddTable", _
                          "table '" & name & "' is one cell; a table needs a " & _
                          "header row and at least one column"
            End If
            data = source.Value
        Case "Worksheet"
            ' A whole sheet, meaning whatever it holds when it is asked. A
            ' range is a rectangle fixed when it was handed over, so a row
            ' typed under it is outside the table; a sheet grows.
            data = OneCellSafe(source.UsedRange)
        Case "ListObject"
            data = ListObjectValues(source)
        Case Else
            If Not IsArray(source) Then
                Err.Raise ERR_PROTOCOL, "SqlBridge.AddTable", _
                          "table '" & name & "' must come from a range, a " & _
                          "list object, or a two-dimensional array"
            End If
            data = source
    End Select

    On Error GoTo NotTwoDimensional
    If UBound(data, 2) < LBound(data, 2) Then GoTo NotTwoDimensional
    On Error GoTo 0
    TableValues = data
    Exit Function

NotTwoDimensional:
    Err.Raise ERR_PROTOCOL, "SqlBridge.AddTable", _
              "table '" & name & "' must be two-dimensional, with the column " & _
              "names in its first row"
End Function

' A worksheet's rectangle, read as a table.
'
' What comes back from a sheet is whatever was in the cells: a caption over
' the top, a blank row somebody liked the look of, a column emptied but not
' deleted, an error where a formula gave up. This works out which of it is
' the table, over the array that was already handed across, so an untidy
' sheet costs no more trips to Excel than a tidy one.
'
' hasHeader says whether the first row is the column names. Left out, it is
' worked out: a row of nothing but numbers and dates is data, and anything
' with a word in it is a header. That is the only kind of missing header
' there is any telling about, because a first row of 1 and Ada looks exactly
' like the header of a table whose first column is called 1.
' rowMap and columnMap say where each served cell came from: row i of the
' answer is row rowMap(i) of the rectangle, and column j is columnMap(j).
' Captions, blank rows and empty columns are dropped on the way through, so
' a write needs telling rather than counting.
Private Function CleanTable(ByRef raw As Variant, ByVal hasHeader As Variant, _
                            ByRef wide As Long, ByRef rowMap() As Long, _
                            ByRef columnMap() As Long) As Variant
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstColumn As Long
    Dim lastColumn As Long
    Dim row As Long
    Dim column As Long
    Dim keep() As Boolean
    Dim kept() As Long
    Dim headerRow As Long
    Dim firstBody As Long
    Dim taken As Long
    Dim into As Long
    Dim out As Variant
    Dim index As Long
    Dim value As Variant
    Dim filled As Long
    Dim blank As Boolean
    Dim heads As Boolean

    wide = 0
    ReDim out(1 To 1, 1 To 1)
    CleanTable = out
    If Not IsArray(raw) Then Exit Function

    firstRow = LBound(raw, 1)
    lastRow = UBound(raw, 1)
    firstColumn = LBound(raw, 2)
    lastColumn = UBound(raw, 2)

    ' A column with nothing anywhere in it, header included, is not a
    ' column: somebody cleared it and left it there.
    ReDim keep(firstColumn To lastColumn)
    For column = firstColumn To lastColumn
        For row = firstRow To lastRow
            If Not BlankCell(raw(row, column)) Then
                keep(column) = True
                Exit For
            End If
        Next row
        If keep(column) Then wide = wide + 1
    Next column
    If wide = 0 Then Exit Function

    ReDim kept(1 To wide)
    index = 1
    For column = firstColumn To lastColumn
        If keep(column) Then
            kept(index) = column
            index = index + 1
        End If
    Next column
    ReDim columnMap(1 To wide)
    For index = 1 To wide
        columnMap(index) = kept(index)
    Next index

    ' A row with one cell on it, over a table wider than one column, is a
    ' caption rather than a header of one column.
    headerRow = 0
    For row = firstRow To lastRow
        filled = 0
        For index = 1 To wide
            If Not BlankCell(raw(row, kept(index))) Then filled = filled + 1
        Next index
        If filled = 0 Then GoTo NextCandidate
        If filled = 1 And wide > 1 And row < lastRow Then GoTo NextCandidate
        headerRow = row
        Exit For
NextCandidate:
    Next row
    If headerRow = 0 Then
        wide = 0
        Exit Function
    End If

    If IsEmpty(hasHeader) Then
        heads = False
        For index = 1 To wide
            value = raw(headerRow, kept(index))
            If Not BlankCell(value) Then
                If IsError(value) Then
                    heads = True
                ElseIf VarType(value) = vbString Then
                    heads = True
                End If
            End If
        Next index
    Else
        heads = CBool(hasHeader)
    End If

    ' Counted first and copied second, over the array both times. A list
    ' of the rows to keep would be a list asked for its nth item once per
    ' row, which is a walk down the whole list every time; and a call per
    ' cell to ask whether it is blank is a call per cell.
    firstBody = headerRow
    If heads Then firstBody = headerRow + 1

    taken = 0
    For row = firstBody To lastRow
        For index = 1 To wide
            value = raw(row, kept(index))
            If IsError(value) Then
                taken = taken + 1
                Exit For
            ElseIf Not IsEmpty(value) Then
                If VarType(value) <> vbString Then
                    taken = taken + 1
                    Exit For
                ElseIf Len(value) > 0 Then
                    taken = taken + 1
                    Exit For
                End If
            End If
        Next index
    Next row

    ReDim out(1 To taken + 1, 1 To wide)
    ReDim rowMap(1 To taken + 1)
    rowMap(1) = headerRow
    For index = 1 To wide
        If heads Then
            out(1, index) = CleanValue(raw(headerRow, kept(index)))
        Else
            out(1, index) = Empty
        End If
    Next index

    into = 1
    For row = firstBody To lastRow
        blank = True
        For index = 1 To wide
            value = raw(row, kept(index))
            If IsError(value) Then
                blank = False
                Exit For
            ElseIf Not IsEmpty(value) Then
                If VarType(value) <> vbString Then
                    blank = False
                    Exit For
                ElseIf Len(value) > 0 Then
                    blank = False
                    Exit For
                End If
            End If
        Next index

        If Not blank Then
            into = into + 1
            rowMap(into) = row
            For index = 1 To wide
                value = raw(row, kept(index))
                If IsError(value) Then
                    out(into, index) = Null
                ElseIf IsEmpty(value) Then
                    out(into, index) = Null
                Else
                    out(into, index) = value
                End If
            Next index
        End If
    Next row
    CleanTable = out
End Function

' Whether a cell holds nothing.
'
' An empty cell holds nothing and an error holds something, badly. A Null is
' neither: a sheet has no way to write one, so a Null in an array is one
' somebody meant, and a row carrying it is a row.
Private Function BlankCell(ByRef value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsEmpty(value) Then
        BlankCell = True
    ElseIf IsNull(value) Then
        BlankCell = False
    ElseIf VarType(value) = vbString Then
        BlankCell = (Len(value) = 0)
    End If
End Function

' A cell as a value a client can be sent. An Excel error is not one: there
' is no such thing on the wire, and nothing is the honest stand-in.
Private Function CleanValue(ByRef value As Variant) As Variant
    If IsError(value) Then
        CleanValue = Null
    ElseIf IsEmpty(value) Then
        CleanValue = Null
    Else
        CleanValue = value
    End If
End Function

' The type of each column, read from every value in it.
'
' One value that is not a whole number drops the column to float, and one
' that is not a number at all drops it to text. A column is only a date, or
' only a bit, when every value present is one.
' A range's values, as an array even when the range is one cell. Excel
' hands a single cell over as the value itself, and an empty sheet is a
' single cell: the table with nothing in it would otherwise be the one shape
' that could not be read.
Private Function OneCellSafe(ByVal target As Object) As Variant
    Dim out(1 To 1, 1 To 1) As Variant
    Dim value As Variant

    value = target.Value
    If IsArray(value) Then
        OneCellSafe = value
        Exit Function
    End If
    out(1, 1) = value
    OneCellSafe = out
End Function

Private Function InferColumns(ByRef data As Variant) As Collection
    Dim columns As Collection
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstColumn As Long
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim value As Variant
    Dim name As String
    Dim seen As Long
    Dim numbers As Long
    Dim whole As Long
    Dim dates As Long
    Dim bits As Long
    Dim widest As Long
    Dim colon As Long
    Dim declared As Variant

    Set columns = New Collection
    firstRow = LBound(data, 1)
    lastRow = UBound(data, 1)
    firstColumn = LBound(data, 2)
    lastColumn = UBound(data, 2)

    For columnIndex = firstColumn To lastColumn
        name = Trim$(CStr(NullToEmpty(data(firstRow, columnIndex))))
        If Len(name) = 0 Then name = "Column" & (columnIndex - firstColumn + 1)

        ' Two columns with one name is one column as far as a client asking
        ' for it by name is concerned, so the second is numbered.
        name = UnusedName(name, columns)

        ' A header written name:type says what the column is, whatever is
        ' or is not under it.
        declared = Empty
        colon = InStr(name, ":")
        If colon > 0 Then
            declared = TypeFromName(Mid$(name, colon + 1))
            name = Left$(name, colon - 1)
        End If

        seen = 0: numbers = 0: whole = 0: dates = 0: bits = 0: widest = 0
        For rowIndex = firstRow + 1 To lastRow
            value = data(rowIndex, columnIndex)
            If Not (IsEmpty(value) Or IsNull(value)) Then
                If Len(CStr(value)) > 0 Then
                    seen = seen + 1
                    If VarType(value) = vbBoolean Then bits = bits + 1
                    If IsDate(value) And VarType(value) = vbDate Then
                        dates = dates + 1
                    End If
                    ' A number, not a string that reads as one: a code
                    ' typed as text keeps its leading zeros, and calling
                    ' the column an integer would send 007 as 7.
                    If IsNumeric(value) And VarType(value) <> vbDate _
                       And VarType(value) <> vbBoolean _
                       And VarType(value) <> vbString Then
                        numbers = numbers + 1
                        If CDbl(value) = Int(CDbl(value)) _
                           And Abs(CDbl(value)) < 2147483647# Then
                            whole = whole + 1
                        End If
                    End If
                    If Len(CStr(value)) > widest Then widest = Len(CStr(value))
                End If
            End If
        Next rowIndex

        If IsArray(declared) Then
            columns.Add ColumnOf(name, CLng(declared(0)), CLng(declared(1)))
        ElseIf seen = 0 Then
            columns.Add ColumnOf(name, TYPE_NVARCHAR, 255)
        ElseIf bits = seen Then
            columns.Add ColumnOf(name, TYPE_BITN)
        ElseIf dates = seen Then
            columns.Add ColumnOf(name, TYPE_DATETIMN)
        ElseIf numbers = seen And whole = seen Then
            columns.Add ColumnOf(name, TYPE_INTN, 4)
        ElseIf numbers = seen Then
            columns.Add ColumnOf(name, TYPE_FLTN, 8)
        Else
            If widest < 1 Then widest = 1
            If widest > NVARCHAR_SIZED_LIMIT Then widest = 0   ' the MAX form
            columns.Add ColumnOf(name, TYPE_NVARCHAR, widest)
        End If
    Next columnIndex

    Set InferColumns = columns
End Function

' A name nothing else in this table has taken.
Private Function UnusedName(ByVal wanted As String, _
                            ByVal columns As Collection) As String
    Dim tried As String
    Dim suffix As Long
    Dim column As Variant
    Dim taken As Boolean

    tried = wanted
    suffix = 1
    Do
        taken = False
        For Each column In columns
            If StrComp(CStr(column(COLUMN_NAME)), tried, vbTextCompare) = 0 Then
                taken = True
                Exit For
            End If
        Next column
        If Not taken Then
            UnusedName = tried
            Exit Function
        End If
        suffix = suffix + 1
        tried = wanted & suffix
    Loop
End Function

Private Function NullToEmpty(ByVal value As Variant) As Variant
    If IsNull(value) Or IsEmpty(value) Then
        NullToEmpty = vbNullString
    Else
        NullToEmpty = value
    End If
End Function

' The table an identifier names, from the session first and the catalog
' after it, because a temp table is this connection's own.
Friend Function LookUpTable(ByVal name As String) As Variant
    Dim entry As Variant
    Dim bare As String
    Dim parts() As String

    ' A client may qualify a name with its database and schema, and one
    ' database is served whatever it names.
    parts = Split(name, ".")
    bare = Replace(Replace(parts(UBound(parts)), "[", ""), "]", "")

    If Left$(bare, 1) = "#" Then
        On Error GoTo Missing
        entry = mTempTables(LCase$(bare))
        On Error GoTo 0
        LookUpTable = entry
        Exit Function
    End If

    On Error GoTo Missing
    entry = mTables(LCase$(bare))
    On Error GoTo 0
    LookUpTable = entry
    Exit Function

Missing:
    RaiseQueryError SQL_ERR_INVALID_OBJECT, "Invalid object name '" & name & "'."
End Function

' New rows under the ones already there, in one handover.
'
' A row at a time is a round trip each, and an Excel table asked to add a
' row does more work still: the whole block goes over at once, and a table
' is told afterwards how much bigger it is.
Private Sub AppendRowsToSheet(ByVal entry As Variant, ByVal name As String, _
                              ByVal rows As Collection)
    Dim target As Object
    Dim band As Object
    Dim rowMap() As Long
    Dim columnMap() As Long
    Dim block As Variant
    Dim rowIndex As Long
    Dim index As Long
    Dim values As Variant
    Dim value As Variant
    Dim wide As Long
    Dim under As Long
    Dim first As Long
    Dim last As Long
    Dim listed As Object
    Dim wanted As Long

    Set target = WritableRange(entry, name, True)
    If SourceKind(entry) = "ListObject" Then
        Set listed = entry(3)
        wanted = listed.Range.Rows.count + rows.count
    End If
    rowMap = entry(6)
    columnMap = entry(7)
    wide = UBound(columnMap)

    ' Under the last row the table occupies, not under the rectangle: a
    ' sheet's rectangle is everything used on it, and a table ending at row
    ' six with a note down at row twelve grows at row seven.
    first = columnMap(1)
    last = columnMap(wide)
    under = rowMap(UBound(rowMap)) + 1

    ' The columns need not sit next to each other, so the block covers from
    ' the first to the last and puts nothing in the gaps.
    ReDim block(1 To rows.count, 1 To last - first + 1)
    For rowIndex = 1 To rows.count
        values = rows(rowIndex)
        For index = 1 To wide
            If index - 1 <= UBound(values) Then
                value = values(LBound(values) + index - 1)
                ' A cell has no Null. A column the INSERT did not name is
                ' left empty, which is what the table reads back as Null.
                If Not IsNull(value) Then
                    block(rowIndex, columnMap(index) - first + 1) = value
                End If
            End If
        Next index
    Next rowIndex

    Set band = target.Cells(under, first).Resize(rows.count, last - first + 1)

    ' Room has to be made if anything is in those cells already. Whole rows,
    ' so what sits beside the table stays beside the same rows of it, and
    ' the band is taken again afterwards: a range follows the cells it
    ' named, and those cells are the ones that just moved down.
    If Occupied(band) Then
        band.EntireRow.Insert
        Set band = target.Cells(under, first) _
                         .Resize(rows.count, last - first + 1)
    End If
    band.Value = block

    ' A table takes in the row under it by itself, but only sometimes: it
    ' depends on a setting, and on how much was written at once. Asked for a
    ' size it already has it does nothing, and told to grow again it would
    ' swallow whatever came next.
    If Not listed Is Nothing Then
        If listed.Range.Rows.count <> wanted Then
            listed.Resize listed.Range.Resize(wanted)
        End If
    End If

    RefreshTable CStr(entry(0))
End Sub

' Whether anything is in these cells, read rather than calculated.
'
' `Application.WorksheetFunction.CountA` answers the same question by way of
' Excel's calculation engine, and over an `EntireRow` it walks sixteen
' thousand cells to do it. This reads the cells the write is about to cover
' and tests them here.
Private Function Occupied(ByVal band As Object) As Boolean
    Dim values As Variant
    Dim row As Long
    Dim column As Long

    values = OneCellSafe(band)
    For row = LBound(values, 1) To UBound(values, 1)
        For column = LBound(values, 2) To UBound(values, 2)
            If Not BlankCell(values(row, column)) Then
                Occupied = True
                Exit Function
            End If
        Next column
    Next row
End Function

' Where a served table's rows live, and whether they can be written to.
'
' A range is the rectangle it names, so a row added to it or taken out of it
' would be a rectangle of a different size: that is a table the caller said
' was those cells and no others. A sheet and an Excel table both say how
' many rows there are by how many there are, and both can take more.
Private Function WritableRange(ByVal entry As Variant, _
                               ByVal name As String, _
                               ByVal changesRowCount As Boolean) As Object
    RefuseUnwritable entry, name, changesRowCount

    Set WritableRange = SourceRange(entry)
    If WritableRange Is Nothing Then
        RaiseQueryError ERR_NUMBER_BASE, _
            "Table '" & name & "' has nowhere to write back to."
    End If
End Function

' Say no before anything is worked out, where the answer is already no.
'
' A table with no source object is its own rows: a temporary table, or one
' served from an array, which is the copy it was given. Both take all three
' statements, and neither has a workbook to reach.
Private Sub RefuseUnwritable(ByVal entry As Variant, ByVal name As String, _
                             ByVal changesRowCount As Boolean)
    If IsViewEntry(entry) Then
        RaiseQueryError ERR_NUMBER_BASE, _
            "'" & name & "' is a view here, and a view is read-only: " & _
            "change the sheet or the table its statement reads."
    End If

    If mReadOnly Then
        RaiseQueryError 3906, _
            "Failed to update database '" & ConfiguredDatabase & "' " & _
            "because the database is read-only."
    End If

    ' An Excel table with its header row turned off is read with its column
    ' names put on top of its rows, so the rows read are one away from the
    ' rows on the sheet, and a write would land on the wrong one.
    If SourceKind(entry) = "ListObject" Then
        If Not entry(3).ShowHeaders Then
            RaiseQueryError ERR_NUMBER_BASE, _
                "Table '" & name & "' is an Excel table with its header " & _
                "row turned off. Turn the header row on to write to it."
        End If
    End If

    If SourceKind(entry) = "Range" And changesRowCount Then
        RaiseQueryError ERR_NUMBER_BASE, _
            "Table '" & name & "' was served as a fixed range, which " & _
            "cannot gain or lose rows. Serve the sheet or an Excel table " & _
            "to do that."
    End If
End Sub

' The rectangle on a sheet a served table's rows live in, header included.
Private Function SourceRange(ByVal entry As Variant) As Object
    Select Case SourceKind(entry)
        Case "Range"
            Set SourceRange = entry(3)
        Case "Worksheet"
            Set SourceRange = entry(3).UsedRange
        Case "ListObject"
            Set SourceRange = entry(3).Range
    End Select
End Function

' What a table was served from, which decides what can be done to it.
Private Function SourceKind(ByVal entry As Variant) As String
    If Not IsArray(entry) Then Exit Function
    If UBound(entry) < 3 Then Exit Function
    If Not IsObject(entry(3)) Then Exit Function
    SourceKind = TypeName(entry(3))
End Function

' Read one table's source again, so the statement after a write sees it.
' Called with the name the table was served under, not the one the statement
' spelled: a client may write dbo.[staff], and there is one entry for it.
' The whole catalog is read once before a batch; this is for the SELECT
' after an INSERT in the same one.
Private Sub RefreshTable(ByVal name As String)
    Dim key As String
    Dim entry As Variant
    Dim raw As Variant
    Dim wide As Long
    Dim rowMap() As Long
    Dim columnMap() As Long

    key = LCase$(Replace(Replace(name, "[", ""), "]", ""))
    On Error GoTo Missing
    entry = mTables(key)
    On Error GoTo 0
    If UBound(entry) < 3 Then Exit Sub
    If Not IsObject(entry(3)) Then Exit Sub

    raw = TableValues(entry(3), CStr(entry(0)))
    entry(5) = raw
    entry(1) = CleanTable(raw, entry(4), wide, rowMap, columnMap)
    entry(6) = rowMap
    entry(7) = columnMap
    If wide = 0 Then
        Set entry(2) = New Collection
    Else
        Set entry(2) = InferColumns(entry(1))
    End If
    mTables.Remove key
    mTables.Add entry, key
    Exit Sub

Missing:
End Sub

' UPDATE name SET column = expression [, ...] [WHERE condition]
'
' The rows are worked out first and written afterwards, a block at a time.
' A cell handed to Excel on its own is a round trip, and a sheet of twenty
' thousand rows is twenty thousand of them; a run of rows that all changed
' goes over as one array.
Private Sub RunUpdate(ByVal written As String, ByVal answers As Collection)
    Dim name As String
    Dim targets As Collection
    Dim sets As Collection
    Dim whereNode As Variant
    Dim entry As Variant
    Dim columns As Collection
    Dim data As Variant
    Dim rowMap() As Long
    Dim columnMap() As Long
    Dim target As Object
    Dim row As Long
    Dim index As Long
    Dim values As Variant
    Dim keep As Variant
    Dim touched() As Boolean
    Dim slots() As Long
    Dim changes As Variant
    Dim affected As Long
    Dim slot As Long
    Dim joined As Boolean
    Dim realName As String
    Dim targetAt As Long
    Dim item As Variant

    Set mTokens = Tokenize(written)
    mCursor = 1
    Expect "UPDATE"
    name = TakeText()
    SkipTableHints
    Expect "SET"

    Set targets = New Collection
    Set sets = New Collection
    Do
        targets.Add TakeText()
        Expect "="
        sets.Add ParseExpression()
    Loop While Accept(",")

    ' UPDATE ... FROM, which takes the rows to set, and what to set them
    ' to, from a join.
    joined = Accept("FROM")
    If joined Then ReadFromClause

    whereNode = Empty
    If Accept("WHERE") Then whereNode = ParseExpression()
    RefuseTheRest "UPDATE"

    realName = name
    If joined Then
        targetAt = TargetInFrom(name)
        If targetAt > 0 Then
            item = mFromItems(targetAt)
            realName = CStr(item(1))
        End If
    End If

    entry = LookUpTable(realName)
    RefuseUnwritable entry, realName, False
    RememberBeforeWriting realName
    Set target = SourceRange(entry)
    data = entry(1)
    If Not target Is Nothing Then
        rowMap = entry(6)
        columnMap = entry(7)
    End If
    Set columns = QualifyColumns(entry(2), BareColumnName(realName), Nothing)

    ' Which column each assignment lands in, before any row is looked at.
    ' SET p.team names the column through the alias a FROM gave the table.
    ReDim slots(1 To targets.count)
    For index = 1 To targets.count
        slot = ColumnIndex(CStr(targets(index)), columns)
        If slot = 0 Then
            slot = ColumnIndex(BareColumnName(CStr(targets(index))), columns)
        End If
        If slot = 0 Then
            RaiseQueryError 207, _
                "Invalid column name '" & targets(index) & "'."
        End If
        slots(index) = slot
    Next index

    ' A table of nothing but a header has no row to touch, and an array
    ' declared from 2 to 1 is not an empty array, it is an error.
    If UBound(data, 1) < 2 Then
        answers.Add AffectedAnswer(0, UPDATE_COMMAND)
        Exit Sub
    End If

    ReDim touched(2 To UBound(data, 1))
    ReDim changes(2 To UBound(data, 1), 1 To targets.count)

    If joined Then
        JoinedTargets name, targetAt, entry, whereNode, sets, touched, _
                      changes, affected
    Else
        For row = 2 To UBound(data, 1)
            values = RowValues(data, row)
            If IsEmpty(whereNode) Then
                keep = True
            Else
                keep = Evaluate(whereNode, values, columns)
            End If
            If Not IsNull(keep) Then
                If CBool(keep) Then
                    touched(row) = True
                    affected = affected + 1
                    For index = 1 To targets.count
                        changes(row, index) = Evaluate(sets(index), values, _
                                                       columns)
                    Next index
                End If
            End If
        Next row
    End If

    If affected > 0 Then
        If target Is Nothing Then
            ' Its rows are all there is of it, so they are what changes.
            ' The columns stay as they were declared: a temporary table's
            ' came from its CREATE TABLE, and a value that does not fit one
            ' is a conversion error rather than a new type.
            For row = 2 To UBound(data, 1)
                If touched(row) Then
                    For index = 1 To targets.count
                        data(row, LBound(data, 2) + slots(index) - 1) = _
                            AsDeclared(changes(row, index), _
                                       columns(slots(index)))
                    Next index
                End If
            Next row
            entry(1) = data
            StoreTable CStr(entry(0)), entry
        Else
            For index = 1 To targets.count
                WriteColumnRuns target, rowMap, columnMap(slots(index)), _
                                touched, changes, index
            Next index
            RefreshTable CStr(entry(0))
        End If
    End If

    answers.Add AffectedAnswer(affected, UPDATE_COMMAND)
End Sub

' One column's changed cells, in as few handovers as Excel will take: a run
' of rows that all changed goes over as one array.
Private Sub WriteColumnRuns(ByVal target As Object, ByRef rowMap() As Long, _
                            ByVal column As Long, ByRef touched() As Boolean, _
                            ByRef changes As Variant, ByVal which As Long)
    Dim row As Long
    Dim runStart As Long
    Dim runEnd As Long
    Dim block As Variant
    Dim index As Long

    row = LBound(touched)
    Do While row <= UBound(touched)
        If Not touched(row) Then
            row = row + 1
        Else
            runStart = row
            runEnd = row
            Do While runEnd < UBound(touched)
                If Not touched(runEnd + 1) Then Exit Do
                ' Only rows that sit under each other on the sheet make a
                ' run; a blank row between them is not part of it.
                If rowMap(runEnd + 1) <> rowMap(runEnd) + 1 Then Exit Do
                runEnd = runEnd + 1
            Loop

            ReDim block(1 To runEnd - runStart + 1, 1 To 1)
            For index = runStart To runEnd
                ' A cell has no Null: SET x = NULL empties it, and an empty
                ' cell is what the table reads back as Null.
                If IsNull(changes(index, which)) Then
                    block(index - runStart + 1, 1) = Empty
                Else
                    block(index - runStart + 1, 1) = changes(index, which)
                End If
            Next index
            target.Cells(rowMap(runStart), column) _
                  .Resize(runEnd - runStart + 1, 1).Value = block

            row = runEnd + 1
        End If
    Loop
End Sub

' One row of a table's array, as the values an expression is evaluated over.
Private Function RowValues(ByRef data As Variant, ByVal row As Long) As Variant
    Dim out As Variant
    Dim column As Long
    Dim first As Long
    Dim last As Long

    first = LBound(data, 2)
    last = UBound(data, 2)
    ReDim out(0 To last - first)
    For column = first To last
        out(column - first) = data(row, column)
    Next column
    RowValues = out
End Function

' DELETE [FROM] name [WHERE condition]
' Whatever an UPDATE or a DELETE says past the parts read here, such as an
' OUTPUT clause. Left unread, it left the WHERE unread too, and DELETE FROM
' t OUTPUT deleted.id WHERE id = 1 emptied the table. Refused before
' anything is written.
Private Sub RefuseTheRest(ByVal command As String)
    If AtEnd() Then Exit Sub
    RaiseQueryError ERR_NUMBER_BASE, _
        command & " with '" & PeekText() & "' where it stands is not " & _
        "supported here, and nothing was changed."
End Sub

' Which item of an UPDATE's or a DELETE's FROM is the table being written:
' the one the name is the alias of, or else the one that names the table.
' Nought when the FROM does not name it at all.
Private Function TargetInFrom(ByVal name As String) As Long
    Dim item As Variant
    Dim index As Long
    Dim bare As String

    bare = Replace(Replace(BareColumnName(name), "[", ""), "]", "")
    For Each item In mFromItems
        index = index + 1
        If StrComp(CStr(item(2)), bare, vbTextCompare) = 0 Then
            TargetInFrom = index
            Exit Function
        End If
    Next item

    index = 0
    For Each item In mFromItems
        index = index + 1
        If StrComp(Replace(Replace(BareColumnName(CStr(item(1))), "[", ""), _
                           "]", ""), bare, vbTextCompare) = 0 Then
            TargetInFrom = index
            Exit Function
        End If
    Next item
End Function

' The rows of a write's table that its FROM joins to something the WHERE
' keeps, and for an UPDATE what each SET makes of the first joined row that
' kept it.
'
' The join is run the way a read runs it, over a copy of the table whose
' rows each carry their place in it, so every row the join keeps says which
' row of the table it came from. A row the join keeps twice is changed once.
Private Sub JoinedTargets(ByVal name As String, ByVal targetAt As Long, _
                          ByVal entry As Variant, ByVal whereNode As Variant, _
                          ByVal sets As Collection, ByRef touched() As Boolean, _
                          ByRef changes As Variant, ByRef affected As Long)
    Dim data As Variant
    Dim own As Collection
    Dim rows As Collection
    Dim values() As Variant
    Dim item As Variant
    Dim width As Long
    Dim row As Long
    Dim position As Long
    Dim combinedColumns As Collection
    Dim combinedRows As Collection
    Dim combined As Variant
    Dim keep As Variant
    Dim at As Long
    Dim place As Variant
    Dim index As Long
    Dim hadTables As Boolean
    Dim number As Long
    Dim source As String
    Dim description As String

    data = entry(1)
    width = UBound(data, 2) - LBound(data, 2) + 1
    Set own = New Collection
    For Each item In entry(2)
        own.Add item
    Next item
    own.Add ColumnOf(PLACE_COLUMN, TYPE_INTN, 4)
    Set rows = New Collection
    For row = 2 To UBound(data, 1)
        ReDim values(0 To width)
        For position = 0 To width - 1
            values(position) = data(row, LBound(data, 2) + position)
        Next position
        values(width) = row
        rows.Add values
    Next row

    hadTables = Not mCommonTables Is Nothing
    If Not hadTables Then Set mCommonTables = New Collection
    StoreCommonTable PLACE_TABLE, own, rows
    On Error GoTo Failed
    PointFromAt name, targetAt

    CombinedSource combinedColumns, combinedRows
    at = ColumnIndex(PLACE_COLUMN, combinedColumns)
    For Each combined In combinedRows
        If IsEmpty(whereNode) Then
            keep = True
        Else
            keep = Evaluate(whereNode, combined, combinedColumns)
        End If
        If Not IsNull(keep) Then
            If CBool(keep) Then
                place = combined(at - 1)
                If Not IsNull(place) And Not IsEmpty(place) Then
                    row = CLng(place)
                    If Not touched(row) Then
                        touched(row) = True
                        affected = affected + 1
                        If Not sets Is Nothing Then
                            For index = 1 To sets.count
                                changes(row, index) = Evaluate(sets(index), _
                                    combined, combinedColumns)
                            Next index
                        End If
                    End If
                End If
            End If
        End If
    Next combined

    ForgetPlaceTable hadTables
    Exit Sub

Failed:
    ' Kept before the tidying, which clears it: an On Error statement
    ' anywhere wipes Err.
    number = Err.Number
    source = Err.Source
    description = Err.Description
    ForgetPlaceTable hadTables
    Err.Raise number, source, description
End Sub

' The FROM pointed at the copy of the table that carries each row's place:
' the item that is the table, or, where the FROM does not name it, the copy
' joined on to everything else, which is how a real server reads a FROM
' that leaves the table out.
Private Sub PointFromAt(ByVal name As String, ByVal targetAt As Long)
    Dim rebuilt As Collection
    Dim item As Variant
    Dim index As Long

    Set rebuilt = New Collection
    For Each item In mFromItems
        index = index + 1
        If index = targetAt Then item(1) = PLACE_TABLE
        rebuilt.Add item
    Next item
    If targetAt = 0 Then
        rebuilt.Add FromItem(FROM_CROSS, PLACE_TABLE, BareColumnName(name), _
                             Empty, New Collection, New Collection, Nothing)
    End If
    Set mFromItems = rebuilt
End Sub

Private Sub ForgetPlaceTable(ByVal hadTables As Boolean)
    If hadTables Then
        On Error Resume Next
        mCommonTables.Remove LCase$(PLACE_TABLE)
        mCommonTables.Remove LCase$(SOURCE_TABLE)
        On Error GoTo 0
    Else
        Set mCommonTables = Nothing
    End If
End Sub

' TRUNCATE TABLE name: every row gone, and no count, which is what a real
' server answers it with.
Private Sub RunTruncate(ByVal written As String)
    Dim name As String

    Set mTokens = Tokenize(written)
    mCursor = 1
    Expect "TRUNCATE"
    Expect "TABLE"
    name = TakeText()
    RefuseTheRest "TRUNCATE"
    RunDelete "DELETE FROM " & name, New Collection
End Sub

' MERGE [INTO] target [[AS] alias] USING source ON condition, then its WHEN
' clauses, each an UPDATE, a DELETE or an INSERT.
'
' The target and the source are joined in full, over copies that say for
' each row which side it came from, so every joined row is matched, a source
' row the target has not got, or a target row the source has not got. Each
' takes the first WHEN clause of its kind whose condition holds. The updates
' land where the rows are, the deletes go next and the inserts last, under
' what is left. A target row two source rows match is refused before any of
' it is written, the way a real server refuses it.
Private Sub RunMerge(ByVal written As String, ByVal answers As Collection)
    Dim name As String
    Dim alias As String
    Dim source As Variant
    Dim onNode As Variant
    Dim clauses As Collection
    Dim clause As Variant
    Dim clauseSlots As Collection
    Dim slotsHere() As Long
    Dim entry As Variant
    Dim data As Variant
    Dim columns As Collection
    Dim target As Object
    Dim rowMap() As Long
    Dim columnMap() As Long
    Dim unionSlots As Collection
    Dim unionOf() As Long
    Dim slot As Long
    Dim growsOrShrinks As Boolean
    Dim hadTables As Boolean
    Dim combinedColumns As Collection
    Dim combinedRows As Collection
    Dim combined As Variant
    Dim placeAt As Long
    Dim sourcedAt As Long
    Dim place As Variant
    Dim hasTarget As Boolean
    Dim hasSource As Boolean
    Dim kind As String
    Dim keep As Variant
    Dim row As Long
    Dim index As Long
    Dim position As Long
    Dim height As Long
    Dim touched() As Boolean
    Dim going() As Boolean
    Dim acted() As Boolean
    Dim changes As Variant
    Dim inserted As Collection
    Dim updated As Long
    Dim deleted As Long
    Dim number As Long
    Dim failedIn As String
    Dim description As String

    ParseMerge written, name, alias, source, onNode, clauses

    entry = LookUpTable(name)
    For Each clause In clauses
        If clause(2) <> "UPDATE" Then growsOrShrinks = True
    Next clause
    RefuseUnwritable entry, name, growsOrShrinks
    Set columns = QualifyColumns(entry(2), BareColumnName(name), Nothing)
    data = entry(1)
    height = UBound(data, 1)

    ' Which column each SET lands in, and every column any of them does.
    Set clauseSlots = New Collection
    Set unionSlots = New Collection
    ReDim unionOf(1 To columns.count)
    For Each clause In clauses
        If clause(2) = "UPDATE" Then
            ReDim slotsHere(1 To clause(3).count)
            For index = 1 To clause(3).count
                slot = ColumnIndex(BareColumnName(CStr(clause(3)(index))), _
                                   columns)
                If slot = 0 Then
                    RaiseQueryError 207, "Invalid column name '" & _
                                         clause(3)(index) & "'."
                End If
                If unionOf(slot) = 0 Then
                    unionSlots.Add slot
                    unionOf(slot) = unionSlots.count
                End If
                slotsHere(index) = slot
            Next index
            clauseSlots.Add slotsHere
        Else
            clauseSlots.Add Empty
        End If
    Next clause

    If height >= 2 Then
        ReDim touched(2 To height)
        ReDim going(2 To height)
        ReDim acted(2 To height)
        ReDim changes(2 To height, 1 To unionSlots.count + 1)
    End If
    Set inserted = New Collection

    hadTables = Not mCommonTables Is Nothing
    If Not hadTables Then Set mCommonTables = New Collection
    On Error GoTo Failed
    KeepMergeCopies entry, source

    Set mFromItems = New Collection
    mFromItems.Add FromItem(FROM_BASE, PLACE_TABLE, alias, Empty, _
                            New Collection, New Collection, Nothing)
    mFromItems.Add FromItem(FROM_FULL, SOURCE_TABLE, CStr(source(2)), onNode, _
                            New Collection, New Collection, Nothing)
    CombinedSource combinedColumns, combinedRows
    placeAt = ColumnIndex(PLACE_COLUMN, combinedColumns)
    sourcedAt = ColumnIndex(SOURCED_COLUMN, combinedColumns)

    For Each combined In combinedRows
        place = combined(placeAt - 1)
        hasTarget = Not (IsNull(place) Or IsEmpty(place))
        hasSource = Not (IsNull(combined(sourcedAt - 1)) Or _
                         IsEmpty(combined(sourcedAt - 1)))
        If hasTarget And hasSource Then
            kind = "MATCHED"
        ElseIf hasSource Then
            kind = "TARGET"
        Else
            kind = "SOURCE"
        End If
        If hasTarget Then row = CLng(place)

        For index = 1 To clauses.count
            clause = clauses(index)
            If clause(0) = kind Then
                keep = True
                If Not IsEmpty(clause(1)) Then
                    keep = Evaluate(clause(1), combined, combinedColumns)
                End If
                If Not IsNull(keep) Then
                    If CBool(keep) Then
                        If clause(2) = "INSERT" Then
                            inserted.Add MergeInsertRow(clause, entry(2), _
                                combined, combinedColumns)
                        Else
                            If acted(row) Then
                                RaiseQueryError 8672, "The MERGE statement " & _
                                    "attempted to UPDATE or DELETE the same " & _
                                    "row more than once. A target row " & _
                                    "matched more than one source row."
                            End If
                            acted(row) = True
                            If clause(2) = "DELETE" Then
                                going(row) = True
                                deleted = deleted + 1
                            Else
                                touched(row) = True
                                updated = updated + 1
                                ' A column this clause does not set keeps
                                ' what it had, because every column any
                                ' clause sets is written back for the row.
                                For position = 1 To unionSlots.count
                                    changes(row, position) = data(row, _
                                        LBound(data, 2) + unionSlots(position) - 1)
                                Next position
                                slotsHere = clauseSlots(index)
                                For position = 1 To clause(4).count
                                    changes(row, unionOf(slotsHere(position))) = _
                                        Evaluate(clause(4)(position), combined, _
                                                 combinedColumns)
                                Next position
                            End If
                        End If
                        Exit For
                    End If
                End If
            End If
        Next index
    Next combined

    ForgetPlaceTable hadTables
    On Error GoTo 0

    RememberBeforeWriting name
    Set target = SourceRange(entry)
    If Not target Is Nothing Then
        rowMap = entry(6)
        columnMap = entry(7)
    End If

    If updated > 0 Then
        If target Is Nothing Then
            For row = 2 To height
                If touched(row) Then
                    For position = 1 To unionSlots.count
                        data(row, LBound(data, 2) + unionSlots(position) - 1) = _
                            changes(row, position)
                    Next position
                End If
            Next row
        Else
            For position = 1 To unionSlots.count
                WriteColumnRuns target, rowMap, columnMap(unionSlots(position)), _
                                touched, changes, position
            Next position
        End If
    End If
    If deleted > 0 Then
        If target Is Nothing Then
            data = WithoutRows(data, going)
        Else
            RemoveRows target, rowMap, going, _
                       (SourceKind(entry) = "ListObject")
        End If
    End If
    If updated > 0 Or deleted > 0 Then
        If target Is Nothing Then
            entry(1) = data
            StoreTable CStr(entry(0)), entry
        Else
            RefreshTable CStr(entry(0))
        End If
    End If
    If inserted.count > 0 Then AppendRowsTo name, inserted

    answers.Add AffectedAnswer(updated + deleted + inserted.count, _
                               MERGE_COMMAND)
    Exit Sub

Failed:
    number = Err.Number
    failedIn = Err.Source
    description = Err.Description
    ForgetPlaceTable hadTables
    Err.Raise number, failedIn, description
End Sub

' MERGE's parts: the target and what it is called in the statement, the
' source as a FROM item, the condition that matches them, and the WHEN
' clauses in the order they were written.
Private Sub ParseMerge(ByVal written As String, ByRef name As String, _
                       ByRef alias As String, ByRef source As Variant, _
                       ByRef onNode As Variant, ByRef clauses As Collection)
    Set mTokens = Tokenize(written)
    mCursor = 1
    Expect "MERGE"
    If PeekIs("TOP") Then
        RaiseQueryError ERR_NUMBER_BASE, "MERGE with TOP is not supported " & _
                                         "here, and nothing was changed."
    End If
    Accept "INTO"
    name = TakeText()
    ' MERGE INTO t WITH (HOLDLOCK) AS target, the way an upsert is written.
    SkipTableHints
    alias = BareColumnName(name)
    If Accept("AS") Then
        alias = TakeText()
    ElseIf PeekKind() = T_NAME And Not PeekIs("USING") Then
        alias = TakeText()
    End If

    Expect "USING"
    source = ReadFromItem(FROM_FULL)
    Expect "ON"
    onNode = ParseExpression()

    Set clauses = New Collection
    Do While Accept("WHEN")
        clauses.Add ReadMergeClause()
    Loop
    If clauses.count = 0 Then
        RaiseQueryError ERR_NUMBER_BASE, "Incorrect syntax near the end " & _
                                         "of the MERGE; expected WHEN."
    End If
    RefuseTheRest "MERGE"
End Sub

' One WHEN clause: which rows it is for, the condition it adds, what it does
' to them, and what that needs: the SET list for an UPDATE, the columns and
' values for an INSERT.
Private Function ReadMergeClause() As Variant
    Dim out(0 To 5) As Variant
    Dim targets As Collection
    Dim nodes As Collection
    Dim names As Collection

    If Accept("MATCHED") Then
        out(0) = "MATCHED"
    Else
        Expect "NOT"
        Expect "MATCHED"
        out(0) = "TARGET"
        If Accept("BY") Then
            If Accept("SOURCE") Then
                out(0) = "SOURCE"
            Else
                Expect "TARGET"
            End If
        End If
    End If
    out(1) = Empty
    If Accept("AND") Then out(1) = ParseExpression()
    Expect "THEN"

    Set targets = New Collection
    Set nodes = New Collection
    Set names = New Collection
    If Accept("UPDATE") Then
        out(2) = "UPDATE"
        Expect "SET"
        Do
            targets.Add TakeText()
            Expect "="
            nodes.Add ParseExpression()
        Loop While Accept(",")
    ElseIf Accept("DELETE") Then
        out(2) = "DELETE"
    Else
        Expect "INSERT"
        out(2) = "INSERT"
        If Accept("(") Then
            Do
                names.Add TakeText()
            Loop While Accept(",")
            Expect ")"
        End If
        If Accept("DEFAULT") Then
            Expect "VALUES"
        Else
            Expect "VALUES"
            Expect "("
            Do
                nodes.Add ParseExpression()
            Loop While Accept(",")
            Expect ")"
        End If
    End If

    ' An INSERT is for a source row the target has not got, and only an
    ' INSERT is.
    If (out(0) = "TARGET") <> (out(2) = "INSERT") Then
        RaiseQueryError ERR_NUMBER_BASE, "An action of type '" & out(2) & _
            "' is not allowed in that WHEN clause of a MERGE statement."
    End If

    Set out(3) = targets
    Set out(4) = nodes
    Set out(5) = names
    ReadMergeClause = out
End Function

' The copies a MERGE joins: the target with each row's place in it, and the
' source with a mark on every row, so a joined row says which sides it has.
Private Sub KeepMergeCopies(ByVal entry As Variant, ByVal source As Variant)
    Dim data As Variant
    Dim own As Collection
    Dim rows As Collection
    Dim values() As Variant
    Dim item As Variant
    Dim width As Long
    Dim row As Long
    Dim position As Long
    Dim sourceColumns As Collection
    Dim sourceRows As Collection
    Dim marked As Collection
    Dim sourced As Collection
    Dim sourceValues As Variant
    Dim grown() As Variant
    Dim sourceWidth As Long

    data = entry(1)
    width = UBound(data, 2) - LBound(data, 2) + 1
    Set own = New Collection
    For Each item In entry(2)
        own.Add item
    Next item
    own.Add ColumnOf(PLACE_COLUMN, TYPE_INTN, 4)
    Set rows = New Collection
    For row = 2 To UBound(data, 1)
        ReDim values(0 To width)
        For position = 0 To width - 1
            values(position) = data(row, LBound(data, 2) + position)
        Next position
        values(width) = row
        rows.Add values
    Next row
    StoreCommonTable PLACE_TABLE, own, rows

    LoadFromItem source, sourceColumns, sourceRows
    Set marked = New Collection
    For Each item In sourceColumns
        marked.Add item
    Next item
    marked.Add ColumnOf(SOURCED_COLUMN, TYPE_INTN, 4)
    sourceWidth = sourceColumns.count
    Set sourced = New Collection
    For Each sourceValues In sourceRows
        ReDim grown(0 To sourceWidth)
        For position = 0 To sourceWidth - 1
            grown(position) = sourceValues(LBound(sourceValues) + position)
        Next position
        grown(sourceWidth) = 1
        sourced.Add grown
    Next sourceValues
    StoreCommonTable SOURCE_TABLE, marked, sourced
End Sub

' The row a WHEN NOT MATCHED inserts, in the target's own column order: the
' columns it names get their values and the rest are NULL. With no column
' list, the values go in by position.
Private Function MergeInsertRow(ByVal clause As Variant, _
                                ByVal ownColumns As Collection, _
                                ByRef combined As Variant, _
                                ByVal combinedColumns As Collection) As Variant
    Dim out() As Variant
    Dim names As Collection
    Dim nodes As Collection
    Dim index As Long
    Dim slot As Long

    ReDim out(0 To ownColumns.count - 1)
    For index = 0 To ownColumns.count - 1
        out(index) = Null
    Next index
    Set names = clause(5)
    Set nodes = clause(4)

    If names.count = 0 Then
        For index = 1 To nodes.count
            If index <= ownColumns.count Then
                out(index - 1) = Evaluate(nodes(index), combined, combinedColumns)
            End If
        Next index
    Else
        For index = 1 To names.count
            slot = ColumnIndex(Replace(Replace(BareColumnName( _
                       CStr(names(index))), "[", ""), "]", ""), ownColumns)
            If slot = 0 Then
                RaiseQueryError 207, "Invalid column name '" & _
                                     names(index) & "'."
            End If
            If index <= nodes.count Then
                out(slot - 1) = Evaluate(nodes(index), combined, combinedColumns)
            End If
        Next index
    End If
    MergeInsertRow = out
End Function

Private Sub RunDelete(ByVal written As String, ByVal answers As Collection)
    Dim name As String
    Dim whereNode As Variant
    Dim entry As Variant
    Dim columns As Collection
    Dim data As Variant
    Dim rowMap() As Long
    Dim target As Object
    Dim going() As Boolean
    Dim row As Long
    Dim values As Variant
    Dim keep As Variant
    Dim affected As Long
    Dim joined As Boolean
    Dim realName As String
    Dim targetAt As Long
    Dim item As Variant

    Set mTokens = Tokenize(written)
    mCursor = 1
    Expect "DELETE"
    Accept "FROM"
    name = TakeText()
    SkipTableHints

    ' DELETE t FROM t JOIN ..., which takes the rows to delete from a join.
    joined = Accept("FROM")
    If joined Then ReadFromClause

    whereNode = Empty
    If Accept("WHERE") Then whereNode = ParseExpression()
    RefuseTheRest "DELETE"

    realName = name
    If joined Then
        targetAt = TargetInFrom(name)
        If targetAt > 0 Then
            item = mFromItems(targetAt)
            realName = CStr(item(1))
        End If
    End If

    entry = LookUpTable(realName)
    RefuseUnwritable entry, realName, True
    RememberBeforeWriting realName
    Set target = SourceRange(entry)
    data = entry(1)
    If Not target Is Nothing Then rowMap = entry(6)
    Set columns = QualifyColumns(entry(2), BareColumnName(realName), Nothing)

    If UBound(data, 1) < 2 Then
        answers.Add AffectedAnswer(0, DELETE_COMMAND)
        Exit Sub
    End If

    ReDim going(2 To UBound(data, 1))
    If joined Then
        JoinedTargets name, targetAt, entry, whereNode, Nothing, going, _
                      Empty, affected
    Else
        For row = 2 To UBound(data, 1)
            values = RowValues(data, row)
            If IsEmpty(whereNode) Then
                keep = True
            Else
                keep = Evaluate(whereNode, values, columns)
            End If
            If Not IsNull(keep) Then
                If CBool(keep) Then
                    going(row) = True
                    affected = affected + 1
                End If
            End If
        Next row
    End If

    If affected > 0 Then
        If target Is Nothing Then
            entry(1) = WithoutRows(data, going)
            StoreTable CStr(entry(0)), entry
        Else
            RemoveRows target, rowMap, going, _
                       (SourceKind(entry) = "ListObject")
            RefreshTable CStr(entry(0))
        End If
    End If

    answers.Add AffectedAnswer(affected, DELETE_COMMAND)
End Sub

' The same rows without the marked ones, for a table that is its own rows.
Private Function WithoutRows(ByRef data As Variant, _
                             ByRef going() As Boolean) As Variant
    Dim out As Variant
    Dim row As Long
    Dim column As Long
    Dim into As Long
    Dim kept As Long

    For row = 2 To UBound(data, 1)
        If Not going(row) Then kept = kept + 1
    Next row

    ' Row one is the header wherever the rows came from, and it stays.
    ReDim out(1 To kept + 1, LBound(data, 2) To UBound(data, 2))
    For column = LBound(data, 2) To UBound(data, 2)
        out(1, column) = data(1, column)
    Next column

    into = 1
    For row = 2 To UBound(data, 1)
        If Not going(row) Then
            into = into + 1
            For column = LBound(data, 2) To UBound(data, 2)
                out(into, column) = data(row, column)
            Next column
        End If
    Next row
    WithoutRows = out
End Function

' The marked rows, taken out in as few operations as Excel will take.
'
' Rows that sit under each other are one area, and the areas are handed over
' a batch at a time, from the bottom up so that the rows still to go keep
' the numbers they had. One range of ten thousand separate rows, unioned and
' deleted once, was measured at four minutes: both the building and the
' delete cost more than linearly in the number of areas, so the number of
' areas is what is kept down.
Private Sub RemoveRows(ByVal target As Object, ByRef rowMap() As Long, _
                       ByRef going() As Boolean, ByVal listed As Boolean)
    Dim row As Long
    Dim runStart As Long
    Dim runEnd As Long
    Dim batch As Object
    Dim piece As Object
    Dim areas As Long

    row = UBound(going)
    Do While row >= LBound(going)
        If Not going(row) Then
            row = row - 1
        Else
            runEnd = row
            runStart = row
            Do While runStart > LBound(going)
                If Not going(runStart - 1) Then Exit Do
                ' Only rows that sit under each other on the sheet make one
                ' area; a row that was blank, and so is not a row of the
                ' table, still sits between them.
                If rowMap(runStart - 1) <> rowMap(runStart) - 1 Then Exit Do
                runStart = runStart - 1
            Loop

            Set piece = target.Rows(rowMap(runStart)) _
                              .Resize(runEnd - runStart + 1)
            If batch Is Nothing Then
                Set batch = piece
            Else
                Set batch = Application.Union(batch, piece)
            End If

            areas = areas + 1
            If areas >= DELETE_AREAS Then
                DeleteRange batch, listed
                Set batch = Nothing
                areas = 0
            End If
            row = runStart - 1
        End If
    Loop
    If Not batch Is Nothing Then DeleteRange batch, listed
End Sub

' A table's rows are deleted through the table, so what is beside it on the
' sheet stays where it is; a sheet's are whole rows.
Private Sub DeleteRange(ByVal going As Object, ByVal listed As Boolean)
    If listed Then
        going.Delete
    Else
        going.EntireRow.Delete
    End If
End Sub

' Keep a table under its name, replacing one already there.
Private Sub StoreTable(ByVal name As String, ByVal entry As Variant)
    Dim key As String
    Dim into As Collection

    key = LCase$(Replace(Replace(name, "[", ""), "]", ""))
    If Left$(key, 1) = "#" Then
        Set into = mTempTables
    Else
        Set into = mTables
    End If

    On Error Resume Next
    into.Remove key
    On Error GoTo 0
    into.Add entry, key

    ' A served table stored, made or written to is a catalog that reads
    ' differently now. A temporary one is not: the system rowsets are built
    ' from the served tables and never mention it, and the Object Explorer's
    ' Tables node is one batch that makes five temporary tables and fills
    ' them, which threw the rowsets away ten times over for nothing.
    If Not into Is mTempTables Then ForgetSystemRowsets
End Sub

' Split at a separator outside quotes, brackets and parentheses.
Private Function SplitTopLevel(ByVal text As String, _
                               ByVal separator As String) As Collection
    Dim parts As Collection
    Dim at As Long
    Dim start As Long
    Dim depth As Long
    Dim ch As String

    Set parts = New Collection
    at = 1
    start = 1

    Do While at <= Len(text)
        ch = Mid$(text, at, 1)
        If ch = "'" Or ch = """" Or ch = "[" Then
            at = PastOne(text, at)
        ElseIf ch = "(" Then
            depth = depth + 1
            at = at + 1
        ElseIf ch = ")" Then
            If depth > 0 Then depth = depth - 1
            at = at + 1
        ElseIf ch = separator And depth = 0 Then
            parts.Add Mid$(text, start, at - start)
            at = at + 1
            start = at
        Else
            at = at + 1
        End If
    Loop
    parts.Add Mid$(text, start)
    Set SplitTopLevel = parts
End Function


Private Function HasTable(ByVal name As String) As Boolean
    Dim entry As Variant
    Dim parts() As String

    parts = Split(name, ".")
    On Error Resume Next
    entry = mTables(LCase$(parts(UBound(parts))))
    HasTable = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


' ============================================================================
' INFORMATION_SCHEMA
'
' A client asks what exists before it asks for anything in particular, and a
' client that gets no answer here shows an empty table picker whatever else
' works.
' ============================================================================

Private Function IsInformationSchema(ByVal name As String) As Boolean
    IsInformationSchema = (InStr(1, name, "INFORMATION_SCHEMA.", vbTextCompare) > 0)
End Function

Private Function InformationSchemaTable(ByVal name As String) As Variant
    Dim which As String
    Dim parts() As String

    parts = Split(UCase$(name), ".")
    which = parts(UBound(parts))

    Select Case which
        Case "TABLES"
            InformationSchemaTable = SchemaTables()
        Case "COLUMNS"
            InformationSchemaTable = SchemaColumns()
        Case "VIEWS"
            InformationSchemaTable = SchemaViews()
        Case "SCHEMATA"
            InformationSchemaTable = SchemaSchemata()
        Case Else
            RaiseQueryError SQL_ERR_INVALID_OBJECT, _
                            "Invalid object name '" & name & "'."
    End Select
End Function

Private Function SchemaTables() As Variant
    Dim data As Variant
    Dim entry As Variant
    Dim index As Long

    ReDim data(1 To mTables.count + 1, 1 To 4)
    data(1, 1) = "TABLE_CATALOG"
    data(1, 2) = "TABLE_SCHEMA"
    data(1, 3) = "TABLE_NAME"
    data(1, 4) = "TABLE_TYPE"

    index = 2
    For Each entry In mTables
        data(index, 1) = ConfiguredDatabase
        data(index, 2) = "dbo"
        data(index, 3) = entry(0)
        data(index, 4) = IIf(IsViewEntry(entry), "VIEW", "BASE TABLE")
        index = index + 1
    Next entry
    SchemaTables = data
End Function

' INFORMATION_SCHEMA.VIEWS: the views, and the statement each was given.
Private Function SchemaViews() As Variant
    Dim data As Variant
    Dim entry As Variant
    Dim index As Long
    Dim total As Long

    For Each entry In mTables
        If IsViewEntry(entry) Then total = total + 1
    Next entry

    ReDim data(1 To total + 1, 1 To 6)
    data(1, 1) = "TABLE_CATALOG"
    data(1, 2) = "TABLE_SCHEMA"
    data(1, 3) = "TABLE_NAME"
    data(1, 4) = "VIEW_DEFINITION"
    data(1, 5) = "CHECK_OPTION"
    data(1, 6) = "IS_UPDATABLE"

    index = 2
    For Each entry In mTables
        If IsViewEntry(entry) Then
            data(index, 1) = ConfiguredDatabase
            data(index, 2) = "dbo"
            data(index, 3) = entry(0)
            data(index, 4) = ViewSql(entry)
            data(index, 5) = "NONE"
            data(index, 6) = "NO"
            index = index + 1
        End If
    Next entry
    SchemaViews = data
End Function

' INFORMATION_SCHEMA.COLUMNS, the twenty-three fields a real server declares
' and in that order. A client asking for CHARACTER_MAXIMUM_LENGTH, which is
' how it works out how wide to draw a text column, used to be told there was
' no such column and stop.
Private Function SchemaColumns() As Variant
    Dim data As Variant
    Dim entry As Variant
    Dim columns As Collection
    Dim column As Variant
    Dim described As Variant
    Dim textual As Boolean
    Dim total As Long
    Dim index As Long
    Dim field As Long
    Dim position As Long

    For Each entry In mTables
        Set columns = EntryColumns(entry)
        total = total + columns.count
    Next entry

    ReDim data(1 To total + 1, 1 To 23)
    data(1, 1) = "TABLE_CATALOG"
    data(1, 2) = "TABLE_SCHEMA"
    data(1, 3) = "TABLE_NAME"
    data(1, 4) = "COLUMN_NAME"
    data(1, 5) = "ORDINAL_POSITION"
    data(1, 6) = "COLUMN_DEFAULT"
    data(1, 7) = "IS_NULLABLE"
    data(1, 8) = "DATA_TYPE"
    data(1, 9) = "CHARACTER_MAXIMUM_LENGTH"
    data(1, 10) = "CHARACTER_OCTET_LENGTH"
    data(1, 11) = "NUMERIC_PRECISION"
    data(1, 12) = "NUMERIC_PRECISION_RADIX"
    data(1, 13) = "NUMERIC_SCALE"
    data(1, 14) = "DATETIME_PRECISION"
    data(1, 15) = "CHARACTER_SET_CATALOG"
    data(1, 16) = "CHARACTER_SET_SCHEMA"
    data(1, 17) = "CHARACTER_SET_NAME"
    data(1, 18) = "COLLATION_CATALOG"
    data(1, 19) = "COLLATION_SCHEMA"
    data(1, 20) = "COLLATION_NAME"
    data(1, 21) = "DOMAIN_CATALOG"
    data(1, 22) = "DOMAIN_SCHEMA"
    data(1, 23) = "DOMAIN_NAME"

    index = 2
    For Each entry In mTables
        Set columns = EntryColumns(entry)
        For position = 1 To columns.count
            column = columns(position)
            described = DescribeColumnType(CLng(column(COLUMN_TYPE)), _
                                           CLng(column(COLUMN_SIZE)))
            textual = (column(COLUMN_TYPE) = TYPE_NVARCHAR)

            For field = 1 To 23
                data(index, field) = Null
            Next field
            data(index, 1) = ConfiguredDatabase
            data(index, 2) = "dbo"
            data(index, 3) = entry(0)
            data(index, 4) = column(COLUMN_NAME)
            data(index, 5) = position
            data(index, 7) = "YES"
            data(index, 8) = SqlTypeName(CLng(column(COLUMN_TYPE)), _
                                         CLng(column(COLUMN_SIZE)))
            If textual Then
                data(index, 9) = described(2)
                data(index, 10) = described(3)
                data(index, 17) = "UNICODE"
                data(index, 18) = "master"
                data(index, 19) = "dbo"
                data(index, 20) = "SQL_Latin1_General_CP1_CI_AS"
            ElseIf column(COLUMN_TYPE) = TYPE_DATETIMN Then
                data(index, 14) = described(4)
            ElseIf column(COLUMN_TYPE) <> TYPE_GUIDN Then
                data(index, 11) = described(2)
                data(index, 12) = described(5)
                data(index, 13) = described(4)
            End If
            index = index + 1
        Next position
    Next entry
    SchemaColumns = data
End Function

Private Function SchemaSchemata() As Variant
    Dim data As Variant

    ReDim data(1 To 2, 1 To 3)
    data(1, 1) = "CATALOG_NAME"
    data(1, 2) = "SCHEMA_NAME"
    data(1, 3) = "SCHEMA_OWNER"
    data(2, 1) = ConfiguredDatabase
    data(2, 2) = "dbo"
    data(2, 3) = "dbo"
    SchemaSchemata = data
End Function

Private Function SqlTypeName(ByVal typeCode As Long, ByVal size As Long) As String
    Select Case typeCode
        Case TYPE_INTN
            SqlTypeName = IIf(size = 8, "bigint", "int")
        Case TYPE_FLTN
            SqlTypeName = "float"
        Case TYPE_BITN
            SqlTypeName = "bit"
        Case TYPE_DATETIMN
            SqlTypeName = "datetime"
        Case TYPE_GUIDN
            SqlTypeName = "uniqueidentifier"
        Case TYPE_INT2
            SqlTypeName = "smallint"
        Case Else
            SqlTypeName = "nvarchar"
    End Select
End Function


' ============================================================================
' Reading SQL
' ============================================================================

' Split a statement into tokens: names, numbers, strings and operators.
'
' Bracketed and quoted identifiers keep whatever is inside them, including
' spaces and reserved words, because that is what the brackets are for.
Private Function Tokenize(ByVal sql As String) As Collection
    Dim tokens As Collection
    Dim cursor As Long
    Dim length As Long
    Dim ch As String
    Dim start As Long
    Dim text As String

    Set tokens = New Collection
    length = Len(sql)
    cursor = 1

    Do While cursor <= length
        ch = Mid$(sql, cursor, 1)

        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            cursor = cursor + 1

        ElseIf ch = "-" And Mid$(sql, cursor, 2) = "--" Then
            Do While cursor <= length And Mid$(sql, cursor, 1) <> vbLf
                cursor = cursor + 1
            Loop

        ElseIf ch = "/" And Mid$(sql, cursor, 2) = "/*" Then
            cursor = cursor + 2
            Do While cursor <= length And Mid$(sql, cursor, 2) <> "*/"
                cursor = cursor + 1
            Loop
            cursor = cursor + 2

        ElseIf ch = "'" Then
            cursor = cursor + 1
            text = vbNullString
            Do While cursor <= length
                If Mid$(sql, cursor, 1) = "'" Then
                    If Mid$(sql, cursor, 2) = "''" Then
                        text = text & "'"
                        cursor = cursor + 2
                    Else
                        cursor = cursor + 1
                        Exit Do
                    End If
                Else
                    text = text & Mid$(sql, cursor, 1)
                    cursor = cursor + 1
                End If
            Loop
            tokens.Add TokenOf(T_STRING, text)

        ElseIf ch = "[" Then
            tokens.Add TokenOf(T_NAME, ReadQualifiedName(sql, cursor))

        ElseIf ch = """" Then
            ' A name in double quotes is one name, parts and all, the way a
            ' bracketed one is. Read a part at a time, "Tbl1002"."id" was a
            ' column called Tbl1002 followed by a stray dot, and a linked
            ' server writes every name it sends that way.
            tokens.Add TokenOf(T_NAME, ReadQualifiedName(sql, cursor))

        ElseIf ch = "0" And UCase$(Mid$(sql, cursor + 1, 1)) = "X" Then
            ' A hex literal. SSMS writes masks and status bits this way, and
            ' the value goes on as the number it stands for.
            start = cursor + 2
            cursor = start
            Do While cursor <= length And IsHexChar(Mid$(sql, cursor, 1))
                cursor = cursor + 1
            Loop
            tokens.Add TokenOf(T_NUMBER, _
                CStr(HexValue(Mid$(sql, start, cursor - start))))

        ElseIf (UCase$(ch) = "N") And Mid$(sql, cursor + 1, 1) = "'" Then
            ' N'text' is text. The prefix says it is Unicode, which
            ' everything here already is.
            cursor = cursor + 1

        ElseIf IsDigitChar(ch) Or (ch = "." And IsDigitChar(Mid$(sql, cursor + 1, 1))) Then
            start = cursor
            Do While cursor <= length And _
                     (IsDigitChar(Mid$(sql, cursor, 1)) Or Mid$(sql, cursor, 1) = ".")
                cursor = cursor + 1
            Loop
            ' An exponent, which is how SQL Server writes a float it pushes
            ' down to a linked server: 9.0000000000000000e+001 for ninety.
            ' Stopping at the e left it as a name and the whole statement a
            ' syntax error.
            If cursor <= length And UCase$(Mid$(sql, cursor, 1)) = "E" Then
                If IsDigitChar(Mid$(sql, cursor + 1, 1)) Then
                    cursor = cursor + 1
                ElseIf InStr("+-", Mid$(sql, cursor + 1, 1)) > 0 And _
                       IsDigitChar(Mid$(sql, cursor + 2, 1)) Then
                    cursor = cursor + 2
                End If
                Do While cursor <= length And IsDigitChar(Mid$(sql, cursor, 1))
                    cursor = cursor + 1
                Loop
            End If
            tokens.Add TokenOf(T_NUMBER, Mid$(sql, start, cursor - start))

        ElseIf IsNameChar(ch) Or ch = "@" Or ch = "#" Then
            tokens.Add TokenOf(T_NAME, ReadQualifiedName(sql, cursor))

        ElseIf InStr("<>!=|", ch) > 0 And _
               InStr("<>=", Mid$(sql, cursor + 1, 1)) > 0 Then
            tokens.Add TokenOf(T_OPERATOR, Mid$(sql, cursor, 2))
            cursor = cursor + 2

        Else
            tokens.Add TokenOf(T_OPERATOR, ch)
            cursor = cursor + 1
        End If
    Loop

    Set Tokenize = tokens
End Function

Private Function TokenOf(ByVal kind As Long, ByVal text As String) As Variant
    Dim out(0 To 1) As Variant

    out(0) = kind
    out(1) = text
    TokenOf = out
End Function

Private Function IsDigitChar(ByVal ch As String) As Boolean
    IsDigitChar = (ch >= "0" And ch <= "9")
End Function

Private Function IsNameChar(ByVal ch As String) As Boolean
    IsNameChar = (UCase$(ch) >= "A" And UCase$(ch) <= "Z") Or ch = "_" Or ch = "."
End Function

' -- the parser's view of the token stream --------------------------------

Private Function AtEnd() As Boolean
    AtEnd = (mCursor > mTokens.count)
End Function

Private Function PeekText() As String
    If AtEnd() Then Exit Function
    PeekText = mTokens(mCursor)(1)
End Function

Private Function PeekKind() As Long
    If AtEnd() Then Exit Function
    PeekKind = mTokens(mCursor)(0)
End Function

Private Function PeekIs(ByVal text As String) As Boolean
    If AtEnd() Then Exit Function

    ' A string is what it says, not what it spells. ISNULL(x, '-') holds a
    ' minus sign and is not a subtraction; matching on the text alone read
    ' it as one and then wanted something to subtract.
    If PeekKind() = T_STRING Then Exit Function

    PeekIs = (StrComp(PeekText(), text, vbTextCompare) = 0)
End Function

Private Function TakeText() As String
    TakeText = PeekText()
    mCursor = mCursor + 1
End Function

Private Function Accept(ByVal text As String) As Boolean
    If PeekIs(text) Then
        mCursor = mCursor + 1
        Accept = True
    End If
End Function

Private Sub Expect(ByVal text As String)
    If Not Accept(text) Then
        RaiseQueryError ERR_NUMBER_BASE, _
            "Incorrect syntax near '" & PeekText() & "'; expected '" & text & "'."
    End If
End Sub

' -- expressions ----------------------------------------------------------

' A node holds its kind and up to three parts, any of which may be a
' Collection. The assignments are written out because VBA will not pass a
' statically dimensioned array into a Variant parameter, so a helper that
' took the array by reference could not be called.
Private Function NodeOf(ByVal kind As Long, ByVal a As Variant, _
                        ByVal b As Variant, ByVal c As Variant) As Variant
    Dim out(0 To 3) As Variant

    out(0) = kind
    If IsObject(a) Then Set out(1) = a Else out(1) = a
    If IsObject(b) Then Set out(2) = b Else out(2) = b
    If IsObject(c) Then Set out(3) = c Else out(3) = c
    NodeOf = out
End Function

Private Function ParseExpression() As Variant
    ParseExpression = ParseOr()
End Function

Private Function ParseOr() As Variant
    Dim left As Variant

    left = ParseAnd()
    Do While Accept("OR")
        left = NodeOf(N_BINARY, "OR", left, ParseAnd())
    Loop
    ParseOr = left
End Function

Private Function ParseAnd() As Variant
    Dim left As Variant

    left = ParseNot()
    Do While Accept("AND")
        left = NodeOf(N_BINARY, "AND", left, ParseNot())
    Loop
    ParseAnd = left
End Function

Private Function ParseNot() As Variant
    If Accept("NOT") Then
        ParseNot = NodeOf(N_UNARY, "NOT", ParseNot(), Empty)
    Else
        ParseNot = ParseComparison()
    End If
End Function

Private Function ParseComparison() As Variant
    Dim left As Variant
    Dim operator As String
    Dim negated As Boolean
    Dim items As Collection
    Dim quantifier As String

    left = ParseAdditive()

    If Accept("IS") Then
        negated = Accept("NOT")
        Expect "NULL"
        ParseComparison = NodeOf(N_UNARY, IIf(negated, "IS NOT NULL", "IS NULL"), _
                                 left, Empty)
        Exit Function
    End If

    negated = Accept("NOT")
    If Accept("LIKE") Then
        Dim pattern As Variant
        pattern = ParseAdditive()
        If Accept("ESCAPE") Then
            ' The pattern and its escape character travel as one right-hand
            ' side, the way BETWEEN's two bounds do, because a node holds an
            ' operator and two operands and nothing more.
            ParseComparison = NodeOf(N_BINARY, _
                IIf(negated, "NOT LIKE ESCAPE", "LIKE ESCAPE"), left, _
                NodeOf(N_BINARY, ",", pattern, ParseAdditive()))
            Exit Function
        End If
        ParseComparison = NodeOf(N_BINARY, IIf(negated, "NOT LIKE", "LIKE"), _
                                 left, pattern)
        Exit Function
    End If
    If Accept("IN") Then
        Expect "("
        If PeekIs("SELECT") Then
            ParseComparison = NodeOf(N_BINARY, _
                                     IIf(negated, "NOT IN READ", "IN READ"), _
                                     left, BracketedTokens())
            Exit Function
        End If
        Set items = New Collection
        Do
            items.Add ParseExpression()
        Loop While Accept(",")
        Expect ")"
        ParseComparison = NodeOf(N_BINARY, IIf(negated, "NOT IN", "IN"), _
                                 left, items)
        Exit Function
    End If
    If Accept("BETWEEN") Then
        Dim low As Variant
        low = ParseAdditive()
        Expect "AND"
        ParseComparison = NodeOf(N_BINARY, IIf(negated, "NOT BETWEEN", "BETWEEN"), _
                                 left, NodeOf(N_BINARY, ",", low, ParseAdditive()))
        Exit Function
    End If
    If negated Then
        RaiseQueryError ERR_NUMBER_BASE, _
            "Incorrect syntax near '" & PeekText() & "'."
    End If

    If PeekKind() = T_OPERATOR Then
        operator = PeekText()
        Select Case operator
            Case "=", "<>", "!=", "<", "<=", ">", ">="
                mCursor = mCursor + 1

                ' x > ANY (SELECT ...) holds where the comparison holds for
                ' one of the rows, and x > ALL where it holds for all of
                ' them. SOME is another spelling of ANY.
                If PeekIs("ANY") Or PeekIs("SOME") Or PeekIs("ALL") Then
                    quantifier = IIf(PeekIs("ALL"), "ALL", "ANY")
                    mCursor = mCursor + 1
                    Expect "("
                    ParseComparison = NodeOf(N_BINARY, _
                        operator & " " & quantifier, left, BracketedTokens())
                    Exit Function
                End If

                ParseComparison = NodeOf(N_BINARY, operator, left, ParseAdditive())
                Exit Function
        End Select
    End If

    ParseComparison = left
End Function

Private Function ParseAdditive() As Variant
    Dim left As Variant
    Dim operator As String

    left = ParseMultiplicative()
    Do
        If PeekIs("+") Then
            operator = "+"
        ElseIf PeekIs("-") Then
            operator = "-"
        ElseIf PeekIs("&") Then
            operator = "&"
        ElseIf PeekIs("|") Then
            operator = "|"
        ElseIf PeekIs("^") Then
            operator = "^"
        Else
            Exit Do
        End If
        mCursor = mCursor + 1
        left = NodeOf(N_BINARY, operator, left, ParseMultiplicative())
    Loop
    ParseAdditive = left
End Function

Private Function ParseMultiplicative() As Variant
    Dim left As Variant
    Dim operator As String

    left = ParseUnary()
    Do
        If PeekIs("*") Then
            operator = "*"
        ElseIf PeekIs("/") Then
            operator = "/"
        ElseIf PeekIs("%") Then
            operator = "%"
        Else
            Exit Do
        End If
        mCursor = mCursor + 1
        left = NodeOf(N_BINARY, operator, left, ParseUnary())
    Loop
    ParseMultiplicative = left
End Function

Private Function ParseUnary() As Variant
    If Accept("-") Then
        ParseUnary = NodeOf(N_UNARY, "-", ParseUnary(), Empty)
    ElseIf Accept("+") Then
        ParseUnary = ParseUnary()
    ElseIf Accept("~") Then
        ' An object tree writes ~i.allow_page_locks to read a flag it stores
        ' the other way round from the way it shows it.
        ParseUnary = NodeOf(N_UNARY, "~", ParseUnary(), Empty)
    Else
        ParseUnary = ParsePrimary()

        ' COLLATE names a collation to compare in. This serves one, so the
        ' clause is read and dropped: taking it as an alias instead turns
        ' the column in front of it into a name nothing has.
        If Accept("COLLATE") Then
            If Not AtEnd() Then mCursor = mCursor + 1
        End If
    End If
End Function

Private Function ParsePrimary() As Variant
    Dim text As String
    Dim node As Variant
    Dim arguments As Collection
    Dim distinct As Boolean

    If AtEnd() Then
        RaiseQueryError ERR_NUMBER_BASE, _
            "Incorrect syntax near the end of the statement."
    End If

    If Accept("(") Then
        If PeekIs("SELECT") Then
            ' A read used as a value. Its tokens are kept rather than its
            ' text, because the tokens are what the parser runs on and the
            ' text is gone by now.
            ParsePrimary = NodeOf(N_SUBQUERY, BracketedTokens(), Empty, Empty)
            Exit Function
        End If
        node = ParseExpression()
        Expect ")"
        ParsePrimary = node
        Exit Function
    End If

    If PeekIs("*") Then
        mCursor = mCursor + 1
        ParsePrimary = NodeOf(N_STAR, vbNullString, Empty, Empty)
        Exit Function
    End If

    Select Case PeekKind()
        Case T_NUMBER
            text = TakeText()
            If InStr(text, ".") > 0 Or InStr(1, text, "e", vbTextCompare) > 0 Then
                ParsePrimary = NodeOf(N_LITERAL, CDbl(Val(text)), Empty, Empty)
            ElseIf Abs(Val(text)) > 2147483647# Then
                ' Past what a Long holds. The collation a login announces is
                ' written 0x0904D00034, which a linked server hands straight
                ' back to collationname, and forcing it into a Long raised
                ' Overflow before the function was even looked up.
                ParsePrimary = NodeOf(N_LITERAL, CLngLng(Val(text)), Empty, Empty)
            Else
                ParsePrimary = NodeOf(N_LITERAL, CLng(Val(text)), Empty, Empty)
            End If

        Case T_STRING
            ParsePrimary = NodeOf(N_LITERAL, TakeText(), Empty, Empty)

        Case T_NAME
            text = TakeText()
            If Right$(text, 1) = "." And PeekIs("*") Then
                ' t.* is every column of that source, and only that source.
                mCursor = mCursor + 1
                ParsePrimary = NodeOf(N_STAR, Left$(text, Len(text) - 1), _
                                      Empty, Empty)
                Exit Function
            End If
            If StrComp(text, "NULL", vbTextCompare) = 0 Then
                ParsePrimary = NodeOf(N_LITERAL, Null, Empty, Empty)
            ElseIf StrComp(text, "CASE", vbTextCompare) = 0 Then
                ParsePrimary = ParseCase()
            ElseIf StrComp(text, "CAST", vbTextCompare) = 0 And PeekIs("(") Then
                ParsePrimary = ParseCast()
            ElseIf StrComp(text, "CONVERT", vbTextCompare) = 0 And PeekIs("(") Then
                ParsePrimary = ParseConvert()
            ElseIf StrComp(text, "TRIM", vbTextCompare) = 0 And PeekIs("(") Then
                ParsePrimary = ParseTrim()
            ElseIf StrComp(text, "EXISTS", vbTextCompare) = 0 And PeekIs("(") Then
                ' Not a function: what follows its bracket is a statement,
                ' and reading it as an argument list makes SELECT a column
                ' name and the token after it a syntax error.
                mCursor = mCursor + 1
                ParsePrimary = NodeOf(N_EXISTS, BracketedTokens(), Empty, Empty)
            ElseIf PeekIs("(") Then
                mCursor = mCursor + 1
                Set arguments = New Collection

                ' An aggregate may say DISTINCT before what it counts.
                ' Reading it as the start of an expression made it a column
                ' name, and the column after it a syntax error.
                distinct = Accept("DISTINCT")
                If Not distinct Then Accept "ALL"

                If Not PeekIs(")") Then
                    Do
                        arguments.Add ParseExpression()
                    Loop While Accept(",")
                End If
                Expect ")"
                node = NodeOf(N_CALL, UCase$(text), arguments, distinct)
                If PeekIs("OVER") Then node = ParseOver(node)
                ParsePrimary = node
            Else
                ParsePrimary = NodeOf(N_COLUMN, text, Empty, Empty)
            End If

        Case Else
            RaiseQueryError ERR_NUMBER_BASE, _
                "Incorrect syntax near '" & PeekText() & "'."
    End Select
End Function

' CASE, in both its forms, flattened into nested calls the evaluator can walk.
Private Function ParseCase() As Variant
    Dim subject As Variant
    Dim hasSubject As Boolean
    Dim arguments As Collection

    Set arguments = New Collection
    If Not PeekIs("WHEN") Then
        subject = ParseExpression()
        hasSubject = True
    End If

    Do While Accept("WHEN")
        If hasSubject Then
            arguments.Add NodeOf(N_BINARY, "=", subject, ParseExpression())
        Else
            arguments.Add ParseExpression()
        End If
        Expect "THEN"
        arguments.Add ParseExpression()
    Loop

    If Accept("ELSE") Then
        arguments.Add ParseExpression()
    Else
        arguments.Add NodeOf(N_LITERAL, Null, Empty, Empty)
    End If
    Expect "END"

    ParseCase = NodeOf(N_CALL, "$CASE", arguments, Empty)
End Function

' -- the statement --------------------------------------------------------

Private Sub ParseSelect()
    Dim node As Variant
    Dim topValue As Variant
    Dim alias As String
    Dim item(0 To 1) As Variant

    Set mSelectItems = New Collection
    Set mOrderBy = New Collection
    Set mGroupBy = New Collection
    Set mFromItems = New Collection
    mFromName = vbNullString
    mWhereNode = Empty
    mHavingNode = Empty
    mTopCount = -1
    mTopPercent = -1
    mOffsetCount = -1
    mFetchCount = -1
    mIsDistinct = False

    ' A read in brackets is the read inside them.
    If Accept("(") Then
        ParseSelect
        Expect ")"
        ParseOrderBy
        Exit Sub
    End If

    Expect "SELECT"
    If Accept("DISTINCT") Then mIsDistinct = True
    If Accept("ALL") Then mIsDistinct = False
    If Accept("TOP") Then
        ' TOP (expression) takes a variable, a parameter or arithmetic, and
        ' is worked out before a row is read. Read as text, TOP (@n) was the
        ' number in "@n", which is none, and answered no rows at all.
        If Accept("(") Then
            topValue = EvaluateHere(ParseExpression())
            Expect ")"
        Else
            topValue = Val(TakeText())
        End If
        ' PERCENT is of the rows the statement ends up with, so it waits for
        ' them. Read and dropped, TOP 50 PERCENT answered fifty rows.
        If Accept("PERCENT") Then
            mTopPercent = CDbl(NullToZero(topValue))
        Else
            mTopCount = CLng(NullToZero(topValue))
        End If
    End If

    Do
        node = ParseExpression()
        alias = vbNullString
        If Accept("AS") Then
            alias = TakeText()
        ElseIf PeekKind() = T_NAME And Not IsClauseWord(PeekText()) Then
            alias = TakeText()
        ElseIf PeekKind() = T_STRING Then
            alias = TakeText()
        End If
        item(0) = node
        item(1) = alias
        mSelectItems.Add item
    Loop While Accept(",")

    If Accept("FROM") Then ReadFromClause

    If Accept("WHERE") Then mWhereNode = ParseExpression()

    If Accept("GROUP") Then
        Expect "BY"
        Do
            mGroupBy.Add ParseExpression()
        Loop While Accept(",")
    End If

    If Accept("HAVING") Then mHavingNode = ParseExpression()

    ParseOrderBy
End Sub

' ORDER BY, with the page OFFSET and FETCH cut from it. Also read after a
' read in brackets, where it orders what the brackets hold, or with a set
' operator before it, the whole statement.
Private Sub ParseOrderBy()
    Dim node As Variant
    Dim descending As Boolean
    Dim order(0 To 1) As Variant

    If Accept("ORDER") Then
        Expect "BY"
        Do
            node = ParseExpression()
            descending = False
            If Accept("DESC") Then
                descending = True
            Else
                Accept "ASC"
            End If
            order(0) = node
            order(1) = descending
            mOrderBy.Add order
        Loop While Accept(",")

        ' OFFSET n ROWS [FETCH NEXT m ROWS ONLY]. Read as a clause boundary
        ' and then dropped, it asked for a page and got the whole answer.
        If Accept("OFFSET") Then
            mOffsetCount = CLng(Val(TakeText()))
            If Not Accept("ROWS") Then Accept "ROW"
            If Accept("FETCH") Then
                If Not Accept("NEXT") Then Accept "FIRST"
                mFetchCount = CLng(Val(TakeText()))
                If Not Accept("ROWS") Then Accept "ROW"
                Accept "ONLY"
            End If
        End If
    End If
End Sub

Private Function IsClauseWord(ByVal text As String) As Boolean
    Select Case UCase$(text)
        Case "FROM", "WHERE", "ORDER", "GROUP", "HAVING", "AS", "AND", "OR", _
             "ASC", "DESC", "INNER", "LEFT", "RIGHT", "FULL", "JOIN", "ON", _
             "UNION", "EXCEPT", "INTERSECT", "OFFSET", "FETCH", "FOR", "OPTION", _
             "CROSS", "OUTER", "APPLY", "VALUES", "PIVOT", "UNPIVOT", "WITH", _
             "INTO", "END", "ELSE", "THEN", "WHEN"
            IsClauseWord = True
    End Select
End Function


' ============================================================================
' Answering
' ============================================================================

' What a client's batch gets back, as the bytes of a token stream.
'
' A call with no connection's variables behind it is a session of its own,
' and a transaction it leaves open is rolled back as it ends, whether the
' batch answered or failed, the way a client's is when it disconnects.
Friend Function AnswerQuery(ByVal sql As String, ByVal tdsVersion As Long, _
                            Optional ByVal parameters As Collection = Nothing, _
                            Optional ByVal variables As Collection = Nothing, _
                            Optional ByVal tempTables As Collection = Nothing, _
                            Optional ByVal session As Collection = Nothing, _
                            Optional ByVal describeOnly As Boolean = False) _
                            As Byte()
    Dim transient As Boolean
    Dim number As Long
    Dim failedIn As String
    Dim description As String

    transient = (variables Is Nothing)
    On Error GoTo Failed
    AnswerQuery = AnswerBatch(sql, tdsVersion, parameters, variables, _
                              tempTables, session, describeOnly)
    If transient Then EndTransientTransaction
    Exit Function

Failed:
    number = Err.Number
    failedIn = Err.Source
    description = Err.Description
    If transient Then EndTransientTransaction
    Err.Raise number, failedIn, description
End Function

' The batch itself, answered with whatever state the caller handed over.
Private Function AnswerBatch(ByVal sql As String, ByVal tdsVersion As Long, _
                            Optional ByVal parameters As Collection = Nothing, _
                            Optional ByVal variables As Collection = Nothing, _
                            Optional ByVal tempTables As Collection = Nothing, _
                            Optional ByVal session As Collection = Nothing, _
                            Optional ByVal describeOnly As Boolean = False) _
                            As Byte()
    Dim statement As String
    Dim pieces As Collection
    Dim piece As Variant
    Dim answers As Collection
    Dim answer As Variant
    Dim raised As Long
    Dim failedIn As String
    Dim message As String

    mQueryErrorNumber = 0
    mCaughtNumber = 0
    mCaughtMessage = vbNullString
    Set mCommonTables = Nothing
    If parameters Is Nothing Then
        Set mParameters = New Collection
    Else
        Set mParameters = parameters
    End If
    If variables Is Nothing Then
        Set mVariables = New Collection
    Else
        Set mVariables = variables
    End If
    If tempTables Is Nothing Then
        Set mTempTables = New Collection
    Else
        Set mTempTables = tempTables
    End If
    Set mSession = session

    ' Read the sheets again before answering. One pass per statement,
     ' rather than one per mention of a table, because a join naming the
     ' same source twice is one question about one moment.
    RefreshTables

    statement = TrimWhitespace(StripComments(sql))
    Do While Right$(statement, 1) = ";"
        statement = TrimWhitespace(Left$(statement, Len(statement) - 1))
    Loop
    If Len(statement) = 0 Then
        AnswerBatch = TokenDone(DONE_FINAL, 0, tdsVersion)
        Exit Function
    End If

    ' A client can call a procedure as a batch as well as through RPC, and
    ' the ODBC drivers do. Only when that is the whole of it: SSMS writes an
    ' EXEC inside a batch that goes on to read something.
    If UCase$(FirstWord(statement)) = "EXEC" _
       Or UCase$(FirstWord(statement)) = "EXECUTE" Then
        If NextWordIn(statement, 1, STATEMENT_STARTS) = 1 _
           And NextWordIn(statement, 5, STATEMENT_STARTS) = 0 Then
            AnswerBatch = AnswerProcedure(ProcedureFromBatch(statement), _
                                          Nothing, tdsVersion)
            Exit Function
        End If
    End If

    Set pieces = SplitStatements(statement)
    Set answers = New Collection

    ' A real server carries a batch on past some errors and ends it at
    ' others. The ones it carries on past are kept among the answers as each
    ' statement runs, and the ones that end it arrive at Stopped.
    mCatching = Not describeOnly
    On Error GoTo Stopped
    For Each piece In pieces
        RunOne CStr(piece), answers
    Next piece
Answered:
    On Error GoTo 0

    ' Nothing that produces rows means no result set at all. A SET answered
    ' with columns makes a client report an invalid cursor state when it
    ' later reads the results it was actually waiting for.
    If answers.count = 0 Then
        AnswerBatch = TokenDone(DONE_FINAL, 0, tdsVersion)
        Exit Function
    End If

    ' The whole batch came to one error and nothing else. Raised rather
    ' than returned: it is the same bytes on the wire, and it is what every
    ' caller already expects of a statement that could not run.
    If answers.count = 1 Then
        answer = answers(1)
        If answer(2) = ANSWER_FAILED Then
            RaiseQueryError CLng(answer(3)), CStr(answer(4))
        End If
    End If
    If describeOnly Then
        ' What sp_prepare answers with: the first result's columns, and
        ' none of its rows.
        For Each answer In answers
            If answer(2) < 0 And answer(2) <> ANSWER_FAILED Then
                AnswerBatch = ResultSet(answer(0), New Collection, tdsVersion)
                Exit Function
            End If
        Next answer
        AnswerBatch = TokenDone(DONE_FINAL, 0, tdsVersion)
        Exit Function
    End If
    AnswerBatch = EncodeAnswers(answers, tdsVersion)
    Exit Function

Stopped:
    ' The batch ends here. What ran before it has answered, and a real
    ' server has sent those results by the time it reaches the error, so
    ' for the errors measured to end a batch that far in they are kept and
    ' the error is the last thing read. Any other error is answered by
    ' itself, the way one a real server finds before it runs anything is.
    raised = Err.Number
    failedIn = Err.Source
    message = Err.Description
    If raised = ERR_QUERY And answers.count > 0 Then
        If EndsTheBatchAfter(mQueryErrorNumber) Then
            answers.Add FailedAnswer(mQueryErrorNumber, message)
            Resume Answered
        End If
    End If
    Err.Raise raised, failedIn, message
End Function

' The procedure an EXEC batch names, without its arguments.
Private Function ProcedureFromBatch(ByVal statement As String) As String
    Dim rest As String
    Dim ends As Long
    Dim index As Long
    Dim ch As String

    rest = TrimWhitespace(Mid$(statement, Len(FirstWord(statement)) + 1))
    ends = Len(rest) + 1
    For index = 1 To Len(rest)
        ch = Mid$(rest, index, 1)
        If ch = " " Or ch = "(" Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            ends = index
            Exit For
        End If
    Next index
    ProcedureFromBatch = Left$(rest, ends - 1)
End Function

Private Function FirstWord(ByVal text As String) As String
    Dim parts() As String

    parts = Split(Trim$(text), " ")
    FirstWord = parts(0)
End Function

' The rows a statement reads, and what they are called.
Private Sub ReadSource(ByRef matching As Collection, _
                       ByRef sourceColumns As Collection)
    Dim rows As Collection

    CombinedSource sourceColumns, rows
    Set matching = FilterRows(rows, sourceColumns)
End Sub

' Everything the FROM names, joined, before the WHERE has looked at it.
Private Sub CombinedSource(ByRef columns As Collection, _
                           ByRef rows As Collection)
    Dim index As Long

    If mFromItems Is Nothing Then
        Set columns = New Collection
        Set rows = SingleRow()
        Exit Sub
    End If
    If mFromItems.count = 0 Then
        Set columns = New Collection
        Set rows = SingleRow()
        Exit Sub
    End If

    LoadFromItem mFromItems(1), columns, rows
    For index = 2 To mFromItems.count
        CombineWith mFromItems(index), columns, rows
    Next index
End Sub

Private Sub ExecuteSelect(ByRef outColumns As Collection, _
                          ByRef outRows As Collection)
    Dim matching As Collection
    Dim sourceColumns As Collection
    Dim hidden As Long

    ReadSource matching, sourceColumns

    ' An aggregate written only in the ORDER BY is worked out with the rest
    ' of them and carried out of sight: the rows going out are the groups,
    ' and a group cannot be summarised again from the one row it became.
    hidden = KeepOrderAggregates()
    Project matching, sourceColumns, outColumns, outRows

    ' A grouped or aggregated answer is ordered by what it produced rather
    ' than by what it read: the rows the ORDER BY would have sorted no longer
    ' exist one to one with the rows going out.
    If Summarising() And mOrderBy.count > 0 Then
        Set outRows = SortRows(outRows, outColumns)
    End If
    If hidden > 0 Then DropHiddenColumns outColumns, outRows, hidden

    ' A page is taken from the rows in the order they end up in, which is
    ' why this waits until the sorting is done.
    TakeTopPercent outRows
    TakePage outRows
End Sub

' The aggregates an ORDER BY names that the select list does not, added to
' the select list under names of their own so the groups work them out, and
' taken off the answer again before it goes out. ORDER BY COUNT(*) DESC over
' a grouped read is the case: a real server sorts by it without showing it.
Private Function KeepOrderAggregates() As Long
    Dim term As Variant
    Dim node As Variant
    Dim added As Long

    If Not Summarising() Then Exit Function

    For Each term In mOrderBy
        node = term(0)
        If NodeHasAggregate(node) Then
            If SelectItemMatching(node) = 0 Then
                added = added + 1
                mSelectItems.Add SelectItemOf(node, "$order" & added)
            End If
        End If
    Next term
    KeepOrderAggregates = added
End Function

Private Function SelectItemOf(ByVal node As Variant, _
                              ByVal alias As String) As Variant
    Dim item(0 To 1) As Variant

    item(0) = node
    item(1) = alias
    SelectItemOf = item
End Function

' Take the columns an ORDER BY needed and nobody asked for back off the
' answer, and off the select list with it.
Private Sub DropHiddenColumns(ByRef columns As Collection, _
                              ByRef rows As Collection, ByVal count As Long)
    Dim kept As Collection
    Dim trimmedRows As Collection
    Dim values As Variant
    Dim trimmed As Variant
    Dim width As Long
    Dim index As Long

    width = columns.count - count
    If width < 1 Then Exit Sub

    Set kept = New Collection
    For index = 1 To width
        kept.Add columns(index)
    Next index

    Set trimmedRows = New Collection
    For Each values In rows
        ReDim trimmed(0 To width - 1)
        For index = 0 To width - 1
            trimmed(index) = values(LBound(values) + index)
        Next index
        trimmedRows.Add trimmed
    Next values

    Set columns = kept
    Set rows = trimmedRows
    For index = 1 To count
        mSelectItems.Remove mSelectItems.count
    Next index
End Sub

' TOP n PERCENT, over the rows in their final order: n per cent of them,
' rounded up.
Private Sub TakeTopPercent(ByRef rows As Collection)
    Dim wanted As Long
    Dim kept As Collection
    Dim values As Variant

    If mTopPercent < 0 Then Exit Sub
    wanted = -Int(-rows.count * mTopPercent / 100)
    If wanted >= rows.count Then Exit Sub

    Set kept = New Collection
    For Each values In rows
        If kept.count >= wanted Then Exit For
        kept.Add values
    Next values
    Set rows = kept
End Sub

' OFFSET and FETCH, over the rows in their final order.
Private Sub TakePage(ByRef rows As Collection)
    Dim page As Collection
    Dim values As Variant
    Dim at As Long

    If mOffsetCount < 0 And mFetchCount < 0 Then Exit Sub

    Set page = New Collection
    For Each values In rows
        at = at + 1
        If mOffsetCount < 0 Or at > mOffsetCount Then
            If mFetchCount >= 0 And page.count >= mFetchCount Then Exit For
            page.Add values
        End If
    Next values
    Set rows = page
End Sub

' Whether this statement collapses its rows, by grouping or by aggregating.
Private Function Summarising() As Boolean
    Summarising = (mGroupBy.count > 0) Or HasAggregate()
End Function

' One empty row, so a select list with no FROM is evaluated exactly once.
Private Function SingleRow() As Collection
    Dim rows As Collection
    Dim blank() As Variant

    ReDim blank(0 To 0)
    Set rows = New Collection
    rows.Add blank
    Set SingleRow = rows
End Function

' Turn the kept rows into the answer's columns and values.
Private Sub Project(ByVal matching As Collection, ByVal sourceColumns As Collection, _
                    ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim item As Variant
    Dim node As Variant
    Dim values As Variant
    Dim produced As Variant
    Dim index As Long
    Dim position As Long
    Dim column As Variant
    Dim taken As Long
    Dim seen As Collection
    Dim computed As Collection
    Dim slots() As Long
    Dim starWanted As Collection
    Dim expressions As Collection
    Dim takes() As Boolean
    Dim at As Long

    If mGroupBy.count > 0 Then
        ProjectGroups matching, sourceColumns, outColumns, outRows
        Exit Sub
    End If
    If HasAggregate() Then
        ProjectAggregates matching, sourceColumns, outColumns, outRows
        Exit Sub
    End If

    ' The shape first, from the select list rather than from any row, because
    ' a query that matches nothing still declares what it would have sent.
    Set computed = New Collection
    For Each item In mSelectItems
        node = item(0)
        If node(0) = N_STAR Then
            For position = 1 To sourceColumns.count
                If StarWants(node, sourceColumns(position)) Then
                    outColumns.Add sourceColumns(position)
                    computed.Add False
                End If
            Next position
        Else
            outColumns.Add ColumnOf(ItemName(item, outColumns.count + 1), _
                                  ItemType(node, sourceColumns), _
                                  ItemSize(node, sourceColumns), _
                                  -1, vbNullString, _
                                  ItemScale(node, sourceColumns))
            ' A column whose type is stated keeps it. Anything else has
            ' none until the values exist.
            computed.Add Not IsArray(StatedType(node, sourceColumns))
        End If
    Next item

    ' What each select item takes from a row, worked out once. A name is
    ' resolved by walking the columns and comparing strings, and doing that
    ' per row is that walk per row for an answer that cannot change: four
    ' named columns over twenty thousand rows cost 418ms where the same four
    ' as a star cost 254, and the star is only faster because it reads the
    ' position it already knows. Which columns a star wants is settled here
    ' too, for the same reason.
    ReDim slots(1 To mSelectItems.count + 1)
    Set starWanted = New Collection
    Set expressions = New Collection
    at = 0
    For Each item In mSelectItems
        at = at + 1
        node = item(0)
        If node(0) = N_STAR Then
            slots(at) = -1
            ReDim takes(1 To sourceColumns.count + 1)
            For position = 1 To sourceColumns.count
                takes(position) = StarWants(node, sourceColumns(position))
            Next position
            starWanted.Add takes
        Else
            If node(0) = N_COLUMN Then
                slots(at) = ColumnIndex(CStr(node(1)), sourceColumns)
            End If
            starWanted.Add Empty
        End If
        ' And an expression is resolved against the columns here too, so a
        ' computed column costs the walk along them once rather than once a
        ' row.
        expressions.Add Prepared(item(0), sourceColumns)
    Next item

    Set seen = New Collection
    For Each values In matching
        If mTopCount >= 0 And taken >= mTopCount Then Exit For

        ReDim produced(0 To outColumns.count - 1)
        index = 0
        at = 0
        For Each item In mSelectItems
            at = at + 1
            If slots(at) = -1 Then
                takes = starWanted(at)
                For position = 0 To sourceColumns.count - 1
                    If takes(position + 1) Then
                        produced(index) = values(position)
                        index = index + 1
                    End If
                Next position
            ElseIf slots(at) > 0 Then
                produced(index) = values(slots(at) - 1)
                index = index + 1
            Else
                produced(index) = Evaluate(expressions(at), values, _
                                           sourceColumns)
                index = index + 1
            End If
        Next item

        If mIsDistinct Then
            If Not RememberRow(seen, produced) Then GoTo NextRow
        End If
        outRows.Add produced
        taken = taken + 1
NextRow:
    Next values

    RetypeComputedColumns outColumns, outRows, computed
End Sub

' A row as one string, so a Collection's own keys decide what counts as the
' same row. DISTINCT and the set operators share it, because they share the
' question.
Private Function RowKey(ByRef produced As Variant) As String
    Dim index As Long

    For index = LBound(produced) To UBound(produced)
        RowKey = RowKey & Chr$(1) & CStr(NullToEmpty(produced(index)))
    Next index
End Function

' Whether this row is new, remembering it if so. A Collection keyed by the
' row's own text is the cheapest set VBA offers.
Private Function RememberRow(ByVal seen As Collection, _
                             ByRef produced As Variant) As Boolean
    On Error GoTo Duplicate
    seen.Add True, RowKey(produced)
    On Error GoTo 0
    RememberRow = True
    Exit Function

Duplicate:
End Function

' Whether this row is already there, leaving the set as it was.
Private Function RowIsIn(ByVal seen As Collection, _
                         ByRef produced As Variant) As Boolean
    Dim found As Variant

    On Error Resume Next
    found = seen(RowKey(produced))
    RowIsIn = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function HasAggregate() As Boolean
    Dim item As Variant

    For Each item In mSelectItems
        If NodeHasAggregate(item(0)) Then
            HasAggregate = True
            Exit Function
        End If
    Next item
End Function

Private Function NodeHasAggregate(ByVal node As Variant) As Boolean
    Dim argument As Variant

    If Not IsArray(node) Then Exit Function

    Select Case node(0)
        Case N_CALL
            If IsAggregateName(CStr(node(1))) Then
                NodeHasAggregate = True
                Exit Function
            End If
            For Each argument In node(2)
                If NodeHasAggregate(argument) Then
                    NodeHasAggregate = True
                    Exit Function
                End If
            Next argument
        Case N_BINARY
            NodeHasAggregate = NodeHasAggregate(node(2)) Or _
                               NodeHasAggregate(node(3))
        Case N_UNARY
            NodeHasAggregate = NodeHasAggregate(node(2))
    End Select
End Function

Private Function IsAggregateName(ByVal name As String) As Boolean
    Select Case name
        Case "COUNT", "SUM", "MIN", "MAX", "AVG", _
             "STDEV", "STDEVP", "VAR", "VARP"
            IsAggregateName = True
    End Select
End Function

' One row, computed over every row the WHERE clause kept.
Private Sub ProjectAggregates(ByVal matching As Collection, _
                              ByVal sourceColumns As Collection, _
                              ByRef outColumns As Collection, _
                              ByRef outRows As Collection)
    Dim item As Variant
    Dim node As Variant
    Dim produced As Variant
    Dim index As Long
    Dim value As Variant
    Dim stated As Variant
    Dim computed As Collection

    Set computed = New Collection
    ReDim produced(0 To mSelectItems.count - 1)
    For Each item In mSelectItems
        node = item(0)
        value = EvaluateAggregate(node, matching, sourceColumns)
        produced(index) = value

        ' A column whose type the statement stated keeps it. The rest are
        ' worked out from the one value they produced, which is all there is
        ' to go on and is wrong whenever that value is null.
        stated = StatedType(node, sourceColumns)
        If IsArray(stated) Then
            outColumns.Add ColumnOf(ItemName(item, index + 1), _
                                    CLng(stated(0)), CLng(stated(1)), _
                                    -1, vbNullString, _
                                    CLng(NullToZero(stated(2))))
            computed.Add False
        Else
            outColumns.Add ColumnOf(ItemName(item, index + 1), _
                                  TypeOfValue(value), SizeOfValue(value))
            computed.Add True
        End If
        index = index + 1
    Next item
    outRows.Add produced
    RetypeComputedColumns outColumns, outRows, computed
End Sub

' One value for a whole group.
'
' The aggregates inside the expression are folded to literals first and the
' rest is then evaluated once, which is what lets COUNT(*) > 2 and
' MAX(score) - MIN(score) mean anything: an aggregate is a value the group
' has, and everything around it is ordinary arithmetic on that value.
Private Function EvaluateAggregate(ByVal node As Variant, _
                                   ByVal matching As Collection, _
                                   ByVal columns As Collection) As Variant
    Dim resolved As Variant
    Dim representative As Variant
    Dim blank() As Variant

    resolved = ResolveAggregates(node, matching, columns)

    If matching.count > 0 Then
        representative = matching(1)
    Else
        ReDim blank(0 To 0)
        representative = blank
    End If
    EvaluateAggregate = Evaluate(resolved, representative, columns)
End Function

' The same tree with every aggregate call replaced by what it computed.
Private Function ResolveAggregates(ByVal node As Variant, _
                                   ByVal matching As Collection, _
                                   ByVal columns As Collection) As Variant
    Dim rebuilt As Collection
    Dim argument As Variant
    Dim inner As Variant

    If Not IsArray(node) Then
        ResolveAggregates = node
        Exit Function
    End If

    Select Case node(0)
        Case N_CALL
            If IsAggregateName(CStr(node(1))) Then
                If node(2).count = 0 Then
                    inner = Empty
                Else
                    inner = node(2)(1)
                End If
                ResolveAggregates = NodeOf(N_LITERAL, _
                    ComputeAggregate(CStr(node(1)), inner, matching, columns, _
                                     (node(3) = True)), _
                    Empty, Empty)
                Exit Function
            End If
            Set rebuilt = New Collection
            For Each argument In node(2)
                rebuilt.Add ResolveAggregates(argument, matching, columns)
            Next argument
            ResolveAggregates = NodeOf(N_CALL, node(1), rebuilt, node(3))

        Case N_BINARY
            If IsObject(node(3)) Then
                ' An IN list, which holds a Collection rather than a node.
                Set rebuilt = New Collection
                For Each argument In node(3)
                    rebuilt.Add ResolveAggregates(argument, matching, columns)
                Next argument
                ResolveAggregates = NodeOf(N_BINARY, node(1), _
                    ResolveAggregates(node(2), matching, columns), rebuilt)
            Else
                ResolveAggregates = NodeOf(N_BINARY, node(1), _
                    ResolveAggregates(node(2), matching, columns), _
                    ResolveAggregates(node(3), matching, columns))
            End If

        Case N_UNARY
            ResolveAggregates = NodeOf(N_UNARY, node(1), _
                ResolveAggregates(node(2), matching, columns), Empty)

        Case Else
            ResolveAggregates = node
    End Select
End Function

' What one aggregate makes of a group.
'
' Nulls are skipped rather than counted, which is why COUNT(*) and
' COUNT(column) disagree: the star counts rows and the column counts values.
Private Function ComputeAggregate(ByVal name As String, ByVal inner As Variant, _
                                  ByVal matching As Collection, _
                                  ByVal columns As Collection, _
                                  Optional ByVal distinct As Boolean = False) As Variant
    Dim values As Variant
    Dim total As Double
    Dim squares As Double
    Dim count As Long
    Dim best As Variant
    Dim current As Variant
    Dim isStar As Boolean
    Dim seen As Collection
    Dim key As String

    If IsEmpty(inner) Then
        isStar = True
    ElseIf IsArray(inner) Then
        isStar = (inner(0) = N_STAR)
    End If

    ' DISTINCT counts each value once. Keyed the way this server compares,
    ' so 2 and '2' are one value and text matches without regard to case.
    If distinct Then Set seen = New Collection

    best = Null
    For Each values In matching
        If isStar Then
            count = count + 1
        Else
            current = Evaluate(inner, values, columns)
            If Not IsNull(current) And distinct Then
                key = ValueKey(current)
                If HasKey(seen, key) Then
                    current = Null
                Else
                    seen.Add True, key
                End If
            End If
            If Not IsNull(current) Then
                count = count + 1
                If IsNumeric(current) Then
                    total = total + CDbl(current)
                    squares = squares + CDbl(current) * CDbl(current)
                End If
                If IsNull(best) Then
                    best = current
                ElseIf name = "MIN" Then
                    If CompareValues(current, best) < 0 Then best = current
                ElseIf name = "MAX" Then
                    If CompareValues(current, best) > 0 Then best = current
                End If
            End If
        End If
    Next values

    Select Case name
        Case "COUNT"
            ComputeAggregate = count
        Case "SUM"
            If count = 0 Then
                ComputeAggregate = Null
            Else
                ComputeAggregate = total
            End If
        Case "AVG"
            If count = 0 Then
                ComputeAggregate = Null
            ElseIf DeclaredWhole(inner, columns) Then
                ComputeAggregate = Fix(total / count)
            Else
                ComputeAggregate = total / count
            End If
        Case "VAR", "STDEV"
            ' Spread over a sample, which divides by one fewer than it holds:
            ' a single value has nothing to divide by and answers NULL.
            If count < 2 Then
                ComputeAggregate = Null
            Else
                ComputeAggregate = Spread(total, squares, count, count - 1, _
                                          name = "STDEV")
            End If
        Case "VARP", "STDEVP"
            ' Over the whole population, which divides by all of it.
            If count = 0 Then
                ComputeAggregate = Null
            Else
                ComputeAggregate = Spread(total, squares, count, count, _
                                          name = "STDEVP")
            End If
        Case Else
            ComputeAggregate = best
    End Select
End Function

' How far values are spread: the variance, or its square root for a standard
' deviation, from their total and the total of their squares. Rounding can
' leave a spread of values that are all the same a hair below nought, which
' no square root takes, so it is held at nought.
Private Function Spread(ByVal total As Double, ByVal squares As Double, _
                        ByVal count As Long, ByVal divisor As Long, _
                        ByVal root As Boolean) As Double
    Dim variance As Double

    variance = (squares - total * total / count) / divisor
    If variance < 0 Then variance = 0
    If root Then
        Spread = Sqr(variance)
    Else
        Spread = variance
    End If
End Function

' Whether an expression is declared a whole number, which makes its average
' one too: a real server divides an int's total by its count in whole
' numbers, so the average of 1 and 2 is 1. Asked of the declaration rather
' than the values, because a sheet hands every number over as a double.
Private Function DeclaredWhole(ByVal node As Variant, _
                               ByVal columns As Collection) As Boolean
    Dim stated As Variant

    If Not IsArray(node) Then Exit Function
    stated = StatedType(node, columns)
    If Not IsArray(stated) Then Exit Function
    DeclaredWhole = (stated(0) = TYPE_INTN Or stated(0) = TYPE_INT2)
End Function

' What a select item is called in the answer.
Private Function ItemName(ByVal item As Variant, ByVal position As Long) As String
    Dim node As Variant

    If Len(CStr(item(1))) > 0 Then
        ItemName = CStr(item(1))
        Exit Function
    End If

    node = item(0)
    If node(0) = N_COLUMN Then
        ItemName = BareColumnName(CStr(node(1)))
        ' A window with no alias has no name, like any other expression.
        If Left$(ItemName, 7) = "$window" Then ItemName = vbNullString
    Else
        ' An expression with no alias has no name, which is what SQL Server
        ' sends: an empty one rather than a made-up one.
        ItemName = vbNullString
    End If
End Function

Private Function BareColumnName(ByVal name As String) As String
    Dim parts() As String

    parts = Split(name, ".")
    BareColumnName = parts(UBound(parts))
End Function

' The type an expression states, as a two-part answer of code and size, or
' nothing when it states none.
'
' A client that writes CAST(x AS bit) has said what the column is, and a
' column read straight out of a source carries the source's type. ISNULL is
' here because SQL Server types it from its first argument and a tree reads
' several of those with a typed getter: ISNULL(c.is_not_for_replication, 0)
' is a bit on a real server, and a workbook that calls it an integer is
' unreadable rather than merely different.
'
' Everything else has no type until its values exist, and is worked out at
' projection time from what came out.
Private Function StatedType(ByVal node As Variant, _
                            ByVal columns As Collection) As Variant
    Dim out(0 To 2) As Variant
    Dim arguments As Collection
    Dim index As Long

    If Not IsArray(node) Then Exit Function

    Select Case node(0)
        Case N_LITERAL
            If VarType(node(1)) = vbDouble Or VarType(node(1)) = vbSingle Then
                ' Written with a point. A literal is the one place where
                ' how a number was written is known, and 7.0 is not 7: it
                ' divides as a real number where 7 divides as a whole one.
                out(0) = TYPE_FLTN
                out(1) = 8
                StatedType = out
                Exit Function
            End If
            If IsNull(node(1)) Then
                ' A select list that says NULL and nothing else about a
                ' column has still said what it is: SQL Server calls a bare
                ' NULL an integer, and a client reading one as text is a
                ' failed read rather than a blank.
                out(0) = TYPE_INTN
                out(1) = 4
            Else
                out(0) = TypeOfValue(node(1))
                out(1) = SizeOfValue(node(1))
            End If
            StatedType = out

        Case N_COLUMN
            index = ColumnIndex(CStr(node(1)), columns)
            If index > 0 Then
                out(0) = columns(index)(COLUMN_TYPE)
                out(1) = columns(index)(COLUMN_SIZE)
                out(2) = NullToZero(columns(index)(COLUMN_SCALE))
                StatedType = out
            Else
                StatedType = ServerVariableType(CStr(node(1)))
            End If

        Case N_CALL
            Set arguments = node(2)
            If CStr(node(1)) = "$CAST" Then
                If arguments.count >= 2 Then
                    StatedType = TypeFromName(CStr(arguments(2)(1)))
                End If
            ElseIf CStr(node(1)) = "ISNULL" Then
                If arguments.count >= 1 Then
                    StatedType = StatedType(arguments(1), columns)

                    ' A bare NULL says nothing about what the column is.
                    ' What goes in its place does, and taking the NULL's own
                    ' type sent a string down a column declared an integer.
                    If arguments.count >= 2 Then
                        If IsBareNull(arguments(1)) Then
                            StatedType = StatedType(arguments(2), columns)
                        End If
                    End If
                End If
            ElseIf CStr(node(1)) = "COUNT" _
                   Or ReturnsWholeNumber(CStr(node(1))) Then
                out(0) = TYPE_INTN
                out(1) = 4
                StatedType = out
            ElseIf CStr(node(1)) = "ERROR_MESSAGE" _
                   Or CStr(node(1)) = "ERROR_PROCEDURE" Then
                ' Text whatever it reads. Typed from its value, the NULL it
                ' is outside a CATCH was one character wide, and sqlcmd
                ' printed N.
                out(0) = TYPE_NVARCHAR
                out(1) = IIf(CStr(node(1)) = "ERROR_MESSAGE", 4000, 128)
                StatedType = out
            ElseIf CStr(node(1)) = "STDEV" Or CStr(node(1)) = "STDEVP" _
                   Or CStr(node(1)) = "VAR" Or CStr(node(1)) = "VARP" Then
                ' A spread is a float whatever it was taken of.
                out(0) = TYPE_FLTN
                out(1) = 8
                StatedType = out
            ElseIf CStr(node(1)) = "SUM" Or CStr(node(1)) = "MIN" _
                   Or CStr(node(1)) = "MAX" Or CStr(node(1)) = "AVG" Then
                ' What comes out is what went in: the sum of a float is a
                ' float, and a client asking for a database's size in pages
                ' wants the double it declared rather than the integer the
                ' total happens to be.
                If arguments.count >= 1 Then
                    StatedType = StatedType(arguments(1), columns)
                End If
            ElseIf CStr(node(1)) = "$CASE" Then
                ' A CASE is as wide as the widest thing it can produce. Its
                ' arguments are pairs of condition and result with the else
                ' value last, so the results are every second one and the
                ' last. Taking the first would call a sum of page counts an
                ' integer because one of its branches says zero.
                For index = 2 To arguments.count - 1 Step 2
                    StatedType = WiderType(StatedType, _
                                           StatedType(arguments(index), columns))
                Next index
                StatedType = WiderType(StatedType, _
                    StatedType(arguments(arguments.count), columns))
            End If

        Case N_SUBQUERY
            StatedType = SubqueryStatedType(node(1))

        Case N_BINARY
            ' Arithmetic keeps the type it works on. A sum of pages plus a
            ' sum of pages is a number, and a client that asked for one
            ' cannot read it as the text an untyped expression would be.
            Select Case CStr(node(1))
                Case "+", "-", "*", "/", "%"
                    StatedType = WiderType(StatedType(node(2), columns), _
                                           StatedType(node(3), columns))
            End Select
    End Select
End Function

' NULL written out, rather than a value that happens to be missing.
Private Function IsBareNull(ByVal node As Variant) As Boolean
    If Not IsArray(node) Then Exit Function
    If node(0) <> N_LITERAL Then Exit Function
    IsBareNull = IsNull(node(1))
End Function

' Whichever of two types can hold the other, for an expression that may
' produce either. Nothing on one side means the other side decides.
Private Function WiderType(ByVal left As Variant, _
                           ByVal right As Variant) As Variant
    If Not IsArray(left) Then
        WiderType = right
    ElseIf Not IsArray(right) Then
        WiderType = left
    ElseIf TypeRank(CLng(right(0)), CLng(right(1))) > _
           TypeRank(CLng(left(0)), CLng(left(1))) Then
        WiderType = right
    Else
        WiderType = left
    End If
End Function

' How much a type can hold. Text is widest because it takes anything, which
' is the safe answer when the branches of an expression disagree.
Private Function TypeRank(ByVal typeCode As Long, ByVal size As Long) As Long
    Select Case typeCode
        Case TYPE_BITN: TypeRank = 1
        Case TYPE_INT2: TypeRank = 3
        Case TYPE_INTN: TypeRank = 1 + size
        Case TYPE_DECIMALN: TypeRank = 20
        Case TYPE_FLTN: TypeRank = 30
        Case TYPE_DATETIMN: TypeRank = 40
        Case TYPE_GUIDN: TypeRank = 50
        Case Else: TypeRank = 60
    End Select
End Function

' The type a read used as a value produces, without running it.
'
' Its source is loaded for its columns and the first thing it selects is
' asked what it is. A tree reads several dates and numbers this way, and one
' that finds no row is otherwise a column of nulls typed as text, which the
' client then cannot read as the date it asked for.
Private Function SubqueryStatedType(ByVal tokens As Collection) As Variant
    Dim kept As Variant
    Dim columns As Collection
    Dim rows As Collection
    Dim item As Variant

    kept = SaveParseState()
    On Error GoTo Restore

    Set mTokens = tokens
    mCursor = 1
    ParseSelect

    Set columns = New Collection
    If Not mFromItems Is Nothing Then
        If mFromItems.count > 0 Then
            LoadFromItem mFromItems(1), columns, rows
        End If
    End If

    If mSelectItems.count > 0 Then
        item = mSelectItems(1)
        SubqueryStatedType = StatedType(item(0), columns)
    End If

Restore:
    RestoreParseState kept
    On Error GoTo 0
End Function

Private Function ItemType(ByVal node As Variant, _
                          ByVal columns As Collection) As Long
    Dim stated As Variant

    stated = StatedType(node, columns)
    If IsArray(stated) Then
        ItemType = CLng(stated(0))
    Else
        ItemType = TYPE_NVARCHAR
    End If
End Function

Private Function ItemSize(ByVal node As Variant, _
                          ByVal columns As Collection) As Long
    Dim stated As Variant

    stated = StatedType(node, columns)
    If IsArray(stated) Then
        ItemSize = CLng(stated(1))
    Else
        ItemSize = NVARCHAR_SIZED_LIMIT
    End If
End Function

Private Function ItemScale(ByVal node As Variant, _
                           ByVal columns As Collection) As Long
    Dim stated As Variant

    stated = StatedType(node, columns)
    If IsArray(stated) Then ItemScale = CLng(NullToZero(stated(2)))
End Function

Private Function TypeOfValue(ByVal value As Variant) As Long
    If IsNull(value) Or IsEmpty(value) Then
        TypeOfValue = TYPE_NVARCHAR
    ElseIf VarType(value) = vbBoolean Then
        TypeOfValue = TYPE_BITN
    ElseIf VarType(value) = vbDate Then
        TypeOfValue = TYPE_DATETIMN
    ElseIf IsNumeric(value) Then
        If CDbl(value) = Int(CDbl(value)) And Abs(CDbl(value)) < 2147483647# Then
            TypeOfValue = TYPE_INTN
        Else
            TypeOfValue = TYPE_FLTN
        End If
    Else
        TypeOfValue = TYPE_NVARCHAR
    End If
End Function

Private Function SizeOfValue(ByVal value As Variant) As Long
    Select Case TypeOfValue(value)
        Case TYPE_INTN
            SizeOfValue = 4
        Case TYPE_FLTN
            SizeOfValue = 8
        Case TYPE_NVARCHAR
            SizeOfValue = NVARCHAR_SIZED_LIMIT
    End Select
End Function

' Which column a name refers to, or zero when none does.
'
' A name written with a source in front of it means that source's column.
' Without one, the first column of that name wins: a real server refuses an
' ambiguous name, and refusing here would break batches that read a column
' present in two sources and mean either.
Private Function ColumnIndex(ByVal name As String, _
                             ByVal columns As Collection) As Long
    Dim bare As String
    Dim qualifier As String
    Dim index As Long
    Dim dot As Long
    Dim clean As String
    Dim column As Variant

    ' A plain name needs no taking apart. This runs for every column of
    ' every row a condition looks at, and the two Replaces and the search
    ' for a dot were three string allocations per row to arrive back at the
    ' name that was passed in.
    If InStr(1, name, ".") = 0 Then
        index = 1
        For Each column In columns
            If StrComp(CStr(column(COLUMN_NAME)), name, _
                       vbTextCompare) = 0 Then
                ColumnIndex = index
                Exit Function
            End If
            index = index + 1
        Next column
        If InStr(1, name, "[") = 0 Then Exit Function
        ' A name holding a bracket is looked for as it stands first, because
        ' the sheet may really call a column [Bracketed]. Only when no column
        ' is spelt that way are the brackets read as quoting and taken off.
    End If

    clean = Replace(Replace(name, "[", ""), "]", "")
    dot = InStrRev(clean, ".")
    If dot > 0 Then
        qualifier = BareColumnName(Left$(clean, dot - 1))
        bare = Mid$(clean, dot + 1)
    Else
        bare = clean
    End If

    ' For Each rather than an index: a Collection asked for its nth item
    ' walks to it, and a catalog view is ninety-eight columns wide.
    index = 1
    For Each column In columns
        If StrComp(CStr(column(COLUMN_NAME)), bare, vbTextCompare) = 0 Then
            If Len(qualifier) = 0 Then
                ColumnIndex = index
                Exit Function
            ElseIf StrComp(CStr(column(COLUMN_QUALIFIER)), qualifier, _
                           vbTextCompare) = 0 Then
                ColumnIndex = index
                Exit Function
            End If
        End If
        index = index + 1
    Next column
End Function


' ============================================================================
' Evaluating one row
' ============================================================================

' The same expression with every column reference it holds turned into the
' position that column sits at.
'
' A name is resolved by walking the columns and comparing strings, and an
' expression evaluated a row at a time did that walk for every row: a WHERE
' over three columns cost 1.6 seconds over fifty thousand rows, almost all
' of it looking the same three names up fifty thousand times each. Prepared
' once against the columns a loop runs over, it costs the walk once.
'
' A name the columns have not got is left as it was, so a parameter, an
' outer row's column and a name that is simply wrong all behave as before.
' What comes back is used for the loop and thrown away with it, so nothing
' else ever sees a position where it expects a name.
Private Function Prepared(ByVal node As Variant, _
                          ByVal columns As Collection) As Variant
    Dim rebuilt As Collection
    Dim argument As Variant
    Dim at As Long
    Dim outerValue As Variant
    Dim foundOuter As Boolean

    ' A part of a node can be a list rather than a node: the values of an
    ' IN, or the tokens of a subquery. Those are handed back as they are,
    ' and an object needs Set to be handed back at all.
    If IsObject(node) Then
        Set Prepared = node
        Exit Function
    End If
    If Not IsArray(node) Then
        Prepared = node
        Exit Function
    End If

    Select Case node(0)
        Case N_COLUMN
            If Left$(CStr(node(1)), 1) = "@" Then
                Prepared = node
                Exit Function
            End If
            at = ColumnIndex(CStr(node(1)), columns)
            If at > 0 Then
                Prepared = NodeOf(N_AT, at, node(1), Empty)
            Else
                ' Not one of this read's columns. It may be one of the reads
                ' this one is running inside, and that row cannot change
                ' while this loop runs, so its value is folded in here
                ' rather than looked up again for every row: a correlated
                ' subquery over three hundred rows asked for the same outer
                ' value three hundred times.
                outerValue = OuterColumnValue(CStr(node(1)), foundOuter)
                If foundOuter Then
                    Prepared = NodeOf(N_LITERAL, outerValue, Empty, Empty)
                Else
                    Prepared = node
                End If
            End If

        Case N_BINARY
            Prepared = NodeOf(N_BINARY, node(1), _
                              Prepared(node(2), columns), _
                              Prepared(node(3), columns))

        Case N_UNARY
            Prepared = NodeOf(N_UNARY, node(1), _
                              Prepared(node(2), columns), node(3))

        Case N_CALL
            Set rebuilt = New Collection
            For Each argument In node(2)
                rebuilt.Add Prepared(argument, columns)
            Next argument
            Prepared = NodeOf(N_CALL, node(1), rebuilt, node(3))

        Case Else
            ' A subquery, an EXISTS, a window or a literal: left alone.
            ' What is inside a subquery is resolved against its own columns
            ' when it runs, not against these.
            Prepared = node
    End Select
End Function

Private Function Evaluate(ByVal node As Variant, ByRef values As Variant, _
                          ByVal columns As Collection) As Variant
    Dim index As Long

    If Not IsArray(node) Then
        Evaluate = Null
        Exit Function
    End If

    Select Case node(0)
        Case N_LITERAL
            Evaluate = node(1)

        Case N_COLUMN
            ' A name opening with @ is a parameter the call carried, not a
            ' column: sp_executesql sends the statement and its values apart.
            If Left$(CStr(node(1)), 1) = "@" Then
                Evaluate = NamedValue(CStr(node(1)))
                Exit Function
            End If
            index = ColumnIndex(CStr(node(1)), columns)
            If index = 0 Then
                Dim outerFound As Boolean
                Evaluate = OuterColumnValue(CStr(node(1)), outerFound)
                If outerFound Then Exit Function
                RaiseQueryError 207, "Invalid column name '" & node(1) & "'."
            End If
            Evaluate = NormalizeValue(values(index - 1))

        Case N_AT
            ' The same value as a name would have reached, without the walk
            ' along the columns to find out which one it is.
            Evaluate = NormalizeValue(values(CLng(node(1)) - 1))

        Case N_STAR
            Evaluate = Null

        Case N_UNARY
            Evaluate = EvaluateUnary(CStr(node(1)), node(2), values, columns)

        Case N_BINARY
            Evaluate = EvaluateBinary(CStr(node(1)), node(2), node(3), _
                                      values, columns)

        Case N_CALL
            Evaluate = EvaluateCall(CStr(node(1)), node(2), values, columns)

        Case N_SUBQUERY
            Evaluate = ScalarSubquery(node(1), values, columns)

        Case N_EXISTS
            Evaluate = ExistsSubquery(node(1), values, columns)

        Case N_WINDOW
            ' Worked out over all the rows before any one of them is
            ' evaluated, and read from there. One still here is somewhere a
            ' window cannot be.
            RaiseQueryError 4108, "Windowed functions can only appear in " & _
                                  "the SELECT or ORDER BY clauses."

        Case Else
            Evaluate = Null
    End Select
End Function

Private Function NormalizeValue(ByVal value As Variant) As Variant
    If IsEmpty(value) Then
        NormalizeValue = Null
    ElseIf VarType(value) = vbString Then
        If Len(value) = 0 Then
            NormalizeValue = Null
        Else
            NormalizeValue = value
        End If
    Else
        NormalizeValue = value
    End If
End Function

Private Function EvaluateUnary(ByVal operator As String, ByVal operand As Variant, _
                               ByRef values As Variant, _
                               ByVal columns As Collection) As Variant
    Dim value As Variant

    value = Evaluate(operand, values, columns)

    Select Case operator
        Case "NOT"
            If IsNull(value) Then
                EvaluateUnary = Null
            Else
                EvaluateUnary = Not CBool(value)
            End If
        Case "-"
            If IsNull(value) Then
                EvaluateUnary = Null
            Else
                EvaluateUnary = -CDbl(value)
            End If
        Case "~"
            ' A bit's complement is the other bit. A number's is what
            ' flipping every one of its bits leaves, which is -(n + 1).
            If IsNull(value) Then
                EvaluateUnary = Null
            ElseIf VarType(value) = vbBoolean Then
                EvaluateUnary = Not CBool(value)
            Else
                EvaluateUnary = -(CLngLng(value) + 1)
            End If
        Case "IS NULL"
            EvaluateUnary = IsNull(value)
        Case "IS NOT NULL"
            EvaluateUnary = Not IsNull(value)
        Case Else
            EvaluateUnary = Null
    End Select
End Function

' One side of a simple operator, without a call into Evaluate where it is a
' prepared column or a literal, which is what most sides are.
Private Function OperandValue(ByVal node As Variant, ByRef values As Variant, _
                              ByVal columns As Collection) As Variant
    If IsArray(node) Then
        Select Case node(0)
            Case N_AT
                OperandValue = NormalizeValue(values(CLng(node(1)) - 1))
                Exit Function
            Case N_LITERAL
                OperandValue = node(1)
                Exit Function
        End Select
    End If
    OperandValue = Evaluate(node, values, columns)
End Function

Private Function EvaluateBinary(ByVal operator As String, ByVal leftNode As Variant, _
                                ByVal rightNode As Variant, ByRef values As Variant, _
                                ByVal columns As Collection) As Variant
    Dim left As Variant
    Dim right As Variant
    Dim item As Variant
    Dim found As Boolean
    Dim bounds As Variant
    Dim escape As Variant
    Dim lookup As Collection
    Dim unknown As Boolean
    Dim missed As Boolean
    Dim quantified As String

    ' A comparison first. It is the commonest thing a WHERE, a join or a
    ' sort asks for, and it walked every other operator's test on the way
    ' to its own case: a dozen string comparisons, once for every row it
    ' was asked about, before the two values had even been fetched.
    Select Case operator
        Case "=", "<>", "!=", "<", "<=", ">", ">="
            left = OperandValue(leftNode, values, columns)
            right = OperandValue(rightNode, values, columns)
            If IsNull(left) Or IsNull(right) Then
                EvaluateBinary = Null
            Else
                EvaluateBinary = ComparisonHolds(operator, left, right)
            End If
            Exit Function
    End Select

    ' AND and OR short-circuit, which matters because the other side may
    ' name a column this row has no value for.
    If operator = "AND" Or operator = "OR" Then
        left = Evaluate(leftNode, values, columns)
        If operator = "AND" Then
            If Not IsNull(left) Then
                If Not CBool(left) Then
                    EvaluateBinary = False
                    Exit Function
                End If
            End If
        Else
            If Not IsNull(left) Then
                If CBool(left) Then
                    EvaluateBinary = True
                    Exit Function
                End If
            End If
        End If
        right = Evaluate(rightNode, values, columns)

        ' An unknown on one side does not always make the answer unknown.
        ' NULL AND FALSE is FALSE and NULL OR TRUE is TRUE: the other side
        ' settles it whichever side the unknown is on, and the short-circuit
        ' above only settles it when the unknown is on the right.
        If Not IsNull(right) Then
            If operator = "AND" Then
                If Not CBool(right) Then
                    EvaluateBinary = False
                    Exit Function
                End If
            ElseIf CBool(right) Then
                EvaluateBinary = True
                Exit Function
            End If
        End If

        If IsNull(left) Or IsNull(right) Then
            EvaluateBinary = Null
        ElseIf operator = "AND" Then
            EvaluateBinary = CBool(left) And CBool(right)
        Else
            EvaluateBinary = CBool(left) Or CBool(right)
        End If
        Exit Function
    End If

    ' Mid$ throughout rather than Left$ and Right$: this procedure has
    ' locals called left and right, and VBA reads Left$ as that variable
    ' with a type character on it, which is a compile error about a
    ' type-declaration character rather than about the name.
    ' Only an operator long enough to carry one. Mid$ counts from one, so
    ' asking a one-character operator for its last four characters asks for
    ' position -2, which is a run-time error rather than an empty string,
    ' and every comparison in every statement goes through here.
    If Len(operator) > 4 Then quantified = Mid$(operator, Len(operator) - 3)
    If quantified = " ANY" Or quantified = " ALL" Then
        left = Evaluate(leftNode, values, columns)
        If IsNull(left) Then
            EvaluateBinary = Null
            Exit Function
        End If

        ' ANY holds as soon as one row does; ALL fails as soon as one does
        ' not. A row that cannot be compared makes the answer unknown, and
        ' an empty read is false for ANY and true for ALL.
        For Each item In SubqueryColumn(rightNode, values, columns)
            If IsNull(item) Then
                unknown = True
            ElseIf ComparisonHolds(Mid$(operator, 1, Len(operator) - 4), _
                                   left, item) Then
                found = True
            Else
                missed = True
            End If
        Next item

        If quantified = " ANY" Then
            If found Then
                EvaluateBinary = True
            ElseIf unknown Then
                EvaluateBinary = Null
            Else
                EvaluateBinary = False
            End If
        Else
            If missed Then
                EvaluateBinary = False
            ElseIf unknown Then
                EvaluateBinary = Null
            Else
                EvaluateBinary = True
            End If
        End If
        Exit Function
    End If

    If operator = "IN READ" Or operator = "NOT IN READ" Then
        left = Evaluate(leftNode, values, columns)
        If IsNull(left) Then
            EvaluateBinary = Null
            Exit Function
        End If

        ' Look the value up rather than walking the candidates. The
        ' candidates are the same for every row of the outer read, and
        ' walking them made the query a pass over the inner table per outer
        ' row.
        Set lookup = SubqueryLookup(rightNode, values, columns)
        If lookup Is Nothing Then
            For Each item In SubqueryColumn(rightNode, values, columns)
                If IsNull(item) Then
                    unknown = True
                ElseIf CompareValues(left, item) = 0 Then
                    found = True
                End If
            Next item
        Else
            found = HasValueKey(lookup, left)
            unknown = HasKey(lookup, NULL_CANDIDATE)
        End If
        If found Then
            EvaluateBinary = (operator = "IN READ")
        ElseIf unknown Then
            EvaluateBinary = Null
        Else
            EvaluateBinary = (operator <> "IN READ")
        End If
        Exit Function
    End If

    If operator = "LIKE ESCAPE" Or operator = "NOT LIKE ESCAPE" Then
        left = Evaluate(leftNode, values, columns)
        bounds = rightNode
        right = Evaluate(bounds(2), values, columns)
        escape = Evaluate(bounds(3), values, columns)
        If IsNull(left) Or IsNull(right) Then
            EvaluateBinary = Null
            Exit Function
        End If
        If IsNull(escape) Then escape = vbNullString
        found = MatchesLike(CStr(left), CStr(right), CStr(escape))
        EvaluateBinary = IIf(operator = "LIKE ESCAPE", found, Not found)
        Exit Function
    End If

    If operator = "IN" Or operator = "NOT IN" Then
        left = Evaluate(leftNode, values, columns)
        If IsNull(left) Then
            EvaluateBinary = Null
            Exit Function
        End If
        ' A NULL among the candidates is a candidate this might be equal
        ' to. Nothing matching is therefore not the same as no match: it is
        ' unknown, and NOT IN over a list holding one is never true.
        For Each item In rightNode
            right = Evaluate(item, values, columns)
            If IsNull(right) Then
                unknown = True
            ElseIf CompareValues(left, right) = 0 Then
                found = True
            End If
        Next item
        If found Then
            EvaluateBinary = (operator = "IN")
        ElseIf unknown Then
            EvaluateBinary = Null
        Else
            EvaluateBinary = (operator <> "IN")
        End If
        Exit Function
    End If

    If operator = "BETWEEN" Or operator = "NOT BETWEEN" Then
        left = Evaluate(leftNode, values, columns)
        bounds = rightNode
        If IsNull(left) Then
            EvaluateBinary = Null
            Exit Function
        End If
        found = (CompareValues(left, Evaluate(bounds(2), values, columns)) >= 0) And _
                (CompareValues(left, Evaluate(bounds(3), values, columns)) <= 0)
        EvaluateBinary = IIf(operator = "BETWEEN", found, Not found)
        Exit Function
    End If

    left = Evaluate(leftNode, values, columns)
    right = Evaluate(rightNode, values, columns)

    ' Anything compared with NULL is unknown, which is neither true nor
    ' false and drops the row.
    If IsNull(left) Or IsNull(right) Then
        Select Case operator
            Case "+", "-", "*", "/", "%"
                EvaluateBinary = Null
            Case Else
                EvaluateBinary = Null
        End Select
        Exit Function
    End If

    Select Case operator
        Case "=":  EvaluateBinary = (CompareValues(left, right) = 0)
        Case "<>", "!=": EvaluateBinary = (CompareValues(left, right) <> 0)
        Case "<":  EvaluateBinary = (CompareValues(left, right) < 0)
        Case "<=": EvaluateBinary = (CompareValues(left, right) <= 0)
        Case ">":  EvaluateBinary = (CompareValues(left, right) > 0)
        Case ">=": EvaluateBinary = (CompareValues(left, right) >= 0)
        Case "LIKE": EvaluateBinary = MatchesLike(CStr(left), CStr(right))
        Case "NOT LIKE": EvaluateBinary = Not MatchesLike(CStr(left), CStr(right))
        Case "+"
            If VarType(left) = vbString Or VarType(right) = vbString Then
                EvaluateBinary = CStr(left) & CStr(right)
            Else
                EvaluateBinary = CDbl(left) + CDbl(right)
            End If
        Case "-": EvaluateBinary = CDbl(left) - CDbl(right)
        Case "*": EvaluateBinary = CDbl(left) * CDbl(right)
        Case "&": EvaluateBinary = BitwiseAnd(left, right)
        Case "|": EvaluateBinary = BitwiseOr(left, right)
        Case "^": EvaluateBinary = BitwiseXor(left, right)
        Case "/"
            If CDbl(right) = 0 Then
                RaiseQueryError 8134, "Divide by zero error encountered."
            End If
            ' Two whole numbers divide to a whole number, truncated toward
            ' zero. SSMS takes @@microsoftversion apart this way, and a
            ' quotient with a fraction on it comes out as the wrong version.
            '
            ' Whole by declaration, not by what the value happens to be:
            ' 7.0 / 2 is 3.5 because of how the 7.0 was written, and a
            ' worksheet's whole numbers arrive as reals whatever they are
            ' declared, so the value alone cannot say.
            If IsWholeNumber(left) And IsWholeNumber(right) Then
                If DividesWhole(leftNode, columns) And _
                   DividesWhole(rightNode, columns) Then
                    EvaluateBinary = TruncatedDivide(CLngLng(left), CLngLng(right))
                Else
                    EvaluateBinary = CDbl(left) / CDbl(right)
                End If
            Else
                EvaluateBinary = CDbl(left) / CDbl(right)
            End If
        Case "%"
            If CDbl(right) = 0 Then
                RaiseQueryError 8134, "Divide by zero error encountered."
            End If
            ' The remainder takes the sign of what was divided, because
            ' the division truncates toward zero rather than downward:
            ' -7 % 3 is -1 and not 2. Int rounds down, so it cannot be
            ' used to work out the quotient here.
            EvaluateBinary = CDbl(left) - CDbl(right) * _
                Fix(CDbl(left) / CDbl(right))
        Case Else
            EvaluateBinary = Null
    End Select
End Function

' SQL's wildcards, translated to the ones VBA's Like understands. The
' characters VBA treats as special have to be escaped on the way through, or
' a pattern holding a bracket matches something nobody asked for.
Private Function MatchesLike(ByVal value As String, ByVal pattern As String, _
                             Optional ByVal escape As String = vbNullString) _
                             As Boolean
    Dim translated As String
    Dim index As Long
    Dim ch As String
    Dim ends As Long
    Dim inside As String

    ' Text is compared padded, and a pattern is text: 'a ' is LIKE 'a' for
    ' the same reason it equals it. Trailing spaces go from both sides,
    ' which is that comparison without having to pad either one.
    value = RTrim$(value)
    pattern = RTrim$(pattern)

    index = 1
    Do While index <= Len(pattern)
        ch = Mid$(pattern, index, 1)
        If Len(escape) > 0 And index < Len(pattern) And ch = escape Then
            ' Whatever follows the escape character stands for itself,
            ' wildcard or not, so a pattern can look for a per cent sign.
            index = index + 1
            translated = translated & LikeLiteral(Mid$(pattern, index, 1))
        ElseIf ch = "%" Then
            translated = translated & "*"
        ElseIf ch = "_" Then
            translated = translated & "?"
        ElseIf ch = "[" Then
            ' A set of characters, which VBA has too. It writes the
            ' negation with an exclamation mark where SQL writes a caret,
            ' and a bracket with nothing to close it is a bracket.
            ends = InStr(index + 1, pattern, "]")
            If ends = 0 Then
                translated = translated & "[[]"
            Else
                inside = Mid$(pattern, index + 1, ends - index - 1)
                If Left$(inside, 1) = "^" Then inside = "!" & Mid$(inside, 2)
                translated = translated & "[" & inside & "]"
                index = ends
            End If
        Else
            translated = translated & LikeLiteral(ch)
        End If
        index = index + 1
    Loop
    MatchesLike = (LCase$(value) Like LCase$(translated))
End Function

' A character that has to survive VBA's own Like, whose wildcards are not
' SQL's.
Private Function LikeLiteral(ByVal ch As String) As String
    Select Case ch
        Case "*", "?", "#": LikeLiteral = "[" & ch & "]"
        Case "[": LikeLiteral = "[[]"
        Case Else: LikeLiteral = ch
    End Select
End Function

Private Function EvaluateCall(ByVal name As String, ByVal arguments As Collection, _
                              ByRef values As Variant, _
                              ByVal columns As Collection) As Variant
    Dim first As Variant
    Dim second As Variant
    Dim index As Long
    Dim handled As Boolean

    If name = "$CASE" Then
        ' Pairs of condition and result, with the else value last.
        For index = 1 To arguments.count - 1 Step 2
            first = Evaluate(arguments(index), values, columns)
            If Not IsNull(first) Then
                If CBool(first) Then
                    EvaluateCall = Evaluate(arguments(index + 1), values, columns)
                    Exit Function
                End If
            End If
        Next index
        EvaluateCall = Evaluate(arguments(arguments.count), values, columns)
        Exit Function
    End If

    ' A client writes a built-in with the database and schema it lives in:
    ' msdb.dbo.fn_syspolicy_is_automation_enabled(). Only the last part
    ' names the function.
    name = BareFunctionName(name)

    ' The two tables of functions first: what depends on the connection,
    ' then what depends only on its arguments.
    EvaluateCall = ContextFunction(name, arguments, values, columns, handled)
    If handled Then Exit Function
    EvaluateCall = PlainFunction(name, arguments, values, columns, handled)
    If handled Then Exit Function

    If arguments.count >= 1 Then first = Evaluate(arguments(1), values, columns)
    If arguments.count >= 2 Then second = Evaluate(arguments(2), values, columns)

    If PropagatesNull(name) Then
        If IsNull(first) Or IsNull(second) Then
            EvaluateCall = Null
            Exit Function
        End If
    End If

    Select Case name
        Case "UPPER": EvaluateCall = UCase$(CStr(first))
        Case "LOWER": EvaluateCall = LCase$(CStr(first))
        Case "LEN": EvaluateCall = Len(RTrim$(CStr(first)))
        Case "LTRIM", "RTRIM", "TRIM"
            EvaluateCall = TrimCall(name, arguments, first, second, values, _
                                    columns)
        Case "ABS": EvaluateCall = Abs(CDbl(first))
        Case "FLOOR": EvaluateCall = Int(CDbl(first))
        Case "CEILING": EvaluateCall = -Int(-CDbl(first))
        Case "SQRT": EvaluateCall = Sqr(CDbl(first))
        Case "ROUND"
            If IsNull(first) Then
                EvaluateCall = Null
            ElseIf IsNull(second) Then
                EvaluateCall = CDbl(first)
            Else
                EvaluateCall = RoundHalfUp(CDbl(first), CLng(second))
            End If
        Case "LEFT": EvaluateCall = Left$(CStr(first), CLng(second))
        Case "RIGHT": EvaluateCall = Right$(CStr(first), CLng(second))
        Case "REPLACE"
            If IsNull(first) Then
                EvaluateCall = Null
            Else
                EvaluateCall = Replace(CStr(first), CStr(second), _
                                       CStr(Evaluate(arguments(3), values, columns)))
            End If
        Case "SUBSTRING"
            If IsNull(first) Then
                EvaluateCall = Null
            Else
                EvaluateCall = Substring(CStr(first), CLng(second), _
                    CLng(Evaluate(arguments(3), values, columns)))
            End If
        Case "CONCAT"
            For index = 1 To arguments.count
                EvaluateCall = CStr(EvaluateCall) & _
                               CStr(NullToEmpty(Evaluate(arguments(index), values, columns)))
            Next index
        Case "ISNULL", "COALESCE"
            For index = 1 To arguments.count
                first = Evaluate(arguments(index), values, columns)
                If Not IsNull(first) Then
                    EvaluateCall = first
                    Exit Function
                End If
            Next index
            EvaluateCall = Null
        Case "$CAST"
            EvaluateCall = CastValue(first, CStr(Evaluate(arguments(2), _
                                                         values, columns)))
        Case "QUOTENAME"
            EvaluateCall = QuoteName(first, second, arguments.count)
        Case "COLLATIONNAME"
            ' The name a binary collation id stands for. A linked server
            ' asks about the one the login response announced, and there is
            ' only the one here.
            If IsNull(first) Then
                EvaluateCall = Null
            Else
                EvaluateCall = "SQL_Latin1_General_CP1_CI_AS"
            End If
        Case "@@VERSION"
            EvaluateCall = "Microsoft SQL Server " & SERVER_VERSION_MAJOR & "." & _
                           SERVER_VERSION_MINOR & "." & SERVER_VERSION_BUILD & _
                           " (vbaSQLBridge)"
        Case Else
            RaiseQueryError 195, "'" & name & "' is not a recognized built-in " & _
                                 "function name."
    End Select
End Function

' The functions that answer NULL when any argument is NULL.
'
' The test has to happen before the case runs. VBA works out a call's
' arguments before the call, so a guard written as
'
'     NullOr(first, UCase$(CStr(first)))
'
' has already run CStr over the Null by the time the guard is reached, and
' CStr of a Null raises. It reads exactly like a guard and never was one:
' twenty-five functions were written that way and every one of them answered
' a NULL argument with "Invalid use of Null".
' LTRIM, RTRIM and TRIM: spaces off one end or both, or with a set of
' characters as the second argument, those. TRIM's third argument, from how
' it was written, says which end.
Private Function TrimCall(ByVal name As String, ByVal arguments As Collection, _
                          ByVal first As Variant, ByVal second As Variant, _
                          ByRef values As Variant, _
                          ByVal columns As Collection) As Variant
    Dim chars As String
    Dim side As String

    chars = " "
    If arguments.count >= 2 Then chars = CStr(second)
    Select Case name
        Case "LTRIM"
            side = "LEADING"
        Case "RTRIM"
            side = "TRAILING"
        Case Else
            side = "BOTH"
            If arguments.count >= 3 Then
                side = CStr(Evaluate(arguments(3), values, columns))
            End If
    End Select
    TrimCall = TrimSet(CStr(first), chars, side)
End Function

' The characters of a set taken off the start, the end or both, matched the
' way this server matches text.
Private Function TrimSet(ByVal text As String, ByVal chars As String, _
                         ByVal side As String) As String
    Dim start As Long
    Dim finish As Long

    start = 1
    finish = Len(text)
    If side <> "TRAILING" Then
        Do While start <= finish
            If InStr(1, chars, Mid$(text, start, 1), vbTextCompare) = 0 Then
                Exit Do
            End If
            start = start + 1
        Loop
    End If
    If side <> "LEADING" Then
        Do While finish >= start
            If InStr(1, chars, Mid$(text, finish, 1), vbTextCompare) = 0 Then
                Exit Do
            End If
            finish = finish - 1
        Loop
    End If
    TrimSet = Mid$(text, start, finish - start + 1)
End Function

Private Function PropagatesNull(ByVal name As String) As Boolean
    Select Case name
        Case "UPPER", "LOWER", "LEN", "LTRIM", "RTRIM", "TRIM", _
             "ABS", "FLOOR", "CEILING", "SQRT", "LEFT", "RIGHT", _
             "REVERSE", "POWER", "SQUARE", "EXP", "LOG", "LOG10", _
             "SPACE", "STR", "CHAR", "NCHAR", "YEAR", "MONTH", "DAY"
            PropagatesNull = True
    End Select
End Function

' Whether one comparison holds between two values that are both there.
'
' The quantified forms put the same operator to every row a read answered,
' so the operator is decided once here rather than written out again.
Private Function ComparisonHolds(ByVal operator As String, _
                                 ByVal left As Variant, _
                                 ByVal right As Variant) As Boolean
    Select Case operator
        Case "=":  ComparisonHolds = (CompareValues(left, right) = 0)
        Case "<>", "!=": ComparisonHolds = (CompareValues(left, right) <> 0)
        Case "<":  ComparisonHolds = (CompareValues(left, right) < 0)
        Case "<=": ComparisonHolds = (CompareValues(left, right) <= 0)
        Case ">":  ComparisonHolds = (CompareValues(left, right) > 0)
        Case ">=": ComparisonHolds = (CompareValues(left, right) >= 0)
    End Select
End Function

' Whether a value's kind is a number, as VarType reports it. Booleans and
' dates are not among them: each compares by a rule of its own.
Private Function IsNumberKind(ByVal kind As Integer) As Boolean
    Select Case kind
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, _
             vbByte, vbLongLong
            IsNumberKind = True
    End Select
End Function

' -1, 0 or 1, comparing numbers as numbers and everything else as text.
Private Function CompareValues(ByVal left As Variant, ByVal right As Variant) As Long
    Dim a As Double
    Dim b As Double
    Dim leftKind As Integer
    Dim rightKind As Integer

    ' Each side's kind once rather than six times over. Every row of every
    ' WHERE, every pair a join tries and every comparison a sort makes comes
    ' through here, so the tests are ordered by how often they answer.
    leftKind = VarType(left)
    rightKind = VarType(right)

    ' Two numbers, which most comparisons are.
    If IsNumberKind(leftKind) And IsNumberKind(rightKind) Then
        a = CDbl(left)
        b = CDbl(right)
        If a < b Then
            CompareValues = -1
        ElseIf a > b Then
            CompareValues = 1
        End If
        Exit Function
    End If

    If leftKind = vbNull And rightKind = vbNull Then Exit Function
    If leftKind = vbNull Then
        CompareValues = -1
        Exit Function
    End If
    If rightKind = vbNull Then
        CompareValues = 1
        Exit Function
    End If

    If leftKind = vbBoolean Or rightKind = vbBoolean Then
        a = IIf(CBool(left), 1, 0)
        b = IIf(CBool(right), 1, 0)
    ElseIf leftKind = vbString And rightKind = vbString Then
        ' Two strings compare as strings, whatever they spell. '10' is less
        ' than '9' the way 'a' is less than 'b', and a column of numbers
        ' typed as text is a column of text: reading both sides as numbers
        ' because both could be read that way sorted it as neither.
        CompareValues = StrComp(RTrim$(CStr(left)), RTrim$(CStr(right)), _
                                vbTextCompare)
        Exit Function
    ElseIf leftKind = vbDate Or rightKind = vbDate Then
        ' A moment against anything is compared as a moment, the way a real
        ' server brings the other side up to datetime. A date is not a
        ' number to IsNumeric, so two dates were compared as the text they
        ' print as, and the first of October came before the ninth of
        ' September; against '2024-9-1' a sheet's date was text as well.
        a = MomentOf(left)
        b = MomentOf(right)
    ElseIf IsNumeric(left) And IsNumeric(right) Then
        a = CDbl(left)
        b = CDbl(right)
    Else
        ' Text compares padded: the shorter side is made up to the longer
        ' with spaces, so 'a' and 'a ' are the same value and a trailing
        ' space never decides anything. Trimming the ends is the same
        ' comparison and does not need the longer length to do it.
        CompareValues = StrComp(RTrim$(CStr(left)), RTrim$(CStr(right)), _
                                vbTextCompare)
        Exit Function
    End If

    If a < b Then
        CompareValues = -1
    ElseIf a > b Then
        CompareValues = 1
    End If
End Function

' A value as the moment it names, in days, to compare with a date. Text is
' read the way a real server reads it into a datetime: the ISO forms with the
' T or without it, a fraction of a second, and yyyymmdd with no separators,
' which VBA would take for a number. A number counts days from 1900-01-01,
' where VBA counts from 1899-12-30, so it lands two days on.
Private Function MomentOf(ByVal value As Variant) As Double
    Dim text As String
    Dim fraction As Double
    Dim dot As Long

    If VarType(value) = vbDate Then
        MomentOf = CDbl(value)
    ElseIf VarType(value) = vbString Then
        text = Trim$(CStr(value))
        If text Like "########" Then
            text = Left$(text, 4) & "-" & Mid$(text, 5, 2) & "-" & Right$(text, 2)
        ElseIf Len(text) > 10 And Mid$(text, 11, 1) = "T" Then
            Mid$(text, 11, 1) = " "
        End If
        dot = InStrRev(text, ".")
        If dot > 0 And dot > InStrRev(text, ":") And InStr(text, ":") > 0 Then
            fraction = Val("0" & Mid$(text, dot))
            text = Left$(text, dot - 1)
        End If
        If Not IsDate(text) Then
            RaiseQueryError 241, "Conversion failed when converting date " & _
                                 "and/or time from character string."
        End If
        MomentOf = CDbl(CDate(text)) + fraction / 86400#
    ElseIf IsNumeric(value) Then
        MomentOf = CDbl(value) + 2
    Else
        MomentOf = CDbl(CDate(value))
    End If
End Function

' ORDER BY, as a merge sort over the rows.
'
' A merge sort rather than anything simpler because the order has to be
' stable: rows that compare equal on the named columns keep the order the
' sheet had them in, which is the only order left to keep.
Private Function SortRows(ByVal rows As Collection, _
                          ByVal columns As Collection) As Collection
    Dim indexes() As Long
    Dim buffer() As Long
    Dim items() As Variant
    Dim sorted As Collection
    Dim index As Long
    Dim values As Variant
    Dim prepared As Variant
    Dim keys As Variant
    Dim terms As Long
    Dim at As Long
    Dim column As Long

    If rows.count < 2 Then
        Set SortRows = rows
        Exit Function
    End If

    ReDim indexes(0 To rows.count - 1)
    ReDim buffer(0 To rows.count - 1)
    ReDim items(0 To rows.count - 1)

    ' For Each, not an index: a Collection asked for its nth item walks to
    ' it, so numbering the rows one at a time reads the whole list once per
    ' row.
    index = 0
    For Each values In rows
        indexes(index) = index
        items(index) = values
        index = index + 1
    Next values

    ' The value each row sorts on, worked out once. Reading it inside the
    ' comparison read it twice per comparison, and twenty thousand rows sort
    ' in a quarter of a million comparisons: half a million conversions of
    ' twenty thousand values.
    prepared = PreparedOrder(columns)
    terms = UBound(prepared, 1) - LBound(prepared, 1) + 1
    ReDim keys(0 To rows.count - 1, 0 To terms - 1)
    For at = 0 To terms - 1
        column = CLng(prepared(at, 2))
        If column > 0 Then
            For index = 0 To rows.count - 1
                values = items(index)
                keys(index, at) = _
                    NormalizeValue(values(LBound(values) + column - 1))
            Next index
        End If
    Next at

    MergeSort indexes, buffer, 0, rows.count - 1, items, keys, columns, _
              prepared

    Set sorted = New Collection
    For index = 0 To rows.count - 1
        sorted.Add items(indexes(index))
    Next index
    Set SortRows = sorted
End Function

' What each ORDER BY term is, worked out once.
'
' A term resolves to the same thing for every row, and resolving it inside
' the comparison resolved it n log n times: twenty thousand rows sort in a
' quarter of a million comparisons, each of which looked a name up in the
' select list and then looked it up again in the columns. Where a term is a
' plain column its position is kept as well, and the comparison reads the
' value straight out of the row.
Private Function PreparedOrder(ByVal columns As Collection) As Variant
    Dim prepared() As Variant
    Dim term As Variant
    Dim node As Variant
    Dim at As Long

    ReDim prepared(0 To mOrderBy.count - 1, 0 To 2)
    For Each term In mOrderBy
        node = OrderTerm(term(0), columns)
        prepared(at, 0) = node
        prepared(at, 1) = term(1)
        prepared(at, 2) = 0
        If IsArray(node) Then
            If node(0) = N_COLUMN Then
                prepared(at, 2) = ColumnIndex(CStr(node(1)), columns)
            End If
        End If
        at = at + 1
    Next term
    PreparedOrder = prepared
End Function

Private Sub MergeSort(ByRef indexes() As Long, ByRef buffer() As Long, _
                      ByVal low As Long, ByVal high As Long, _
                      ByRef items() As Variant, ByRef keys As Variant, _
                      ByVal columns As Collection, _
                      ByRef prepared As Variant)
    Dim middle As Long
    Dim left As Long
    Dim right As Long
    Dim at As Long

    If low >= high Then Exit Sub
    middle = (low + high) \ 2
    MergeSort indexes, buffer, low, middle, items, keys, columns, prepared
    MergeSort indexes, buffer, middle + 1, high, items, keys, columns, prepared

    left = low
    right = middle + 1
    at = low
    Do While left <= middle And right <= high
        If CompareRows(indexes(right), indexes(left), items, keys, _
                       columns, prepared) < 0 Then
            buffer(at) = indexes(right)
            right = right + 1
        Else
            buffer(at) = indexes(left)
            left = left + 1
        End If
        at = at + 1
    Loop
    Do While left <= middle
        buffer(at) = indexes(left)
        left = left + 1
        at = at + 1
    Loop
    Do While right <= high
        buffer(at) = indexes(right)
        right = right + 1
        at = at + 1
    Loop
    For at = low To high
        indexes(at) = buffer(at)
    Next at
End Sub

Private Function CompareRows(ByVal leftAt As Long, ByVal rightAt As Long, _
                             ByRef items() As Variant, ByRef keys As Variant, _
                             ByVal columns As Collection, _
                             ByRef prepared As Variant) As Long
    Dim at As Long
    Dim outcome As Long

    For at = LBound(prepared, 1) To UBound(prepared, 1)
        If prepared(at, 2) > 0 Then
            outcome = CompareValues(keys(leftAt, at), keys(rightAt, at))
        Else
            outcome = CompareValues( _
                Evaluate(prepared(at, 0), items(leftAt), columns), _
                Evaluate(prepared(at, 0), items(rightAt), columns))
        End If
        If prepared(at, 1) Then outcome = -outcome
        If outcome <> 0 Then
            CompareRows = outcome
            Exit Function
        End If
    Next at
End Function


' ============================================================================
' RPC
'
' Clients do not send everything as a batch. Anything with a parameter, and
' every catalog query .NET issues, arrives as a call to sp_executesql:
'
'     proc: well-known id 10, ExecuteSql
'     param 1 ''      nvarchar  the statement
'     param 2 ''      nvarchar  '@id int'
'     param 3 '@id'   intn      1
'
' Reading a value requires knowing its type, because the value does not
' describe itself: the width and the length prefix both come from the type
' declaration in front of it. An unsupported type therefore stops the parse
' rather than being skipped, since there is no way to know how many bytes it
' occupied.
' ============================================================================

' The statement an RPC carries, with its named parameters collected. A call
' to anything but sp_executesql answers with the procedure's name, because
' which procedures exist is a question about the catalog rather than about
' the protocol.
Private Function ParseRpc(ByRef data() As Byte, _
                          ByVal parameters As Collection, _
                          ByRef names As String, _
                          ByRef calledProcedure As String) As String
    Dim available As Long
    Dim at As Long
    Dim nameLength As Long
    Dim procedureId As Long
    Dim procedure As String
    Dim nameChars As Long
    Dim parameterName As String
    Dim typeId As Long
    Dim value As Variant
    Dim ordinal As Long
    Dim positional As Long
    Dim statement As String
    Dim declared As Collection
    Dim status As Long
    Dim declaredSize As Long

    available = ByteCount(data)
    Set mOutputs = New Collection

    ' The ALL_HEADERS block arrived in TDS 7.2. A 7.1 client starts at the
    ' procedure name, and reading its name length as a header length
    ' produces a seven-megabyte block inside a hundred-byte packet.
    If Not UnsignedLess(mTdsVersion, TDS_VERSION_72) Then
        If available < 4 Then
            Err.Raise ERR_PROTOCOL, "SqlBridge.ParseRpc", _
                      "an RPC of " & available & " bytes is too short for an " & _
                      "ALL_HEADERS length"
        End If
        at = ReadUInt32LE(data, 0)
        If at < 4 Or at > available Then
            Err.Raise ERR_PROTOCOL, "SqlBridge.ParseRpc", _
                      "the RPC declares a " & at & "-byte ALL_HEADERS block, " & _
                      "which does not fit in " & available & " bytes"
        End If
    End If

    nameLength = ReadUInt16LE(data, at)
    at = at + 2
    If nameLength = 65535 Then
        procedureId = ReadUInt16LE(data, at)
        at = at + 2
        procedure = WellKnownProcedure(procedureId)
    Else
        procedure = BytesToText(data, at, nameLength * 2)
        at = at + nameLength * 2
    End If

    at = at + 2                                  ' option flags

    Do While at < available
        nameChars = data(at)
        at = at + 1
        parameterName = BytesToText(data, at, nameChars * 2)
        at = at + nameChars * 2
        status = data(at)                        ' 1 for an OUTPUT parameter
        at = at + 1
        typeId = data(at)
        at = at + 1
        declaredSize = RpcDeclaredSize(data, at, typeId)
        value = ReadRpcValue(data, at, typeId)

        ordinal = ordinal + 1
        If Len(parameterName) > 0 Then
            Bind parameters, names, parameterName, value, typeId
            If (status And 1) <> 0 Then
                KeepOutput ordinal - 1, parameterName, parameterName, _
                           typeId, declaredSize
            End If
        ElseIf (nameLength <> 65535 And _
                BareProcedureName(procedure) <> "sp_executesql") Or _
               procedure = "sp_unprepare" Or _
               Left$(procedure, 9) = "sp_cursor" Then
            ' A procedure called by name takes its arguments in order. Only
            ' sp_executesql reads the first as a statement and the second as
            ' declarations, and reading every call that way turned a catalog
            ' procedure's table name into the statement: the procedure was
            ' never called, and "people" was run as a batch in its place.
            ' sp_unprepare and the cursor procedures come by number and take
            ' theirs in order too.
            Bind parameters, names, "#" & ordinal, value, typeId
            If procedure = "sp_unprepare" Then ForgetPrepared value
        ElseIf procedure = "sp_execute" Then
            ' sp_execute(@handle, values...): the handle stands for a
            ' statement prepared earlier on this connection, and the values
            ' after it are for that statement's declarations, in order.
            If ordinal = 1 Then
                FindPrepared value, statement, declared
            Else
                BindPositional parameters, names, declared, positional, _
                               value, typeId
            End If
        ElseIf (procedure = "sp_prepexec" Or procedure = "sp_prepare") And _
               ordinal <= 3 Then
            ' sp_prepexec(@handle OUTPUT, @params, @stmt, values...): the
            ' handle, the declarations, the statement, and only then the
            ' values. Read the way sp_executesql is, the handle was the
            ' statement and a batch called 0 was run in its place. sp_prepare
            ' takes the same three and then its options, with no values.
            If ordinal = 2 Then
                Set declared = DeclarationNames(CStr(NullToEmpty(value)))
            ElseIf ordinal = 3 Then
                statement = CStr(NullToEmpty(value))
            End If
        ElseIf procedure = "sp_prepare" Then
            ' @options, which asks for nothing answered differently here.
        ElseIf ordinal = 1 Then
            ' sp_executesql puts the statement in the first parameter, which
            ' is unnamed, and the declarations in the second.
            statement = CStr(NullToEmpty(value))
        ElseIf ordinal = 2 Then
            Set declared = DeclarationNames(CStr(NullToEmpty(value)))
        Else
            ' Measured against MSOLEDBSQL: it sends the values after the
            ' declarations with no names on them at all, so the only thing
            ' that says which is which is their order against the
            ' declaration string. A client that does name them is answered
            ' by the branch above.
            BindPositional parameters, names, declared, positional, _
                           value, typeId
            If (status And 1) <> 0 And Not declared Is Nothing Then
                If positional <= declared.count Then
                    KeepOutput ordinal - 1, "", CStr(declared(positional)), _
                               typeId, declaredSize
                End If
            End If
        End If
    Loop

    Select Case procedure
        Case "sp_prepexec", "sp_prepare", "sp_execute"
            ' A prepared statement is run like any other and answered as
            ' the call it came in, so which it was is kept, and the handle
            ' it goes by: a new one for a statement being prepared, the one
            ' it was prepared under for a statement being run again.
            calledProcedure = procedure
            ParseRpc = statement
            If procedure <> "sp_execute" Then
                mCallHandle = RememberPrepared(statement, declared)
            End If
        Case Else
            If Len(statement) > 0 Then
                ParseRpc = statement
            Else
                calledProcedure = procedure
                ParseRpc = "EXEC " & procedure
            End If
    End Select
End Function

' An OUTPUT parameter, kept to be handed back once the statement has run:
' where it sat in the call, the name it was sent under, the name it is
' bound to, and the type it goes back as.
Private Sub KeepOutput(ByVal ordinal As Long, ByVal sentName As String, _
                       ByVal boundName As String, ByVal typeId As Long, _
                       ByVal declaredSize As Long)
    Dim typeCode As Long
    Dim size As Long

    OutputType typeId, declaredSize, typeCode, size
    mOutputs.Add Array(ordinal, sentName, LCase$(boundName), typeCode, size)
End Sub

Private Function HasOutputs() As Boolean
    If mOutputs Is Nothing Then Exit Function
    HasOutputs = (mOutputs.count > 0)
End Function

' The type an OUTPUT parameter goes back as: the nullable form of the type
' it was declared, since an output can always come back NULL. A number with
' places goes back as a float, and what this does not carry as itself goes
' back as text.
Private Sub OutputType(ByVal typeId As Long, ByVal declaredSize As Long, _
                       ByRef typeCode As Long, ByRef size As Long)
    Select Case typeId
        Case &H26                               ' intn, its width declared
            typeCode = TYPE_INTN: size = declaredSize
        Case &H30
            typeCode = TYPE_INTN: size = 1
        Case &H34
            typeCode = TYPE_INTN: size = 2
        Case &H38
            typeCode = TYPE_INTN: size = 4
        Case &H7F
            typeCode = TYPE_INTN: size = 8
        Case &H32, &H68
            typeCode = TYPE_BITN: size = 1
        Case &H3A, &H3D, &H6F
            typeCode = TYPE_DATETIMN: size = 8
        Case &H3B, &H3E, &H6D, &H6A, &H6C, &H3C, &H6E, &H7A
            typeCode = TYPE_FLTN: size = 8
        Case &HE7, &HEF
            typeCode = TYPE_NVARCHAR: size = declaredSize \ 2
        Case &HA7, &HAF
            typeCode = TYPE_NVARCHAR: size = declaredSize
        Case Else
            typeCode = TYPE_NVARCHAR: size = NVARCHAR_SIZED_LIMIT
    End Select
    If typeCode = TYPE_INTN Then
        Select Case size
            Case 1, 2, 4, 8
            Case Else
                size = 4
        End Select
    End If
    If typeCode = TYPE_NVARCHAR Then
        If size <= 0 Or size > NVARCHAR_SIZED_LIMIT Then
            size = NVARCHAR_SIZED_LIMIT
        End If
    End If
End Sub

' The size a parameter's type declares, read without moving past it: one
' byte for the nullable numbers, two for text and binary, and the width of
' a fixed type.
Private Function RpcDeclaredSize(ByRef data() As Byte, ByVal at As Long, _
                                 ByVal typeId As Long) As Long
    Select Case typeId
        Case &H24, &H26, &H68, &H6D, &H6E, &H6F, &H6A, &H6C
            RpcDeclaredSize = data(at)
        Case &HA7, &HAF, &HE7, &HEF, &HA5, &HAD
            RpcDeclaredSize = ReadUInt16LE(data, at)
        Case Else
            RpcDeclaredSize = FixedWidthOf(typeId)
    End Select
End Function

' A value sent without a name, bound to the declaration at its position.
Private Sub BindPositional(ByVal parameters As Collection, _
                           ByRef names As String, _
                           ByVal declared As Collection, _
                           ByRef positional As Long, _
                           ByVal value As Variant, ByVal typeId As Long)
    positional = positional + 1
    If declared Is Nothing Then Exit Sub
    If positional <= declared.count Then
        Bind parameters, names, CStr(declared(positional)), value, typeId
    End If
End Sub

' A statement prepared on this connection, kept with its declarations under
' the handle its client is given for it. sp_prepare and sp_prepexec hand one
' out, sp_execute runs what one stands for again with new values, and
' sp_unprepare lets it go.
Private Function RememberPrepared(ByVal statement As String, _
                                  ByVal declared As Collection) As Long
    Dim entry(0 To 1) As Variant

    If mPrepared Is Nothing Then Set mPrepared = New Collection
    mLastHandle = mLastHandle + 1
    entry(0) = statement
    Set entry(1) = declared
    mPrepared.Add entry, "h" & mLastHandle
    RememberPrepared = mLastHandle
End Function

' What a handle was prepared as. A handle this connection never gave out
' leaves the statement empty, and the call is refused when it is answered,
' where a refusal is an answer rather than a dropped connection.
Private Sub FindPrepared(ByVal handle As Variant, ByRef statement As String, _
                         ByRef declared As Collection)
    Dim entry As Variant

    mCallHandle = CLng(NullToZero(handle))
    If mPrepared Is Nothing Then Exit Sub
    On Error Resume Next
    entry = mPrepared("h" & mCallHandle)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    statement = CStr(entry(0))
    Set declared = entry(1)
End Sub

Private Sub ForgetPrepared(ByVal handle As Variant)
    If mPrepared Is Nothing Then Exit Sub
    On Error Resume Next
    mPrepared.Remove "h" & CLng(NullToZero(handle))
    On Error GoTo 0
End Sub

Private Function WellKnownProcedure(ByVal procedureId As Long) As String
    Select Case procedureId
        Case 1: WellKnownProcedure = "sp_cursor"
        Case 2: WellKnownProcedure = "sp_cursoropen"
        Case 3: WellKnownProcedure = "sp_cursorprepare"
        Case 4: WellKnownProcedure = "sp_cursorexecute"
        Case 5: WellKnownProcedure = "sp_cursorprepexec"
        Case 6: WellKnownProcedure = "sp_cursorunprepare"
        Case 7: WellKnownProcedure = "sp_cursorfetch"
        Case 8: WellKnownProcedure = "sp_cursoroption"
        Case 9: WellKnownProcedure = "sp_cursorclose"
        Case 10: WellKnownProcedure = "sp_executesql"
        Case 11: WellKnownProcedure = "sp_prepare"
        Case 12: WellKnownProcedure = "sp_execute"
        Case 13: WellKnownProcedure = "sp_prepexec"
        Case 14: WellKnownProcedure = "sp_prepexecrpc"
        Case 15: WellKnownProcedure = "sp_unprepare"
        Case Else
            Err.Raise ERR_PROTOCOL, "SqlBridge.WellKnownProcedure", _
                      "unknown well-known procedure id " & procedureId
    End Select
End Function

' One parameter's declaration and value, advancing the cursor past both.
Private Function ReadRpcValue(ByRef data() As Byte, ByRef at As Long, _
                              ByVal typeId As Long) As Variant
    Dim width As Long
    Dim length As Long
    Dim declared As Long

    width = FixedWidthOf(typeId)
    If width >= 0 Then
        ReadRpcValue = DecodeNumber(data, at, width, typeId)
        at = at + width
        Exit Function
    End If

    Select Case typeId
        Case &H24, &H26, &H68, &H6D, &H6E, &H6F      ' one length byte each
            at = at + 1                              ' the declared maximum
            length = data(at)
            at = at + 1
            If length = 0 Then
                ReadRpcValue = Null
            Else
                ReadRpcValue = DecodeNumber(data, at, length, typeId)
            End If
            at = at + length

        Case &H6A, &H6C                              ' decimal and numeric
            at = at + 3                              ' maximum, precision, scale
            length = data(at)
            at = at + 1
            If length = 0 Then
                ReadRpcValue = Null
            Else
                ReadRpcValue = DecodeDecimal(data, at, length)
            End If
            at = at + length

        Case &HA7, &HAF, &HE7, &HEF, &HA5, &HAD      ' char and binary
            declared = ReadUInt16LE(data, at)
            at = at + 2
            If typeId = &HA7 Or typeId = &HAF Or typeId = &HE7 Or typeId = &HEF Then
                at = at + 5                          ' collation
            End If
            If declared = 65535 Then
                ReadRpcValue = ReadChunked(data, at, typeId)
            Else
                length = ReadUInt16LE(data, at)
                at = at + 2
                If length = 65535 Then
                    ReadRpcValue = Null
                Else
                    ReadRpcValue = DecodeText(data, at, length, typeId)
                    at = at + length
                End If
            End If

        Case &H23, &H63, &H22                        ' text, ntext and image
            ' Deprecated for twenty years and still sent: SSMS passes a
            ' filter to a catalog query as ntext. A four-byte maximum, a
            ' collation for the two text ones, then a four-byte length. A row
            ' sent the other way puts a pointer and a timestamp in front of
            ' that length; a parameter, which is a value rather than a place,
            ' does not.
            at = at + 4
            If typeId = &H23 Or typeId = &H63 Then at = at + 5
            length = ReadUInt32LE(data, at)
            at = at + 4
            If length = -1 Then
                ReadRpcValue = Null
            Else
                ReadRpcValue = DecodeText(data, at, length, _
                                          IIf(typeId = &H63, &HE7, &HA7))
                at = at + length
            End If

        Case Else
            Err.Raise ERR_PROTOCOL, "SqlBridge.ReadRpcValue", _
                      "an RPC parameter has type 0x" & Hex$(typeId) & _
                      ", which this cannot read; its length is unknown, so " & _
                      "the rest of the call cannot be parsed either"
    End Select
End Function

' The width a type implies, or -1 when the value carries its own length.
Private Function FixedWidthOf(ByVal typeId As Long) As Long
    Select Case typeId
        Case &H1F: FixedWidthOf = 0                  ' NULL
        Case &H30: FixedWidthOf = 1                  ' INT1
        Case &H32: FixedWidthOf = 1                  ' BIT
        Case &H34: FixedWidthOf = 2                  ' INT2
        Case &H38: FixedWidthOf = 4                  ' INT4
        Case &H3A: FixedWidthOf = 4                  ' DATETIME4
        Case &H3B: FixedWidthOf = 4                  ' FLT4
        Case &H3C: FixedWidthOf = 8                  ' MONEY
        Case &H3D: FixedWidthOf = 8                  ' DATETIME
        Case &H3E: FixedWidthOf = 8                  ' FLT8
        Case &H7A: FixedWidthOf = 4                  ' MONEY4
        Case &H7F: FixedWidthOf = 8                  ' INT8
        Case Else: FixedWidthOf = -1
    End Select
End Function

' The chunked form the MAX types use: a total length, then length-prefixed
' chunks, then a zero.
Private Function ReadChunked(ByRef data() As Byte, ByRef at As Long, _
                             ByVal typeId As Long) As Variant
    Dim total As String
    Dim size As Long
    Dim high As Long
    Dim low As Long

    low = ReadUInt32LE(data, at)
    high = ReadUInt32LE(data, at + 4)
    at = at + 8
    If low = -1 And high = -1 Then
        ReadChunked = Null
        Exit Function
    End If

    Do
        size = ReadUInt32LE(data, at)
        at = at + 4
        If size = 0 Then Exit Do
        total = total & DecodeText(data, at, size, typeId)
        at = at + size
    Loop
    ReadChunked = total
End Function

Private Function DecodeText(ByRef data() As Byte, ByVal at As Long, _
                            ByVal length As Long, ByVal typeId As Long) As Variant
    Dim raw() As Byte
    Dim index As Long
    Dim out As String

    If length <= 0 Then
        DecodeText = vbNullString
        Exit Function
    End If

    If typeId = &HE7 Or typeId = &HEF Then
        DecodeText = BytesToText(data, at, length)
        Exit Function
    End If

    ' The single-byte types. Read as the code page the client's collation
    ' names would be better; every value seen here has been ASCII.
    raw = MidBytes(data, at, length)
    For index = 0 To UBound(raw)
        out = out & Chr$(raw(index))
    Next index
    DecodeText = out
End Function

Private Function DecodeNumber(ByRef data() As Byte, ByVal at As Long, _
                              ByVal width As Long, ByVal typeId As Long) As Variant
    Dim raw() As Byte
    Dim value As LongLong
    Dim index As Long
    Dim asDouble As Double
    Dim asSingle As Single

    If width = 0 Then
        DecodeNumber = Null
        Exit Function
    End If
    raw = MidBytes(data, at, width)

    Select Case typeId
        Case &H32, &H68                              ' bit
            DecodeNumber = (raw(0) <> 0)
            Exit Function
        Case &H3B                                    ' float, four bytes
            CopyMemory asSingle, raw(0), 4^
            DecodeNumber = CDbl(asSingle)
            Exit Function
        Case &H3E                                    ' float, eight bytes
            CopyMemory asDouble, raw(0), 8^
            DecodeNumber = asDouble
            Exit Function
        Case &H6D                                    ' FLTN, either width
            If width = 4 Then
                CopyMemory asSingle, raw(0), 4^
                DecodeNumber = CDbl(asSingle)
            Else
                CopyMemory asDouble, raw(0), 8^
                DecodeNumber = asDouble
            End If
            Exit Function
        Case &H24                                    ' a guid stays bytes
            DecodeNumber = ToHex(raw)
            Exit Function
        Case &H3A, &H3D, &H6F                        ' datetime
            DecodeNumber = DecodeDateTime(raw, width)
            Exit Function
    End Select

    ' Everything else is a little-endian signed integer. INT1 is the one
    ' exception: tinyint is unsigned, and 200 is 200 rather than -56.
    For index = UBound(raw) To 0 Step -1
        value = value * 256^ + raw(index)
    Next index
    If width > 1 And raw(UBound(raw)) >= 128 Then
        value = value - (256^ ^ width)
    End If
    DecodeNumber = value
End Function

Private Function DecodeDateTime(ByRef raw() As Byte, ByVal width As Long) As Variant
    Dim days As Long
    Dim ticks As Long
    Dim index As Long

    If width = 4 Then
        days = CLng(raw(0)) + CLng(raw(1)) * &H100
        ticks = (CLng(raw(2)) + CLng(raw(3)) * &H100) * 60
        DecodeDateTime = CDate(days + 2 + ticks / 86400#)
        Exit Function
    End If

    days = ReadUInt32LE(raw, 0)
    For index = 7 To 4 Step -1
        ticks = ticks * 256 + raw(index)
    Next index
    DecodeDateTime = CDate(days + 2 + (ticks / 300#) / 86400#)
End Function

' Decimal and numeric: a sign byte, then the value little-endian, scaled by
' the scale that came with the declaration. The scale is not read back here,
' so this answers the unscaled magnitude and its sign.
Private Function DecodeDecimal(ByRef data() As Byte, ByVal at As Long, _
                               ByVal length As Long) As Variant
    Dim raw() As Byte
    Dim value As Double
    Dim index As Long

    raw = MidBytes(data, at, length)
    For index = UBound(raw) To 1 Step -1
        value = value * 256# + raw(index)
    Next index
    If raw(0) = 0 Then value = -value
    DecodeDecimal = value
End Function

Private Function BytesToText(ByRef data() As Byte, ByVal at As Long, _
                             ByVal length As Long) As String
    Dim raw() As Byte

    If length <= 0 Then Exit Function
    raw = MidBytes(data, at, length)
    BytesToText = raw
End Function


' What a name opening with @ stands for: a parameter the call carried, a
' variable the batch declared, or one of the server's own.
'
' A missing one is not an error: a client that declares a parameter and never
' sends it means it to be null.
Private Function NamedValue(ByVal name As String) As Variant
    Dim value As Variant
    Dim key As String

    key = LCase$(name)

    If Left$(key, 2) = "@@" Then
        NamedValue = ServerVariable(key)
        Exit Function
    End If

    ' Resume Next rather than a handler apiece: an error raised while a
    ' handler is already running is not trapped by the next one, so a name
    ' that is neither a parameter nor a variable escaped as a failure.
    On Error Resume Next

    value = mParameters(key)
    If Err.Number = 0 Then
        On Error GoTo 0
        NamedValue = value
        Exit Function
    End If

    Err.Clear
    value = mVariables(key)
    If Err.Number = 0 Then
        On Error GoTo 0
        NamedValue = value
        Exit Function
    End If

    Err.Clear
    On Error GoTo 0
    NamedValue = Null
End Function


' What a type byte is called, for a log line that has to be read by somebody
' working out why a driver's parameter did not arrive.
Private Function TypeNameOfId(ByVal typeId As Long) As String
    Select Case typeId
        Case &H1F: TypeNameOfId = "null"
        Case &H22: TypeNameOfId = "image"
        Case &H23: TypeNameOfId = "text"
        Case &H24: TypeNameOfId = "uniqueidentifier"
        Case &H26: TypeNameOfId = "intn"
        Case &H30: TypeNameOfId = "tinyint"
        Case &H32: TypeNameOfId = "bit"
        Case &H34: TypeNameOfId = "smallint"
        Case &H38: TypeNameOfId = "int"
        Case &H3A: TypeNameOfId = "smalldatetime"
        Case &H3B: TypeNameOfId = "real"
        Case &H3C: TypeNameOfId = "money"
        Case &H3D: TypeNameOfId = "datetime"
        Case &H3E: TypeNameOfId = "float"
        Case &H63: TypeNameOfId = "ntext"
        Case &H68: TypeNameOfId = "bitn"
        Case &H6A: TypeNameOfId = "decimal"
        Case &H6C: TypeNameOfId = "numeric"
        Case &H6D: TypeNameOfId = "floatn"
        Case &H6E: TypeNameOfId = "moneyn"
        Case &H6F: TypeNameOfId = "datetimn"
        Case &H7A: TypeNameOfId = "smallmoney"
        Case &H7F: TypeNameOfId = "bigint"
        Case &HA5: TypeNameOfId = "varbinary"
        Case &HA7: TypeNameOfId = "varchar"
        Case &HAD: TypeNameOfId = "binary"
        Case &HAF: TypeNameOfId = "char"
        Case &HE7: TypeNameOfId = "nvarchar"
        Case &HEF: TypeNameOfId = "nchar"
        Case Else: TypeNameOfId = "type 0x" & Hex$(typeId)
    End Select
End Function


' Remember one parameter under the name the statement will use for it.
Private Sub Bind(ByVal parameters As Collection, ByRef names As String, _
                 ByVal name As String, ByVal value As Variant, _
                 ByVal typeId As Long)
    If Len(names) > 0 Then names = names & ", "
    names = names & name & " " & TypeNameOfId(typeId)

    On Error Resume Next
    parameters.Remove LCase$(name)
    On Error GoTo 0
    parameters.Add value, LCase$(name)
End Sub

' The names out of a declaration string such as "@P1 float,@P2 nvarchar(50)".
'
' Split at the commas outside parentheses, because decimal(18,2) has one
' inside and splitting on every comma turns one parameter into two.
Private Function DeclarationNames(ByVal text As String) As Collection
    Dim names As Collection
    Dim depth As Long
    Dim start As Long
    Dim index As Long
    Dim ch As String

    Set names = New Collection
    start = 1
    For index = 1 To Len(text) + 1
        If index > Len(text) Then
            ch = ","
        Else
            ch = Mid$(text, index, 1)
        End If
        If ch = "(" Then
            depth = depth + 1
        ElseIf ch = ")" Then
            depth = depth - 1
        ElseIf ch = "," And depth = 0 Then
            AddDeclaredName names, Mid$(text, start, index - start)
            start = index + 1
        End If
    Next index
    Set DeclarationNames = names
End Function

Private Sub AddDeclaredName(ByVal names As Collection, ByVal part As String)
    Dim trimmed As String
    Dim space As Long

    trimmed = Trim$(part)
    If Len(trimmed) = 0 Then Exit Sub

    space = InStr(trimmed, " ")
    If space > 0 Then trimmed = Left$(trimmed, space - 1)
    names.Add trimmed
End Sub


' ============================================================================
' Catalog procedures
'
' A client asks what exists before it asks for anything in particular, and it
' does not ask in SQL. MSOLEDBSQL calls [master].[sys].sp_tables_rowset2 and
' sp_columns_100_rowset2 by name; the ODBC drivers call sp_tables and
' sp_columns; both families describe the same catalog in different layouts.
'
' Every layout here was read off SQL Server 2025 by running the procedure and
' taking its result metadata rather than from documentation, by way of
' pySQLbridge. sp_tables_rowset2 ends with a TABLE_PROPID and two datetimes
' that no summary of it mentions, and its TABLE_GUID is a uniqueidentifier: a
' provider handed nvarchar where it expects one of those does not report a
' mismatch, it fails.
' ============================================================================

' What a client's procedure call gets back.
Friend Function AnswerProcedure(ByVal procedure As String, _
                                ByVal arguments As Collection, _
                                ByVal tdsVersion As Long, _
                                Optional ByVal asRpc As Boolean = False, _
                                Optional ByVal variables As Collection = Nothing, _
                                Optional ByVal tempTables As Collection = Nothing, _
                                Optional ByVal session As Collection = Nothing) _
                                As Byte()
    Dim columns As Collection
    Dim rows As Collection
    Dim encoded() As Byte

    mQueryErrorNumber = 0
    Set mCallOutputs = Nothing

    ' The calling connection's own state, for a procedure that writes: an
    ' INSERT through a linked server's cursor goes inside that connection's
    ' transaction. Called from a batch, the batch's is in place already.
    If Not variables Is Nothing Then Set mVariables = variables
    If Not tempTables Is Nothing Then Set mTempTables = tempTables
    If Not session Is Nothing Then Set mSession = session

    ' The cursor procedures read their arguments and answer with what they
    ' did, which a rowset cannot carry.
    Select Case BareProcedureName(procedure)
        Case "sp_cursoropen", "sp_cursor", "sp_cursorclose"
            AnswerProcedure = AnswerCursor(procedure, arguments, tdsVersion)
            Exit Function
    End Select

    Set columns = New Collection
    Set rows = New Collection
    ProcedureResult procedure, columns, rows
    Set rows = RestrictedToNamed(procedure, arguments, columns, rows)
    encoded = ResultSet(columns, rows, tdsVersion)
    If asRpc Then
        encoded = ClosedAsCall(encoded, RowCountOf(rows), _
                               columns.count > 0, tdsVersion)
    End If
    AnswerProcedure = encoded
End Function

' The rows of a catalog rowset about the one table a call names, when it
' names one.
'
' A linked server asks each rowset about the table it is about to use, by
' name, and takes more than one row back as more than one table of that
' name: Msg 7315, "contains multiple tables that match the name". With one
' table served that never showed. Measured against a real server, the name
' comes first in every rowset procedure without a 2 on the end, and the ones
' with it take no table name at all.
Private Function RestrictedToNamed(ByVal procedure As String, _
                                   ByVal arguments As Collection, _
                                   ByVal columns As Collection, _
                                   ByVal rows As Collection) As Collection
    Dim bare As String
    Dim wanted As Variant
    Dim at As Long
    Dim index As Long
    Dim row As Variant
    Dim kept As Collection

    Set RestrictedToNamed = rows
    bare = BareProcedureName(procedure)
    If Right$(bare, 7) <> "_rowset" And Right$(bare, 10) <> "_rowset_64" Then
        Exit Function
    End If
    wanted = CallArgument(arguments, "@table_name")
    If IsNull(wanted) Then wanted = CallArgument(arguments, "#1")
    If IsNull(wanted) Then Exit Function
    If Len(CStr(wanted)) = 0 Then Exit Function

    For index = 1 To columns.count
        If UCase$(CStr(columns(index)(COLUMN_NAME))) = "TABLE_NAME" Then
            at = index
        End If
    Next index
    If at = 0 Then Exit Function

    Set kept = New Collection
    For Each row In rows
        If StrComp(CStr(NullToEmpty(row(LBound(row) + at - 1))), _
                   CStr(wanted), vbTextCompare) = 0 Then kept.Add row
    Next row
    Set RestrictedToNamed = kept
End Function

' A prepared statement, run: sp_prepexec preparing and running it in one
' call, or sp_execute running one prepared earlier by its handle. The
' connection keeps what each handle stands for. sp_prepexec hands its new
' handle back after the rows, which is where a real server puts it.
Friend Function AnswerPrepared(ByVal statement As String, _
                               ByVal handle As Long, _
                               ByVal handingBack As Boolean, _
                               ByVal tdsVersion As Long, _
                               ByVal parameters As Collection, _
                               ByVal variables As Collection, _
                               ByVal tempTables As Collection, _
                               ByVal session As Collection) As Byte()
    Dim encoded() As Byte

    If Not handingBack And Len(statement) = 0 Then
        RaiseQueryError 8179, "Could not find prepared statement with " & _
                              "handle " & handle & "."
    End If
    encoded = AnswerQuery(statement, tdsVersion, parameters, variables, _
                          tempTables, session)
    Set mCallOutputs = Nothing
    If handingBack Then
        Set mCallOutputs = New Collection
        mCallOutputs.Add CallOutput(0, TYPE_INTN, 4, handle)
    End If
    AnswerPrepared = ClosedAsCall(encoded, mRowCount, True, tdsVersion)
End Function

' sp_executesql with OUTPUT parameters, answered as the call it came in as:
' the rows, then each OUTPUT parameter as the statement left it, which is
' where a real server hands them back. SqlClient asks for the transaction
' count this way after a transaction request fails, and casting the value
' it never got to an int threw a NullReferenceException in the client, in
' place of the error the server had sent.
Friend Function AnswerWithOutputs(ByVal statement As String, _
                                  ByVal tdsVersion As Long, _
                                  ByVal parameters As Collection, _
                                  ByVal outputs As Collection, _
                                  ByVal variables As Collection, _
                                  ByVal tempTables As Collection, _
                                  ByVal session As Collection) As Byte()
    Dim encoded() As Byte
    Dim wanted As Variant
    Dim value As Variant

    encoded = AnswerQuery(statement, tdsVersion, parameters, variables, _
                          tempTables, session)
    Set mCallOutputs = New Collection
    For Each wanted In outputs
        value = Null
        If HasKey(parameters, CStr(wanted(2))) Then
            value = parameters(CStr(wanted(2)))
        End If
        If CLng(wanted(3)) = TYPE_DATETIMN And Not IsNull(value) _
           And Not IsEmpty(value) Then
            value = CDate(MomentOf(value))
        End If
        mCallOutputs.Add CallOutput(CLng(wanted(0)), CLng(wanted(3)), _
                                    CLng(wanted(4)), value, CStr(wanted(1)))
    Next wanted
    AnswerWithOutputs = ClosedAsCall(encoded, mRowCount, True, tdsVersion)
End Function

' sp_prepare: the columns a statement would answer with, without its rows,
' and the handle it is kept under.
'
' A linked server asks this of every OPENQUERY before it runs it, and takes
' the pass-through query's columns from nothing else: answered without
' them, the real server reports that the object has no columns. A read is
' run to find them and its rows are dropped. Anything else is described as
' answering nothing, because running a write to describe it would make the
' write.
Friend Function AnswerDescribed(ByVal statement As String, _
                                ByVal handle As Long, _
                                ByVal tdsVersion As Long, _
                                ByVal parameters As Collection, _
                                ByVal variables As Collection, _
                                ByVal tempTables As Collection, _
                                ByVal session As Collection) As Byte()
    Dim encoded() As Byte

    If OnlyReads(statement) Then
        encoded = AnswerQuery(statement, tdsVersion, parameters, variables, _
                              tempTables, session, True)
    Else
        encoded = TokenDone(DONE_FINAL, 0, tdsVersion)
    End If
    Set mCallOutputs = New Collection
    mCallOutputs.Add CallOutput(0, TYPE_INTN, 4, handle)
    AnswerDescribed = ClosedAsCall(encoded, 0, _
        encoded(LBound(encoded)) = TOKEN_COL_METADATA, tdsVersion)
End Function

' Whether every statement in a batch only reads, which is what may be run to
' find out what the batch would answer with.
Private Function OnlyReads(ByVal sql As String) As Boolean
    Dim written As String
    Dim piece As Variant
    Dim ends As String

    written = TrimWhitespace(StripComments(sql))
    Do While Right$(written, 1) = ";"
        written = TrimWhitespace(Left$(written, Len(written) - 1))
    Loop
    If Len(written) = 0 Then Exit Function

    For Each piece In SplitStatements(written)
        Select Case UCase$(WordAt(CStr(piece), SkipSpaceAt(CStr(piece), 1), _
                                  ends))
            Case "SELECT", "WITH"
            Case Else
                Exit Function
        End Select
    Next piece
    OnlyReads = True
End Function

' The server-side cursor calls a linked server makes an INSERT with.
'
' Measured against a real server: sp_cursoropen over "select * from" the
' table, answered with the table's columns and a hidden ROWSTAT after them,
' no rows, and the cursor's handle, scroll and concurrency options and row
' count as output parameters; then sp_cursor once per row, operation 4, with
' the row's values named after their columns; then sp_cursorclose. An UPDATE
' or a DELETE through a linked server is pushed down as a statement instead,
' so no other operation is answered.
Private Function AnswerCursor(ByVal procedure As String, _
                              ByVal arguments As Collection, _
                              ByVal tdsVersion As Long) As Byte()
    Const CURSOR_INSERT As Long = 4
    Dim name As String
    Dim entry As Variant
    Dim columns As Collection
    Dim column As Variant
    Dim rows As Collection
    Dim values() As Variant
    Dim index As Long
    Dim handle As Long
    Dim operation As Long
    Dim out() As Byte
    Dim used As Long
    Dim part() As Byte
    Dim encoded() As Byte
    Dim copied As Variant

    Select Case BareProcedureName(procedure)
        Case "sp_cursoropen"
            name = CursorTable(CStr(NullToEmpty(CallArgument(arguments, "#2"))))
            RefreshTables
            entry = LookUpTable(name)

            ' The table's own columns, flagged as a real server flags them.
            ' With the computed bit a select list's columns carry, each was
            ' refused as a column nobody may write: Msg 7344, "did not have
            ' permission to write to the column".
            Set columns = New Collection
            For Each column In entry(2)
                copied = column
                copied(COLUMN_FLAGS) = TABLE_COLUMN_FLAGS
                columns.Add copied
            Next column
            columns.Add ColumnOf("ROWSTAT", TYPE_INTN, 4, 0)

            part = ColMetadata(columns, tdsVersion)
            AppendBytes out, used, part
            part = CursorTabName(CStr(entry(0)))
            AppendBytes out, used, part
            part = CursorColInfo(columns.count)
            AppendBytes out, used, part
            part = TokenDoneCounted(DONE_COUNT, 0, tdsVersion)
            AppendBytes out, used, part

            If mCursors Is Nothing Then Set mCursors = New Collection
            mLastCursor = mLastCursor + 1
            mCursors.Add name, "c" & mLastCursor

            Set mCallOutputs = New Collection
            mCallOutputs.Add CallOutput(0, TYPE_INTN, 4, mLastCursor)
            mCallOutputs.Add CallOutput(2, TYPE_INTN, 4, _
                CursorOption(CallArgument(arguments, "#3"), 2))
            mCallOutputs.Add CallOutput(3, TYPE_INTN, 4, _
                CursorOption(CallArgument(arguments, "#4"), 4))
            mCallOutputs.Add CallOutput(4, TYPE_INTN, 4, -1)
            encoded = LeftBytes(out, used)
            AnswerCursor = ClosedAsCall(encoded, 0, True, tdsVersion)

        Case "sp_cursor"
            handle = CLng(NullToZero(CallArgument(arguments, "#1")))
            operation = CLng(NullToZero(CallArgument(arguments, "#2")))
            If operation <> CURSOR_INSERT Then
                RaiseQueryError ERR_NUMBER_BASE, _
                    "sp_cursor operation " & operation & " is not " & _
                    "supported: only an insert is written through a cursor."
            End If
            name = OpenCursorTable(handle, CallArgument(arguments, "#4"))
            entry = LookUpTable(name)
            RefuseUnwritable entry, name, True

            ' The values come named after the columns they are for, so a
            ' column the INSERT left out is simply not among them.
            ReDim values(0 To entry(2).count - 1)
            index = 0
            For Each column In entry(2)
                values(index) = CallArgument(arguments, _
                                             "@" & CStr(column(COLUMN_NAME)))
                index = index + 1
            Next column
            Set rows = New Collection
            rows.Add values
            AppendRowsTo name, rows

            encoded = TokenDoneCounted(DONE_COUNT Or DONE_FINAL, 1, _
                                       tdsVersion, INSERT_COMMAND)
            AnswerCursor = ClosedAsCall(encoded, 1, True, tdsVersion)

        Case "sp_cursorclose"
            handle = CLng(NullToZero(CallArgument(arguments, "#1")))
            If Not mCursors Is Nothing Then
                On Error Resume Next
                mCursors.Remove "c" & handle
                On Error GoTo 0
            End If
            encoded = TokenDone(DONE_FINAL, 0, tdsVersion)
            AnswerCursor = ClosedAsCall(encoded, 0, False, tdsVersion)
    End Select
End Function

' An argument a call was given, by position ("#2") or by name ("@id"), or
' Null where it was not given one.
Private Function CallArgument(ByVal arguments As Collection, _
                              ByVal key As String) As Variant
    CallArgument = Null
    If arguments Is Nothing Then Exit Function
    On Error Resume Next
    CallArgument = arguments(LCase$(key))
    On Error GoTo 0
End Function

' The table a cursor is opened over, out of the statement it is opened with.
Private Function CursorTable(ByVal statement As String) As String
    Set mTokens = Tokenize(statement)
    mCursor = 1
    Do
        If mCursor > mTokens.count Then
            RaiseQueryError ERR_NUMBER_BASE, _
                "A cursor is answered here only over a table, not over: " & _
                statement
        End If
        If Accept("FROM") Then Exit Do
        mCursor = mCursor + 1
    Loop
    CursorTable = TakeText()
End Function

' The table an open cursor was opened over, or the one the call names itself
' when the handle is not one given out here.
Private Function OpenCursorTable(ByVal handle As Long, _
                                 ByVal named As Variant) As String
    If Not mCursors Is Nothing Then
        On Error Resume Next
        OpenCursorTable = mCursors("c" & handle)
        On Error GoTo 0
    End If
    If Len(OpenCursorTable) = 0 Then OpenCursorTable = CStr(NullToEmpty(named))
    If Len(OpenCursorTable) = 0 Then
        RaiseQueryError ERR_NUMBER_BASE, "The cursor " & handle & " is not open."
    End If
End Function

' A cursor's scroll or concurrency option as the real server hands it back:
' the kind asked for, without the flags saying which others would do. To a
' linked server's request that is 2, dynamic, and 4, optimistic.
Private Function CursorOption(ByVal requested As Variant, _
                              ByVal fallback As Long) As Long
    CursorOption = CLng(NullToZero(requested)) And &HFFF&
    If CursorOption = 0 Then CursorOption = fallback
End Function

' TABNAME for a cursor over one table: its database, schema and name.
Private Function CursorTabName(ByVal table As String) As Byte()
    Dim body() As Byte
    Dim bodyUsed As Long
    Dim out() As Byte
    Dim used As Long
    Dim part As Variant
    Dim text() As Byte

    AppendByte body, bodyUsed, 3
    For Each part In Array(ConfiguredDatabase, "dbo", table)
        AppendUInt16LE body, bodyUsed, Len(CStr(part))
        text = Utf16Bytes(CStr(part))
        AppendBytes body, bodyUsed, text
    Next part
    body = LeftBytes(body, bodyUsed)

    AppendByte out, used, TOKEN_TABNAME
    AppendUInt16LE out, used, bodyUsed
    AppendBytes out, used, body
    CursorTabName = LeftBytes(out, used)
End Function

' COLINFO for the same: every column read from that one table, and the last,
' ROWSTAT, hidden and computed rather than read.
Private Function CursorColInfo(ByVal count As Long) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim index As Long

    AppendByte out, used, TOKEN_COLINFO
    AppendUInt16LE out, used, count * 3
    For index = 1 To count - 1
        AppendByte out, used, index
        AppendByte out, used, 1
        AppendByte out, used, 0
    Next index
    AppendByte out, used, count
    AppendByte out, used, 0
    AppendByte out, used, &H14
    CursorColInfo = LeftBytes(out, used)
End Function

' The same answer, closed the way a real server closes a procedure called by
' RPC rather than the way it closes a batch.
'
' The DONE a result set ends with is the last thing in it and always the
' same width, so it is taken off and the three closing tokens put on.
Private Function ClosedAsCall(ByRef encoded() As Byte, ByVal rowCount As Long, _
                              ByVal hasRows As Boolean, _
                              ByVal tdsVersion As Long) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim doneWidth As Long
    Dim doneAt As Long
    Dim command As Long
    Dim returned As Variant
    Dim part() As Byte

    doneWidth = 9
    If Not UnsignedLess(tdsVersion, WIDENED_IN) Then doneWidth = 13
    doneAt = ByteCount(encoded) - doneWidth
    out = LeftBytes(encoded, doneAt)
    used = ByteCount(out)

    ' A write closes on its own command rather than a SELECT's. That is how
    ' a real server closes an UPDATE a linked server pushes down to it:
    ' 0xC5, and the number of rows it changed.
    command = ReadUInt16LE(encoded, doneAt + 3)
    Select Case command
        Case INSERT_COMMAND, UPDATE_COMMAND, DELETE_COMMAND, MERGE_COMMAND
        Case Else
            command = SELECT_COMMAND
    End Select

    If hasRows Then
        AppendByte out, used, TOKEN_DONEINPROC
        AppendUInt16LE out, used, DONE_MORE Or DONE_COUNT
        AppendUInt16LE out, used, command
        AppendUInt32LE out, used, rowCount
        If doneWidth = 13 Then AppendUInt32LE out, used, 0
    End If

    AppendByte out, used, TOKEN_RETURNSTATUS
    AppendUInt32LE out, used, 0

    ' Output parameters go between the return status and the end of the
    ' call, which is where a real server puts them.
    If Not mCallOutputs Is Nothing Then
        For Each returned In mCallOutputs
            part = TokenReturnValue(returned, tdsVersion)
            AppendBytes out, used, part
        Next returned
    End If

    AppendByte out, used, TOKEN_DONEPROC
    AppendUInt16LE out, used, DONE_FINAL
    AppendUInt16LE out, used, RPC_COMMAND
    AppendUInt32LE out, used, 0
    If doneWidth = 13 Then AppendUInt32LE out, used, 0

    ClosedAsCall = LeftBytes(out, used)
End Function

' One output parameter a call hands back: where it sat, what it is, its
' value, and the name it was sent under when it was sent with one.
Private Function CallOutput(ByVal ordinal As Long, ByVal typeCode As Long, _
                            ByVal size As Long, ByVal value As Variant, _
                            Optional ByVal name As String = "") As Variant
    Dim out(0 To 4) As Variant

    out(0) = ordinal
    out(1) = typeCode
    out(2) = size
    out(3) = value
    out(4) = name
    CallOutput = out
End Function

' RETURNVALUE, laid out as a real server lays it out: ordinal, the name the
' parameter was sent under or none, the output status, a user type, flags,
' the type, then the value the way a row would carry it. SqlClient finds
' its parameter by the name, and one handed back nameless stayed null.
Private Function TokenReturnValue(ByVal returned As Variant, _
                                  ByVal tdsVersion As Long) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim name() As Byte
    Dim info() As Byte

    AppendByte out, used, TOKEN_RETURNVALUE
    AppendUInt16LE out, used, CLng(returned(0))
    AppendByte out, used, Len(CStr(returned(4)))
    If Len(CStr(returned(4))) > 0 Then
        name = Utf16Bytes(CStr(returned(4)))
        AppendBytes out, used, name
    End If
    AppendByte out, used, 1                   ' an output parameter
    If UnsignedLess(tdsVersion, WIDENED_IN) Then
        AppendUInt16LE out, used, 0           ' user type
    Else
        AppendUInt32LE out, used, 0
    End If
    AppendUInt16LE out, used, 0               ' flags
    info = TypeInfo(CLng(returned(1)), CLng(returned(2)), 0)
    AppendBytes out, used, info
    AppendEncodedValue out, used, CLng(returned(1)), CLng(returned(2)), _
                       returned(3)
    TokenReturnValue = LeftBytes(out, used)
End Function

' What a procedure answers, as columns and rows.
'
' Separate from the encoding because a batch inserts a procedure's rows into
' a temp table: SSMS fills one from xp_msver while connecting and reads the
' version back out of it.
Friend Sub ProcedureResult(ByVal procedure As String, _
                           ByRef outColumns As Collection, _
                           ByRef outRows As Collection)
    Dim bare As String

    bare = BareProcedureName(procedure)

    Select Case bare
        Case "sp_tables_rowset", "sp_tables_rowset2"
            TablesRowset outColumns, outRows

        Case "sp_tables_info_rowset", "sp_tables_info_rowset2"
            TablesInfoRowset False, outColumns, outRows

        Case "sp_tables_info_90_rowset", "sp_tables_info_90_rowset2", _
             "sp_tables_info_90_rowset_64"
            TablesInfoRowset True, outColumns, outRows

        Case "sp_columns_rowset", "sp_columns_rowset2", _
             "sp_columns_90_rowset", "sp_columns_90_rowset2", _
             "sp_columns_100_rowset", "sp_columns_100_rowset2"
            ColumnsRowset outColumns, outRows

        Case "sp_schemata_rowset"
            SchemataRowset outColumns, outRows

        Case "sp_catalogs_rowset"
            CatalogsRowset outColumns, outRows

        Case "sp_databases"
            DatabasesRowset outColumns, outRows

        Case "sp_tables"
            OdbcTables outColumns, outRows

        Case "sp_columns", "sp_columns_90", "sp_columns_100"
            OdbcColumns outColumns, outRows

        Case "xp_msver"
            ServerVersionRowset outColumns, outRows

        Case "sp_reset_connection"
            ' How a pooled connection is handed to its next user. Nothing
            ' here outlives a batch, so there is nothing to reset, and a
            ' real server answers it with no rows.

        Case "sp_indexes_rowset", "sp_indexes_90_rowset", _
             "sp_indexes_100_rowset", "sp_indexes_rowset2", _
             "sp_indexes_90_rowset2", "sp_indexes_100_rowset2"
            IndexesRowset outColumns, outRows

        Case "sp_check_constbytable_rowset", "sp_check_constbytable_rowset2"
            CheckConstraintsRowset outColumns, outRows

        Case "sp_table_statistics2_rowset", "sp_table_statistics_rowset"
            TableStatisticsRowset outColumns, outRows

        Case "sp_getschemalock"
            ' A linked server locks a table's schema before it reads and
            ' takes the lock's handle and the schema's version back as
            ' output parameters. Nothing here changes a schema under a
            ' reader, so the lock is only its handle.
            Set mCallOutputs = New Collection
            mCallOutputs.Add CallOutput(0, TYPE_INTN, 4, 1)
            mCallOutputs.Add CallOutput(1, TYPE_INTN, 8, 1^)

        Case "sp_oledb_ro_usrname"
            ' Whether the database is read-only, and who the caller is in
            ' it. MSOLEDBSQL asks as it connects when its service components
            ' are off, and a login it could not find this for was refused.
            ReadOnlyAndUserRowset outColumns, outRows

        Case "sp_releaseschemalock", "sp_unprepare"
            ' Letting go of the lock and of a prepared statement. Neither
            ' was held, and a real server answers both with no rows.

        Case Else
            ' Named rather than answered with an empty rowset. A provider
            ' handed the wrong shape fails in a way that names nothing, and
            ' a procedure this does not serve should say so.
            RaiseQueryError 2812, "Could not find stored procedure '" & _
                                  procedure & "'."
    End Select
End Sub

' The name without its database, schema and brackets.
Private Function BareProcedureName(ByVal procedure As String) As String
    Dim parts() As String
    Dim bare As String

    bare = Replace(Replace(procedure, "[", ""), "]", "")
    parts = Split(bare, ".")
    BareProcedureName = LCase$(Trim$(parts(UBound(parts))))
End Function

' sp_oledb_ro_usrname, which a real server defines as
' SELECT substring('NY', is_read_only + 1, 1), user_name(): two columns with
' no names, one character and a user, answered here from ReadOnly and from
' the same user USER_NAME() answers with.
Private Sub ReadOnlyAndUserRowset(ByRef outColumns As Collection, _
                                  ByRef outRows As Collection)
    Dim row(0 To 1) As Variant

    Set outColumns = New Collection
    outColumns.Add CatalogColumn(vbNullString, TYPE_NVARCHAR, 1)
    outColumns.Add CatalogColumn(vbNullString, TYPE_NVARCHAR, 128)
    row(0) = IIf(mReadOnly, "Y", "N")
    row(1) = "dbo"
    Set outRows = New Collection
    outRows.Add row
End Sub

' -- the OLE DB family ----------------------------------------------------

Private Sub TablesRowset(ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim entry As Variant
    Dim row As Variant

    Set columns = New Collection
    columns.Add CatalogColumn("TABLE_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_TYPE", TYPE_NVARCHAR, 30)
    columns.Add CatalogColumn("TABLE_GUID", TYPE_GUIDN, 16)
    columns.Add CatalogColumn("DESCRIPTION", TYPE_NVARCHAR, 1)
    columns.Add CatalogColumn("TABLE_PROPID", TYPE_INTN, 4)
    columns.Add CatalogColumn("DATE_CREATED", TYPE_DATETIMN, 8)
    columns.Add CatalogColumn("DATE_MODIFIED", TYPE_DATETIMN, 8)

    Set rows = New Collection
    For Each entry In mTables
        ReDim row(0 To 8)
        row(0) = ConfiguredDatabase
        row(1) = "dbo"
        row(2) = entry(0)
        row(3) = IIf(IsViewEntry(entry), "VIEW", "TABLE")
        row(4) = Null
        row(5) = Null
        row(6) = Null
        row(7) = Null
        row(8) = Null
        rows.Add row
    Next entry

    Set outColumns = columns
    Set outRows = rows
End Sub

' The same tables with the bookmark and cardinality columns.
'
' The 90 spellings carry a TABLE_FLAGS column the others do not, and a
' provider that asked for fifteen columns and got fourteen reports a
' catastrophic failure rather than a mismatch.
Private Sub TablesInfoRowset(ByVal withFlags As Boolean, _
                             ByRef outColumns As Collection, _
                             ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim entry As Variant
    Dim row As Variant
    Dim width As Long
    Dim index As Long

    Set columns = New Collection
    columns.Add CatalogColumn("TABLE_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_TYPE", TYPE_NVARCHAR, 30)
    columns.Add CatalogColumn("TABLE_GUID", TYPE_GUIDN, 16)
    columns.Add CatalogColumn("BOOKMARKS", TYPE_BITN, 1)
    columns.Add CatalogColumn("BOOKMARK_TYPE", TYPE_INTN, 4)
    columns.Add CatalogColumn("BOOKMARK_DATATYPE", TYPE_INTN, 2)
    columns.Add CatalogColumn("BOOKMARK_MAXIMUM_LENGTH", TYPE_INTN, 4)
    columns.Add CatalogColumn("BOOKMARK_INFORMATION", TYPE_INTN, 4)
    columns.Add CatalogColumn("TABLE_VERSION", TYPE_INTN, 8)
    columns.Add CatalogColumn("CARDINALITY", TYPE_INTN, 4)
    columns.Add CatalogColumn("DESCRIPTION", TYPE_NVARCHAR, 1)
    columns.Add CatalogColumn("TABLE_PROPID", TYPE_INTN, 4)
    If withFlags Then columns.Add CatalogColumn("TABLE_FLAGS", TYPE_INTN, 4)

    width = columns.count
    Set rows = New Collection
    For Each entry In mTables
        ReDim row(0 To width - 1)
        For index = 0 To width - 1
            row(index) = Null
        Next index
        row(0) = ConfiguredDatabase
        row(1) = "dbo"
        row(2) = entry(0)
        row(3) = IIf(IsViewEntry(entry), "VIEW", "TABLE")
        ' The bookmark and version columns as a real server fills them for
        ' an ordinary table, read off one for each procedure: the 90 one
        ' describes an eight-byte bookmark (DBTYPE_UI8) and the older one a
        ' four-byte one (DBTYPE_UI4). This had nulls in all of them.
        row(5) = True
        row(6) = 1
        row(7) = IIf(withFlags, 21, 19)
        row(8) = IIf(withFlags, 8, 4)
        row(9) = 0
        row(10) = 1^
        row(11) = TableRowCount(entry)
        If withFlags Then row(14) = 0
        rows.Add row
    Next entry

    Set outColumns = columns
    Set outRows = rows
End Sub

' Every column, in the OLE DB layout rather than the ODBC one.
'
' Forty-two columns, of which this fills the dozen that mean anything for a
' read-only source. The rest are declared with the types the real procedure
' declares and left null, because a provider reads them by position.
Private Sub ColumnsRowset(ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim entry As Variant
    Dim tableColumns As Collection
    Dim column As Variant
    Dim row As Variant
    Dim described As Variant
    Dim position As Long
    Dim index As Long
    Dim textual As Boolean

    Set columns = New Collection
    columns.Add CatalogColumn("TABLE_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("COLUMN_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("COLUMN_GUID", TYPE_GUIDN, 16)
    columns.Add CatalogColumn("COLUMN_PROPID", TYPE_INTN, 4)
    columns.Add CatalogColumn("ORDINAL_POSITION", TYPE_INTN, 4)
    columns.Add CatalogColumn("COLUMN_HASDEFAULT", TYPE_BITN, 1)
    columns.Add CatalogColumn("COLUMN_DEFAULT", TYPE_NVARCHAR, 2000)
    columns.Add CatalogColumn("COLUMN_FLAGS", TYPE_INTN, 4)
    columns.Add CatalogColumn("IS_NULLABLE", TYPE_BITN, 1)
    columns.Add CatalogColumn("DATA_TYPE", TYPE_INT2, 2)
    columns.Add CatalogColumn("TYPE_GUID", TYPE_GUIDN, 16)
    columns.Add CatalogColumn("CHARACTER_MAXIMUM_LENGTH", TYPE_INTN, 4)
    columns.Add CatalogColumn("CHARACTER_OCTET_LENGTH", TYPE_INTN, 4)
    columns.Add CatalogColumn("NUMERIC_PRECISION", TYPE_INTN, 2)
    columns.Add CatalogColumn("NUMERIC_SCALE", TYPE_INTN, 2)
    columns.Add CatalogColumn("DATETIME_PRECISION", TYPE_INTN, 4)
    columns.Add CatalogColumn("CHARACTER_SET_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("CHARACTER_SET_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("CHARACTER_SET_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("COLLATION_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("COLLATION_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("COLLATION_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DOMAIN_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DOMAIN_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DOMAIN_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DESCRIPTION", TYPE_NVARCHAR, 1)
    columns.Add CatalogColumn("COLUMN_LCID", TYPE_INTN, 4)
    columns.Add CatalogColumn("COLUMN_COMPFLAGS", TYPE_INTN, 4)
    columns.Add CatalogColumn("COLUMN_SORTID", TYPE_INTN, 4)
    columns.Add CatalogColumn("COLUMN_TDSCOLLATION", TYPE_BIGBINARY, 5)
    columns.Add CatalogColumn("IS_COMPUTED", TYPE_BITN, 1)
    columns.Add CatalogColumn("SS_XML_SCHEMACOLLECTION_CATALOGNAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("SS_XML_SCHEMACOLLECTION_SCHEMANAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("SS_XML_SCHEMACOLLECTIONNAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("SS_UDT_CATALOGNAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("SS_UDT_SCHEMANAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("SS_UDT_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("SS_UDT_ASSEMBLY_TYPENAME", TYPE_NVARCHAR, 4000)
    columns.Add CatalogColumn("SS_IS_SPARSE", TYPE_BITN, 1)
    columns.Add CatalogColumn("SS_IS_COLUMN_SET", TYPE_BITN, 1)

    Set rows = New Collection
    For Each entry In mTables
        Set tableColumns = EntryColumns(entry)
        For position = 1 To tableColumns.count
            column = tableColumns(position)
            described = DescribeColumnType(CLng(column(COLUMN_TYPE)), _
                                           CLng(column(COLUMN_SIZE)))
            textual = (column(COLUMN_TYPE) = TYPE_NVARCHAR)

            ReDim row(0 To columns.count - 1)
            For index = 0 To columns.count - 1
                row(index) = Null
            Next index
            row(0) = ConfiguredDatabase
            row(1) = "dbo"
            row(2) = entry(0)
            row(3) = column(COLUMN_NAME)
            row(6) = position
            row(7) = False
            ' Readable and nullable, and writable unless the bridge was told
            ' otherwise: what a real server's flags say of an ordinary
            ' column, 116 for an int and 100 for text.
            row(9) = DBCOLUMNFLAGS_ISNULLABLE Or DBCOLUMNFLAGS_MAYBENULL Or _
                     IIf(described(7), DBCOLUMNFLAGS_ISFIXEDLENGTH, 0) Or _
                     IIf(mReadOnly, 0, DBCOLUMNFLAGS_WRITE)
            row(10) = True
            row(11) = described(6)
            If textual Then
                row(13) = described(2)
                row(14) = described(3)
                ' Which collation a text column compares by, in every form a
                ' provider might read it in, as a real server gives them.
                ' Left null, a linked server refused the whole rowset: it
                ' supports column level collation and was given none.
                row(18) = "master"
                row(19) = "dbo"
                row(20) = "iso_1"
                row(21) = "master"
                row(22) = "dbo"
                row(23) = "SQL_Latin1_General_CP1_CI_AS"
                row(28) = 1033
                row(29) = 196609
                row(31) = "0904D00034"
            Else
                row(15) = described(2)
                row(16) = described(4)
            End If
            ' The sort id, which a real server gives every column, text or
            ' not.
            row(30) = 52
            row(32) = False
            row(40) = False
            row(41) = False
            rows.Add row
        Next position
    Next entry

    Set outColumns = columns
    Set outRows = rows
End Sub

' A workbook's tables have no indexes. The columns are the real procedure's,
' read off a capture, because a provider reads them by position even when
' there are no rows to read.
Private Sub IndexesRowset(ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim columns As Collection

    Set columns = New Collection
    columns.Add CatalogColumn("TABLE_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("INDEX_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("INDEX_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("INDEX_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("PRIMARY_KEY", TYPE_BITN, 1)
    columns.Add CatalogColumn("UNIQUE", TYPE_BITN, 1)
    columns.Add CatalogColumn("CLUSTERED", TYPE_BITN, 1)
    columns.Add CatalogColumn("TYPE", TYPE_INTN, 2)
    columns.Add CatalogColumn("FILL_FACTOR", TYPE_INTN, 4)
    columns.Add CatalogColumn("INITIAL_SIZE", TYPE_INTN, 4)
    columns.Add CatalogColumn("NULLS", TYPE_INTN, 4)
    columns.Add CatalogColumn("SORT_BOOKMARKS", TYPE_BITN, 1)
    columns.Add CatalogColumn("AUTO_UPDATE", TYPE_BITN, 1)
    columns.Add CatalogColumn("NULL_COLLATION", TYPE_INTN, 4)
    columns.Add CatalogColumn("ORDINAL_POSITION", TYPE_INTN, 4)
    columns.Add CatalogColumn("COLUMN_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("COLUMN_GUID", TYPE_GUIDN, 16)
    columns.Add CatalogColumn("COLUMN_PROPID", TYPE_INTN, 4)
    columns.Add CatalogColumn("COLLATION", TYPE_INTN, 2)
    columns.Add CatalogColumn("CARDINALITY", TYPE_INTN, 4)
    columns.Add CatalogColumn("PAGES", TYPE_INTN, 4)
    columns.Add CatalogColumn("FILTER_CONDITION", TYPE_NVARCHAR, _
                              NVARCHAR_SIZED_LIMIT + 1)
    columns.Add CatalogColumn("INTEGRATED", TYPE_BITN, 1)
    columns.Add CatalogColumn("STATUS", TYPE_INTN, 4)

    Set outColumns = columns
    Set outRows = New Collection
End Sub

' Nor check constraints.
Private Sub CheckConstraintsRowset(ByRef outColumns As Collection, _
                                   ByRef outRows As Collection)
    Dim columns As Collection

    Set columns = New Collection
    columns.Add CatalogColumn("TABLE_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("CONSTRAINT_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("CONSTRAINT_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("CONSTRAINT_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("CHECK_CLAUSE", TYPE_NVARCHAR, 4000)
    columns.Add CatalogColumn("DESCRIPTION", TYPE_NVARCHAR, 1)

    Set outColumns = columns
    Set outRows = New Collection
End Sub

' Nor statistics, which a linked server asks for to cost the read it is
' about to make.
Private Sub TableStatisticsRowset(ByRef outColumns As Collection, _
                                  ByRef outRows As Collection)
    Dim columns As Collection

    Set columns = New Collection
    columns.Add CatalogColumn("TABLE_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("STATISTICS_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("STATISTICS_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("STATISTICS_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("STATISTICS_TYPE", TYPE_INTN, 2)
    columns.Add CatalogColumn("COLUMN_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("COLUMN_GUID", TYPE_GUIDN, 16)
    columns.Add CatalogColumn("COLUMN_PROPID", TYPE_INTN, 4)
    columns.Add CatalogColumn("ORDINAL_POSITION", TYPE_INTN, 4)
    columns.Add CatalogColumn("SAMPLE_PCT", TYPE_INTN, 2)
    columns.Add CatalogColumn("LAST_UPDATE_TIME", TYPE_DATETIMN, 8)
    columns.Add CatalogColumn("NO_OF_RANGES", TYPE_INTN, 4)
    columns.Add CatalogColumn("COLUMN_CARDINALITY", TYPE_INTN, 8)
    columns.Add CatalogColumn("TUPLE_CARDINALITY", TYPE_INTN, 8)
    columns.Add CatalogColumn("TABLE_CARDINALITY", TYPE_INTN, 8)
    columns.Add CatalogColumn("AVG_COLUMN_LENGTH", TYPE_INTN, 4)

    Set outColumns = columns
    Set outRows = New Collection
End Sub

Private Sub SchemataRowset(ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim row As Variant

    Set columns = New Collection
    columns.Add CatalogColumn("CATALOG_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("SCHEMA_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("SCHEMA_OWNER", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DEFAULT_CHARACTER_SET_CATALOG", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DEFAULT_CHARACTER_SET_SCHEMA", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DEFAULT_CHARACTER_SET_NAME", TYPE_NVARCHAR, 128)

    Set rows = New Collection
    ReDim row(0 To 5)
    row(0) = ConfiguredDatabase
    row(1) = "dbo"
    row(2) = "dbo"
    row(3) = Null
    row(4) = Null
    row(5) = Null
    rows.Add row

    Set outColumns = columns
    Set outRows = rows
End Sub

Private Sub CatalogsRowset(ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim row As Variant

    Set columns = New Collection
    columns.Add CatalogColumn("CATALOG_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DESCRIPTION", TYPE_NVARCHAR, 1)

    Set rows = New Collection
    ReDim row(0 To 1)
    row(0) = ConfiguredDatabase
    row(1) = Null
    rows.Add row

    Set outColumns = columns
    Set outRows = rows
End Sub

' -- the ODBC family ------------------------------------------------------

Private Sub DatabasesRowset(ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim row As Variant

    Set columns = New Collection
    columns.Add CatalogColumn("DATABASE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DATABASE_SIZE", TYPE_INTN, 4)
    columns.Add CatalogColumn("REMARKS", TYPE_NVARCHAR, 254)

    Set rows = New Collection
    ReDim row(0 To 2)
    row(0) = ConfiguredDatabase
    row(1) = 0
    row(2) = Null
    rows.Add row

    Set outColumns = columns
    Set outRows = rows
End Sub

Private Sub OdbcTables(ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim entry As Variant
    Dim row As Variant

    Set columns = New Collection
    columns.Add CatalogColumn("TABLE_QUALIFIER", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_OWNER", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_TYPE", TYPE_NVARCHAR, 32)
    columns.Add CatalogColumn("REMARKS", TYPE_NVARCHAR, 254)

    Set rows = New Collection
    For Each entry In mTables
        ReDim row(0 To 4)
        row(0) = ConfiguredDatabase
        row(1) = "dbo"
        row(2) = entry(0)
        row(3) = IIf(IsViewEntry(entry), "VIEW", "TABLE")
        row(4) = Null
        rows.Add row
    Next entry

    Set outColumns = columns
    Set outRows = rows
End Sub

Private Sub OdbcColumns(ByRef outColumns As Collection, ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim entry As Variant
    Dim tableColumns As Collection
    Dim column As Variant
    Dim described As Variant
    Dim row As Variant
    Dim position As Long

    Set columns = New Collection
    columns.Add CatalogColumn("TABLE_QUALIFIER", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_OWNER", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("TABLE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("COLUMN_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("DATA_TYPE", TYPE_INT2, 2)
    columns.Add CatalogColumn("TYPE_NAME", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("PRECISION", TYPE_INTN, 4)
    columns.Add CatalogColumn("LENGTH", TYPE_INTN, 4)
    columns.Add CatalogColumn("SCALE", TYPE_INTN, 2)
    columns.Add CatalogColumn("RADIX", TYPE_INTN, 2)
    columns.Add CatalogColumn("NULLABLE", TYPE_INTN, 2)
    columns.Add CatalogColumn("REMARKS", TYPE_NVARCHAR, 254)
    columns.Add CatalogColumn("COLUMN_DEF", TYPE_NVARCHAR, 4000)
    columns.Add CatalogColumn("SQL_DATA_TYPE", TYPE_INTN, 2)
    columns.Add CatalogColumn("SQL_DATETIME_SUB", TYPE_INTN, 2)
    columns.Add CatalogColumn("CHAR_OCTET_LENGTH", TYPE_INTN, 4)
    columns.Add CatalogColumn("ORDINAL_POSITION", TYPE_INTN, 4)
    columns.Add CatalogColumn("IS_NULLABLE", TYPE_NVARCHAR, 10)
    columns.Add CatalogColumn("SS_DATA_TYPE", TYPE_INTN, 1)

    Set rows = New Collection
    For Each entry In mTables
        Set tableColumns = EntryColumns(entry)
        For position = 1 To tableColumns.count
            column = tableColumns(position)
            described = DescribeColumnType(CLng(column(COLUMN_TYPE)), _
                                           CLng(column(COLUMN_SIZE)))
            ReDim row(0 To 18)
            row(0) = ConfiguredDatabase
            row(1) = "dbo"
            row(2) = entry(0)
            row(3) = column(COLUMN_NAME)
            row(4) = described(1)
            row(5) = described(0)
            row(6) = described(2)
            row(7) = described(3)
            row(8) = described(4)
            row(9) = described(5)
            row(10) = 1                            ' every column is nullable
            row(11) = Null
            row(12) = Null
            row(13) = described(1)
            row(14) = Null
            row(15) = described(3)
            row(16) = position
            row(17) = "YES"
            row(18) = Null
            rows.Add row
        Next position
    Next entry

    Set outColumns = columns
    Set outRows = rows
End Sub

' -- what a column looks like to a client ---------------------------------

' Name, ODBC code, precision, length, scale, radix, OLE DB type, fixed width.
'
' Both code systems, because the ODBC procedures and the OLE DB ones want
' different numbers for the same column: nvarchar is -9 to ODBC and 130 to
' OLE DB, and -9 is not a DBTYPE at all.
Private Function DescribeColumnType(ByVal typeCode As Long, _
                                    ByVal size As Long) As Variant
    Dim out(0 To 7) As Variant

    Select Case typeCode
        Case TYPE_INTN
            If size = 8 Then
                out(0) = "bigint": out(1) = SQL_BIGINT
                out(2) = 19: out(3) = 8: out(6) = DBTYPE_I8
            ElseIf size = 2 Then
                out(0) = "smallint": out(1) = SQL_INTEGER
                out(2) = 5: out(3) = 2: out(6) = DBTYPE_I2
            ElseIf size = 1 Then
                out(0) = "tinyint": out(1) = SQL_INTEGER
                out(2) = 3: out(3) = 1: out(6) = DBTYPE_I2
            Else
                out(0) = "int": out(1) = SQL_INTEGER
                out(2) = 10: out(3) = 4: out(6) = DBTYPE_I4
            End If
            out(4) = 0: out(5) = 10: out(7) = True

        Case TYPE_INT2
            out(0) = "smallint": out(1) = SQL_INTEGER
            out(2) = 5: out(3) = 2: out(4) = 0: out(5) = 10
            out(6) = DBTYPE_I2: out(7) = True

        Case TYPE_FLTN
            out(0) = "float": out(1) = SQL_FLOAT
            out(2) = 53: out(3) = 8: out(4) = Null: out(5) = 2
            out(6) = DBTYPE_R8: out(7) = True

        Case TYPE_BITN
            out(0) = "bit": out(1) = SQL_BIT
            out(2) = 1: out(3) = 1: out(4) = 0: out(5) = Null
            out(6) = DBTYPE_BOOL: out(7) = True

        Case TYPE_DATETIMN
            out(0) = "datetime": out(1) = SQL_TYPE_TIMESTAMP
            out(2) = 23: out(3) = 16: out(4) = 3: out(5) = Null
            out(6) = DBTYPE_DBTIMESTAMP: out(7) = True

        Case TYPE_GUIDN
            out(0) = "uniqueidentifier": out(1) = SQL_GUID
            out(2) = 36: out(3) = 16: out(4) = Null: out(5) = Null
            out(6) = DBTYPE_GUID: out(7) = True

        Case Else
            ' A MAX column has no size to declare, and a client shown one
            ' asks for it in chunks rather than by length.
            If size <= 0 Or size > NVARCHAR_SIZED_LIMIT Then
                out(0) = "nvarchar": out(1) = SQL_WLONGVARCHAR
                out(2) = 1073741823#: out(3) = 2147483646#
            Else
                out(0) = "nvarchar": out(1) = SQL_WVARCHAR
                out(2) = size: out(3) = size * 2
            End If
            out(4) = Null: out(5) = Null
            out(6) = DBTYPE_WSTR: out(7) = False
    End Select

    DescribeColumnType = out
End Function

Private Function TableRowCount(ByVal entry As Variant) As Long
    Dim data As Variant

    data = entry(1)
    ' A view keeps no rows of its own, and counting them would mean running
    ' it for every rowset that asks how big the tables are.
    If Not IsArray(data) Then Exit Function
    TableRowCount = UBound(data, 1) - LBound(data, 1)
End Function

' A catalog rowset's columns are flagged the way SQL Server flags its own,
' not the way a SELECT result is. The difference is real: a SELECT of a CAST
' expression carries the computed bit, and a provider reading a catalog
' rowset that claims to be computed is being told something untrue about it.
Private Function CatalogColumn(ByVal name As String, ByVal typeCode As Long, _
                               ByVal size As Long) As Variant
    CatalogColumn = ColumnOf(name, typeCode, size, CATALOG_COLUMN_FLAGS)
End Function


' Sixteen bytes from a guid written as text, in the mixed-endian layout
' Windows uses: the first three fields little-endian and the last two big.
Private Function GuidBytes(ByVal text As String) As Byte()
    Dim clean As String
    Dim raw() As Byte
    Dim out(0 To 15) As Byte
    Dim order As Variant
    Dim index As Long

    clean = Replace(Replace(Replace(text, "-", ""), "{", ""), "}", "")
    If Len(clean) <> 32 Then
        GuidBytes = FromHex(String$(32, "0"))
        Exit Function
    End If
    raw = FromHex(clean)

    order = Array(3, 2, 1, 0, 5, 4, 7, 6, 8, 9, 10, 11, 12, 13, 14, 15)
    For index = 0 To 15
        out(index) = raw(order(index))
    Next index
    GuidBytes = out
End Function


' One row per distinct set of grouping values, in the order the groups were
' first seen. That order is what a client gets when it asks for none.
Private Sub ProjectGroups(ByVal matching As Collection, _
                          ByVal sourceColumns As Collection, _
                          ByRef outColumns As Collection, _
                          ByRef outRows As Collection)
    Dim groups As Collection
    Dim keys As Collection
    Dim rows As Collection
    Dim values As Variant
    Dim term As Variant
    Dim key As String
    Dim item As Variant
    Dim node As Variant
    Dim produced As Variant
    Dim index As Long
    Dim value As Variant
    Dim keep As Variant
    Dim taken As Long
    Dim computed As Collection
    Dim says As Variant

    Set groups = New Collection
    Set keys = New Collection

    For Each values In matching
        key = vbNullString
        For Each term In mGroupBy
            key = key & Chr$(1) & CStr(NullToEmpty( _
                Evaluate(term, values, sourceColumns)))
        Next term

        On Error Resume Next
        Set rows = groups(key)
        On Error GoTo 0
        If rows Is Nothing Then
            Set rows = New Collection
            groups.Add rows, key
            keys.Add key
        End If
        rows.Add values
        Set rows = Nothing
    Next values

    ' The shape comes from the select list, so a query that matched nothing
    ' still declares what it would have sent. A column whose type the
    ' statement states keeps it, the way a whole-table aggregate's does:
    ' COUNT is an integer and the sum of a float is a float whether or not
    ' any group produced one, and a grouped answer that matched nothing used
    ' to declare every column of it as text.
    Set computed = New Collection
    For Each item In mSelectItems
        says = StatedType(item(0), sourceColumns)
        If IsArray(says) Then
            outColumns.Add ColumnOf(ItemName(item, outColumns.count + 1), _
                                    CLng(says(0)), CLng(says(1)), _
                                    -1, vbNullString, _
                                    CLng(NullToZero(says(2))))
            computed.Add False
        Else
            outColumns.Add ColumnOf(ItemName(item, outColumns.count + 1), _
                                    TYPE_NVARCHAR, NVARCHAR_SIZED_LIMIT)
            computed.Add True
        End If
    Next item

    For Each item In keys
        Set rows = groups(CStr(item))

        If Not IsEmpty(mHavingNode) Then
            keep = EvaluateAggregate(mHavingNode, rows, sourceColumns)
            If IsNull(keep) Then GoTo NextGroup
            If Not CBool(keep) Then GoTo NextGroup
        End If
        If mTopCount >= 0 And taken >= mTopCount Then Exit For

        ReDim produced(0 To mSelectItems.count - 1)
        index = 0
        For Each node In mSelectItems
            value = EvaluateAggregate(node(0), rows, sourceColumns)
            produced(index) = value
            ' The declaration is fixed once, from the first group that
            ' produced a value, because every row is encoded against it.
            ' Only for a column the statement did not already say the type
            ' of: guessing from a value would undo what it said.
            If taken = 0 And IsComputedColumn(computed, index + 1) Then
                RetypeColumn outColumns, index + 1, value
            End If
            index = index + 1
        Next node
        outRows.Add produced
        taken = taken + 1
NextGroup:
    Next item

    RetypeComputedColumns outColumns, outRows, computed
End Sub

' Give a column the type its first value calls for, keeping its name.
Private Sub RetypeColumn(ByVal columns As Collection, ByVal position As Long, _
                         ByVal value As Variant)
    Dim existing As Variant

    If IsNull(value) Or IsEmpty(value) Then Exit Sub
    existing = columns(position)
    ReplaceColumn columns, position, _
        ColumnOf(CStr(existing(COLUMN_NAME)), TypeOfValue(value), _
                 SizeOfValue(value))
End Sub

Private Sub ReplaceColumn(ByVal columns As Collection, ByVal position As Long, _
                          ByVal replacement As Variant)
    columns.Remove position
    If position > columns.count Then
        columns.Add replacement
    Else
        columns.Add replacement, , position
    End If
End Sub

' Declare a computed column from the values it produced.
'
' A column worked out by an expression has no type until its values exist,
' and until then the only safe guess is the widest text a column can be. A
' client believes that guess in both directions: sqlcmd pads every value out
' to four thousand characters, and SMO reads a number declared as text,
' tries to cast it to an integer, and gives up on the connection.
Private Sub RetypeComputedColumns(ByVal columns As Collection, _
                                  ByVal rows As Collection, _
                                  ByVal computed As Collection)
    Dim index As Long
    Dim column As Variant
    Dim values As Variant
    Dim value As Variant
    Dim seen As Long
    Dim numbers As Long
    Dim whole As Long
    Dim dates As Long
    Dim bits As Long
    Dim widest As Long

    For index = 1 To columns.count
        If IsComputedColumn(computed, index) Then
            column = columns(index)
            seen = 0: numbers = 0: whole = 0: dates = 0: bits = 0: widest = 1

            For Each values In rows
                value = values(LBound(values) + index - 1)
                If Not (IsNull(value) Or IsEmpty(value)) Then
                    seen = seen + 1
                    If VarType(value) = vbBoolean Then bits = bits + 1
                    If VarType(value) = vbDate Then dates = dates + 1
                    ' A number, not a string that reads as one: a code
                    ' typed as text keeps its leading zeros, and calling
                    ' the column an integer would send 007 as 7.
                    If IsNumeric(value) And VarType(value) <> vbDate _
                       And VarType(value) <> vbBoolean _
                       And VarType(value) <> vbString Then
                        numbers = numbers + 1
                        If CDbl(value) = Int(CDbl(value)) _
                           And Abs(CDbl(value)) < 2147483647# Then
                            whole = whole + 1
                        End If
                    End If
                    If Len(CStr(value)) > widest Then widest = Len(CStr(value))
                End If
            Next values

            If seen = 0 Then
                ReplaceColumn columns, index, _
                    ColumnOf(CStr(column(COLUMN_NAME)), TYPE_NVARCHAR, 1)
            ElseIf bits = seen Then
                ReplaceColumn columns, index, _
                    ColumnOf(CStr(column(COLUMN_NAME)), TYPE_BITN, 1)
            ElseIf dates = seen Then
                ReplaceColumn columns, index, _
                    ColumnOf(CStr(column(COLUMN_NAME)), TYPE_DATETIMN, 8)
            ElseIf numbers = seen And whole = seen Then
                ReplaceColumn columns, index, _
                    ColumnOf(CStr(column(COLUMN_NAME)), TYPE_INTN, 4)
            ElseIf numbers = seen Then
                ReplaceColumn columns, index, _
                    ColumnOf(CStr(column(COLUMN_NAME)), TYPE_FLTN, 8)
            Else
                ' Past what a size field can describe, the MAX form is the
                ' only way to carry the value at all.
                If widest > NVARCHAR_SIZED_LIMIT Then widest = 0
                ReplaceColumn columns, index, _
                    ColumnOf(CStr(column(COLUMN_NAME)), TYPE_NVARCHAR, widest)
            End If
        End If
    Next index
End Sub

' Nothing said means every column was computed, which is what the grouped
' and aggregated paths produce.
Private Function IsComputedColumn(ByVal computed As Collection, _
                                  ByVal index As Long) As Boolean
    If computed Is Nothing Then
        IsComputedColumn = True
    ElseIf index > computed.count Then
        IsComputedColumn = True
    Else
        IsComputedColumn = CBool(computed(index))
    End If
End Function


' What an ORDER BY term actually names.
'
' A whole number is a position in the select list rather than a value: every
' row would compare equal on the literal, so sorting by one would do nothing
' at all and say nothing about it.
' Which select item goes out under this name, or zero.
Private Function SelectItemNamed(ByVal name As String) As Long
    Dim position As Long
    Dim bare As String

    bare = Replace(Replace(name, "[", ""), "]", "")
    For position = 1 To mSelectItems.count
        If StrComp(ItemName(mSelectItems(position), position), bare, _
                   vbTextCompare) = 0 Then
            SelectItemNamed = position
            Exit Function
        End If
    Next position
End Function

' Which select item is written the same way as this expression, or nought.
' Written the same way rather than worked out the same way: ORDER BY
' SUM(total) and SUM(total) AS revenue are one expression twice, which is
' what a client means by it.
Private Function SelectItemMatching(ByVal node As Variant) As Long
    Dim index As Long
    Dim item As Variant

    If mSelectItems Is Nothing Then Exit Function
    For index = 1 To mSelectItems.count
        item = mSelectItems(index)
        If SameNode(item(0), node) Then
            SelectItemMatching = index
            Exit Function
        End If
    Next index
End Function

' Whether two parsed expressions are the same expression.
Private Function SameNode(ByVal left As Variant, ByVal right As Variant) As Boolean
    Dim index As Long

    If IsArray(left) <> IsArray(right) Then Exit Function
    If Not IsArray(left) Then
        SameNode = SamePart(left, right)
        Exit Function
    End If
    If UBound(left) <> UBound(right) Then Exit Function
    For index = LBound(left) To UBound(left)
        If Not SamePart(left(index), right(index)) Then Exit Function
    Next index
    SameNode = True
End Function

' One part of a node against the other's: a nested node, a list of them, or
' a value. Text compares the way this server compares text, so SUM(Total)
' and SUM(total) are the same expression.
Private Function SamePart(ByVal left As Variant, ByVal right As Variant) As Boolean
    Dim index As Long

    If IsObject(left) Or IsObject(right) Then
        If Not (IsObject(left) And IsObject(right)) Then Exit Function
        If left Is Nothing Or right Is Nothing Then
            SamePart = (left Is Nothing And right Is Nothing)
            Exit Function
        End If
        If Not (TypeOf left Is Collection) Then Exit Function
        If Not (TypeOf right Is Collection) Then Exit Function
        If left.count <> right.count Then Exit Function
        For index = 1 To left.count
            If Not SameNode(left(index), right(index)) Then Exit Function
        Next index
        SamePart = True
        Exit Function
    End If

    If IsArray(left) Or IsArray(right) Then
        SamePart = SameNode(left, right)
        Exit Function
    End If
    If IsNull(left) Or IsNull(right) Then
        SamePart = (IsNull(left) And IsNull(right))
        Exit Function
    End If
    If IsEmpty(left) Or IsEmpty(right) Then
        SamePart = (IsEmpty(left) And IsEmpty(right))
        Exit Function
    End If
    If VarType(left) <> VarType(right) Then Exit Function
    If VarType(left) = vbString Then
        SamePart = (StrComp(CStr(left), CStr(right), vbTextCompare) = 0)
    Else
        SamePart = (left = right)
    End If
End Function

Private Function OrderTerm(ByVal node As Variant, _
                           ByVal columns As Collection) As Variant
    Dim position As Long
    Dim item As Variant
    Dim bare As String

    OrderTerm = node
    If Not IsArray(node) Then Exit Function

    ' A name in ORDER BY may be a column of what was read or the name a
    ' select item was given. The source is asked first, as SQL Server asks
    ' it; only a name it has no column for is looked for among the names
    ' the select list hands out. Without this an object tree cannot order
    ' by the alias it just wrote: ORDER BY [ID] over column_id AS [ID].
    If node(0) = N_COLUMN Then
        If ColumnIndex(CStr(node(1)), columns) > 0 Then Exit Function

        ' A grouped answer's columns carry no source name with them, so
        ' ORDER BY p.team has to find the column called team.
        bare = BareColumnName(CStr(node(1)))
        If ColumnIndex(bare, columns) > 0 Then
            OrderTerm = NodeOf(N_COLUMN, bare, Empty, Empty)
            Exit Function
        End If

        If Not Summarising() Then
            position = SelectItemNamed(CStr(node(1)))
            If position = 0 Then position = SelectItemNamed(bare)
            If position > 0 Then
                item = mSelectItems(position)
                OrderTerm = item(0)
            End If
        End If
        Exit Function
    End If

    ' A grouped answer is sorted over the rows that went out, so an
    ' aggregate written in the ORDER BY is the column the same aggregate
    ' went out as: ORDER BY SUM(total) DESC sorts by the SUM(total) in the
    ' select list. Taken as it stands it looked for a column the summarised
    ' rows no longer have and refused the statement as a missing column.
    If Summarising() Then
        position = SelectItemMatching(node)
        If position > 0 Then
            If position <= columns.count Then
                OrderTerm = NodeOf(N_COLUMN, _
                                   CStr(columns(position)(COLUMN_NAME)), _
                                   Empty, Empty)
                Exit Function
            End If
        End If
    End If

    If node(0) <> N_LITERAL Then Exit Function
    If IsNull(node(1)) Then Exit Function
    If Not IsNumeric(node(1)) Then Exit Function

    position = CLng(node(1))
    If position < 1 Or position > mSelectItems.count Then
        RaiseQueryError 108, "The ORDER BY position number " & position & _
            " is out of range of the number of items in the select list."
    End If

    ' In a summarised answer the sort runs over the rows that went out, so
    ' the position is the column itself; otherwise it is the expression that
    ' would have produced it.
    If Summarising() Then
        OrderTerm = NodeOf(N_COLUMN, _
                           CStr(columns(position)(COLUMN_NAME)), Empty, Empty)
    Else
        item = mSelectItems(position)
        OrderTerm = item(0)
    End If
End Function


' ============================================================================
' What the server says about itself
'
' A client asks all of this before it will show a single table. The values
' agree with each other on purpose: a client that read the version two ways
' and found them different would be right to complain.
' ============================================================================

Private Function IsHexChar(ByVal ch As String) As Boolean
    IsHexChar = (ch >= "0" And ch <= "9") Or _
                (UCase$(ch) >= "A" And UCase$(ch) <= "F")
End Function

Private Function HexValue(ByVal digits As String) As LongLong
    Dim value As LongLong
    Dim index As Long
    Dim ch As String

    For index = 1 To Len(digits)
        ch = UCase$(Mid$(digits, index, 1))
        If ch >= "0" And ch <= "9" Then
            value = value * 16 + (Asc(ch) - 48)
        Else
            value = value * 16 + (Asc(ch) - 55)
        End If
    Next index
    HexValue = value
End Function

' Whether this side of a division is an integer by declaration. Anything
' that declares nothing is taken as one, which is what division did before
' there was anything better to go on.
Private Function DividesWhole(ByVal node As Variant, _
                              ByVal columns As Collection) As Boolean
    Dim stated As Variant

    stated = StatedType(node, columns)
    If Not IsArray(stated) Then
        DividesWhole = True
        Exit Function
    End If
    DividesWhole = (CLng(stated(0)) = TYPE_INTN Or CLng(stated(0)) = TYPE_INT2)
End Function

Private Function IsWholeNumber(ByVal value As Variant) As Boolean
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbBoolean Or VarType(value) = vbDate Then Exit Function
    If Not IsNumeric(value) Then Exit Function
    IsWholeNumber = (CDbl(value) = Int(CDbl(value)))
End Function

' Integer division that truncates toward zero, the way T-SQL does.
Private Function TruncatedDivide(ByVal a As LongLong, ByVal b As LongLong) As LongLong
    Dim quotient As LongLong

    quotient = Abs(a) \ Abs(b)
    If (a < 0) <> (b < 0) Then quotient = -quotient
    TruncatedDivide = quotient
End Function

Private Function BitwiseAnd(ByVal left As Variant, ByVal right As Variant) As Variant
    If IsNull(left) Or IsNull(right) Then
        BitwiseAnd = Null
    Else
        BitwiseAnd = CLngLng(left) And CLngLng(right)
    End If
End Function

Private Function BitwiseOr(ByVal left As Variant, ByVal right As Variant) As Variant
    If IsNull(left) Or IsNull(right) Then
        BitwiseOr = Null
    Else
        BitwiseOr = CLngLng(left) Or CLngLng(right)
    End If
End Function

Private Function BitwiseXor(ByVal left As Variant, ByVal right As Variant) As Variant
    If IsNull(left) Or IsNull(right) Then
        BitwiseXor = Null
    Else
        BitwiseXor = CLngLng(left) Xor CLngLng(right)
    End If
End Function

' CAST(expression AS type), which a client writes around nearly everything
' it reads. The type is kept as text and applied to the value.
Private Function ParseCast() As Variant
    Dim node As Variant
    Dim typeName As String
    Dim arguments As Collection

    Expect "("
    node = ParseExpression()
    Expect "AS"
    typeName = TakeText()
    If PeekIs("(") Then typeName = typeName & BracketedText()
    Expect ")"

    Set arguments = New Collection
    arguments.Add node
    arguments.Add NodeOf(N_LITERAL, typeName, Empty, Empty)
    ParseCast = NodeOf(N_CALL, "$CAST", arguments, Empty)
End Function

' A bracketed size, put back together as it was written, so that a type
' name carries the width the client asked for.
Private Function BracketedText() As String
    Dim out As String

    Expect "("
    Do While Not AtEnd()
        If Accept(")") Then Exit Do
        out = out & TakeText()
    Loop
    BracketedText = "(" & out & ")"
End Function

Private Function CastValue(ByVal value As Variant, ByVal typeName As String) As Variant
    Dim text As String
    Dim width As Long

    If IsNull(value) Or IsEmpty(value) Then
        CastValue = Null
        Exit Function
    End If

    Select Case BareTypeName(typeName)
        Case "int", "integer", "bigint", "smallint", "tinyint"
            If IsNumeric(value) Then
                ' Toward zero, not downward: casting -2.7 to an integer is
                ' -2, and Int would make it -3.
                CastValue = CLngLng(Fix(CDbl(value)))
            ElseIf VarType(value) = vbBoolean Then
                CastValue = IIf(CBool(value), 1, 0)
            Else
                CastValue = Null
            End If
        Case "bit"
            If IsNumeric(value) Then
                CastValue = (CDbl(value) <> 0)
            Else
                CastValue = CBool(value)
            End If
        Case "decimal", "numeric", "money", "smallmoney"
            ' Decimal rather than Double, because the point of asking for a
            ' decimal is the digits a Double would round away.
            CastValue = CDec(value)
        Case "float", "real"
            CastValue = CDbl(value)
        Case "datetime", "datetime2", "smalldatetime", "date"
            ' Read the way a comparison reads it, so the ISO forms VBA's
            ' CDate cannot take are dates here too, and a number is days
            ' from 1900 rather than from VBA's 1899.
            CastValue = CDate(MomentOf(value))
        Case Else
            ' A cast to nvarchar(3) is three characters. Answering with the
            ' whole string sent more bytes than the column it was declared
            ' in could hold, and a client reading it stopped with an
            ' internal error rather than a wrong value.
            text = CStr(value)
            width = DeclaredWidth(typeName)
            If width > 0 And Len(text) > width Then text = Left$(text, width)
            CastValue = text
    End Select
End Function

' A type name without whatever it was written with in brackets.
Private Function BareTypeName(ByVal written As String) As String
    Dim bare As String
    Dim opening As Long

    bare = LCase$(Trim$(written))
    opening = InStr(bare, "(")
    If opening > 0 Then bare = Trim$(Left$(bare, opening - 1))
    BareTypeName = bare
End Function

' How many characters a type name asks for, or zero where it does not ask
' or asks for the most there is.
Private Function DeclaredWidth(ByVal written As String) As Long
    Dim inside As String
    Dim opening As Long

    opening = InStr(written, "(")
    If opening = 0 Then Exit Function
    inside = LCase$(Trim$(Mid$(written, opening + 1)))
    If Left$(inside, 3) = "max" Then Exit Function
    DeclaredWidth = CLng(Val(inside))
End Function

' QUOTENAME(name [, quote]), which wraps a name so a client can put it back
' into a statement. The quote character is doubled inside, which is what
' makes the result safe to paste.
Private Function QuoteName(ByVal value As Variant, ByVal quote As Variant, _
                           ByVal given As Long) As Variant
    Dim text As String
    Dim mark As String

    If IsNull(value) Then
        QuoteName = Null
        Exit Function
    End If

    text = CStr(value)
    If given < 2 Then
        QuoteName = "[" & Replace(text, "]", "]]") & "]"
        Exit Function
    End If

    mark = CStr(quote)
    If mark = "[" Or mark = "]" Then
        QuoteName = "[" & Replace(text, "]", "]]") & "]"
    Else
        QuoteName = mark & Replace(text, mark, mark & mark) & mark
    End If
End Function

' How many rows the statement before this one answered with or changed.
'
' Kept per batch rather than per connection, which is where a real server
' keeps it: the first read of a batch says nothing rather than what the
' batch before it did.
Private Sub RememberRowCount(ByVal answers As Collection, _
                             ByVal before As Long)
    Dim index As Long
    Dim answer As Variant

    mRowCount = 0
    For index = before + 1 To answers.count
        answer = answers(index)
        If answer(2) >= 0 Then
            mRowCount = CLng(answer(2))
        ElseIf Not answer(1) Is Nothing Then
            mRowCount = answer(1).count
        End If
    Next index
End Sub

' What a variable written with two at-signs is declared as, where the value
' alone would say something else.
'
' @@SPID is a smallint on a real server and its value is a small whole
' number, which taken at face value is an int. sqlcmd never minded; the
' linked-server provider asks @@SPID to check the connection is usable, is
' handed four bytes where it bound two, and hangs up saying the physical
' connection is not usable.
Private Function ServerVariableType(ByVal name As String) As Variant
    Dim out(0 To 2) As Variant

    Select Case LCase$(name)
        Case "@@spid"
            out(0) = TYPE_INT2
            out(1) = 2
            out(2) = 0
            ServerVariableType = out
    End Select
End Function

' The variables written with two at-signs, which a client reads instead of
' asking a procedure.
Private Function ServerVariable(ByVal name As String) As Variant
    Select Case LCase$(name)
        Case "@@version"
            ServerVariable = "Microsoft SQL Server " & SERVER_VERSION_MAJOR & _
                "." & SERVER_VERSION_MINOR & "." & SERVER_VERSION_BUILD & _
                " (RTM) - vbaSQLBridge, serving a workbook"
        Case "@@microsoftversion"
            ' Major, minor and build packed the way a client takes them
            ' apart: shifted right by 0x1000000 and 0x10000 and masked.
            ServerVariable = CLngLng(SERVER_VERSION_MAJOR) * 16777216^ + _
                             CLngLng(SERVER_VERSION_MINOR) * 65536^ + _
                             SERVER_VERSION_BUILD
        Case "@@servername"
            ServerVariable = ConfiguredServerName
        Case "@@servicename"
            ServerVariable = "MSSQLSERVER"
        Case "@@language"
            ServerVariable = DEFAULT_LANGUAGE
        Case "@@spid"
            ServerVariable = TDS_FIRST_SESSION_ID
        Case "@@rowcount"
            ServerVariable = mRowCount
        Case "@@trancount"
            ' The calling connection's own count. Nought whatever had just
            ' happened, IF @@TRANCOUNT > 0 COMMIT never committed.
            ServerVariable = TranCount()
        Case "@@error"
            ' The connection's last error, where it answered nought always:
            ' IF @@ERROR <> 0 after a statement never fired.
            ServerVariable = ErrorNumber()
        Case "@@fetch_status"
            ServerVariable = 0
        Case "@@max_precision"
            ServerVariable = 38
        Case Else
            ServerVariable = Null
    End Select
End Function

' What SERVERPROPERTY answers.
'
' SMO reads Edition to find out whether it is talking to Azure and
' EngineEdition to find out what kind of server this is, and it asks both
' before it will show anything at all.
Private Function ServerProperty(ByVal name As String) As Variant
    Select Case UCase$(Trim$(name))
        Case "EDITION": ServerProperty = "Developer Edition (64-bit)"
        Case "ENGINEEDITION": ServerProperty = 3
        Case "PRODUCTVERSION"
            ServerProperty = SERVER_VERSION_MAJOR & "." & _
                             SERVER_VERSION_MINOR & "." & _
                             SERVER_VERSION_BUILD & ".0"
        Case "PRODUCTLEVEL": ServerProperty = "RTM"
        Case "PRODUCTMAJORVERSION": ServerProperty = CStr(SERVER_VERSION_MAJOR)
        Case "PRODUCTMINORVERSION": ServerProperty = CStr(SERVER_VERSION_MINOR)
        Case "PRODUCTBUILD": ServerProperty = CStr(SERVER_VERSION_BUILD)
        Case "PRODUCTUPDATELEVEL": ServerProperty = Null
        Case "MACHINENAME", "SERVERNAME", "COMPUTERNAMEPHYSICALNETBIOS"
            ServerProperty = ConfiguredServerName
        Case "INSTANCENAME": ServerProperty = Null
        Case "ISCLUSTERED", "ISHADRENABLED", "ISSINGLEUSER"
            ServerProperty = 0
        Case "ISINTEGRATEDSECURITYONLY": ServerProperty = 1
        Case "ISXTPSUPPORTED": ServerProperty = 1
        Case "ISFULLTEXTINSTALLED": ServerProperty = 0
        Case "COLLATION": ServerProperty = "SQL_Latin1_General_CP1_CI_AS"
        Case "COLLATIONID": ServerProperty = 872468488
        Case "COMPARISONSTYLE": ServerProperty = 196609
        Case "LCID": ServerProperty = 1033
        Case "SQLCHARSET": ServerProperty = 1
        Case "SQLCHARSETNAME": ServerProperty = "iso_1"
        Case "SQLSORTORDER": ServerProperty = 52
        Case "SQLSORTORDERNAME": ServerProperty = "nocase_iso"
        Case "PATHSEPARATOR": ServerProperty = "\"
        Case "BUILDCLRVERSION": ServerProperty = "v4.0.30319"
        Case "INSTANCEDEFAULTDATAPATH", "INSTANCEDEFAULTLOGPATH"
            ' Where a server would put a file it was asked to make. This is
            ' asked before anything is created, and answering nothing leaves
            ' the dialog that asked with an empty box.
            ServerProperty = DEFAULT_DATA_PATH

        Case "ERRORLOGFILENAME"
            ServerProperty = DEFAULT_DATA_PATH & "LOG\ERRORLOG"
        Case "RESOURCEVERSION"
            ServerProperty = SERVER_VERSION_MAJOR & "." & _
                             SERVER_VERSION_MINOR & "." & SERVER_VERSION_BUILD
        Case "RESOURCELASTUPDATEDATETIME": ServerProperty = UtcNow()
        Case "LICENSETYPE": ServerProperty = "DISABLED"
        Case "NUMLICENSES": ServerProperty = Null
        ' The features this does not have. No rather than nothing: a client
        ' asks whether a feature is installed and reads the answer as a yes
        ' or a no, where NULL is neither and is what a real server gives only
        ' for a property it has never heard of.
        Case "ISADVANCEDANALYTICSINSTALLED", "ISBIGDATACLUSTER", "ISLOCALDB", _
             "ISPOLYBASEINSTALLED", "ISTEMPDBMETADATAMEMORYOPTIMIZED", _
             "SUSPENDEDDATABASECOUNT", "FILESTREAMCONFIGUREDLEVEL", _
             "FILESTREAMEFFECTIVELEVEL", "ISSERVERSUSPENDEDFORSNAPSHOTBACKUP"
            ServerProperty = 0
        Case Else
            ServerProperty = Null
    End Select
End Function


' ============================================================================
' The sys views a client reads while connecting
'
' Enough of each for the question being asked, which is always about the
' server rather than about the data. A view this does not have is answered
' with its columns and no rows, because a client that asked is looking for
' something this cannot have: there are no availability groups behind a
' worksheet.
' ============================================================================

' Whether a name belongs to the server rather than to the workbook.
'
' Anything under the sys schema, and the handful of tables a client reads out
' of master by their old names.
Private Function IsSystemView(ByVal name As String) As Boolean
    Dim parts() As String
    Dim bare As String

    If InStr(1, name, "sys.", vbTextCompare) > 0 Then
        IsSystemView = True
        Exit Function
    End If

    parts = Split(Replace(Replace(name, "[", ""), "]", ""), ".")
    bare = LCase$(parts(UBound(parts)))
    Select Case bare
        Case "spt_values", "sysobjects", "syscolumns", "sysusers", "sysfiles"
            IsSystemView = True
            Exit Function
    End Select

    ' A name in one of the databases a real server always has is the
    ' server's business, not a user's. An object tree reads several of them
    ' while drawing itself, msdb.dbo.syspolicy_configuration among them, and
    ' a name it asks for and does not get ends the connection. A name that
    ' is nobody's system database still raises, so a misspelt table is
    ' still a misspelt table.
    If UBound(parts) >= 1 Then
        Select Case LCase$(parts(0))
            Case "master", "msdb", "tempdb", "model"
                IsSystemView = True
        End Select
    End If
End Function

Private Function SystemViewTable(ByVal name As String) As Variant
    Dim which As String
    Dim declared As String
    Dim parts() As String
    Dim rows As Collection
    Dim row As Variant
    Dim entry As Variant
    Dim tableColumns As Collection
    Dim column As Variant
    Dim described As Variant
    Dim position As Long

    parts = Split(LCase$(Replace(Replace(name, "[", ""), "]", "")), ".")
    which = parts(UBound(parts))
    Set rows = New Collection

    Select Case which
        Case "dm_os_host_info"
            ReDim row(0 To 6)
            row(0) = "Windows"
            row(1) = "Windows"
            row(2) = "10.0"
            row(3) = ""
            row(4) = 4
            row(5) = 1033
            row(6) = "x64"
            rows.Add row
            SystemViewTable = BuildView( _
                "host_platform,host_distribution,host_release," & _
                "host_service_pack_level,host_sku,os_language_version," & _
                "host_architecture", rows)

        Case "databases"
            ' What a real server reports for a database with nothing unusual
            ' about it, and the truth about this one where they differ: it
            ' is read only, it has no owner because this keeps no
            ' principals, and it was created when the workbook started.
            ReDim row(0 To 97)
            row(0) = ConfiguredDatabase     ' name
            row(1) = 1                      ' database_id
            row(2) = Null                   ' source_database_id
            row(3) = Null                   ' owner_sid
            row(4) = UtcNow()               ' create_date
            row(5) = 170                    ' compatibility_level
            row(6) = "SQL_Latin1_General_CP1_CI_AS" ' collation_name
            row(7) = 0                      ' user_access
            row(8) = "MULTI_USER"           ' user_access_desc
            row(9) = mReadOnly              ' is_read_only
            row(10) = False                 ' is_auto_close_on
            row(11) = False                 ' is_auto_shrink_on
            row(12) = 0                     ' state
            row(13) = "ONLINE"              ' state_desc
            row(14) = False                 ' is_in_standby
            row(15) = False                 ' is_cleanly_shutdown
            row(16) = False                 ' is_supplemental_logging_enabled
            row(17) = 0                     ' snapshot_isolation_state
            row(18) = "OFF"                 ' snapshot_isolation_state_desc
            row(19) = False                 ' is_read_committed_snapshot_on
            row(20) = 3                     ' recovery_model
            row(21) = "SIMPLE"              ' recovery_model_desc
            row(22) = 2                     ' page_verify_option
            row(23) = "CHECKSUM"            ' page_verify_option_desc
            row(24) = True                  ' is_auto_create_stats_on
            ' is_auto_create_stats_incremental_on
            row(25) = False
            row(26) = True                  ' is_auto_update_stats_on
            row(27) = False                 ' is_auto_update_stats_async_on
            row(28) = False                 ' is_ansi_null_default_on
            row(29) = False                 ' is_ansi_nulls_on
            row(30) = False                 ' is_ansi_padding_on
            row(31) = False                 ' is_ansi_warnings_on
            row(32) = False                 ' is_arithabort_on
            row(33) = False                 ' is_concat_null_yields_null_on
            row(34) = False                 ' is_numeric_roundabort_on
            row(35) = False                 ' is_quoted_identifier_on
            row(36) = False                 ' is_recursive_triggers_on
            row(37) = False                 ' is_cursor_close_on_commit_on
            row(38) = False                 ' is_local_cursor_default
            row(39) = False                 ' is_fulltext_enabled
            row(40) = False                 ' is_trustworthy_on
            row(41) = True                  ' is_db_chaining_on
            row(42) = False                 ' is_parameterization_forced
            row(43) = False                 ' is_master_key_encrypted_by_server
            row(44) = False                 ' is_query_store_on
            row(45) = False                 ' is_published
            row(46) = False                 ' is_subscribed
            row(47) = False                 ' is_merge_published
            row(48) = False                 ' is_distributor
            row(49) = False                 ' is_sync_with_backup
            row(50) = StableGuid(ConfiguredDatabase, "service broker")
            row(51) = False                 ' is_broker_enabled
            row(52) = 0                     ' log_reuse_wait
            row(53) = "NOTHING"             ' log_reuse_wait_desc
            row(54) = False                 ' is_date_correlation_on
            row(55) = False                 ' is_cdc_enabled
            row(56) = False                 ' is_encrypted
            row(57) = False                 ' is_honor_broker_priority_on
            row(58) = Null                  ' replica_id
            row(59) = Null                  ' group_database_id
            row(60) = Null                  ' resource_pool_id
            row(61) = Null                  ' default_language_lcid
            row(62) = Null                  ' default_language_name
            row(63) = Null                  ' default_fulltext_language_lcid
            row(64) = Null                  ' default_fulltext_language_name
            row(65) = Null                  ' is_nested_triggers_on
            row(66) = Null                  ' is_transform_noise_words_on
            row(67) = Null                  ' two_digit_year_cutoff
            row(68) = 0                     ' containment
            row(69) = "NONE"                ' containment_desc
            row(70) = 0                     ' target_recovery_time_in_seconds
            row(71) = 0                     ' delayed_durability
            row(72) = "DISABLED"            ' delayed_durability_desc
            ' is_memory_optimized_elevate_to_snapshot_on
            row(73) = False
            row(74) = False                 ' is_federation_member
            row(75) = False                 ' is_remote_data_archive_enabled
            row(76) = True                  ' is_mixed_page_allocation_on
            ' is_temporal_history_retention_enabled
            row(77) = True
            row(78) = 0                     ' catalog_collation_type
            row(79) = "DATABASE_DEFAULT"    ' catalog_collation_type_desc
            row(80) = ConfiguredDatabase    ' physical_database_name
            row(81) = False                 ' is_result_set_caching_on
            ' is_accelerated_database_recovery_on
            row(82) = False
            row(83) = False                 ' is_tempdb_spill_to_remote_store
            row(84) = False                 ' is_stale_page_detection_on
            row(85) = True                  ' is_memory_optimized_enabled
            row(86) = False                 ' is_data_retention_enabled
            row(87) = False                 ' is_ledger_on
            row(88) = False                 ' is_change_feed_enabled
            row(89) = False                 ' is_data_lake_replication_enabled
            row(90) = False                 ' is_event_stream_enabled
            row(91) = 0                     ' data_compaction
            row(92) = "UNSUPPORTED"         ' data_compaction_desc
            row(93) = 0                     ' data_lake_log_publishing
            row(94) = "UNSUPPORTED"         ' data_lake_log_publishing_desc
            row(95) = False                 ' is_vorder_enabled
            ' is_proactive_statistics_refresh_on
            row(96) = False
            row(97) = False                 ' is_optimized_locking_on
            rows.Add row
            declared = "name,database_id:int," & _
                "source_database_id:int,owner_sid," & _
                "create_date:datetime," & _
                "compatibility_level:tinyint," & _
                "collation_name,user_access:tinyint," & _
                "user_access_desc,is_read_only:bit," & _
                "is_auto_close_on:bit," & _
                "is_auto_shrink_on:bit,state:tinyint," & _
                "state_desc,is_in_standby:bit," & _
                "is_cleanly_shutdown:bit," & _
                "is_supplemental_logging_enabled:bit," & _
                "snapshot_isolation_state:tinyint,"
            declared = declared & "snapshot_isolation_state_desc," & _
                "is_read_committed_snapshot_on:bit," & _
                "recovery_model:tinyint," & _
                "recovery_model_desc," & _
                "page_verify_option:tinyint," & _
                "page_verify_option_desc," & _
                "is_auto_create_stats_on:bit," & _
                "is_auto_create_stats_incremental_on:bit," & _
                "is_auto_update_stats_on:bit," & _
                "is_auto_update_stats_async_on:bit," & _
                "is_ansi_null_default_on:bit," & _
                "is_ansi_nulls_on:bit,"
            declared = declared & "is_ansi_padding_on:bit," & _
                "is_ansi_warnings_on:bit," & _
                "is_arithabort_on:bit," & _
                "is_concat_null_yields_null_on:bit," & _
                "is_numeric_roundabort_on:bit," & _
                "is_quoted_identifier_on:bit," & _
                "is_recursive_triggers_on:bit," & _
                "is_cursor_close_on_commit_on:bit," & _
                "is_local_cursor_default:bit," & _
                "is_fulltext_enabled:bit," & _
                "is_trustworthy_on:bit," & _
                "is_db_chaining_on:bit,"
            declared = declared & "is_parameterization_forced:bit," & _
                "is_master_key_encrypted_by_server:bit," & _
                "is_query_store_on:bit," & _
                "is_published:bit,is_subscribed:bit," & _
                "is_merge_published:bit," & _
                "is_distributor:bit," & _
                "is_sync_with_backup:bit," & _
                "service_broker_guid:uniqueidentifier," & _
                "is_broker_enabled:bit," & _
                "log_reuse_wait:tinyint," & _
                "log_reuse_wait_desc," & _
                "is_date_correlation_on:bit,"
            declared = declared & "is_cdc_enabled:bit,is_encrypted:bit," & _
                "is_honor_broker_priority_on:bit," & _
                "replica_id:uniqueidentifier," & _
                "group_database_id:uniqueidentifier," & _
                "resource_pool_id:int," & _
                "default_language_lcid:smallint," & _
                "default_language_name," & _
                "default_fulltext_language_lcid:int," & _
                "default_fulltext_language_name," & _
                "is_nested_triggers_on:bit," & _
                "is_transform_noise_words_on:bit," & _
                "two_digit_year_cutoff:smallint,"
            declared = declared & "containment:tinyint,containment_desc," & _
                "target_recovery_time_in_seconds:int," & _
                "delayed_durability:int," & _
                "delayed_durability_desc," & _
                "is_memory_optimized_elevate_to_snapshot_on:bit," & _
                "is_federation_member:bit," & _
                "is_remote_data_archive_enabled:bit," & _
                "is_mixed_page_allocation_on:bit," & _
                "is_temporal_history_retention_enabled:bit," & _
                "catalog_collation_type:int," & _
                "catalog_collation_type_desc," & _
                "physical_database_name,"
            declared = declared & "is_result_set_caching_on:bit," & _
                "is_accelerated_database_recovery_on:bit," & _
                "is_tempdb_spill_to_remote_store:bit," & _
                "is_stale_page_detection_on:bit," & _
                "is_memory_optimized_enabled:bit," & _
                "is_data_retention_enabled:bit," & _
                "is_ledger_on:bit," & _
                "is_change_feed_enabled:bit," & _
                "is_data_lake_replication_enabled:bit," & _
                "is_event_stream_enabled:bit," & _
                "data_compaction:tinyint," & _
                "data_compaction_desc,"
            declared = declared & "data_lake_log_publishing:tinyint," & _
                "data_lake_log_publishing_desc," & _
                "is_vorder_enabled:bit," & _
                "is_proactive_statistics_refresh_on:bit," & _
                "is_optimized_locking_on:bit"
            SystemViewTable = BuildView(declared, rows)

        Case "database_files", "master_files"
            ' SMO reads the master data file's path out of this to work out
            ' where the server keeps its files, and the log file's beside
            ' it. Nothing here has files, so the answer is where a server
            ' would have put them. The Databases node joins this on the
            ' first file of the first filegroup, which is why the data file
            ' carries file_id 1 and data_space_id 1 and the log carries
            ' neither.
            AddFileRow rows, 1, "master", "master.mdf", 0, "ROWS", 1
            AddFileRow rows, 2, "mastlog", "mastlog.ldf", 1, "LOG", 0
            SystemViewTable = BuildView( _
                "database_id,file_id,name,physical_name,type,type_desc," & _
                "state,state_desc,data_space_id,size,max_size,growth", rows)

        Case "views", "all_views"
            ' What fills the Views node. SMO reads all_views rather than
            ' views, and reads every field of it: a name under sys. that
            ' nothing here answers is an empty table rather than an error,
            ' so the node drew itself with no views in it and said nothing
            ' was wrong. The fields are the twenty-three a real server
            ' declares, in its order, read off the one on this machine.
            For Each entry In mTables
                If Not IsViewEntry(entry) Then GoTo NextView
                ReDim row(0 To 22)
                row(0) = entry(0)
                row(1) = TableObjectId(LCase$(CStr(entry(0))))
                row(2) = Null                       ' principal_id
                row(3) = 1                          ' schema_id: dbo
                row(4) = 0                          ' parent_object_id
                row(5) = "V"
                row(6) = "VIEW"
                row(7) = UtcNow()                   ' create_date
                row(8) = UtcNow()                   ' modify_date
                row(9) = False                      ' is_ms_shipped
                row(10) = False                     ' is_published
                row(11) = False                     ' is_schema_published
                row(12) = False                     ' is_replicated
                row(13) = False                     ' has_replication_filter
                row(14) = False                     ' has_opaque_metadata
                row(15) = False                     ' has_unchecked_assembly_data
                row(16) = False                     ' with_check_option
                row(17) = False                     ' is_date_correlation_view
                row(18) = False                     ' is_tracked_by_cdc
                row(19) = False                     ' has_snapshot
                row(20) = 0                         ' ledger_view_type
                row(21) = "NONE"                    ' ledger_view_type_desc
                row(22) = False                     ' is_dropped_ledger_view
                rows.Add row
NextView:
            Next entry
            SystemViewTable = BuildView( _
                "name,object_id,principal_id,schema_id,parent_object_id," & _
                "type,type_desc,create_date,modify_date,is_ms_shipped," & _
                "is_published,is_schema_published,is_replicated," & _
                "has_replication_filter,has_opaque_metadata," & _
                "has_unchecked_assembly_data,with_check_option," & _
                "is_date_correlation_view,is_tracked_by_cdc,has_snapshot," & _
                "ledger_view_type,ledger_view_type_desc," & _
                "is_dropped_ledger_view", rows)

        Case "tables", "objects", "all_objects"
            ' What fills the Tables node of an object tree. A client reads
            ' the object id here and looks the columns up by it, so the
            ' number has to be the one sys.columns reports as well.
            For Each entry In mTables
                If Not ObjectBelongs(which, entry) Then GoTo NextObject
                ReDim row(0 To 34)
                row(0) = entry(0)
                row(1) = TableObjectId(LCase$(CStr(entry(0))))
                row(2) = 1                          ' schema_id: dbo
                row(3) = 0                          ' parent_object_id
                row(4) = IIf(IsViewEntry(entry), "V", "U")
                row(5) = IIf(IsViewEntry(entry), "VIEW", "USER_TABLE")
                row(6) = UtcNow()                   ' create_date
                row(7) = UtcNow()                   ' modify_date
                row(8) = False                      ' is_ms_shipped
                row(9) = Null                       ' principal_id
                row(10) = True                      ' uses_ansi_nulls
                row(11) = False                     ' is_memory_optimized
                row(12) = False                     ' is_external
                row(13) = False                     ' is_filetable
                row(14) = 0                         ' ledger_view_id
                row(15) = False                     ' is_dropped_ledger_table
                row(16) = 0                         ' lock_escalation
                row(17) = "TABLE"                   ' lock_escalation_desc
                row(18) = 0                         ' temporal_type
                row(19) = "NON_TEMPORAL_TABLE"
                row(20) = Null                      ' history_table_id
                row(21) = False                     ' is_node
                row(22) = False                     ' is_edge
                row(23) = 0                         ' durability
                row(24) = "SCHEMA_AND_DATA"         ' durability_desc
                row(25) = False                     ' is_replicated
                row(26) = False                     ' is_tracked_by_cdc
                row(27) = 0                         ' text_in_row_limit
                row(28) = False                     ' large_value_types_out_of_row
                row(29) = 1                         ' filestream_data_space_id
                row(30) = Null                      ' data_retention_period
                row(31) = 0                         ' data_retention_period_unit
                row(32) = 0                         ' ledger_type
                row(33) = "NONE"                    ' ledger_type_desc
                row(34) = 1                         ' lob_data_space_id
                rows.Add row
NextObject:
            Next entry
            SystemViewTable = BuildView( _
                "name,object_id,schema_id,parent_object_id,type,type_desc," & _
                "create_date,modify_date,is_ms_shipped,principal_id," & _
                "uses_ansi_nulls,is_memory_optimized,is_external,is_filetable," & _
                "ledger_view_id,is_dropped_ledger_table,lock_escalation," & _
                "lock_escalation_desc,temporal_type,temporal_type_desc," & _
                "history_table_id,is_node,is_edge,durability,durability_desc," & _
                "is_replicated,is_tracked_by_cdc,text_in_row_limit," & _
                "large_value_types_out_of_row,filestream_data_space_id," & _
                "data_retention_period,data_retention_period_unit," & _
                "ledger_type,ledger_type_desc,lob_data_space_id", rows)

        Case "columns", "all_columns"
            For Each entry In mTables
                Set tableColumns = EntryColumns(entry)
                For position = 1 To tableColumns.count
                    column = tableColumns(position)
                    described = DescribeColumnType(CLng(column(COLUMN_TYPE)), _
                                                   CLng(column(COLUMN_SIZE)))
                    ReDim row(0 To 31)
                    row(0) = TableObjectId(LCase$(CStr(entry(0))))
                    row(1) = column(COLUMN_NAME)
                    row(2) = position
                    row(3) = SystemTypeId(CLng(column(COLUMN_TYPE)))
                    row(4) = SystemTypeId(CLng(column(COLUMN_TYPE)))
                    row(5) = CLng(described(3))     ' max_length, in bytes
                    row(6) = CLng(NullToZero(described(2)))
                    row(7) = CLng(NullToZero(described(4)))
                    row(8) = True                   ' is_nullable
                    row(9) = False                  ' is_sparse
                    row(10) = False                 ' is_identity
                    row(11) = False                 ' is_computed
                    row(12) = False                 ' is_rowguidcol
                    row(13) = False                 ' is_filestream
                    row(14) = False                 ' is_column_set
                    row(15) = 0                     ' ledger_view_column_type
                    row(16) = Null                  ' collation_name
                    row(17) = 0                     ' generated_always_type
                    row(18) = 0                     ' default_object_id
                    row(19) = 0                     ' rule_object_id
                    row(20) = 0                     ' xml_collection_id
                    row(21) = True                  ' is_ansi_padded
                    row(22) = False                 ' is_xml_document
                    row(23) = False                 ' is_hidden
                    row(24) = False                 ' is_masked
                    row(25) = False                 ' is_dropped_ledger_column
                    row(26) = Null                  ' column_encryption_key_id
                    row(27) = Null                  ' encryption_type
                    row(28) = Null                  ' encryption_algorithm_name
                    row(29) = Null                  ' graph_type
                    row(30) = Null                  ' vector_dimensions
                    row(31) = Null                  ' vector_base_type_desc
                    rows.Add row
                Next position
            Next entry
            SystemViewTable = BuildView( _
                "object_id,name,column_id,system_type_id,user_type_id," & _
                "max_length,precision,scale,is_nullable,is_sparse," & _
                "is_identity,is_computed,is_rowguidcol,is_filestream," & _
                "is_column_set,ledger_view_column_type,collation_name," & _
                "generated_always_type:tinyint,default_object_id," & _
                "rule_object_id,xml_collection_id,is_ansi_padded," & _
                "is_xml_document,is_hidden,is_masked," & _
                "is_dropped_ledger_column,column_encryption_key_id:int," & _
                "encryption_type:int,encryption_algorithm_name,graph_type:int," & _
                "vector_dimensions:int,vector_base_type_desc", rows)

        Case "types", "all_types"
            ' The types this serves, as a real server lists them. An empty
            ' view here joins to nothing and the tree draws every column
            ' with no type against it.
            AddTypeRow rows, "int", 56, 4, 10, 0
            AddTypeRow rows, "smallint", 52, 2, 5, 0
            AddTypeRow rows, "bit", 104, 1, 1, 0
            AddTypeRow rows, "float", 62, 8, 53, 0
            AddTypeRow rows, "datetime", 61, 8, 23, 3
            AddTypeRow rows, "uniqueidentifier", 36, 16, 0, 0
            AddTypeRow rows, "nvarchar", 231, 8000, 0, 0
            SystemViewTable = BuildView(EmptyViewColumns("types"), rows)

        Case "spt_values"
            ' A client works the page size out of this before it will report
            ' how much space a table takes. Eight kilobytes is what a real
            ' server answers, and nothing here occupies any of them.
            ReDim row(0 To 5)
            row(0) = Null
            row(1) = 1
            row(2) = "E"
            row(3) = 8192
            row(4) = 8192
            row(5) = 0
            rows.Add row
            SystemViewTable = BuildView( _
                "name,number,type,low,high,status", rows)

        Case "database_recovery_status"
            ' One row, because a client reads the two identifiers in it
            ' into a type that has no null. There is nothing to recover
            ' from, so the sequence numbers stay empty.
            ReDim row(0 To 6)
            row(0) = 1                          ' database_id
            row(1) = StableGuid(ConfiguredDatabase, "database")
            row(2) = StableGuid(ConfiguredDatabase, "family")
            row(3) = Null                       ' last_log_backup_lsn
            row(4) = StableGuid(ConfiguredDatabase, "fork")
            row(5) = StableGuid(ConfiguredDatabase, "first fork")
            row(6) = Null                       ' fork_point_lsn
            rows.Add row
            SystemViewTable = BuildView( _
                "database_id:int,database_guid:uniqueidentifier," & _
                "family_guid:uniqueidentifier," & _
                "last_log_backup_lsn:numeric(25,0)," & _
                "recovery_fork_guid:uniqueidentifier," & _
                "first_recovery_fork_guid:uniqueidentifier," & _
                "fork_point_lsn:numeric(25,0)", rows)

        Case "partitions"
            ' One partition per table, because that is what a heap is.
            '
            ' A client works a database's used space out by summing over
            ' this joined to sys.allocation_units, and a sum over no rows is
            ' a null where it expects a number. A real database always has
            ' at least one partition, and so does this one.
            For Each entry In mTables
                ReDim row(0 To 11)
                row(0) = PartitionIdOf(CStr(entry(0)))
                row(1) = TableObjectId(LCase$(CStr(entry(0))))
                row(2) = 0                      ' index_id: the heap
                row(3) = 1                      ' partition_number
                row(4) = PartitionIdOf(CStr(entry(0)))    ' hobt_id
                row(5) = TableRowCount(entry)
                row(6) = Null                   ' filestream_filegroup_id
                row(7) = 0                      ' data_compression: none
                row(8) = "NONE"
                row(9) = False                  ' xml_compression
                row(10) = "OFF"
                row(11) = 0                     ' used_page_count
                rows.Add row
            Next entry
            SystemViewTable = BuildView( _
                "partition_id:bigint,object_id:int,index_id:int," & _
                "partition_number:int,hobt_id:bigint,rows:bigint," & _
                "filestream_filegroup_id:smallint," & _
                "data_compression:tinyint,data_compression_desc," & _
                "xml_compression:bit,xml_compression_desc," & _
                "used_page_count:bigint", rows)

        Case "allocation_units"
            ' The store behind each partition. A workbook keeps its rows in
            ' a worksheet rather than in pages, so it occupies none of them,
            ' which is a count of zero rather than no answer.
            For Each entry In mTables
                ReDim row(0 To 7)
                row(0) = PartitionIdOf(CStr(entry(0)))
                row(1) = 1                      ' type: in row data
                row(2) = "IN_ROW_DATA"
                row(3) = PartitionIdOf(CStr(entry(0)))   ' container_id
                row(4) = 1                      ' data_space_id
                row(5) = 0                      ' total_pages
                row(6) = 0                      ' used_pages
                row(7) = 0                      ' data_pages
                rows.Add row
            Next entry
            SystemViewTable = BuildView( _
                "allocation_unit_id:bigint,type:tinyint,type_desc," & _
                "container_id:bigint,data_space_id:int,total_pages:bigint," & _
                "used_pages:bigint,data_pages:bigint", rows)

        Case "configurations", "sysconfigures", "syscurconfigs"
            ' The five a client reads by position while building a database,
            ' in the order it reads them. A setting it cannot find is not a
            ' default it falls back to: it is a row that is not there, and
            ' the read stops at "There is no row at position 0".
            AddConfiguration rows, 115, "nested triggers", 1
            AddConfiguration rows, 124, "default language", 0
            AddConfiguration rows, 1126, "default full-text language", 1033
            AddConfiguration rows, 1127, "two digit year cutoff", 2049
            AddConfiguration rows, 1555, "transform noise words", 0
            SystemViewTable = BuildView( _
                "configuration_id:int,name,description," & _
                "value:bigint,minimum:bigint,maximum:bigint," & _
                "value_in_use:bigint,is_dynamic:bit,is_advanced:bit", rows)

        Case "syslanguages"
            ' One language, which is the one this collation is in. The
            ' default language is read out of here by its langid.
            ReDim row(0 To 4)
            row(0) = 0                          ' langid
            row(1) = "us_english"
            row(2) = "English"                  ' alias
            row(3) = 1033                       ' lcid
            row(4) = 7                          ' msglangid
            rows.Add row
            SystemViewTable = BuildView( _
                "langid:smallint,name,alias,lcid:int,msglangid:smallint", rows)

        Case "fulltext_languages"
            ReDim row(0 To 1)
            row(0) = 1033                       ' lcid
            row(1) = "English"
            rows.Add row
            SystemViewTable = BuildView("lcid:int,name", rows)

        Case "syspolicy_configuration"
            ' What the policy store reads before it will report a health
            ' state for anything in the tree.
            '
            ' Answering with no rows is not the same as answering no: the
            ' store reads Enabled with a getter for a bit, and a missing row
            ' reaches it as a null it cannot cast, which surfaces at login
            ' as "Specified cast is not valid". Policy management is off
            ' here, which is both the honest answer and the one that stops
            ' SSMS asking to evaluate policies against a worksheet.
            AddNamedValue rows, "Enabled", 0
            AddNamedValue rows, "HistoryRetentionInDays", 0
            AddNamedValue rows, "LogOnSuccess", 0
            AddNamedValue rows, "PurgeHistoryJobGuid", Null
            SystemViewTable = BuildView("name,current_value", rows)

        Case "schemas"
            ReDim row(0 To 2)
            row(0) = "dbo"
            row(1) = 1
            row(2) = 1
            rows.Add row
            SystemViewTable = BuildView("name,schema_id,principal_id", rows)

        Case "indexes"
            ' One row per table, describing it as the heap it is.
            '
            ' The tree's own query inner joins this, so a table with no row
            ' here is a table the tree never draws. A real server has no
            ' such table either: a table with no clustered index is a heap,
            ' and a heap is index_id 0.
            For Each entry In mTables
                ReDim row(0 To 18)
                row(0) = TableObjectId(LCase$(CStr(entry(0))))
                row(1) = Null                       ' name
                row(2) = 0                          ' index_id
                row(3) = 0                          ' type
                row(4) = "HEAP"                     ' type_desc
                row(5) = 1                          ' data_space_id
                row(6) = False                      ' is_primary_key
                row(7) = False                      ' is_unique_constraint
                row(8) = False                      ' is_unique
                row(9) = False                      ' is_disabled
                row(10) = False                     ' is_hypothetical
                row(11) = False                     ' is_padded
                row(12) = 0                         ' fill_factor
                row(13) = True                      ' allow_page_locks
                row(14) = True                      ' allow_row_locks
                row(15) = False                     ' ignore_dup_key
                row(16) = False                     ' has_filter
                row(17) = Null                      ' filter_definition
                row(18) = 0                         ' compression_delay
                rows.Add row
            Next entry
            SystemViewTable = BuildView( _
                "object_id,name,index_id,type,type_desc,data_space_id," & _
                "is_primary_key,is_unique_constraint,is_unique,is_disabled," & _
                "is_hypothetical,is_padded,fill_factor,allow_page_locks," & _
                "allow_row_locks,ignore_dup_key,has_filter," & _
                "filter_definition,compression_delay:int", rows)

        Case Else
            ' Everything else under sys is present and empty.
            '
            ' A workbook has no change tracking, no file tables, no external
            ' data sources and no ledger. An object tree asks about all of
            ' them while drawing itself, and a name it asks for and does not
            ' get ends the connection; one that answers with no rows is a
            ' server that has none of that, which is the truth.
            SystemViewTable = BuildView(EmptyViewColumns(which), rows)
    End Select
End Function

' The partition a table's rows live in, and the allocation unit holding it.
' One number does for both, because there is one of each per table.
Private Function PartitionIdOf(ByVal name As String) As Variant
    PartitionIdOf = CDbl(TableObjectId(LCase$(name))) * 65536#
End Function

' A stable identifier for something that has none of its own.
'
' A real database carries several GUIDs, and a client reads them into a type
' that cannot hold a null: a missing one is "Specified cast is not valid"
' rather than a blank field. These are made out of the name, so the same
' workbook answers with the same identifier every time it is started, and
' two different workbooks do not claim to be the same database.
Private Function StableGuid(ByVal seed As String, ByVal purpose As String) As String
    Dim digits As String
    Dim total As Long
    Dim index As Long
    Dim text As String

    text = LCase$(seed & "/" & purpose)
    For index = 1 To 32
        total = (total * 31 + index + _
                 Asc(Mid$(text, ((index - 1) Mod Len(text)) + 1, 1))) And &HFFFFFF
        digits = digits & Mid$("0123456789abcdef", (total Mod 16) + 1, 1)
    Next index
    StableGuid = Mid$(digits, 1, 8) & "-" & Mid$(digits, 9, 4) & "-" & _
                 Mid$(digits, 13, 4) & "-" & Mid$(digits, 17, 4) & "-" & _
                 Mid$(digits, 21, 12)
End Function

' One row of sys.database_files, in the order that view declares.
Private Sub AddFileRow(ByVal into As Collection, ByVal fileId As Long, _
                       ByVal fileName As String, ByVal leaf As String, _
                       ByVal kind As Long, ByVal kindText As String, _
                       ByVal space As Long)
    Dim row As Variant

    ReDim row(0 To 11)
    row(0) = 1                          ' database_id
    row(1) = fileId
    row(2) = fileName
    row(3) = DEFAULT_DATA_PATH & leaf
    row(4) = kind
    row(5) = kindText
    row(6) = 0                          ' state
    row(7) = "ONLINE"                   ' state_desc
    row(8) = space                      ' data_space_id
    row(9) = 1024                       ' size, in eight kilobyte pages
    row(10) = -1                        ' max_size: unlimited
    row(11) = 10                        ' growth
    into.Add row
End Sub

' One row of sys.configurations. A setting nothing here can change reads
' the same as its value in use.
Private Sub AddConfiguration(ByVal into As Collection, ByVal number As Long, _
                             ByVal name As String, ByVal value As Long)
    Dim row As Variant

    ReDim row(0 To 8)
    row(0) = number
    row(1) = name
    row(2) = name                       ' description
    row(3) = value
    row(4) = 0                          ' minimum
    row(5) = 2147483647                 ' maximum
    row(6) = value                      ' value_in_use
    row(7) = False                      ' is_dynamic
    row(8) = False                      ' is_advanced
    into.Add row
End Sub

' One name-and-value row, for the configuration views that are shaped that
' way.
Private Sub AddNamedValue(ByVal into As Collection, ByVal name As String, _
                          ByVal value As Variant)
    Dim row As Variant

    ReDim row(0 To 1)
    row(0) = name
    row(1) = value
    into.Add row
End Sub

' One row of sys.types, in the order EmptyViewColumns("types") names.
'
' The fifth parameter is not called scale. Scale is a VB drawing statement,
' so a procedure that declares one stops resolving and every call to it is
' reported as a Sub that was never defined.
Private Sub AddTypeRow(ByVal into As Collection, ByVal typeName As String, _
                       ByVal typeId As Long, ByVal widest As Long, _
                       ByVal digits As Long, ByVal places As Long)
    Dim row As Variant

    ReDim row(0 To 10)
    row(0) = typeName
    row(1) = typeId                     ' system_type_id
    row(2) = typeId                     ' user_type_id
    row(3) = 4                          ' schema_id, sys
    row(4) = widest                     ' max_length
    row(5) = digits                     ' precision
    row(6) = places                     ' scale
    row(7) = True                       ' is_nullable
    row(8) = False                      ' is_user_defined
    row(9) = False                      ' is_table_type
    row(10) = False                     ' is_assembly_type
    into.Add row
End Sub

' What a system view with no rows still has to declare.
'
' An empty view is not free of its shape: a left join against it puts its
' columns on every row it keeps, and the client reads them by name. Each
' list is what an object tree asks that view for, read out of this server's
' own log of an SMO run rather than guessed. A column carrying a type is one
' a client reads with a typed getter, which a column of nothing but nulls
' would otherwise answer as text.
Private Function EmptyViewColumns(ByVal which As String) As String
    Select Case which
        Case "types", "all_types"
            EmptyViewColumns = "name,system_type_id,user_type_id," & _
                "schema_id,max_length,precision,scale,is_nullable," & _
                "is_user_defined,is_table_type,is_assembly_type"

        Case "data_spaces", "partition_schemes"
            EmptyViewColumns = "data_space_id:int,name,type,type_desc," & _
                "is_default:bit,is_system:bit"

        Case "database_principals", "server_principals"
            EmptyViewColumns = "principal_id:int,name,type,type_desc,sid," & _
                "default_schema_name,is_fixed_role:bit," & _
                "create_date:datetime,modify_date:datetime"

        Case "periods"
            EmptyViewColumns = "object_id:int,name,period_type:tinyint," & _
                "period_type_desc,start_column_id:int," & _
                "end_column_id:int"

        Case "extended_properties"
            EmptyViewColumns = "class:tinyint,class_desc,major_id:int," & _
                "minor_id:int,name,value"

        Case "partitions"
            EmptyViewColumns = "partition_id:bigint,object_id:int," & _
                "index_id:int,partition_number:int,rows:bigint," & _
                "used_page_count,data_compression:tinyint," & _
                "data_compression_desc"

        Case "allocation_units"
            EmptyViewColumns = "allocation_unit_id:bigint,type:tinyint," & _
                "type_desc,container_id:bigint,data_space_id:int," & _
                "total_pages:bigint,used_pages:bigint," & _
                "data_pages:bigint"

        Case "dm_db_partition_stats"
            EmptyViewColumns = "partition_id:bigint,object_id:int," & _
                "index_id:int,partition_number:int," & _
                "in_row_data_page_count:bigint,used_page_count:bigint," & _
                "reserved_page_count:bigint,row_count:bigint"

        Case "internal_tables"
            EmptyViewColumns = "object_id:int,name,internal_type:tinyint," & _
                "internal_type_desc,parent_id:int,parent_minor_id:int"

        Case "dm_db_xtp_table_memory_stats"
            EmptyViewColumns = "object_id,memory_allocated_for_table_kb," & _
                "memory_used_by_table_kb," & _
                "memory_allocated_for_indexes_kb," & _
                "memory_used_by_indexes_kb"

        Case "external_tables"
            EmptyViewColumns = "object_id,name,data_source_id," & _
                "file_format_id,location,reject_type,reject_value," & _
                "reject_sample_value,distribution_type," & _
                "remote_object_name,remote_schema_name,sharding_col_id"

        Case "external_data_sources"
            EmptyViewColumns = "data_source_id,name,location,type," & _
                "type_desc,credential_id,database_name,shard_map_name"

        Case "external_file_formats"
            EmptyViewColumns = "file_format_id,name,format_type"

        Case "filetables"
            EmptyViewColumns = "object_id,is_enabled,directory_name," & _
                "filename_collation_id,filename_collation_name"

        Case "change_tracking_tables"
            EmptyViewColumns = "object_id,is_track_columns_updated_on," & _
                "begin_version,min_valid_version,cleanup_version"

        Case "table_types"
            EmptyViewColumns = "name,type_table_object_id:int," & _
                "is_memory_optimized:bit,schema_id:int," & _
                "user_type_id:int,system_type_id:tinyint," & _
                "is_user_defined:bit,is_table_type:bit"

        Case "index_columns"
            EmptyViewColumns = "object_id:int,index_id:int," & _
                "index_column_id:int,column_id:int,key_ordinal:tinyint," & _
                "is_descending_key:bit,is_included_column:bit"

        Case "computed_columns"
            EmptyViewColumns = "object_id:int,name,column_id:int," & _
                "definition,is_persisted:bit,is_computed:bit"

        Case "identity_columns"
            EmptyViewColumns = "object_id:int,name,column_id:int," & _
                "seed_value,increment_value,last_value," & _
                "is_not_for_replication:bit"

        Case "masked_columns"
            EmptyViewColumns = "object_id:int,name,column_id:int," & _
                "is_masked:bit,masking_function"

        Case "sensitivity_classifications"
            EmptyViewColumns = "class,class_desc,major_id,minor_id,label," & _
                "label_id,information_type,information_type_id," & _
                "rank:int,rank_desc"

        Case "column_encryption_keys"
            EmptyViewColumns = "name,column_encryption_key_id:int," & _
                "create_date:datetime,modify_date:datetime"

        Case "xml_schema_collections"
            EmptyViewColumns = "xml_collection_id:int,schema_id:int," & _
                "principal_id:int,name,create_date:datetime," & _
                "modify_date:datetime"

        Case "default_constraints"
            EmptyViewColumns = "object_id:int,name,parent_object_id:int," & _
                "parent_column_id:int,definition,is_system_named:bit"

        Case "foreign_key_columns"
            EmptyViewColumns = "constraint_object_id:int," & _
                "constraint_column_id:int,parent_object_id:int," & _
                "parent_column_id:int,referenced_object_id:int," & _
                "referenced_column_id:int"

        Case "database_files", "master_files"
            EmptyViewColumns = "database_id:int,file_id:int,name," & _
                "physical_name,type:tinyint,type_desc,size:int," & _
                "max_size:int,growth:int,state:tinyint,state_desc," & _
                "data_space_id:int"

        Case "parameters", "all_parameters"
            EmptyViewColumns = "object_id:int,name,parameter_id:int," & _
                "system_type_id:tinyint,user_type_id:int," & _
                "max_length:smallint,precision:tinyint,scale:tinyint," & _
                "is_output:bit,is_xml_document:bit," & _
                "xml_collection_id:int,vector_dimensions:int," & _
                "vector_base_type_desc,default_value," & _
                "has_default_value:bit"

        Case "sql_modules", "all_sql_modules", "system_sql_modules"
            EmptyViewColumns = "object_id:int,definition," & _
                "uses_ansi_nulls:bit,uses_quoted_identifier:bit," & _
                "is_schema_bound:bit,is_recompiled:bit," & _
                "null_on_null_input:bit,execute_as_principal_id:int," & _
                "uses_native_compilation:bit,is_inlineable:bit," & _
                "inline_type:bit"

        Case "assemblies"
            EmptyViewColumns = "assembly_id:int,name,principal_id:int," & _
                "clr_name,permission_set:tinyint,is_visible:bit," & _
                "create_date:datetime"

        Case "assembly_modules"
            EmptyViewColumns = "object_id:int,assembly_id:int," & _
                "assembly_class,assembly_method," & _
                "execute_as_principal_id:int,null_on_null_input:bit"

        Case "certificates", "asymmetric_keys", "symmetric_keys"
            EmptyViewColumns = "name,sid,principal_id,thumbprint,subject," & _
                "issuer_name,start_date,expiry_date"

        Case "database_permissions", "server_permissions"
            EmptyViewColumns = "class,class_desc,major_id,minor_id," & _
                "grantee_principal_id,grantor_principal_id,type," & _
                "permission_name,state,state_desc"

        Case "hash_indexes"
            EmptyViewColumns = "object_id,index_id,name,bucket_count"

        Case "index_resumable_operations"
            EmptyViewColumns = "object_id,index_id,name,sql_text," & _
                "last_max_dop_used,partition_number,state,state_desc," & _
                "percent_complete"

        Case "key_constraints"
            EmptyViewColumns = "object_id:int,name,parent_object_id:int," & _
                "schema_id:int,type,type_desc,unique_index_id:int," & _
                "is_system_named:bit"

        Case "procedures"
            EmptyViewColumns = "object_id:int,name,schema_id:int,type," & _
                "type_desc,is_auto_executed:bit,is_ms_shipped:bit," & _
                "create_date:datetime,modify_date:datetime"

        Case "selective_xml_index_paths"
            EmptyViewColumns = "object_id,index_id,path_id,name,path," & _
                "path_type,path_type_desc,xml_component_id," & _
                "xquery_type_description,is_xquery_type_inferred," & _
                "xquery_max_length,is_xquery_max_length_inferred," & _
                "is_node,system_type_id,user_type_id,max_length," & _
                "precision,scale"

        Case "spatial_index_tessellations"
            EmptyViewColumns = "object_id,index_id,tessellation_scheme," & _
                "bounding_box_xmin,bounding_box_ymin,bounding_box_xmax," & _
                "bounding_box_ymax,level_1_grid,level_1_grid_desc," & _
                "level_2_grid,level_2_grid_desc,level_3_grid," & _
                "level_3_grid_desc,level_4_grid,level_4_grid_desc," & _
                "cells_per_object"

        Case "spatial_indexes"
            EmptyViewColumns = "object_id:int,index_id:int,name," & _
                "spatial_index_type:int,spatial_index_type_desc," & _
                "tessellation_scheme"

        Case "stats"
            EmptyViewColumns = "object_id:int,name,stats_id:int," & _
                "auto_created:bit,auto_drop:bit,user_created:bit," & _
                "no_recompute:bit,has_filter:bit,filter_definition," & _
                "is_temporary:bit,is_incremental:bit"

        Case "trigger_events"
            EmptyViewColumns = "object_id:int,type:int,type_desc," & _
                "is_first:bit,is_last:bit,event_group_type:int," & _
                "event_group_type_desc"

        Case "triggers"
            EmptyViewColumns = "object_id:int,name,parent_class:tinyint," & _
                "parent_class_desc,parent_id:int,type,type_desc," & _
                "create_date:datetime,modify_date:datetime," & _
                "is_ms_shipped:bit,is_disabled:bit," & _
                "is_not_for_replication:bit,is_instead_of_trigger:bit"

        Case "xml_indexes"
            EmptyViewColumns = "object_id:int,index_id:int,name," & _
                "using_xml_index_id:int,secondary_type," & _
                "secondary_type_desc,xml_index_type:tinyint," & _
                "xml_index_type_description,path_id:int"

        Case "json_indexes", "vector_indexes"
            EmptyViewColumns = "object_id,index_id," & _
                "optimize_for_array_search,distance_metric," & _
                "vector_index_type"

        Case "filetable_system_defined_objects"
            EmptyViewColumns = "object_id,parent_object_id"

        Case "dm_os_host_info"
            EmptyViewColumns = "host_platform,host_distribution," & _
                "host_release,host_service_pack_level,host_sku," & _
                "os_language_version"

        Case "dm_exec_sessions", "dm_exec_connections", "dm_exec_requests"
            EmptyViewColumns = "session_id,connection_id,login_name," & _
                "host_name,program_name,status,database_id," & _
                "client_net_address,net_transport,protocol_type"

        Case "change_tracking_databases"
            EmptyViewColumns = "database_id:int," & _
                "is_auto_cleanup_on:tinyint,retention_period:int," & _
                "retention_period_units:tinyint," & _
                "retention_period_units_desc," & _
                "max_cleanup_version:bigint"

        Case "database_recovery_status"
            EmptyViewColumns = "database_id:int," & _
                "database_guid:uniqueidentifier," & _
                "family_guid:uniqueidentifier," & _
                "last_log_backup_lsn:numeric(25,0)," & _
                "recovery_fork_guid:uniqueidentifier," & _
                "first_recovery_fork_guid:uniqueidentifier," & _
                "fork_point_lsn:numeric(25,0)"

        Case "database_mirroring"
            EmptyViewColumns = "database_id:int," & _
                "mirroring_guid:uniqueidentifier," & _
                "mirroring_state:tinyint,mirroring_state_desc," & _
                "mirroring_role:tinyint,mirroring_role_desc," & _
                "mirroring_role_sequence:int," & _
                "mirroring_safety_level:tinyint," & _
                "mirroring_safety_level_desc," & _
                "mirroring_safety_sequence:int,mirroring_partner_name," & _
                "mirroring_partner_instance,mirroring_witness_name," & _
                "mirroring_witness_state:tinyint," & _
                "mirroring_witness_state_desc," & _
                "mirroring_failover_lsn:numeric(25,0)," & _
                "mirroring_connection_timeout:int," & _
                "mirroring_redo_queue:int,mirroring_redo_queue_type," & _
                "mirroring_end_of_log_lsn:numeric(25,0)," & _
                "mirroring_replication_lsn:numeric(25,0)"

        Case "database_filestream_options"
            EmptyViewColumns = "database_id:int," & _
                "non_transacted_access:tinyint," & _
                "non_transacted_access_desc,directory_name"

        Case "fulltext_catalogs"
            EmptyViewColumns = "fulltext_catalog_id:int,name,path," & _
                "is_default:bit,is_accent_sensitivity_on:bit," & _
                "data_space_id:int,file_id:int,principal_id:int," & _
                "is_importing:bit"

        Case "system_objects"
            EmptyViewColumns = "name,object_id,principal_id,schema_id," & _
                "parent_object_id,type,type_desc,create_date," & _
                "modify_date,is_ms_shipped"

        Case "availability_replicas"
            EmptyViewColumns = "replica_id,group_id,replica_metadata_id," & _
                "replica_server_name,owner_sid,endpoint_url," & _
                "availability_mode,availability_mode_desc," & _
                "failover_mode,failover_mode_desc"

        Case "availability_groups", "dm_hadr_database_replica_states"
            EmptyViewColumns = "group_id,group_database_id,name," & _
                "database_id,replica_id,is_local,synchronization_state," & _
                "synchronization_state_desc"

        Case "server_resource_stats"
            EmptyViewColumns = "start_time,end_time,virtual_core_count," & _
                "hardware_generation,sku,reserved_storage_mb," & _
                "storage_space_used_mb"

        Case "configurations"
            EmptyViewColumns = "configuration_id,name,value,minimum," & _
                "maximum,value_in_use,description,is_dynamic," & _
                "is_advanced"

        Case "dm_tran_persistent_version_store_stats"
            EmptyViewColumns = "database_id:int,pvs_filegroup_id:smallint," & _
                "persistent_version_store_size_kb:bigint," & _
                "online_index_version_store_size_kb:bigint," & _
                "current_aborted_transaction_count:bigint," & _
                "oldest_active_transaction_id:bigint," & _
                "oldest_active_transaction_global_id:bigint," & _
                "oldest_aborted_transaction_id:bigint," & _
                "min_transaction_timestamp:bigint," & _
                "online_index_min_transaction_timestamp:bigint," & _
                "secondary_low_water_mark:bigint," & _
                "offrow_version_cleaner_start_time:datetime," & _
                "offrow_version_cleaner_end_time:datetime," & _
                "aborted_version_cleaner_start_time:datetime," & _
                "aborted_version_cleaner_end_time:datetime," & _
                "pvs_off_row_page_skipped_low_water_mark:bigint," & _
                "pvs_off_row_page_skipped_transaction_not_cleaned:bigint," & _
                "pvs_off_row_page_skipped_oldest_active_xdesid:bigint," & _
                "pvs_off_row_page_skipped_min_useful_xts:bigint," & _
                "pvs_off_row_page_skipped_oldest_snapshot:bigint," & _
                "pvs_off_row_page_skipped_oldest_aborted_xdesid:bigint"

        Case "sysprocesses"
            EmptyViewColumns = "spid,dbid,uid,cpu,physical_io,login_time," & _
                "last_batch,status,hostname,program_name,cmd,loginame," & _
                "nt_username"

        Case "database_query_store_options"
            EmptyViewColumns = "desired_state:smallint,desired_state_desc," & _
                "actual_state:smallint,actual_state_desc," & _
                "readonly_reason:int,current_storage_size_mb:bigint," & _
                "flush_interval_seconds:bigint," & _
                "interval_length_minutes:bigint," & _
                "max_storage_size_mb:bigint," & _
                "stale_query_threshold_days:bigint," & _
                "max_plans_per_query:bigint," & _
                "query_capture_mode:smallint,query_capture_mode_desc," & _
                "capture_policy_execution_count:int," & _
                "capture_policy_total_compile_cpu_time_ms:bigint," & _
                "capture_policy_total_execution_cpu_time_ms:bigint," & _
                "capture_policy_stale_threshold_hours:int," & _
                "size_based_cleanup_mode:smallint," & _
                "size_based_cleanup_mode_desc," & _
                "wait_stats_capture_mode:smallint," & _
                "wait_stats_capture_mode_desc," & _
                "actual_state_additional_info"

        Case "filegroups"
            EmptyViewColumns = "name,data_space_id:int,type,type_desc," & _
                "is_default:bit,is_system:bit," & _
                "filegroup_guid:uniqueidentifier,log_filegroup_id:int," & _
                "is_read_only:bit,is_autogrow_all_files:bit"

        Case Else
            EmptyViewColumns = "name,object_id,schema_id,principal_id," & _
                "type,type_desc,create_date,modify_date,is_ms_shipped"
    End Select
End Function

' A view as the two-dimensional array every other source arrives as: names
' in the first row, values under them.
Private Function BuildView(ByVal namesCsv As String, _
                           ByVal rows As Collection) As Variant
    Dim names() As String
    Dim data As Variant
    Dim index As Long
    Dim rowIndex As Long
    Dim values As Variant

    ' A name may carry the type it is, written after a colon. Without one
    ' a column of nothing but nulls reads as text, and a client that asks
    ' for it as a number gets a cast error instead of a null.
    names = Split(namesCsv, ",")
    ReDim data(1 To rows.count + 1, 1 To UBound(names) + 1)

    For index = 0 To UBound(names)
        data(1, index + 1) = Trim$(names(index))
    Next index

    ' For Each rather than an index. A Collection asked for its nth item
    ' walks to it, so a rowset of three hundred rows read the list forty
    ' thousand links deep to copy it out.
    rowIndex = 1
    For Each values In rows
        For index = 0 To UBound(names)
            If index <= UBound(values) Then
                data(rowIndex + 1, index + 1) = values(LBound(values) + index)
            End If
        Next index
        rowIndex = rowIndex + 1
    Next values
    BuildView = data
End Function

' xp_msver: what the server says about itself, one property per row.
'
' SSMS fills a temp table from this while connecting and reads the version
' out of it. The values are this process's where they can be, and say what
' answered them where they cannot.
Private Sub ServerVersionRowset(ByRef outColumns As Collection, _
                                ByRef outRows As Collection)
    Dim columns As Collection
    Dim rows As Collection
    Dim version As String

    version = SERVER_VERSION_MAJOR & "." & SERVER_VERSION_MINOR & "." & _
              SERVER_VERSION_BUILD & ".0"

    Set columns = New Collection
    columns.Add CatalogColumn("Index", TYPE_INTN, 4)
    columns.Add CatalogColumn("Name", TYPE_NVARCHAR, 128)
    columns.Add CatalogColumn("Internal_Value", TYPE_INTN, 4)
    columns.Add CatalogColumn("Character_Value", TYPE_NVARCHAR, 4000)

    Set rows = New Collection
    AddVersionRow rows, 1, "ProductName", Null, "Microsoft SQL Server"
    AddVersionRow rows, 2, "ProductVersion", Null, version
    AddVersionRow rows, 3, "Language", 1033, "English (United States)"
    AddVersionRow rows, 4, "Platform", Null, "NT x64"
    AddVersionRow rows, 5, "Comments", Null, "vbaSQLBridge"
    AddVersionRow rows, 6, "CompanyName", Null, "vbaSQLBridge"
    AddVersionRow rows, 7, "FileDescription", Null, "vbaSQLBridge in Excel"
    AddVersionRow rows, 8, "FileVersion", Null, version
    AddVersionRow rows, 9, "InternalName", Null, "vbaSQLBridge"
    AddVersionRow rows, 10, "LegalCopyright", Null, ""
    AddVersionRow rows, 11, "LegalTrademarks", Null, ""
    AddVersionRow rows, 12, "OriginalFilename", Null, "vbaSQLBridge"
    AddVersionRow rows, 13, "PrivateBuild", Null, ""
    AddVersionRow rows, 14, "SpecialBuild", Null, ""
    AddVersionRow rows, 15, "WindowsVersion", Null, "10.0"
    AddVersionRow rows, 16, "ProcessorCount", 1, "1"
    AddVersionRow rows, 17, "ProcessorActiveMask", Null, ""
    AddVersionRow rows, 18, "ProcessorType", Null, "x64"
    AddVersionRow rows, 19, "PhysicalMemory", Null, ""
    AddVersionRow rows, 20, "Product ID", Null, "vbaSQLBridge"

    Set outColumns = columns
    Set outRows = rows
End Sub

Private Sub AddVersionRow(ByVal rows As Collection, ByVal ordinal As Long, _
                          ByVal name As String, ByVal internal As Variant, _
                          ByVal value As String)
    Dim row As Variant

    ReDim row(0 To 3)
    row(0) = ordinal
    row(1) = name
    row(2) = internal
    row(3) = value
    rows.Add row
End Sub


' ============================================================================
' Batches
'
' A client asks what it is talking to before it will show anything, and it
' asks in one batch: make a temp table, fill it from a procedure, declare a
' variable, choose between two statements, then read something worked out
' from all of it. Nothing but the reads produces rows, and every read in the
' batch goes back, because a client counts them: SSMS opens a connection
' with three in one batch and takes the third.
'
' The splitting follows pySQLbridge, which worked it out against the batches
' SSMS actually sends. T-SQL needs no semicolon between statements, so a word
' that nothing else can be followed by is where the next one begins.
' ============================================================================

Private Function IsWordIn(ByVal word As String, ByVal list As String) As Boolean
    IsWordIn = (InStr(1, list, "|" & UCase$(word) & "|", vbBinaryCompare) > 0)
End Function

' Everything outside comments. A comment can hold a semicolon or a quote, so
' it has to go before anything else looks at the text.
Private Function StripComments(ByVal sql As String) As String
    Dim out As String
    Dim at As Long
    Dim length As Long
    Dim ch As String

    length = Len(sql)
    at = 1
    Do While at <= length
        ch = Mid$(sql, at, 1)
        If ch = "'" Or ch = """" Then
            Dim closing As Long
            closing = SkipQuoted(sql, at, ch)
            out = out & Mid$(sql, at, closing - at)
            at = closing
        ElseIf ch = "[" Then
            closing = InStr(at, sql, "]")
            If closing = 0 Then closing = length
            out = out & Mid$(sql, at, closing - at + 1)
            at = closing + 1
        ElseIf Mid$(sql, at, 2) = "--" Then
            Do While at <= length And Mid$(sql, at, 1) <> vbLf
                at = at + 1
            Loop
            out = out & " "
        ElseIf Mid$(sql, at, 2) = "/*" Then
            at = at + 2
            Do While at <= length And Mid$(sql, at, 2) <> "*/"
                at = at + 1
            Loop
            at = at + 2
            out = out & " "
        Else
            out = out & ch
            at = at + 1
        End If
    Loop
    StripComments = out
End Function

' Past a quoted run, doubled quotes and all.
Private Function SkipQuoted(ByVal sql As String, ByVal at As Long, _
                            ByVal quote As String) As Long
    Dim length As Long

    length = Len(sql)
    at = at + 1
    Do While at <= length
        If Mid$(sql, at, 1) = quote Then
            If Mid$(sql, at, 2) = quote & quote Then
                at = at + 2
            Else
                SkipQuoted = at + 1
                Exit Function
            End If
        Else
            at = at + 1
        End If
    Loop
    SkipQuoted = length + 1
End Function

' Past a quoted run or a bracketed name, or one character.
Private Function PastOne(ByVal sql As String, ByVal at As Long) As Long
    Dim ch As String
    Dim closing As Long

    ch = Mid$(sql, at, 1)
    If ch = "'" Or ch = """" Then
        PastOne = SkipQuoted(sql, at, ch)
    ElseIf ch = "[" Then
        closing = InStr(at, sql, "]")
        PastOne = IIf(closing = 0, Len(sql) + 1, closing + 1)
    Else
        PastOne = at + 1
    End If
End Function

Private Function SkipSpaceAt(ByVal sql As String, ByVal at As Long) As Long
    Do While at <= Len(sql)
        Select Case Mid$(sql, at, 1)
            Case " ", vbTab, vbCr, vbLf
                at = at + 1
            Case Else
                Exit Do
        End Select
    Loop
    SkipSpaceAt = at
End Function

' The word starting at this position, with where it ends.
Private Function WordAt(ByVal sql As String, ByVal at As Long, _
                        ByRef wordEnd As String) As String
    Dim ends As Long
    Dim ch As String

    wordEnd = vbNullString

    ' A name in brackets is one word, brackets included. Stopping at the
    ' opening bracket returns nothing at all, and a CREATE TABLE whose
    ' column is written [location] then declares no such column: the table
    ' is built without it and the first row that reads it is the first
    ' anyone hears of it.
    If Mid$(sql, at, 1) = "[" Then
        ends = InStr(at, sql, "]")
        If ends = 0 Then Exit Function
        WordAt = Mid$(sql, at, ends - at + 1)
        wordEnd = CStr(ends + 1)
        Exit Function
    End If

    ends = at
    Do While ends <= Len(sql)
        ch = Mid$(sql, ends, 1)
        If (UCase$(ch) >= "A" And UCase$(ch) <= "Z") Or ch = "_" _
           Or ch = "@" Or ch = "#" Or (ch >= "0" And ch <= "9" And ends > at) Then
            ends = ends + 1
        Else
            Exit Do
        End If
    Loop
    If ends = at Then Exit Function
    WordAt = Mid$(sql, at, ends - at)
    wordEnd = CStr(ends)
End Function

' Where the next of these words begins, skipping anything quoted, or zero.
Private Function NextWordIn(ByVal sql As String, ByVal at As Long, _
                            ByVal wanted As String) As Long
    Dim depth As Long
    Dim ch As String
    Dim word As String
    Dim ends As String

    Do While at <= Len(sql)
        ch = Mid$(sql, at, 1)
        If ch = "'" Or ch = """" Or ch = "[" Then
            at = PastOne(sql, at)
        ElseIf ch = "(" Then
            depth = depth + 1
            at = at + 1
        ElseIf ch = ")" Then
            If depth > 0 Then depth = depth - 1
            at = at + 1
        ElseIf depth = 0 Then
            word = WordAt(sql, at, ends)
            If Len(word) > 0 Then
                If IsWordIn(word, wanted) Then
                    If Not PartOfAStatement(sql, word, CLng(ends)) Then
                        NextWordIn = at
                        Exit Function
                    End If
                End If
                at = CLng(ends)
            Else
                at = at + 1
            End If
        Else
            at = at + 1
        End If
    Loop
End Function

' Whether this word continues the statement rather than starting one.
'
' INSERT ... EXEC is one statement: the rows the procedure returns are what
' is inserted. SELECT begins a statement after a DECLARE or a SET, which end
' where their value does, and after a statement that already holds a SELECT
' of its own; a set operator keeps two of them together. SET is the other way
' round, beginning a statement everywhere except in an UPDATE.
' Whether the word here opens a block that runs as one statement.
'
' BEGIN TRANSACTION opens nothing: it is a statement in its own right, and
' reading it as a block swallows everything up to some unrelated END.
Private Function OpensABlock(ByVal sql As String, ByVal at As Long, _
                             ByVal word As String) As Boolean
    Dim ends As String
    Dim after As String

    If UCase$(word) <> "BEGIN" Then Exit Function

    WordAt sql, at, ends
    If Len(ends) = 0 Then Exit Function
    after = UCase$(WordAt(sql, SkipSpaceAt(sql, CLng(ends)), ends))

    Select Case after
        Case "TRANSACTION", "TRAN", "DISTRIBUTED", "CONVERSATION", "DIALOG"
        Case Else
            OpensABlock = True
    End Select
End Function

' Whether a word at the top level of a batch begins a statement.
Private Function StartsAStatement(ByVal sql As String, ByVal word As String, _
                                  ByVal after As Long) As Boolean
    If Not IsWordIn(word, STATEMENT_STARTS) Then Exit Function
    StartsAStatement = Not PartOfAStatement(sql, word, after)
End Function

' Whether WITH or SAVE, where it stands, is part of the statement around it
' rather than the start of another.
'
' WITH begins a statement only as a common table expression's list, WITH
' name AS (...) or WITH name (columns) AS (...). After a table it is a hint,
' and elsewhere WITH TIES, WITH ROLLUP or WITH MARK. Taken for the start of
' a statement, it cut the one it was in: SELECT ... FROM t WITH (NOLOCK)
' WHERE x answered every row without its WHERE and then failed on the rest.
' SAVE begins one only before TRAN or TRANSACTION, and is as likely a column.
Private Function PartOfAStatement(ByVal sql As String, ByVal word As String, _
                                  ByVal after As Long) As Boolean
    Dim at As Long
    Dim ends As String
    Dim name As String

    Select Case UCase$(word)
        Case "WITH"
            at = SkipSpaceAt(sql, after)
            name = WordAt(sql, at, ends)
            If Len(name) = 0 Then
                PartOfAStatement = True
                Exit Function
            End If
            at = SkipSpaceAt(sql, CLng(ends))
            If Mid$(sql, at, 1) <> "(" Then
                PartOfAStatement = (UCase$(WordAt(sql, at, ends)) <> "AS")
            End If
        Case "SAVE"
            Select Case UCase$(WordAt(sql, SkipSpaceAt(sql, after), ends))
                Case "TRAN", "TRANSACTION"
                Case Else
                    PartOfAStatement = True
            End Select
    End Select
End Function

Private Function BelongsToIt(ByVal soFar As String, ByVal word As String) As Boolean
    Dim before As String
    Dim trailing As String
    Dim ends As String

    before = Trim$(soFar)

    ' The first statement after a WITH list belongs to it: the names the
    ' list declares are there for that statement alone. Split from it, WITH
    ' ... UPDATE ran as a list with nothing after it, and the UPDATE on its
    ' own named a table that did not exist.
    If MainStatementWord(before) = "WITH" Then
        BelongsToIt = IsWordIn(word, "|SELECT|INSERT|UPDATE|DELETE|MERGE|")
        Exit Function
    End If

    ' A MERGE holds an UPDATE SET, an INSERT or a DELETE in each of its WHEN
    ' clauses, and ends at the semicolon it has to end with.
    If MainStatementWord(before) = "MERGE" Then
        BelongsToIt = IsWordIn(word, "|UPDATE|SET|INSERT|DELETE|")
        Exit Function
    End If

    If UCase$(word) = "SELECT" Then
        trailing = LastWordOf(before)
        If IsWordIn(trailing, JOINS_TWO_SELECTS) Then
            BelongsToIt = True
            Exit Function
        End If
        If UCase$(trailing) = "ALL" Then
            If IsWordIn(SecondLastWordOf(before), JOINS_TWO_SELECTS) Then
                BelongsToIt = True
                Exit Function
            End If
        End If
        If IsWordIn(WordAt(before, 1, ends), HOLDS_NO_SELECT) Then Exit Function
        BelongsToIt = (NextWordIn(before, 1, "|SELECT|") = 0)
        Exit Function
    End If

    If UCase$(word) = "BEGIN" Then
        ' BEGIN CATCH belongs to the END TRY in front of it. The two halves
        ' of a guarded block are one statement, and split between them the
        ' handler is a statement opening with the word CATCH that nothing
        ' runs, so the failure the client guarded against goes unhandled.
        If UCase$(LastWordOf(before)) = "TRY" Then
            BelongsToIt = (UCase$(SecondLastWordOf(before)) = "END")
        End If
        Exit Function
    End If

    If UCase$(word) = "SET" Then
        BelongsToIt = (MainStatementWord(before) = "UPDATE")
        Exit Function
    End If

    If UCase$(word) = "EXEC" Or UCase$(word) = "EXECUTE" Then
        BelongsToIt = (MainStatementWord(before) = "INSERT")
    End If
End Function

' Which statement a piece of a batch is: its first word, or for one that
' opens with a WITH list, the first statement after the list.
Private Function MainStatementWord(ByVal soFar As String) As String
    Dim ends As String
    Dim at As Long

    MainStatementWord = UCase$(WordAt(soFar, SkipSpaceAt(soFar, 1), ends))
    If MainStatementWord <> "WITH" Then Exit Function
    at = NextWordIn(soFar, 1, "|SELECT|INSERT|UPDATE|DELETE|MERGE|")
    If at > 0 Then MainStatementWord = UCase$(WordAt(soFar, at, ends))
End Function

' Split of an empty string is an array with nothing in it, not an array
' holding one empty piece, so the last word of nothing has to be asked for
' carefully.
Private Function LastWordOf(ByVal text As String) As String
    Dim parts() As String

    parts = Split(Trim$(Replace(Replace(text, vbCr, " "), vbLf, " ")), " ")
    If UBound(parts) < LBound(parts) Then Exit Function
    LastWordOf = parts(UBound(parts))
End Function

Private Function SecondLastWordOf(ByVal text As String) As String
    Dim parts() As String

    parts = Split(Trim$(Replace(Replace(text, vbCr, " "), vbLf, " ")), " ")
    If UBound(parts) < LBound(parts) + 1 Then Exit Function
    SecondLastWordOf = parts(UBound(parts) - 1)
End Function

' One batch split into the statements it holds.
Private Function SplitStatements(ByVal sql As String) As Collection
    Dim found As Collection
    Dim at As Long
    Dim start As Long
    Dim depth As Long
    Dim ch As String
    Dim word As String
    Dim ends As String
    Dim breaks As Boolean

    Set found = New Collection
    at = 1
    start = 1

    Do While at <= Len(sql)
        ch = Mid$(sql, at, 1)
        If ch = "'" Or ch = """" Or ch = "[" Then
            at = PastOne(sql, at)
        ElseIf ch = ";" Then
            AddStatement found, Mid$(sql, start, at - start)
            at = at + 1
            start = at
        ElseIf depth = 0 Then
            word = WordAt(sql, at, ends)
            If Len(word) > 0 Then
                If UCase$(word) = "IF" Then
                    ' An IF holds its branches, ELSE and all. What comes
                    ' after it is a statement of its own.
                    If at > start Then
                        AddStatement found, Mid$(sql, start, at - start)
                        start = at
                    End If
                    at = EndOfIf(sql, at)
                    AddStatement found, Mid$(sql, start, at - start)
                    start = at
                ElseIf UCase$(word) = "WHILE" Then
                    ' A WHILE holds its body the same way. Split off, the
                    ' body is a statement of its own and runs once, which
                    ' is a loop that does not loop.
                    If at > start Then
                        AddStatement found, Mid$(sql, start, at - start)
                        start = at
                    End If
                    at = EndOfWhile(sql, at)
                    AddStatement found, Mid$(sql, start, at - start)
                    start = at
                ElseIf OpensABlock(sql, at, word) Then
                    ' A block is one statement too, everything it holds
                    ' included. Splitting inside it runs the statements of a
                    ' branch on their own: the handler of a BEGIN TRY breaks
                    ' at the IF inside it and then runs whatever the failure
                    ' was going to be handled with, having had no failure.
                    If at > start Then
                        AddStatement found, Mid$(sql, start, at - start)
                        start = at
                    End If
                    at = EndOfBranch(sql, at)
                    AddStatement found, Mid$(sql, start, at - start)
                    start = at
                Else
                    ' Nested rather than one condition, because VBA evaluates
                    ' both sides of And: the guard has to actually guard.
                    breaks = False
                    If at > start Then
                        If StartsAStatement(sql, word, CLng(ends)) Then
                            breaks = Not BelongsToIt(Mid$(sql, start, at - start), _
                                                     word)
                        End If
                    End If
                    If breaks Then
                        AddStatement found, Mid$(sql, start, at - start)
                        start = at
                    End If
                    at = CLng(ends)
                End If
            ElseIf ch = "(" Then
                depth = depth + 1
                at = at + 1
            ElseIf ch = ")" Then
                If depth > 0 Then depth = depth - 1
                at = at + 1
            Else
                at = at + 1
            End If
        Else
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                If depth > 0 Then depth = depth - 1
            End If
            word = WordAt(sql, at, ends)
            If Len(word) > 0 Then
                at = CLng(ends)
            Else
                at = at + 1
            End If
        End If
    Loop

    AddStatement found, Mid$(sql, start)
    Set SplitStatements = found
End Function

Private Sub AddStatement(ByVal found As Collection, ByVal text As String)
    Dim trimmed As String

    trimmed = TrimWhitespace(text)
    If Len(trimmed) > 0 Then found.Add trimmed
End Sub

' Trim, including the line breaks Trim leaves behind.
'
' VBA's Trim removes spaces and nothing else, so a statement that began on
' the line after a semicolon kept its newline, had no first word, and was
' run by nothing: a batch of three reads answered with one.
Private Function TrimWhitespace(ByVal text As String) As String
    Dim first As Long
    Dim last As Long

    first = 1
    last = Len(text)
    Do While first <= last
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(text, first, 1)) = 0 Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(text, last, 1)) = 0 Then Exit Do
        last = last - 1
    Loop
    If last >= first Then TrimWhitespace = Mid$(text, first, last - first + 1)
End Function

' Where an IF stops, branches and all.
' Where a WHILE ends: past its condition, then past the one statement that
' is its body.
Private Function EndOfWhile(ByVal sql As String, ByVal at As Long) As Long
    Dim ends As String

    WordAt sql, at, ends
    at = NextWordIn(sql, CLng(ends), STATEMENT_STARTS)
    If at = 0 Then
        EndOfWhile = Len(sql) + 1
        Exit Function
    End If
    EndOfWhile = EndOfBranch(sql, at)
End Function

Private Function EndOfIf(ByVal sql As String, ByVal at As Long) As Long
    Dim ends As String
    Dim otherwise As Long
    Dim between As String

    WordAt sql, at, ends
    at = CLng(ends)

    at = NextWordIn(sql, at, STATEMENT_STARTS)
    If at = 0 Then
        EndOfIf = Len(sql) + 1
        Exit Function
    End If
    at = EndOfBranch(sql, at)

    otherwise = NextWordIn(sql, at, "|ELSE|")
    If otherwise > 0 Then
        between = Mid$(sql, at, otherwise - at)
        If Len(Trim$(between)) = 0 Then
            WordAt sql, otherwise, ends
            at = EndOfBranch(sql, CLng(ends))
        End If
    End If
    EndOfIf = at
End Function

' Where one branch of an IF stops, nesting and all.
'
' A BEGIN block runs to its own END however many blocks and CASEs are inside
' it, which is what tells this IF's ELSE from an inner one.
' Where "END <word>" begins, for the word given, or zero.
Private Function WordEndingAt(ByVal sql As String, ByVal at As Long, _
                              ByVal closer As String) As Long
    Dim ends As String
    Dim after As Long

    Do
        at = NextWordIn(sql, at, "|END|")
        If at = 0 Then Exit Function
        WordAt sql, at, ends
        If Len(ends) = 0 Then Exit Function
        after = SkipSpaceAt(sql, CLng(ends))
        If UCase$(WordAt(sql, after, ends)) = closer Then
            WordEndingAt = at
            Exit Function
        End If
        at = after
    Loop
End Function

' Where a block really ends.
'
' END TRY closes only half of a guarded block: the BEGIN CATCH under it is
' part of the same statement, and splitting between them leaves a statement
' opening with the word CATCH that nothing will run.
Private Function EndOfGuardedBlock(ByVal sql As String, ByVal at As Long) As Long
    Dim ends As String
    Dim after As Long

    EndOfGuardedBlock = at
    after = SkipSpaceAt(sql, at)
    If UCase$(WordAt(sql, after, ends)) <> "TRY" Then Exit Function
    at = CLng(ends)
    EndOfGuardedBlock = at

    after = SkipSpaceAt(sql, at)
    If UCase$(WordAt(sql, after, ends)) <> "BEGIN" Then Exit Function
    after = SkipSpaceAt(sql, CLng(ends))
    If UCase$(WordAt(sql, after, ends)) <> "CATCH" Then Exit Function

    at = WordEndingAt(sql, CLng(ends), "CATCH")
    If at = 0 Then
        EndOfGuardedBlock = Len(sql) + 1
        Exit Function
    End If
    WordAt sql, at, ends                            ' END
    after = SkipSpaceAt(sql, CLng(ends))
    WordAt sql, after, ends                         ' CATCH
    If Len(ends) > 0 Then EndOfGuardedBlock = CLng(ends)
End Function

Private Function EndOfBranch(ByVal sql As String, ByVal at As Long) As Long
    Dim word As String
    Dim ends As String
    Dim depth As Long
    Dim cases As Long
    Dim found As Long

    at = SkipSpaceAt(sql, at)
    word = WordAt(sql, at, ends)

    ' A BEGIN that opens a block, which BEGIN TRAN does not: taken for one,
    ' IF @@TRANCOUNT = 0 BEGIN TRAN looked for an END that was never coming
    ' and took the rest of the batch for its branch.
    If OpensABlock(sql, at, word) Then
        Do While at <= Len(sql)
            word = WordAt(sql, at, ends)
            If Len(word) = 0 Then
                at = PastOne(sql, at)
            Else
                Select Case UCase$(word)
                    Case "BEGIN"
                        depth = depth + 1
                    Case "CASE"
                        ' A CASE ends with END too, and that END closes no
                        ' block.
                        cases = cases + 1
                    Case "END"
                        If cases > 0 Then
                            cases = cases - 1
                        Else
                            depth = depth - 1
                            If depth = 0 Then
                                EndOfBranch = EndOfGuardedBlock(sql, CLng(ends))
                                Exit Function
                            End If
                        End If
                End Select
                at = CLng(ends)
            End If
        Loop
        EndOfBranch = Len(sql) + 1
        Exit Function
    End If

    If Len(word) > 0 Then at = CLng(ends)
    found = NextWordIn(sql, at, STATEMENT_STARTS & "ELSE|")
    EndOfBranch = IIf(found = 0, Len(sql) + 1, found)
End Function


' ============================================================================
' Running a batch
' ============================================================================

Private Function AnswerOf(ByVal columns As Collection, _
                          ByVal rows As Collection) As Variant
    Dim out(0 To 3) As Variant

    Set out(0) = columns
    Set out(1) = rows
    out(2) = -1
    out(3) = SELECT_COMMAND
    AnswerOf = out
End Function

' A statement that failed with the batch carrying on: no columns and no
' rows, its number and its words.
Private Function FailedAnswer(ByVal number As Long, _
                              ByVal message As String) As Variant
    Dim out(0 To 4) As Variant

    Set out(0) = Nothing
    Set out(1) = Nothing
    out(2) = ANSWER_FAILED
    out(3) = number
    out(4) = message
    FailedAnswer = out
End Function

' A statement that changed rows rather than answering with any: no columns,
' a count for the client to report as rows affected, and which of the three
' writes it was, which is how the client knows the count is one.
Private Function AffectedAnswer(ByVal affected As Long, _
                                ByVal command As Long) As Variant
    Dim out(0 To 3) As Variant

    Set out(0) = Nothing
    Set out(1) = Nothing
    out(2) = affected
    out(3) = command
    AffectedAnswer = out
End Function

' Every result set the batch produced, one after another. All but the last
' say more is coming, which is how a client knows to keep reading.
Private Function EncodeAnswers(ByVal answers As Collection, _
                               ByVal tdsVersion As Long) As Byte()
    Dim out() As Byte
    Dim used As Long
    Dim index As Long
    Dim answer As Variant
    Dim part() As Byte

    For index = 1 To answers.count
        answer = answers(index)
        If answer(2) = ANSWER_FAILED Then
            ' A statement that failed with the batch still running: its
            ' error, and a DONE saying both that it failed and whether more
            ' is coming. Without the second a client stops at the error and
            ' reads what follows as the reply to whatever it asks next.
            part = TokenError(CLng(answer(3)), CStr(answer(4)), ServerName, _
                              16, tdsVersion)
            AppendBytes out, used, part
            part = TokenDone(DONE_ERROR Or _
                             IIf(index < answers.count, DONE_MORE, DONE_FINAL), _
                             0, tdsVersion)
        ElseIf answer(2) >= 0 Then
            part = TokenDoneCounted( _
                DONE_COUNT Or IIf(index < answers.count, DONE_MORE, _
                                  DONE_FINAL), _
                CLng(answer(2)), tdsVersion, CLng(answer(3)))
        Else
            part = ResultSet(answer(0), answer(1), tdsVersion, _
                             index < answers.count)
        End If
        AppendBytes out, used, part
    Next index
    EncodeAnswers = LeftBytes(out, used)
End Function

' Run one statement of a batch, keeping whatever it produced.
Private Sub RunStatement(ByVal written As String, ByVal answers As Collection)
    Dim head As String
    Dim ends As String
    Dim before As Long

    before = answers.count

    ' What the last statement's reads answered says nothing about this one:
    ' a temporary table it looked at may have been filled since.
    Set mSubqueryAnswers = Nothing
    Set mSubqueryLookups = Nothing
    Set mExistsLookups = Nothing
    Set mLoadedSources = Nothing

    head = UCase$(WordAt(written, SkipSpaceAt(written, 1), ends))

    ' A read written in brackets, as each side of (SELECT ...) UNION
    ' (SELECT ...) is. An open bracket is not a word, so the statement went
    ' the way a SET does: completed, with no rows and nothing said.
    If Len(head) = 0 Then
        If Mid$(written, SkipSpaceAt(written, 1), 1) = "(" Then head = "SELECT"
    End If

    Select Case head
        Case "DECLARE"
            RunDeclare written

        Case "SET"
            RunSet written

        Case "COMMIT"
            RunCommit

        Case "ROLLBACK"
            RunRollback written

        Case "SAVE"
            RunSave written

        Case "IF"
            RunIf written, answers

        Case "WHILE"
            RunWhile written, answers

        Case "BEGIN"
            RunBlock written, answers

        Case "CREATE"
            RunCreate written

        Case "DROP"
            RunDrop written

        Case "INSERT"
            RunInsert written, answers

        Case "SELECT"
            RunRead written, answers

        Case "WITH"
            RunWith written, answers

        Case "EXEC", "EXECUTE"
            RunExec written, answers

        Case "UPDATE"
            RunUpdate written, answers

        Case "DELETE"
            RunDelete written, answers

        Case "MERGE"
            RunMerge written, answers

        Case "TRUNCATE"
            RunTruncate written

        Case "ALTER"
            ' Named rather than ignored: a client told a change was made and
            ' then shown the table as it was has been lied to.
            RaiseQueryError ERR_NUMBER_BASE, _
                head & " is not supported."

        Case "GRANT", "REVOKE", "DENY"
            ' Completed with nothing said, a client was told permissions had
            ' changed that nothing here keeps. Who may connect is Windows's
            ' to say, and whoever it lets in may read, and write unless the
            ' bridge is read-only.
            RaiseQueryError ERR_NUMBER_BASE, head & " is not supported " & _
                "here: there are no permissions to change, and nothing " & _
                "was changed."

        Case Else
            ' USE, PRINT, RETURN, WAITFOR and the SET options a client
            ' opens with. Nothing to do and nothing to say about it.
    End Select

    RememberRowCount answers, before
    ForgetErrorAfter head, written
End Sub

' DECLARE @name type [= value] [, @name type [= value]]
Private Sub RunDeclare(ByVal written As String)
    Dim rest As String
    Dim parts As Collection
    Dim piece As Variant
    Dim name As String
    Dim ends As String
    Dim assigned As Long

    WordAt written, 1, ends
    rest = Mid$(written, CLng(ends))
    Set parts = SplitTopLevel(rest, ",")

    For Each piece In parts
        name = WordAt(Trim$(CStr(piece)), 1, ends)
        If Left$(name, 1) = "@" Then
            assigned = InStr(CStr(piece), "=")
            If assigned > 0 Then
                SetVariable name, EvaluateStandalone(Mid$(CStr(piece), assigned + 1))
            Else
                SetVariable name, Null
            End If
        End If
    Next piece
End Sub

' SET @name = expression, or one of the session options a client opens with.
Private Sub RunSet(ByVal written As String)
    Dim rest As String
    Dim ends As String
    Dim name As String
    Dim assigned As Long

    WordAt written, 1, ends
    rest = TrimWhitespace(Mid$(written, CLng(ends)))
    If Left$(rest, 1) <> "@" Then Exit Sub

    name = WordAt(rest, 1, ends)
    assigned = InStr(rest, "=")
    If assigned = 0 Then Exit Sub
    SetVariable name, EvaluateStandalone(Mid$(rest, assigned + 1))
End Sub

' IF condition statement [ELSE statement]
Private Sub RunIf(ByVal written As String, ByVal answers As Collection)
    Dim ends As String
    Dim at As Long
    Dim conditionText As String
    Dim taken As String
    Dim otherwise As Long
    Dim branchEnd As Long
    Dim outcome As Variant

    WordAt written, 1, ends
    at = NextWordIn(written, CLng(ends), STATEMENT_STARTS)
    If at = 0 Then Exit Sub

    conditionText = Mid$(written, CLng(ends), at - CLng(ends))
    branchEnd = EndOfBranch(written, at)
    taken = Mid$(written, at, branchEnd - at)

    outcome = EvaluateStandalone(conditionText)
    ' Worked out without failing, so the last error is cleared here rather
    ' than when the IF ends: measured, an IF whose branch failed still reads
    ' that failure afterwards, and one that ran nothing reads nought.
    KeepErrorNumber 0
    If Not IsNull(outcome) Then
        If CBool(outcome) Then
            RunInner taken, answers
            Exit Sub
        End If
    End If

    otherwise = NextWordIn(written, branchEnd, "|ELSE|")
    If otherwise = 0 Then Exit Sub
    If Len(TrimWhitespace(Mid$(written, branchEnd, otherwise - branchEnd))) > 0 _
       Then Exit Sub

    WordAt written, otherwise, ends
    RunInner Mid$(written, CLng(ends)), answers
End Sub

' WHILE condition statement
'
' The body runs again for as long as the condition holds. Bounded, because
' this runs on the timer that keeps the workbook answering: a condition that
' never turns false would hold Excel and there would be no way to reach it.
Private Sub RunWhile(ByVal written As String, ByVal answers As Collection)
    Dim ends As String
    Dim at As Long
    Dim conditionText As String
    Dim body As String
    Dim outcome As Variant
    Dim rounds As Long

    WordAt written, 1, ends
    at = NextWordIn(written, CLng(ends), STATEMENT_STARTS)
    If at = 0 Then Exit Sub

    conditionText = Mid$(written, CLng(ends), at - CLng(ends))
    body = Mid$(written, at, EndOfBranch(written, at) - at)

    Do
        outcome = EvaluateStandalone(conditionText)
        KeepErrorNumber 0
        If IsNull(outcome) Then Exit Do
        If Not CBool(outcome) Then Exit Do

        rounds = rounds + 1
        If rounds > WHILE_ROUNDS Then
            RaiseQueryError ERR_NUMBER_BASE, _
                "A WHILE went round " & WHILE_ROUNDS & " times without " & _
                "its condition turning false."
        End If
        RunInner body, answers
    Loop
End Sub

' BEGIN ... END, which is a statement holding statements.
Private Sub RunBlock(ByVal written As String, ByVal answers As Collection)
    Dim inner As String
    Dim ends As String
    Dim closing As Long
    Dim after As Long
    Dim opening As String

    WordAt written, 1, ends
    after = SkipSpaceAt(written, CLng(ends))
    opening = UCase$(WordAt(written, after, ends))
    If opening = "TRY" Then
        RunGuarded written, answers
        Exit Sub
    End If
    ' BEGIN DISTRIBUTED TRANSACTION too, which a real server hands to a
    ' coordinator this has not got, and which begins one here like any other.
    If opening = "TRANSACTION" Or opening = "TRAN" Or _
       opening = "DISTRIBUTED" Then
        RunBeginTransaction written
        Exit Sub
    End If

    WordAt written, 1, ends
    inner = Mid$(written, CLng(ends))

    closing = InStrRev(UCase$(inner), "END")
    If closing > 0 Then inner = Left$(inner, closing - 1)
    RunInner inner, answers
End Sub

' BEGIN TRANSACTION, COMMIT and ROLLBACK.
'
' A workbook has no log to undo from, so what stands in for one is a copy:
' the first write to a table inside a transaction keeps what that table
' looked like, and a rollback puts it back. Values, not formatting and not
' formulas, which is the part of a real rollback this cannot do.
Private Sub RunBeginTransaction(ByVal written As String)
    BeginNamed TransactionNameIn(written)
End Sub

' One level more, and at the first level a transaction of that name.
Private Sub BeginNamed(ByVal name As String)
    Dim count As Long

    count = TranCount() + 1
    KeepTranValue TRAN_COUNT, count

    ' Only the outermost transaction's name is kept, which is all a real
    ' server keeps: rolling back to an inner one's is refused there.
    If count = 1 Then
        KeepTranValue TRAN_NAME, name
        ForgetTranList TRAN_SNAPSHOTS
        ForgetTranList TRAN_SAVEPOINTS
    End If
End Sub

' COMMIT [TRAN[SACTION] [name] | WORK]: one level fewer, and at the last one
' the transaction is over. A name is read and changes nothing, which,
' measured, is what a real server does with it.
Private Sub RunCommit()
    Dim count As Long

    count = TranCount()
    If count = 0 Then
        RaiseQueryError 3902, "The COMMIT TRANSACTION request has no " & _
                              "corresponding BEGIN TRANSACTION."
    End If
    If count = 1 Then
        EndTransaction
    Else
        KeepTranValue TRAN_COUNT, count - 1
    End If
End Sub

' The transaction-manager requests a linked server and a client's API send,
' run as the statements are: a transaction begun through the API is the one
' BEGIN TRAN counts, and a savepoint marked through it is found by the same
' rules. Ending one that is not open is not an error from there: a linked
' server rolls back after every read whether it wrote or not. The count
' before and after is what says whether a transaction began or ended.
Friend Sub TransactionRequested(ByVal request As Long, ByVal name As String, _
                                ByVal variables As Collection, _
                                ByVal tempTables As Collection, _
                                ByRef before As Long, ByRef after As Long)
    Set mVariables = variables
    Set mTempTables = tempTables
    mQueryErrorNumber = 0
    before = TranCount()
    Select Case request
        Case TM_BEGIN_XACT
            BeginNamed name
        Case TM_COMMIT_XACT
            If before > 0 Then RunCommit
        Case TM_ROLLBACK_XACT
            If before > 0 Then RollBackNamed name
        Case TM_SAVE_XACT
            SaveNamed name
    End Select
    after = TranCount()
End Sub

' ROLLBACK [TRAN[SACTION] [name] | WORK].
'
' With no name, or the outermost transaction's, every table written since
' the outermost BEGIN is put back and the transaction ends: a real server
' unwinds nested transactions to the outermost one, and so does this. With
' a savepoint's name it goes back to the newest savepoint of that name and
' the transaction stays open. Measured on a real server, a savepoint is
' found before a transaction of the same name, a savepoint's name matches
' in any case and a transaction's only as it was written, and anything
' else is refused.
Private Sub RunRollback(ByVal written As String)
    If TranCount() = 0 Then
        RaiseQueryError 3903, _
            "The ROLLBACK TRANSACTION request has no corresponding BEGIN " & _
            "TRANSACTION."
    End If
    RollBackNamed TransactionNameIn(written)
End Sub

' Back to the newest savepoint of a name, or with no name or the outermost
' transaction's, to before the transaction began.
Private Sub RollBackNamed(ByVal name As String)
    Dim savepoints As Collection
    Dim mark As Variant
    Dim at As Long
    Dim index As Long

    If Len(name) > 0 Then
        Set savepoints = TranList(TRAN_SAVEPOINTS)
        For index = savepoints.count To 1 Step -1
            mark = savepoints(index)
            If StrComp(CStr(mark(0)), name, vbTextCompare) = 0 Then
                at = index
                Exit For
            End If
        Next index
        If at > 0 Then
            RollBackTo at
            Exit Sub
        End If
        If StrComp(name, TranName(), vbBinaryCompare) <> 0 Then
            RaiseQueryError 6401, "Cannot roll back " & name & ". No " & _
                "transaction or savepoint of that name was found."
        End If
    End If
    RollBackEverything
End Sub

' SAVE TRAN[SACTION] name: a point the transaction can be rolled back to.
' The tables it has written so far are copied as they are now, because
' rolling back to here puts them back to this rather than to before the
' transaction began.
Private Sub RunSave(ByVal written As String)
    SaveNamed TransactionNameIn(written)
End Sub

Private Sub SaveNamed(ByVal name As String)
    Dim mark(0 To 2) As Variant
    Dim copies As Collection
    Dim kept As Variant

    If TranCount() = 0 Then
        RaiseQueryError 628, "Cannot issue SAVE TRANSACTION when there is " & _
                             "no active transaction."
    End If

    Set copies = New Collection
    For Each kept In TranList(TRAN_SNAPSHOTS)
        copies.Add TableCopy(CStr(kept(0)))
    Next kept
    mark(0) = name
    mark(1) = TranList(TRAN_SNAPSHOTS).count
    Set mark(2) = copies
    TranList(TRAN_SAVEPOINTS).Add mark
End Sub

' Every table written since the outermost BEGIN, put back as it was, and
' the transaction over.
Private Sub RollBackEverything()
    Dim snapshots As Collection
    Dim index As Long

    Set snapshots = TranList(TRAN_SNAPSHOTS)
    For index = snapshots.count To 1 Step -1
        RestoreSnapshot snapshots(index)
    Next index
    EndTransaction
End Sub

' Back to a savepoint: a table first written after it goes back to how it
' was before, one written before it to how it was when the savepoint was
' marked, and the savepoint and every one marked after it are let go, the
' way a real server lets them go. The transaction stays open, and so do
' the tables it holds.
Private Sub RollBackTo(ByVal at As Long)
    Dim savepoints As Collection
    Dim snapshots As Collection
    Dim mark As Variant
    Dim kept As Variant
    Dim index As Long

    Set savepoints = TranList(TRAN_SAVEPOINTS)
    Set snapshots = TranList(TRAN_SNAPSHOTS)
    mark = savepoints(at)
    For index = snapshots.count To CLng(mark(1)) + 1 Step -1
        RestoreSnapshot snapshots(index)
        snapshots.Remove index
    Next index
    For Each kept In mark(2)
        RestoreSnapshot kept
    Next kept
    For index = savepoints.count To at Step -1
        savepoints.Remove index
    Next index
End Sub

' The transaction over, however it ended: the count back to nought, its
' name, savepoints and copies gone, and the tables it held let go.
Private Sub EndTransaction()
    Dim key As Variant

    If Not mHolds Is Nothing Then
        For Each key In TranList(TRAN_HOLDS)
            On Error Resume Next
            mHolds.Remove CStr(key)
            On Error GoTo 0
        Next key
    End If
    KeepTranValue TRAN_COUNT, 0
    KeepTranValue TRAN_NAME, vbNullString
    ForgetTranList TRAN_SNAPSHOTS
    ForgetTranList TRAN_SAVEPOINTS
    ForgetTranList TRAN_HOLDS
End Sub

' Before a write: refuse it if another connection's open transaction holds
' the table, and inside a transaction of this connection's own, hold the
' table and keep what it looks like before the first write to it. Nothing
' to keep twice: the copy wanted is the one from before any of the writes.
'
' A real server would make the write wait until the other transaction
' ended. This one cannot wait, so it refuses the way a real server does
' when a lock is not granted in time. Writing over the table instead would
' have the other connection's rollback put back a table this write had
' changed, and the write would be lost without a word.
Private Sub RememberBeforeWriting(ByVal name As String)
    Dim entry As Variant
    Dim key As String
    Dim owner As String
    Dim holdable As Boolean

    entry = LookUpTable(name)
    key = LCase$(CStr(entry(0)))

    ' A temporary table is its connection's alone: nobody else can hold it,
    ' and two connections' temporary tables of one name are two tables.
    holdable = (Left$(key, 1) <> "#")
    If holdable Then
        owner = HolderOf(key)
        If Len(owner) > 0 And owner <> Holder() Then
            RaiseQueryError 1222, "Lock request time out period exceeded."
        End If
    End If

    If TranCount() = 0 Then Exit Sub
    If holdable And Len(owner) = 0 Then
        If mHolds Is Nothing Then Set mHolds = New Collection
        mHolds.Add Holder(), key
        TranList(TRAN_HOLDS).Add key
    End If
    If HasKey(TranList(TRAN_SNAPSHOTS), key) Then Exit Sub
    TranList(TRAN_SNAPSHOTS).Add TableCopy(name), key
End Sub

' What a table holds now, in the form RestoreSnapshot puts back.
Private Function TableCopy(ByVal name As String) As Variant
    Dim entry As Variant
    Dim kept(0 To 4) As Variant
    Dim target As Object

    entry = LookUpTable(name)
    kept(0) = LCase$(CStr(entry(0)))
    Set target = SourceRange(entry)
    If target Is Nothing Then
        ' Its rows are all there is of it, so a copy of them is the copy.
        kept(1) = "rows"
        kept(2) = entry(1)
        kept(3) = CStr(entry(0))
    Else
        kept(1) = SourceKind(entry)
        Set kept(2) = target.Worksheet
        kept(3) = target.Address
        kept(4) = OneCellSafe(target)
    End If
    TableCopy = kept
End Function

' Which connection holds a table, or nothing.
Private Function HolderOf(ByVal key As String) As String
    If mHolds Is Nothing Then Exit Function
    On Error Resume Next
    HolderOf = CStr(mHolds(key))
    On Error GoTo 0
End Function

' The connection a call is for, as the server tells them apart: by the
' variables it keeps, which are its own for as long as it is connected.
Private Function Holder() As String
    Holder = "c" & CStr(ObjPtr(mVariables))
End Function

' ----------------------------------------------------------------------------
' A connection's transaction, kept with its variables
'
' A connection's variables are what the server is handed for every call it
' makes, so that is where its transaction is kept, under names no statement
' can declare. Kept once for the whole server, a transaction was everybody's:
' one client's ROLLBACK put back another's writes, and a linked server, which
' wraps every read in a transaction and rolls it back, could undo a write that
' landed in between.
' ----------------------------------------------------------------------------

Private Function TranCount() As Long
    If mVariables Is Nothing Then Exit Function
    On Error Resume Next
    TranCount = CLng(mVariables(TRAN_COUNT))
    On Error GoTo 0
End Function

Private Function TranName() As String
    If mVariables Is Nothing Then Exit Function
    On Error Resume Next
    TranName = CStr(mVariables(TRAN_NAME))
    On Error GoTo 0
End Function

Private Sub KeepTranValue(ByVal key As String, ByVal value As Variant)
    On Error Resume Next
    mVariables.Remove key
    On Error GoTo 0
    mVariables.Add value, key
End Sub

' One of the lists a transaction keeps, made the first time it is asked for.
Private Function TranList(ByVal key As String) As Collection
    On Error Resume Next
    Set TranList = mVariables(key)
    On Error GoTo 0
    If TranList Is Nothing Then
        Set TranList = New Collection
        mVariables.Add TranList, key
    End If
End Function

Private Sub ForgetTranList(ByVal key As String)
    On Error Resume Next
    mVariables.Remove key
    On Error GoTo 0
End Sub

' The name a transaction statement carries, if it carries one: the word
' after TRAN or TRANSACTION, a name in brackets, or a variable holding one.
' WORK carries none, and WITH MARK after a name says how a real server
' should log it, which there is no log here to say to.
Private Function TransactionNameIn(ByVal written As String) As String
    Dim at As Long
    Dim ends As String
    Dim word As String

    at = SkipSpaceAt(written, 1)
    Do While at <= Len(written)
        word = UCase$(WordAt(written, at, ends))
        If Len(word) = 0 Or word = "WORK" Then Exit Function
        at = SkipSpaceAt(written, CLng(ends))
        If word = "TRAN" Or word = "TRANSACTION" Then Exit Do
    Loop
    If at > Len(written) Then Exit Function

    word = WordAt(written, at, ends)
    If Left$(word, 1) = "[" And Right$(word, 1) = "]" Then
        word = Mid$(word, 2, Len(word) - 2)
    ElseIf Left$(word, 1) = "@" Then
        word = CStr(NullToEmpty(NamedValue(word)))
    End If
    TransactionNameIn = word
End Function

' A call with no connection behind it is a session of its own, and a
' transaction it leaves open ends with it, the way one ends when a client
' disconnects. Left open, the tables it held stayed held by nobody.
Private Sub EndTransientTransaction()
    If TranCount() = 0 Then Exit Sub
    On Error Resume Next
    RollBackEverything
    If Err.Number <> 0 Then
        Err.Clear
        EndTransaction
    End If
    On Error GoTo 0
End Sub

' A connection that has gone with a transaction still open: what it changed
' is put back and what it held let go, which is what a real server does when
' a connection ends. Nothing raises from here, since the connection is
' closing either way.
Friend Sub EndAbandoned(ByVal variables As Collection, _
                        ByVal tempTables As Collection, _
                        ByVal session As Collection)
    Set mVariables = variables
    Set mTempTables = tempTables
    Set mSession = session
    If TranCount() = 0 Then Exit Sub

    On Error Resume Next
    RollBackEverything
    If Err.Number <> 0 Then
        Note "could not put back what a closed connection changed: " & _
             Err.Description
        ' Let go of what it held regardless, or nobody could write to it.
        Err.Clear
        EndTransaction
    Else
        Note "rolled back the transaction a closed connection left open"
    End If
    On Error GoTo 0
End Sub

Private Sub RestoreSnapshot(ByVal kept As Variant)
    Dim sheet As Object
    Dim target As Object
    Dim current As Object
    Dim entry As Variant
    Dim listed As Object
    Dim wide As Long
    Dim high As Long
    Dim tall As Long

    If kept(1) = "rows" Then
        entry = LookUpTable(CStr(kept(3)))
        entry(1) = kept(2)
        StoreTable CStr(entry(0)), entry
        Exit Sub
    End If

    Set sheet = kept(2)
    Set target = sheet.Range(CStr(kept(3)))
    high = UBound(kept(4), 1)
    wide = UBound(kept(4), 2)

    ' Whatever is there now goes first: a delete moved rows up and an insert
    ' moved them down, so writing the old block over the top would leave
    ' whatever the change left behind underneath it. As far down as the
    ' table reaches now or reached then, and no further. This offset the
    ' whole block by the height of the sheet, which for any table of more
    ' than one row is past the last row, so no rollback of a sheet ever got
    ' this far; and reaching the sheet's last row would have cleared
    ' whatever sat under the table as well.
    tall = high
    entry = LookUpTable(CStr(kept(0)))
    Set current = SourceRange(entry)
    If Not current Is Nothing Then
        If current.row + current.Rows.count - target.row > tall Then
            tall = current.row + current.Rows.count - target.row
        End If
    End If
    target.Resize(tall, wide).ClearContents
    target.Resize(high, wide).Value = kept(4)

    If SourceKindOfSheet(CStr(kept(0)), listed) Then
        listed.Resize target.Resize(high, wide)
    End If
    RefreshTable CStr(kept(0))
End Sub

' The Excel table a served name stands for, when it stands for one.
Private Function SourceKindOfSheet(ByVal name As String, _
                                   ByRef listed As Object) As Boolean
    Dim entry As Variant

    On Error GoTo Missing
    entry = LookUpTable(name)
    On Error GoTo 0
    If SourceKind(entry) <> "ListObject" Then Exit Function
    Set listed = entry(3)
    SourceKindOfSheet = True
    Exit Function

Missing:
End Function

' BEGIN TRY ... END TRY BEGIN CATCH ... END CATCH
'
' A client writes this around a question it expects some servers not to be
' able to answer: SSMS asks for sys.json_indexes and sys.vector_indexes this
' way while opening a table's Indexes node. Swallowing what the try part
' raises is the whole point of it, so a statement that fails in here is not
' the connection failing.
Private Sub RunGuarded(ByVal written As String, ByVal answers As Collection)
    Dim ends As String
    Dim at As Long
    Dim after As Long
    Dim tryEnd As Long
    Dim catchAt As Long
    Dim catchEnd As Long
    Dim body As String
    Dim handler As String
    Dim failed As Boolean
    Dim kept As Boolean
    Dim caughtNumber As Long
    Dim caughtMessage As String
    Dim outerNumber As Long
    Dim outerMessage As String

    WordAt written, 1, ends                         ' BEGIN
    after = SkipSpaceAt(written, CLng(ends))
    WordAt written, after, ends                     ' TRY
    at = CLng(ends)

    tryEnd = WordEndingAt(written, at, "TRY")
    If tryEnd = 0 Then
        RunInner Mid$(written, at), answers
        Exit Sub
    End If
    body = Mid$(written, at, tryEnd - at)

    catchAt = NextWordIn(written, tryEnd, "|CATCH|")
    If catchAt > 0 Then
        WordAt written, catchAt, ends
        catchEnd = WordEndingAt(written, CLng(ends), "CATCH")
        If catchEnd = 0 Then catchEnd = Len(written) + 1
        handler = Mid$(written, CLng(ends), catchEnd - CLng(ends))
    End If

    ' Nothing inside a TRY is carried past, however deep: the point of
    ' writing one is that its CATCH gets the error. What the TRY answered
    ' before it failed stays, because a real server has sent it by then.
    kept = mCatching
    mCatching = False
    On Error Resume Next
    RunInner body, answers
    failed = (Err.Number <> 0)
    If failed Then
        If Err.Number = ERR_QUERY Then
            caughtNumber = mQueryErrorNumber
        Else
            caughtNumber = ERR_NUMBER_BASE
        End If
        caughtMessage = Err.Description
    End If
    Err.Clear
    On Error GoTo 0
    mCatching = kept

    If failed Then
        ' What the CATCH reads: @@ERROR and ERROR_NUMBER() are the error
        ' that sent it there, and a statement that failed leaves no count.
        KeepErrorNumber caughtNumber
        mRowCount = 0
        outerNumber = mCaughtNumber
        outerMessage = mCaughtMessage
        mCaughtNumber = caughtNumber
        mCaughtMessage = caughtMessage
        If Len(TrimWhitespace(handler)) > 0 Then RunInner handler, answers
        mCaughtNumber = outerNumber
        mCaughtMessage = outerMessage
    End If
End Sub

' Whatever this branch holds, which may itself be several statements.
Private Sub RunInner(ByVal written As String, ByVal answers As Collection)
    Dim pieces As Collection
    Dim piece As Variant
    Dim trimmed As String
    Dim ends As String

    trimmed = TrimWhitespace(written)
    If Len(trimmed) = 0 Then Exit Sub

    If OpensABlock(trimmed, 1, WordAt(trimmed, 1, ends)) Then
        RunBlock trimmed, answers
        Exit Sub
    End If

    Set pieces = SplitStatements(trimmed)
    For Each piece In pieces
        RunOne CStr(piece), answers
    Next piece
End Sub

' CREATE TABLE #name (column type [, ...])
Private Sub RunCreate(ByVal written As String)
    Dim ends As String
    Dim at As Long
    Dim word As String
    Dim name As String
    Dim opening As Long
    Dim closing As Long
    Dim columns As Collection
    Dim data As Variant
    Dim index As Long
    Dim entry(0 To 2) As Variant

    at = 1
    word = WordAt(written, at, ends)                ' CREATE
    at = SkipSpaceAt(written, CLng(ends))
    word = WordAt(written, at, ends)                ' TABLE
    If UCase$(word) <> "TABLE" Then Exit Sub

    at = SkipSpaceAt(written, CLng(ends))
    name = WordAt(written, at, ends)

    opening = InStr(CLng(ends), written, "(")
    closing = InStrRev(written, ")")
    If opening = 0 Or closing <= opening Then
        RaiseQueryError ERR_NUMBER_BASE, _
            "Incorrect syntax near '" & name & "'; a table needs columns."
    End If

    Set columns = DeclaredColumns(Mid$(written, opening + 1, closing - opening - 1))
    ReDim data(1 To 1, 1 To columns.count)
    For index = 1 To columns.count
        data(1, index) = columns(index)(COLUMN_NAME)
    Next index

    entry(0) = name
    entry(1) = data
    Set entry(2) = columns
    StoreTable name, entry
End Sub

' The columns a CREATE TABLE declares.
Private Function DeclaredColumns(ByVal body As String) As Collection
    Dim columns As Collection
    Dim parts As Collection
    Dim piece As Variant
    Dim name As String
    Dim ends As String
    Dim typeName As String
    Dim declared As Variant

    Set columns = New Collection
    Set parts = SplitTopLevel(body, ",")

    For Each piece In parts
        name = WordAt(TrimWhitespace(CStr(piece)), 1, ends)
        If Len(name) > 0 Then
            typeName = TrimWhitespace( _
                Mid$(TrimWhitespace(CStr(piece)), CLng(ends)))
            declared = TypeFromName(typeName)
            columns.Add ColumnOf(Trim$(Replace(Replace(name, "[", ""), "]", "")), _
                                 CLng(declared(0)), CLng(declared(1)))
        End If
    Next piece
    Set DeclaredColumns = columns
End Function

' A written type name as the type code and size this encodes it with.
Private Function TypeFromName(ByVal written As String) As Variant
    Dim out(0 To 2) As Variant
    Dim bare As String
    Dim size As Long
    Dim places As Long
    Dim comma As Long
    Dim inside As String
    Dim opening As Long

    bare = LCase$(Trim$(written))
    opening = InStr(bare, "(")
    If opening > 0 Then
        inside = Mid$(bare, opening + 1)
        size = CLng(Val(inside))
        ' decimal(18,2): the second number is how many of the digits sit
        ' after the point.
        comma = InStr(inside, ",")
        If comma > 0 Then places = CLng(Val(Mid$(inside, comma + 1)))
        bare = Trim$(Left$(bare, opening - 1))
    End If
    bare = Trim$(Split(bare & " ", " ")(0))

    Select Case bare
        Case "int", "integer"
            out(0) = TYPE_INTN: out(1) = 4
        Case "bigint"
            out(0) = TYPE_INTN: out(1) = 8
        Case "smallint"
            out(0) = TYPE_INTN: out(1) = 2
        Case "tinyint"
            out(0) = TYPE_INTN: out(1) = 1
        Case "bit"
            out(0) = TYPE_BITN: out(1) = 1
        Case "decimal", "numeric", "money", "smallmoney"
            out(0) = TYPE_DECIMALN
            If size <= 0 Or size > 38 Then
                out(1) = 18
            Else
                out(1) = size
            End If
            out(2) = places
        Case "float", "real"
            out(0) = TYPE_FLTN: out(1) = 8
        Case "datetime", "datetime2", "smalldatetime", "date", "time"
            out(0) = TYPE_DATETIMN: out(1) = 8
        Case "uniqueidentifier"
            out(0) = TYPE_GUIDN: out(1) = 16
        Case "sysname"
            out(0) = TYPE_NVARCHAR: out(1) = 128
        Case Else
            ' Everything else is text, which is what a client can always read.
            out(0) = TYPE_NVARCHAR
            If size <= 0 Or size > NVARCHAR_SIZED_LIMIT Then
                out(1) = NVARCHAR_SIZED_LIMIT
            Else
                out(1) = size
            End If
    End Select
    TypeFromName = out
End Function

' DROP TABLE #name
Private Sub RunDrop(ByVal written As String)
    Dim ends As String
    Dim at As Long
    Dim word As String
    Dim name As String

    at = 1
    WordAt written, at, ends                        ' DROP
    at = SkipSpaceAt(written, CLng(ends))
    word = WordAt(written, at, ends)
    If UCase$(word) <> "TABLE" Then Exit Sub

    at = SkipSpaceAt(written, CLng(ends))
    name = WordAt(written, at, ends)

    On Error Resume Next
    If Left$(name, 1) = "#" Then
        mTempTables.Remove LCase$(name)
    Else
        mTables.Remove LCase$(name)
    End If
    On Error GoTo 0
End Sub

' INSERT [INTO] #name [(columns)] SELECT ... | VALUES ... | EXEC ...
Private Sub RunInsert(ByVal written As String, ByVal answers As Collection)
    Dim at As Long
    Dim ends As String
    Dim word As String
    Dim name As String
    Dim rest As String
    Dim head As String
    Dim source As Collection
    Dim answer As Variant

    at = 1
    WordAt written, at, ends                        ' INSERT
    at = SkipSpaceAt(written, CLng(ends))
    word = WordAt(written, at, ends)
    If UCase$(word) = "INTO" Then
        at = SkipSpaceAt(written, CLng(ends))
        word = WordAt(written, at, ends)
    End If
    name = word
    at = CLng(ends)

    ' A column list after the name is read and ignored: the values go in by
    ' position, which is what every batch here does.
    at = SkipSpaceAt(written, at)
    If Mid$(written, at, 1) = "(" Then at = MatchingBracket(written, at) + 1

    ' TrimWhitespace, not Trim: an INSERT whose column list ends the line
    ' leaves the SELECT under it starting with a newline, and a statement
    ' whose first character is a newline has no first word.
    rest = TrimWhitespace(Mid$(written, at))
    head = UCase$(WordAt(rest, 1, ends))

    ' Where the rows come from, which is not what the client is told: an
    ' INSERT reports how many rows went in, never the rows themselves.
    Set source = New Collection
    Select Case head
        Case "SELECT", "WITH"
            RunRead rest, source
        Case "EXEC", "EXECUTE"
            RunExec rest, source
        Case "VALUES"
            AppendValuesTo name, rest, answers
            Exit Sub
        Case Else
            RaiseQueryError ERR_NUMBER_BASE, _
                "Incorrect syntax near '" & head & "' in an INSERT."
    End Select

    If source.count = 0 Then Exit Sub
    answer = source(1)
    ' The EXEC it read from failed and the batch carried on past it: the
    ' INSERT fails with that error rather than inserting nothing and
    ' reading the count of rows that never came.
    If answer(2) = ANSWER_FAILED Then
        RaiseQueryError CLng(answer(3)), CStr(answer(4))
    End If
    AppendRowsTo name, answer(1)
    answers.Add AffectedAnswer(answer(1).count, INSERT_COMMAND)
End Sub

' Where the bracket opening here closes.
Private Function MatchingBracket(ByVal text As String, ByVal at As Long) As Long
    Dim depth As Long
    Dim ch As String

    Do While at <= Len(text)
        ch = Mid$(text, at, 1)
        If ch = "'" Or ch = """" Or ch = "[" Then
            at = PastOne(text, at)
        Else
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    MatchingBracket = at
                    Exit Function
                End If
            End If
            at = at + 1
        End If
    Loop
    MatchingBracket = Len(text)
End Function

Private Sub AppendValuesTo(ByVal name As String, ByVal written As String, _
                           ByVal answers As Collection)
    Dim rows As Collection
    Dim ends As String

    WordAt written, 1, ends
    Set rows = ReadValueRows(Mid$(written, CLng(ends)))
    AppendRowsTo name, rows
    answers.Add AffectedAnswer(rows.count, INSERT_COMMAND)
End Sub

' The rows of a VALUES list, each evaluated as an expression.
Private Function ReadValueRows(ByVal body As String) As Collection
    Dim rows As Collection
    Dim groups As Collection
    Dim group As Variant
    Dim parts As Collection
    Dim piece As Variant
    Dim values As Variant
    Dim index As Long
    Dim inner As String

    Set rows = New Collection
    Set groups = SplitTopLevel(body, ",")

    ' Each group is one bracketed row. SplitTopLevel keeps a bracketed run
    ' whole, so the commas inside a row are not split on.
    For Each group In groups
        inner = Trim$(CStr(group))
        If Left$(inner, 1) = "(" Then inner = Mid$(inner, 2, Len(inner) - 2)
        Set parts = SplitTopLevel(inner, ",")
        ReDim values(0 To parts.count - 1)
        For index = 1 To parts.count
            values(index - 1) = EvaluateStandalone(CStr(parts(index)))
        Next index
        rows.Add values
    Next group
    Set ReadValueRows = rows
End Function

' Put rows into a table that already exists, growing its array.
'
' The rows live in the same two-dimensional array a worksheet hands over, so
' adding to one means building it again: VBA can only preserve the last
' dimension, and the rows are the first.
Private Sub AppendRowsTo(ByVal name As String, ByVal rows As Collection)
    Dim entry As Variant
    Dim data As Variant
    Dim grown As Variant
    Dim width As Long
    Dim height As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim values As Variant
    Dim value As Variant
    Dim declared As Collection

    If rows Is Nothing Then Exit Sub
    If rows.count = 0 Then Exit Sub

    RememberBeforeWriting name
    entry = LookUpTable(name)

    ' A table with a sheet behind it is written to the sheet. Growing the
    ' copy instead would answer the next question with rows the workbook
    ' does not have.
    If Len(SourceKind(entry)) > 0 Then
        AppendRowsToSheet entry, name, rows
        Exit Sub
    End If

    data = entry(1)
    width = UBound(data, 2) - LBound(data, 2) + 1
    height = UBound(data, 1) - LBound(data, 1) + 1

    ReDim grown(1 To height + rows.count, 1 To width)
    For rowIndex = 1 To height
        For columnIndex = 1 To width
            grown(rowIndex, columnIndex) = _
                data(LBound(data, 1) + rowIndex - 1, LBound(data, 2) + columnIndex - 1)
        Next columnIndex
    Next rowIndex

    If IsObject(entry(2)) Then Set declared = entry(2)
    For rowIndex = 1 To rows.count
        values = rows(rowIndex)
        For columnIndex = 1 To width
            If columnIndex - 1 <= UBound(values) Then
                value = values(LBound(values) + columnIndex - 1)
                If Not declared Is Nothing Then
                    If columnIndex <= declared.count Then
                        value = AsDeclared(value, declared(columnIndex))
                    End If
                End If
                grown(height + rowIndex, columnIndex) = value
            End If
        Next columnIndex
    Next rowIndex

    entry(1) = grown
    StoreTable CStr(entry(0)), entry
End Sub

' A value as the column it goes into declares it. Text or a number going
' into a datetime column is read as a date there and then, the way a real
' server converts it on the way in, so it compares as one afterwards: kept
' as text, a datetime column of ISO dates compared by spelling, and
' '2024-10-01' came before '2024-9-1'.
Private Function AsDeclared(ByVal value As Variant, _
                            ByVal column As Variant) As Variant
    AsDeclared = value
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If VarType(value) = vbDate Then Exit Function
    If CLng(column(COLUMN_TYPE)) = TYPE_DATETIMN Then
        AsDeclared = CDate(MomentOf(value))
    End If
End Function

' A read, which may be a read into variables rather than into rows.
Private Sub RunRead(ByVal written As String, ByVal answers As Collection)
    Dim outColumns As Collection
    Dim outRows As Collection

    Set mTokens = Tokenize(written)
    mCursor = 1
    ParseSelect

    If AssignsToVariables() Then
        RefuseTheRestOfARead
        AssignFromRead
        Exit Sub
    End If

    RefuseTheRestOfARead
    ExecuteWholeSelect outColumns, outRows
    RefuseTheRestOfARead
    answers.Add AnswerOf(outColumns, outRows)
End Sub

' Whatever a read says past the parts read here. OPTION's hints change how a
' real server gets to its answer and never what the answer is, so they are
' passed over. Anything else changes the answer, and a read that stopped
' short of it answered as though it were not there: PIVOT gave the rows
' back unpivoted.
Private Sub RefuseTheRestOfARead()
    If Accept("OPTION") Then
        If Accept("(") Then BracketedTokens
    End If
    If AtEnd() Then Exit Sub
    If FollowsSetOperator() Then Exit Sub

    Select Case UCase$(PeekText())
        Case "PIVOT", "UNPIVOT"
            RaiseQueryError ERR_NUMBER_BASE, _
                UCase$(PeekText()) & " is not supported here."
        Case "FOR"
            mCursor = mCursor + 1
            RaiseQueryError ERR_NUMBER_BASE, _
                "FOR " & UCase$(PeekText()) & " is not supported here."
        Case Else
            RaiseQueryError ERR_NUMBER_BASE, _
                "Incorrect syntax near '" & PeekText() & "'."
    End Select
End Sub

' WITH name [(columns)] AS (read) [, ...] and the statement the names are
' for.
'
' Each read is run once, in order, and its rows kept under its name for the
' length of the statement, so a later one can read an earlier one and the
' statement reads them the way it reads a table. A real server works a
' common table expression out again wherever it is named; a read gives the
' same rows either way.
Private Sub RunWith(ByVal written As String, ByVal answers As Collection)
    Dim at As Long
    Dim ends As String
    Dim closing As Long
    Dim names As Collection
    Dim lists As Collection
    Dim bodies As Collection
    Dim columnNames As Collection
    Dim piece As Variant
    Dim rest As String
    Dim index As Long

    Set names = New Collection
    Set lists = New Collection
    Set bodies = New Collection

    at = SkipSpaceAt(written, 1)
    WordAt written, at, ends                        ' WITH
    Do
        at = SkipSpaceAt(written, CLng(ends))
        names.Add Replace(Replace(WordAt(written, at, ends), "[", ""), "]", "")
        If Len(names(names.count)) = 0 Then
            RaiseQueryError ERR_NUMBER_BASE, "Incorrect syntax near 'WITH'."
        End If
        at = SkipSpaceAt(written, CLng(ends))

        Set columnNames = Nothing
        If Mid$(written, at, 1) = "(" Then
            closing = MatchingBracket(written, at)
            Set columnNames = New Collection
            For Each piece In SplitTopLevel(Mid$(written, at + 1, _
                                                closing - at - 1), ",")
                columnNames.Add Trim$(CStr(piece))
            Next piece
            at = SkipSpaceAt(written, closing + 1)
        End If
        lists.Add columnNames

        If UCase$(WordAt(written, at, ends)) <> "AS" Then
            RaiseQueryError ERR_NUMBER_BASE, "Incorrect syntax near '" & _
                names(names.count) & "'; expected AS."
        End If
        at = SkipSpaceAt(written, CLng(ends))
        If Mid$(written, at, 1) <> "(" Then
            RaiseQueryError ERR_NUMBER_BASE, _
                "Incorrect syntax near 'AS'; expected '('."
        End If
        closing = MatchingBracket(written, at)
        bodies.Add Mid$(written, at + 1, closing - at - 1)

        at = SkipSpaceAt(written, closing + 1)
        If Mid$(written, at, 1) <> "," Then Exit Do
        ends = CStr(at + 1)
    Loop
    rest = Mid$(written, at)

    Set mCommonTables = New Collection
    On Error GoTo Failed
    For index = 1 To names.count
        KeepCommonTable CStr(names(index)), lists(index), _
                        CStr(bodies(index)), RecursionLimit(rest)
    Next index
    RunStatement rest, answers
    Set mCommonTables = Nothing
    Exit Sub

Failed:
    Set mCommonTables = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' One common table expression, run and kept under its name.
'
' One that reads itself is recursive, and is run a level at a time: the
' branch before its first UNION ALL once, then the rest over the rows the
' level before produced, until a level produces none. A level past the
' limit is refused the way a real server refuses it, rather than left to
' go round for ever.
Private Sub KeepCommonTable(ByVal name As String, _
                            ByVal columnNames As Collection, _
                            ByVal body As String, ByVal limit As Long)
    Dim columns As Collection
    Dim rows As Collection
    Dim level As Collection
    Dim levelColumns As Collection
    Dim joinAt As Long
    Dim ends As String
    Dim recursive As String
    Dim depth As Long
    Dim row As Variant

    joinAt = TopLevelUnionAll(body)
    If joinAt = 0 Or Not NamesItself(body, name) Then
        ReadInto body, columns, rows
        StoreCommonTable name, QualifyColumns(columns, "", columnNames), rows
        Exit Sub
    End If

    WordAt body, SkipSpaceAt(body, joinAt + 5), ends       ' ALL
    recursive = Mid$(body, CLng(ends))
    ReadInto Left$(body, joinAt - 1), columns, level
    Set columns = QualifyColumns(columns, "", columnNames)

    Set rows = New Collection
    Do
        For Each row In level
            rows.Add row
        Next row
        StoreCommonTable name, columns, level
        ReadInto recursive, levelColumns, level
        If level.count = 0 Then Exit Do
        depth = depth + 1
        If depth > limit Then
            RaiseQueryError 530, "The statement terminated. The maximum " & _
                "recursion " & limit & " has been exhausted before " & _
                "statement completion."
        End If
    Loop
    StoreCommonTable name, columns, rows
End Sub

' How many levels a recursive common table expression may go down: what
' OPTION (MAXRECURSION n) on the statement says, or a real server's hundred.
' Nought means no limit, which here is the most a real server accepts, so a
' level that never runs out still ends.
Private Function RecursionLimit(ByVal statement As String) As Long
    Dim at As Long

    RecursionLimit = 100
    at = InStr(1, statement, "MAXRECURSION", vbTextCompare)
    If at = 0 Then Exit Function
    RecursionLimit = CLng(Val(Mid$(statement, at + 12)))
    If RecursionLimit <= 0 Then RecursionLimit = 32767
End Function

' A read run from its text for its columns and rows, with the parser left
' where it was.
Private Sub ReadInto(ByVal written As String, ByRef columns As Collection, _
                     ByRef rows As Collection)
    Dim kept As Variant

    ' What the reads before this one answered says nothing about this one:
    ' each level of a recursive expression is a different set of rows under
    ' the same name.
    Set mSubqueryAnswers = Nothing
    Set mSubqueryLookups = Nothing
    Set mExistsLookups = Nothing
    Set mLoadedSources = Nothing

    kept = SaveParseState()
    On Error GoTo Restore
    Set mTokens = Tokenize(written)
    mCursor = 1
    ParseSelect
    ExecuteWholeSelect columns, rows
    RefuseTheRestOfARead
    RestoreParseState kept
    Exit Sub

Restore:
    RestoreParseState kept
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Whether a read takes rows from a table of this name.
Private Function NamesItself(ByVal body As String, ByVal name As String) As Boolean
    Dim tokens As Collection
    Dim index As Long
    Dim before As String

    Set tokens = Tokenize(body)
    For index = 2 To tokens.count
        If tokens(index)(0) = T_NAME Then
            before = UCase$(CStr(tokens(index - 1)(1)))
            If before = "FROM" Or before = "JOIN" Then
                If StrComp(Replace(Replace(BareColumnName( _
                           CStr(tokens(index)(1))), "[", ""), "]", ""), _
                           name, vbTextCompare) = 0 Then
                    NamesItself = True
                    Exit Function
                End If
            End If
        End If
    Next index
End Function

' Where the first UNION ALL outside any brackets starts, or nought.
Private Function TopLevelUnionAll(ByVal text As String) As Long
    Dim at As Long
    Dim depth As Long
    Dim ch As String
    Dim ends As String
    Dim startsWord As Boolean

    at = 1
    Do While at <= Len(text)
        ch = Mid$(text, at, 1)
        If ch = "'" Or ch = """" Or ch = "[" Then
            at = PastOne(text, at)
        Else
            If ch = "(" Then
                depth = depth + 1
            ElseIf ch = ")" Then
                depth = depth - 1
            ElseIf depth = 0 And UCase$(ch) = "U" Then
                ' Mid$ from nought raises, and Or does not stop at its left.
                If at > 1 Then
                    startsWord = Not IsNameChar(Mid$(text, at - 1, 1))
                Else
                    startsWord = True
                End If
                If startsWord And UCase$(WordAt(text, at, ends)) = "UNION" Then
                    If UCase$(WordAt(text, SkipSpaceAt(text, CLng(ends)), _
                                     ends)) = "ALL" Then
                        TopLevelUnionAll = at
                        Exit Function
                    End If
                End If
            End If
            at = at + 1
        End If
    Loop
End Function

Private Sub StoreCommonTable(ByVal name As String, _
                             ByVal columns As Collection, _
                             ByVal rows As Collection)
    Dim kept(0 To 1) As Variant

    Set kept(0) = columns
    Set kept(1) = rows
    On Error Resume Next
    mCommonTables.Remove LCase$(name)
    On Error GoTo 0
    mCommonTables.Add kept, LCase$(name)
End Sub

' A common table expression the statement began with, by name.
Private Function CommonTable(ByVal name As String, _
                             ByRef columns As Collection, _
                             ByRef rows As Collection) As Boolean
    Dim kept As Variant

    If mCommonTables Is Nothing Then Exit Function
    If InStr(name, ".") > 0 Then Exit Function
    On Error Resume Next
    kept = mCommonTables(LCase$(Replace(Replace(name, "[", ""), "]", "")))
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0
    Set columns = kept(0)
    Set rows = kept(1)
    CommonTable = True
End Function

' Whether a set operator carries the read on past where it just stopped.
Private Function FollowsSetOperator() As Boolean
    FollowsSetOperator = PeekIs("UNION") Or PeekIs("EXCEPT") _
                         Or PeekIs("INTERSECT")
End Function

' One read, and the rest of it when a set operator carries it on.
'
' A UNION is read after the SELECT it follows, so anything that runs a
' nested SELECT and stops where that SELECT stopped keeps the first branch
' and drops the others without saying so. A derived table over a UNION ALL
' of two rows answered with one of them, and COUNT(*) over that derived
' table answered one.
Private Sub ExecuteWholeSelect(ByRef outColumns As Collection, _
                               ByRef outRows As Collection)
    Set outColumns = New Collection
    Set outRows = New Collection
    ExecuteSelect outColumns, outRows
    If FollowsSetOperator() Then CombineReads outColumns, outRows
End Sub

' The rest of a statement that joined two reads together.
'
' Each side is parsed and run in turn and the rows are put side by side under
' the first side's column names, which is where SQL Server takes them from
' too. UNION drops the rows that repeat, UNION ALL keeps them, EXCEPT keeps
' what the right side does not have and INTERSECT keeps only what it does.
Private Sub CombineReads(ByRef outColumns As Collection, _
                         ByRef outRows As Collection)
    Dim operator As String
    Dim keepAll As Boolean
    Dim nextColumns As Collection
    Dim nextRows As Collection
    Dim ordering As Collection
    Dim values As Variant

    Do While PeekIs("UNION") Or PeekIs("EXCEPT") Or PeekIs("INTERSECT")
        operator = UCase$(TakeText())
        keepAll = Accept("ALL")

        ParseSelect

        ' The ORDER BY at the end of the statement names the columns going
        ' out, not this side's. Sorting here would sort one side's rows by a
        ' name that side has not got.
        Set ordering = mOrderBy
        Set mOrderBy = New Collection
        Set nextColumns = New Collection
        Set nextRows = New Collection
        ExecuteSelect nextColumns, nextRows
        Set mOrderBy = ordering

        If nextColumns.count <> outColumns.count Then
            RaiseQueryError 205, _
                "All queries combined using a UNION, INTERSECT or EXCEPT " & _
                "operator must have an equal number of expressions in " & _
                "their target lists."
        End If

        Select Case operator
            Case "UNION"
                For Each values In nextRows
                    outRows.Add values
                Next values
                If Not keepAll Then Set outRows = DistinctRows(outRows)
            Case "EXCEPT"
                Set outRows = DistinctRows(RowsNotIn(outRows, nextRows))
            Case Else
                Set outRows = DistinctRows(RowsAlsoIn(outRows, nextRows))
        End Select

        ' It belongs to the whole statement, and now that both sides are
        ' together the columns it names are there to sort by. A position
        ' names one of those columns rather than one of the last side's
        ' select list, which is what it would have meant on its own.
        If ordering.count > 0 Then
            Set mOrderBy = OrderOverOutput(ordering, outColumns)
            Set outRows = SortRows(outRows, outColumns)
            Set mOrderBy = ordering
        End If
    Loop
End Sub

' An ORDER BY read against the columns going out.
'
' ORDER BY 1 after a UNION means the first column of the answer. Left as a
' position it would be looked up in the last side's select list, which is a
' different list with different names in it.
Private Function OrderOverOutput(ByVal ordering As Collection, _
                                 ByVal outColumns As Collection) As Collection
    Dim over As Collection
    Dim term As Variant
    Dim node As Variant
    Dim position As Long
    Dim rebuilt(0 To 1) As Variant

    Set over = New Collection
    For Each term In ordering
        node = term(0)
        If IsArray(node) Then
            If node(0) = N_LITERAL Then
                If Not IsNull(node(1)) Then
                    If IsNumeric(node(1)) Then
                        position = CLng(node(1))
                        If position >= 1 And position <= outColumns.count Then
                            node = NodeOf(N_COLUMN, _
                                CStr(outColumns(position)(COLUMN_NAME)), _
                                Empty, Empty)
                        End If
                    End If
                End If
            End If
        End If
        rebuilt(0) = node
        rebuilt(1) = term(1)
        over.Add rebuilt
    Next term
    Set OrderOverOutput = over
End Function

' The rows of the left side that no row of the right side matches.
Private Function RowsNotIn(ByVal rows As Collection, _
                           ByVal others As Collection) As Collection
    Dim kept As Collection
    Dim seen As Collection
    Dim values As Variant

    Set seen = New Collection
    Set kept = New Collection
    For Each values In others
        RememberRow seen, values
    Next values
    For Each values In rows
        If Not RowIsIn(seen, values) Then kept.Add values
    Next values
    Set RowsNotIn = kept
End Function

' The rows of the left side that some row of the right side matches.
Private Function RowsAlsoIn(ByVal rows As Collection, _
                            ByVal others As Collection) As Collection
    Dim kept As Collection
    Dim seen As Collection
    Dim values As Variant

    Set seen = New Collection
    Set kept = New Collection
    For Each values In others
        RememberRow seen, values
    Next values
    For Each values In rows
        If RowIsIn(seen, values) Then kept.Add values
    Next values
    Set RowsAlsoIn = kept
End Function

Private Function DistinctRows(ByVal rows As Collection) As Collection
    Dim kept As Collection
    Dim seen As Collection
    Dim values As Variant

    Set seen = New Collection
    Set kept = New Collection
    For Each values In rows
        If RememberRow(seen, values) Then kept.Add values
    Next values
    Set DistinctRows = kept
End Function

' EXEC procedure, whose rows are an answer like any other, or EXEC of a
' statement written as text.
Private Sub RunExec(ByVal written As String, ByVal answers As Collection)
    Dim name As String
    Dim outColumns As Collection
    Dim outRows As Collection
    Dim rest As String
    Dim ends As String
    Dim inner As String

    WordAt written, 1, ends
    rest = TrimWhitespace(Mid$(written, CLng(ends)))

    ' EXEC ('select ...') is a statement a client built as a string, which
    ' SSMS uses to ask a question it does not want to fail on older servers.
    If Left$(rest, 1) = "(" Then
        inner = QuotedInside(rest)
        If Len(inner) > 0 Then
            RunInner inner, answers
            Exit Sub
        End If
    End If

    name = ProcedureFromBatch(written)

    ' A registry read, which answers into the variable it was handed rather
    ' than with a row.
    If BareProcedureName(name) = "xp_instance_regread" _
       Or BareProcedureName(name) = "xp_regread" Then
        RunRegistryRead Mid$(rest, Len(name) + 1)
        Exit Sub
    End If

    ' The same thing named outright: EXEC sp_executesql N'...'. The
    ' declarations and values after it are for parameters the statement
    ' carries, and a statement written as a batch has none.
    If BareProcedureName(name) = "sp_executesql" Then
        inner = FirstQuotedIn(Mid$(rest, Len(name) + 1))
        If Len(inner) > 0 Then
            RunInner inner, answers
            Exit Sub
        End If
    End If
    Set outColumns = New Collection
    Set outRows = New Collection
    ProcedureResult name, outColumns, outRows
    answers.Add AnswerOf(outColumns, outRows)
End Sub

' EXEC master.dbo.xp_instance_regread @root, @path, N'Name', @out OUTPUT
'
' A client reads a dozen of the server's settings this way while building
' its own picture of one, and a procedure it cannot find ends the batch:
' asking for the rows of a table is what fails, several menus above.
'
' Nothing here keeps a registry, so what is answered is what this server can
' say about itself and nothing where it cannot. The query is written to
' expect that: most of the values it reads go through ISNULL on the way out.
Private Sub RunRegistryRead(ByVal arguments As String)
    Dim parts As Collection
    Dim wanted As String
    Dim target As String
    Dim last As String
    Dim where As String
    Dim ends As String

    Set parts = SplitTopLevel(arguments, ",")
    If parts.count < 4 Then Exit Sub

    wanted = TrimWhitespace(CStr(parts(3)))
    If Left$(wanted, 1) = "N" Then wanted = Mid$(wanted, 2)
    wanted = Replace(TrimWhitespace(wanted), "'", "")

    last = TrimWhitespace(CStr(parts(parts.count)))
    target = WordAt(last, 1, ends)
    If Left$(target, 1) <> "@" Then Exit Sub

    ' The name alone does not say which protocol Enabled belongs to; the
    ' path handed in does, and it is a variable set earlier in the batch.
    where = CStr(NullToEmpty(EvaluateStandalone(CStr(parts(2)))))

    SetVariable target, RegistryValue(wanted, where)
End Sub

' What this server would have kept under that name.
Private Function RegistryValue(ByVal wanted As String, _
                               ByVal where As String) As Variant
    Select Case UCase$(wanted)
        Case "AUDITLEVEL"
            ' Nothing is audited, which is what none means.
            RegistryValue = 0
        Case "NUMERRORLOGS"
            ' The log is a range on a worksheet, not a file on disk.
            RegistryValue = 0
        Case "LOGINMODE"
            ' Windows authentication and no other, which is the truth: this
            ' never sees a password and could not check one.
            RegistryValue = 1
        Case "SQLPATH", "SQLDATAROOT"
            RegistryValue = INSTALL_PATH
        Case "BACKUPDIRECTORY"
            RegistryValue = DEFAULT_DATA_PATH
        Case "ENABLED"
            ' Whether a protocol is on. The name alone does not say which
            ' one, but the path does, and only one of them is ever on: the
            ' client asking is already talking over it.
            RegistryValue = IIf(Right$(UCase$(where), 4) = "\TCP", 1, 0)
        Case Else
            ' A setting this has no answer for, which the statement reading
            ' it is written to expect.
            RegistryValue = Null
    End Select
End Function

' Whether every item of the select list assigns to a variable.
Private Function AssignsToVariables() As Boolean
    Dim item As Variant
    Dim node As Variant

    If mSelectItems.count = 0 Then Exit Function

    For Each item In mSelectItems
        node = item(0)
        If Not IsArray(node) Then Exit Function
        If node(0) <> N_BINARY Then Exit Function
        If node(1) <> "=" Then Exit Function
        If Not IsArray(node(2)) Then Exit Function
        If node(2)(0) <> N_COLUMN Then Exit Function
        If Left$(CStr(node(2)(1)), 1) <> "@" Then Exit Function
    Next item
    AssignsToVariables = True
End Function

' SELECT @x = expression [FROM ...], which produces no rows.
'
' The last row read wins, which is what a real server does: a variable
' assigned from a read that matched nothing keeps whatever it had.
Private Sub AssignFromRead()
    Dim matching As Collection
    Dim sourceColumns As Collection
    Dim values As Variant
    Dim item As Variant
    Dim node As Variant
    Dim blank() As Variant

    ReadSource matching, sourceColumns
    If matching.count > 0 Then values = matching(matching.count)

    For Each item In mSelectItems
        node = item(0)
        If NodeHasAggregate(node(3)) Then
            ' An aggregate is over the rows rather than over one of them.
            ' Taking the last row and asking it for a COUNT asked the
            ' expression evaluator for a function it has never had, and a
            ' count over no rows at all is nought rather than nothing.
            SetVariable CStr(node(2)(1)), _
                EvaluateAggregate(node(3), matching, sourceColumns)
        ElseIf matching.count > 0 Then
            SetVariable CStr(node(2)(1)), _
                Evaluate(node(3), values, sourceColumns)
        End If
    Next item
End Sub

Private Sub SetVariable(ByVal name As String, ByVal value As Variant)
    Dim key As String

    key = LCase$(Trim$(name))

    ' A parameter is set where it is bound, so an OUTPUT parameter hands back
    ' what the statement made it. Set as a variable of the connection, it
    ' outlived the call, and the parameter, which is found first, still read
    ' as it arrived.
    If Not mParameters Is Nothing Then
        If HasKey(mParameters, key) Then
            mParameters.Remove key
            mParameters.Add value, key
            Exit Sub
        End If
    End If

    On Error Resume Next
    mVariables.Remove key
    On Error GoTo 0
    mVariables.Add value, key
End Sub

' One expression, evaluated with no row behind it.
Private Function EvaluateStandalone(ByVal written As String) As Variant
    Dim keptTokens As Collection
    Dim keptCursor As Long
    Dim node As Variant
    Dim blank() As Variant
    Dim columns As Collection

    Set keptTokens = mTokens
    keptCursor = mCursor

    Set mTokens = Tokenize(written)
    mCursor = 1
    node = ParseExpression()

    Set mTokens = keptTokens
    mCursor = keptCursor

    ReDim blank(0 To 0)
    Set columns = New Collection
    EvaluateStandalone = Evaluate(node, blank, columns)
End Function

' An expression already parsed, evaluated with no row behind it.
Private Function EvaluateHere(ByVal node As Variant) As Variant
    Dim blank() As Variant
    Dim columns As Collection

    ReDim blank(0 To 0)
    Set columns = New Collection
    EvaluateHere = Evaluate(node, blank, columns)
End Function


' ============================================================================
' More than one source
'
' A join and a CROSS APPLY both put two sets of columns beside each other and
' decide which pairs of rows survive. Everything after that, the WHERE, the
' projection, the ordering, works on the combined space and needs to know
' nothing about how it was built.
'
' A column carries the name of the source it came from, so dtb.name and
' avar.name are two columns rather than one asked for twice.
' ============================================================================

Private Function FromItem(ByVal kind As Long, ByVal name As String, _
                          ByVal alias As String, ByVal onNode As Variant, _
                          ByVal columnNames As Collection, _
                          ByVal valueRows As Collection, _
                          Optional ByVal derived As Collection = Nothing) As Variant
    Dim out(0 To 6) As Variant

    out(0) = kind
    out(1) = name
    out(2) = alias
    If IsObject(onNode) Then Set out(3) = onNode Else out(3) = onNode
    Set out(4) = columnNames
    Set out(5) = valueRows
    Set out(6) = derived
    FromItem = out
End Function

' FROM source [alias] { , source | JOIN source [alias] ON condition |
'                         CROSS APPLY ... }
Private Sub ReadFromClause()
    Dim item As Variant

    Set mFromItems = New Collection
    item = ReadFromItem(FROM_BASE)
    mFromItems.Add item
    mFromName = CStr(item(1))

    Do
        If Accept("CROSS") Then
            If Accept("APPLY") Then
                mFromItems.Add ReadFromItem(FROM_APPLY)
            ElseIf Accept("JOIN") Then
                mFromItems.Add ReadFromItem(FROM_CROSS)
            Else
                Exit Do
            End If
        ElseIf Accept("OUTER") Then
            Expect "APPLY"
            mFromItems.Add ReadFromItem(FROM_OUTER_APPLY)
        ElseIf Accept("INNER") Then
            Expect "JOIN"
            mFromItems.Add ReadJoinItem(FROM_INNER)
        ElseIf Accept("LEFT") Then
            Accept "OUTER"
            Expect "JOIN"
            mFromItems.Add ReadJoinItem(FROM_LEFT)
        ElseIf Accept("RIGHT") Then
            Accept "OUTER"
            Expect "JOIN"
            mFromItems.Add ReadJoinItem(FROM_RIGHT)
        ElseIf Accept("FULL") Then
            Accept "OUTER"
            Expect "JOIN"
            mFromItems.Add ReadJoinItem(FROM_FULL)
        ElseIf Accept("JOIN") Then
            mFromItems.Add ReadJoinItem(FROM_INNER)
        ElseIf Accept(",") Then
            ' A comma between two sources is a cross join written the way it
            ' was written before there was a word for it, and SSMS still
            ' writes it: FROM master.sys.databases AS dtb, sys.tables AS tbl.
            mFromItems.Add ReadFromItem(FROM_CROSS)
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function ReadJoinItem(ByVal kind As Long) As Variant
    Dim item As Variant
    Dim onNode As Variant

    item = ReadFromItem(kind)
    If Accept("ON") Then
        onNode = ParseExpression()
        If IsObject(onNode) Then Set item(3) = onNode Else item(3) = onNode
    End If
    ReadJoinItem = item
End Function

' One source: a name, or a bracketed VALUES list, with the alias and column
' names a client may give it.
Private Function ReadFromItem(ByVal kind As Long) As Variant
    Dim name As String
    Dim alias As String
    Dim columnNames As Collection
    Dim valueRows As Collection
    Dim derived As Collection

    Set columnNames = New Collection
    Set valueRows = New Collection

    If PeekIs("(") Then
        mCursor = mCursor + 1
        If PeekIs("SELECT") Then
            ' A read used as a table. Its tokens are kept and run when the
            ' source is loaded, the same way a read used as a value is.
            Set derived = BracketedTokens()
        Else
            Expect "VALUES"
            Set valueRows = ReadValuesNodes()
            Expect ")"
            If kind = FROM_BASE Then kind = FROM_VALUES
        End If
    Else
        name = TakeText()
    End If

    ' A bracket straight after a name holds a function's arguments or an
    ' older table hint, (NOLOCK) with no WITH. Read as a column list, it
    ' renamed the table's first column NOLOCK. Neither says which rows.
    If Len(name) > 0 And PeekIs("(") Then
        mCursor = mCursor + 1
        BracketedTokens
    End If

    ' An alias, with or without the AS that may introduce it.
    If Accept("AS") Then
        alias = TakeText()
    ElseIf PeekKind() = T_NAME Then
        If Not IsClauseWord(PeekText()) Then alias = TakeText()
    End If

    SkipTableHints

    ' Column names for a derived table or a VALUES list, which have no name
    ' of their own; after a table's alias, the older hint again.
    If PeekIs("(") Then
        mCursor = mCursor + 1
        If Len(name) > 0 Then
            BracketedTokens
        Else
            Do
                columnNames.Add TakeText()
            Loop While Accept(",")
            Expect ")"
        End If
    End If

    If Len(alias) = 0 Then alias = BareColumnName(name)
    ReadFromItem = FromItem(kind, name, alias, Empty, columnNames, valueRows, _
                            derived)
End Function

' WITH (NOLOCK) and the other table hints, which say how a real server should
' lock or find the rows and never which rows they are, so are passed over.
' Refused as the rest of a read, WITH (NOLOCK) failed a query a real server
' answers.
Private Sub SkipTableHints()
    If Not PeekIs("WITH") Then Exit Sub
    If mCursor + 1 > mTokens.count Then Exit Sub
    If mTokens(mCursor + 1)(1) <> "(" Then Exit Sub
    mCursor = mCursor + 2
    BracketedTokens
End Sub

' The rows of a VALUES list, each held as expressions rather than values:
' a CROSS APPLY evaluates them against the row it is applied to.
Private Function ReadValuesNodes() As Collection
    Dim rows As Collection
    Dim row As Collection

    Set rows = New Collection
    Do
        Expect "("
        Set row = New Collection
        Do
            row.Add ParseExpression()
        Loop While Accept(",")
        Expect ")"
        rows.Add row
    Loop While Accept(",")
    Set ReadValuesNodes = rows
End Function

' Every column of one source, named for the source it came from.
Private Function QualifyColumns(ByVal columns As Collection, _
                                ByVal alias As String, _
                                ByVal columnNames As Collection) As Collection
    Dim out As Collection
    Dim index As Long
    Dim column As Variant
    Dim name As String

    Set out = New Collection
    For index = 1 To columns.count
        column = columns(index)
        If Not columnNames Is Nothing Then
            If index <= columnNames.count Then
                name = Replace(Replace(CStr(columnNames(index)), "[", ""), "]", "")
            Else
                name = CStr(column(COLUMN_NAME))
            End If
        Else
            name = CStr(column(COLUMN_NAME))
        End If
        out.Add ColumnOf(name, CLng(column(COLUMN_TYPE)), _
                         CLng(column(COLUMN_SIZE)), CLng(column(COLUMN_FLAGS)), _
                         alias, CLng(NullToZero(column(COLUMN_SCALE))))
    Next index
    Set QualifyColumns = out
End Function

' The rows and columns of one named source.
Private Sub LoadFromItem(ByVal item As Variant, ByRef columns As Collection, _
                         ByRef rows As Collection)
    Dim data As Variant
    Dim entry As Variant
    Dim plain As Collection
    Dim key As String
    Dim loaded As Variant

    If Not item(6) Is Nothing Then
        DerivedSource item, columns, rows
        Exit Sub
    End If

    ' A VALUES list, first in the FROM or joined to what came before it.
    ' Only the first was read as one, and a joined list was looked up as a
    ' table with no name.
    If item(0) = FROM_VALUES Or item(5).count > 0 Then
        Set columns = ValuesColumns(item)
        Set rows = ValuesRows(item, Nothing, Nothing)
        TypeFromValues columns, rows
        Exit Sub
    End If

    ' A common table expression the statement began with, which was run
    ' before the statement was and is read here the way a table is. Not
    ' remembered with the other sources: a recursive one reads a different
    ' set of rows under the same name at every level.
    If CommonTable(CStr(item(1)), plain, rows) Then
        Set columns = QualifyColumns(plain, CStr(item(2)), item(4))
        Exit Sub
    End If

    ' A named source loaded once per statement. A read inside another read
    ' is run for every row of the outer one, and each run was rebuilding the
    ' whole inner table: a thousand rows against a thousand meant a thousand
    ' tables built to answer one question. Nothing changes a source in the
    ' middle of a statement, so what was loaded is what is there.
    key = LCase$(CStr(item(1))) & "|" & LCase$(CStr(item(2)))
    If Not mLoadedSources Is Nothing Then
        On Error Resume Next
        loaded = mLoadedSources(key)
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            Set columns = loaded(0)
            Set rows = loaded(1)
            Exit Sub
        End If
        Err.Clear
        On Error GoTo 0
    End If

    If IsInformationSchema(CStr(item(1))) Then
        If Not BuiltRowset(CStr(item(1)), plain, rows) Then
            data = InformationSchemaTable(CStr(item(1)))
            Set plain = InferColumns(data)
            Set rows = RowsOf(data)
            RememberRowset CStr(item(1)), plain, rows
        End If
        Set columns = QualifyColumns(plain, CStr(item(2)), item(4))
        If mColumnsOnly Then Set rows = New Collection
        Exit Sub
    ElseIf IsSystemView(CStr(item(1))) Then
        If Not BuiltRowset(CStr(item(1)), plain, rows) Then
            data = SystemViewTable(CStr(item(1)))
            Set plain = InferColumns(data)
            Set rows = RowsOf(data)
            RememberRowset CStr(item(1)), plain, rows
        End If
        Set columns = QualifyColumns(plain, CStr(item(2)), item(4))
        If mColumnsOnly Then Set rows = New Collection
        Exit Sub
    Else
        entry = LookUpTable(CStr(item(1)))
        If IsViewEntry(entry) Then
            ' A view answers with whatever its statement answers now, except
            ' where all that was asked for is the shape of the answer: the
            ' columns of a view are kept, and reading them back is what
            ' stops a catalog query running a join of fifty thousand rows
            ' once for every view standing on it.
            If mColumnsOnly Then
                Set plain = EntryColumns(entry)
                Set rows = New Collection
            Else
                ReadView entry, plain, rows
            End If
            Set columns = QualifyColumns(plain, CStr(item(2)), item(4))
            If Not mColumnsOnly Then RememberSource key, columns, rows
            Exit Sub
        End If
        data = entry(1)
        Set plain = entry(2)
    End If

    Set columns = QualifyColumns(plain, CStr(item(2)), item(4))
    If mColumnsOnly Then
        ' No rows, so nothing downstream has any to work over: a join makes
        ' no pairs, a WHERE keeps nothing, and what comes out the far end is
        ' the column list the statement would have produced. The rows are
        ' not remembered either, because the statement that asked for the
        ' shape may go on to want the table itself.
        Set rows = New Collection
        Exit Sub
    End If
    Set rows = RowsOf(data)
    RememberSource key, columns, rows
End Sub

' The columns a VALUES list declares, typed from its first row.
Private Function ValuesColumns(ByVal item As Variant) As Collection
    Dim columns As Collection
    Dim names As Collection
    Dim first As Collection
    Dim index As Long
    Dim name As String

    Set columns = New Collection
    Set names = item(4)
    If item(5).count = 0 Then
        Set ValuesColumns = columns
        Exit Function
    End If
    Set first = item(5)(1)

    For index = 1 To first.count
        If index <= names.count Then
            name = Replace(Replace(CStr(names(index)), "[", ""), "]", "")
        Else
            name = "Column" & index
        End If
        ' Typed as text and retyped from the values once they exist, which is
        ' what a list of literals in a statement amounts to.
        columns.Add ColumnOf(name, TYPE_NVARCHAR, NVARCHAR_SIZED_LIMIT, _
                             SELECT_COLUMN_FLAGS, CStr(item(2)))
    Next index
    Set ValuesColumns = columns
End Function

' The columns of a VALUES list, typed from its values once they exist, the
' way a computed column is, and still named for the list. Left as text, a
' list of dates was declared a column of text, and MAX of it went out as
' the text a date prints as rather than as a datetime.
Private Sub TypeFromValues(ByVal columns As Collection, ByVal rows As Collection)
    Dim kept As Collection
    Dim index As Long
    Dim before As Variant
    Dim after As Variant

    Set kept = New Collection
    For index = 1 To columns.count
        kept.Add columns(index)
    Next index
    RetypeComputedColumns columns, rows, Nothing
    For index = 1 To columns.count
        before = kept(index)
        after = columns(index)
        ReplaceColumn columns, index, _
            ColumnOf(CStr(after(COLUMN_NAME)), CLng(after(COLUMN_TYPE)), _
                     CLng(after(COLUMN_SIZE)), CLng(before(COLUMN_FLAGS)), _
                     CStr(before(COLUMN_QUALIFIER)))
    Next index
End Sub

' The rows of a VALUES list, evaluated against the row it is applied to.
Private Function ValuesRows(ByVal item As Variant, ByVal outerValues As Variant, _
                            ByVal outerColumns As Collection) As Collection
    Dim rows As Collection
    Dim row As Collection
    Dim values As Variant
    Dim index As Long
    Dim blank() As Variant
    Dim columns As Collection

    Set rows = New Collection
    If outerColumns Is Nothing Then
        ReDim blank(0 To 0)
        Set columns = New Collection
    Else
        Set columns = outerColumns
    End If

    For Each row In item(5)
        ReDim values(0 To row.count - 1)
        For index = 1 To row.count
            If outerColumns Is Nothing Then
                values(index - 1) = Evaluate(row(index), blank, columns)
            Else
                values(index - 1) = Evaluate(row(index), outerValues, columns)
            End If
        Next index
        rows.Add values
    Next row
    Set ValuesRows = rows
End Function

' Put a second source beside the first, keeping the pairs that survive.
Private Sub CombineWith(ByVal item As Variant, ByRef columns As Collection, _
                        ByRef rows As Collection)
    Dim rightColumns As Collection
    Dim rightRows As Collection
    Dim combined As Collection
    Dim joined As Collection
    Dim leftValues As Variant
    Dim rightValues As Variant
    Dim merged As Variant
    Dim matched As Boolean
    Dim keepsRight As Boolean
    Dim keptRight() As Boolean
    Dim rightAt As Long
    Dim outcome As Variant
    Dim index As Long
    Dim width As Long
    Dim buckets As Collection
    Dim candidates As Collection
    Dim leftKey As Variant
    Dim condition As Variant

    ' A CROSS APPLY is evaluated once per row, because that is the point of
    ' it: the values on the right refer to the row on the left.
    If item(0) = FROM_APPLY Or item(0) = FROM_OUTER_APPLY Then
        If Not item(6) Is Nothing Then
            ApplyRead item, columns, rows
            Exit Sub
        End If
        Set rightColumns = ValuesColumns(item)
        Set combined = JoinedColumns(columns, rightColumns)
        Set joined = New Collection
        For Each leftValues In rows
            Set rightRows = ValuesRows(item, leftValues, columns)
            For Each rightValues In rightRows
                joined.Add MergedRow(leftValues, rightValues, _
                                     rightColumns.count)
            Next rightValues
        Next leftValues
        Set columns = combined
        Set rows = joined
        Exit Sub
    End If

    LoadFromItem item, rightColumns, rightRows
    Set combined = JoinedColumns(columns, rightColumns)
    width = rightColumns.count
    Set joined = New Collection
    condition = item(3)
    If Not IsEmpty(condition) Then condition = Prepared(condition, combined)

    ' A join that keeps the rows on the right has to know which of them
    ' were kept, and a bucket hands back a subset with no way to say where
    ' in the whole each row sat. Those two walk every pair, which is what
    ' every join did before there were buckets.
    keepsRight = (item(0) = FROM_RIGHT Or item(0) = FROM_FULL)
    If keepsRight Then
        If rightRows.count > 0 Then ReDim keptRight(1 To rightRows.count)
    Else
        ' The rows on the right, grouped by whatever the condition matches
        ' them on. Nothing when the condition has no such thing, and then
        ' every pair is tried as before.
        Set buckets = JoinBuckets(item(3), columns, rightColumns, rightRows, _
                                  leftKey)
        If Not buckets Is Nothing Then leftKey = Prepared(leftKey, columns)
    End If

    For Each leftValues In rows
        matched = False
        rightAt = 0

        If buckets Is Nothing Then
            Set candidates = rightRows
        Else
            Set candidates = JoinCandidates(buckets, leftKey, leftValues, _
                                            columns)
        End If

        For Each rightValues In candidates
            rightAt = rightAt + 1
            merged = MergedRow(leftValues, rightValues, width)
            If IsEmpty(item(3)) Then
                joined.Add merged
                matched = True
                If keepsRight Then keptRight(rightAt) = True
            Else
                ' The whole condition, still, on the rows the key kept: the
                ' lookup drops pairs the equality would have dropped and
                ' decides nothing else. Resolved against the joined columns
                ' once, before the pairs.
                outcome = Evaluate(condition, merged, combined)
                If Not IsNull(outcome) Then
                    If CBool(outcome) Then
                        joined.Add merged
                        matched = True
                        If keepsRight Then keptRight(rightAt) = True
                    End If
                End If
            End If
        Next rightValues

        ' A left join keeps the row it started with, with nothing beside it.
        If Not matched Then
            If item(0) = FROM_LEFT Or item(0) = FROM_FULL Then
                joined.Add MergedRow(leftValues, Empty, width)
            End If
        End If
    Next leftValues

    ' And a right join keeps the rows on the right that nothing matched,
    ' with nothing in front of them.
    If keepsRight Then
        rightAt = 0
        For Each rightValues In rightRows
            rightAt = rightAt + 1
            If Not keptRight(rightAt) Then
                joined.Add MergedRow(NullRow(columns.count), rightValues, _
                                     width)
            End If
        Next rightValues
    End If

    Set columns = combined
    Set rows = joined
End Sub

' CROSS APPLY (SELECT ...) and OUTER APPLY: the read run once for each row
' on the left, seeing that row the way a correlated subquery sees the row it
' runs inside, and its rows put beside it. OUTER APPLY keeps a left row the
' read answered nothing for, with nothing beside it. Only a VALUES list was
' applied before.
Private Sub ApplyRead(ByVal item As Variant, ByRef columns As Collection, _
                      ByRef rows As Collection)
    Dim innerColumns As Collection
    Dim innerRows As Collection
    Dim rightColumns As Collection
    Dim joined As Collection
    Dim leftValues As Variant
    Dim rightValues As Variant
    Dim width As Long

    Set joined = New Collection
    For Each leftValues In rows
        RunSubquery item(6), leftValues, columns, innerColumns, innerRows
        If rightColumns Is Nothing Then
            Set rightColumns = QualifyColumns(innerColumns, CStr(item(2)), _
                                              item(4))
            width = rightColumns.count
        End If
        If innerRows.count = 0 Then
            If item(0) = FROM_OUTER_APPLY Then
                joined.Add MergedRow(leftValues, Empty, width)
            End If
        Else
            For Each rightValues In innerRows
                joined.Add MergedRow(leftValues, rightValues, width)
            Next rightValues
        End If
    Next leftValues

    ' The shape of what the read answers, even with no row on the left to
    ' run it for.
    If rightColumns Is Nothing Then
        RunSubquery item(6), NullRow(columns.count), columns, innerColumns, _
                    innerRows
        Set rightColumns = QualifyColumns(innerColumns, CStr(item(2)), item(4))
    End If
    Set columns = JoinedColumns(columns, rightColumns)
    Set rows = joined
End Sub

' A row of nothing, as wide as the columns in front of a join.
Private Function NullRow(ByVal width As Long) As Variant
    Dim out As Variant
    Dim index As Long

    ReDim out(0 To width - 1)
    For index = 0 To width - 1
        out(index) = Null
    Next index
    NullRow = out
End Function

Private Function JoinedColumns(ByVal left As Collection, _
                               ByVal right As Collection) As Collection
    Dim out As Collection
    Dim column As Variant

    Set out = New Collection
    For Each column In left
        out.Add column
    Next column
    For Each column In right
        out.Add column
    Next column
    Set JoinedColumns = out
End Function

Private Function MergedRow(ByVal leftValues As Variant, _
                           ByVal rightValues As Variant, _
                           ByVal width As Long) As Variant
    Dim out As Variant
    Dim leftWidth As Long
    Dim index As Long

    leftWidth = UBound(leftValues) - LBound(leftValues) + 1
    ReDim out(0 To leftWidth + width - 1)

    For index = 0 To leftWidth - 1
        out(index) = leftValues(LBound(leftValues) + index)
    Next index
    For index = 0 To width - 1
        If IsEmpty(rightValues) Then
            out(leftWidth + index) = Null
        ElseIf index <= UBound(rightValues) - LBound(rightValues) Then
            out(leftWidth + index) = rightValues(LBound(rightValues) + index)
        Else
            out(leftWidth + index) = Null
        End If
    Next index
    MergedRow = out
End Function

' Every row of a two-dimensional source, header row dropped.
Private Function RowsOf(ByRef data As Variant) As Collection
    Dim rows As Collection
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim first As Long
    Dim last As Long
    Dim values As Variant

    Set rows = New Collection
    first = LBound(data, 2)
    last = UBound(data, 2)

    For rowIndex = LBound(data, 1) + 1 To UBound(data, 1)
        ReDim values(0 To last - first)
        For columnIndex = first To last
            values(columnIndex - first) = data(rowIndex, columnIndex)
        Next columnIndex
        rows.Add values
    Next rowIndex
    Set RowsOf = rows
End Function

' The rows a WHERE clause keeps, in the order an ORDER BY asks for.
Private Function FilterRows(ByVal rows As Collection, _
                            ByRef columns As Collection) As Collection
    Dim kept As Collection
    Dim values As Variant
    Dim keep As Variant
    Dim condition As Variant

    If IsEmpty(mWhereNode) Then
        Set kept = rows
    Else
        Set kept = New Collection
        ' Resolved against the columns once rather than once a row.
        condition = Prepared(mWhereNode, columns)
        For Each values In rows
            keep = Evaluate(condition, values, columns)
            If Not IsNull(keep) Then
                If CBool(keep) Then kept.Add values
            End If
        Next values
    End If

    ' Window functions are worked out over every row the WHERE kept, before
    ' any is sorted, projected or cut short by a TOP, which is when a real
    ' server works them out. Each becomes a column of its own on the rows,
    ' so an ORDER BY can sort on one.
    If HasWindow() Then
        If Summarising() Then
            RaiseQueryError ERR_NUMBER_BASE, "A window function over a " & _
                "grouped or aggregated read is not supported here."
        End If
        ComputeWindows kept, columns
    End If

    If mOrderBy.count > 0 And Not Summarising() Then
        Set kept = SortRows(kept, columns)
    End If
    Set FilterRows = kept
End Function


' ============================================================================
' Window functions
'
' ROW_NUMBER, RANK and the rest, and an aggregate with OVER after it, are
' worked out over the rows around each row rather than over a group: every
' row keeps its place in the answer, and the window adds a value to it.
' ============================================================================

' OVER ([PARTITION BY ...] [ORDER BY ...] [ROWS | RANGE frame]) after a call.
Private Function ParseOver(ByVal callNode As Variant) As Variant
    Dim partitions As Collection
    Dim orders As Collection
    Dim order(0 To 1) As Variant
    Dim spec(0 To 4) As Variant
    Dim kind As String
    Dim startBound As Long
    Dim endBound As Long

    Expect "OVER"
    Expect "("
    Set partitions = New Collection
    Set orders = New Collection

    If Accept("PARTITION") Then
        Expect "BY"
        Do
            partitions.Add ParseExpression()
        Loop While Accept(",")
    End If

    If Accept("ORDER") Then
        Expect "BY"
        Do
            order(0) = ParseExpression()
            order(1) = False
            If Accept("DESC") Then
                order(1) = True
            Else
                Accept "ASC"
            End If
            orders.Add order
        Loop While Accept(",")
    End If

    If Accept("ROWS") Then
        kind = "ROWS"
    ElseIf Accept("RANGE") Then
        kind = "RANGE"
    End If
    If Len(kind) > 0 Then
        If Accept("BETWEEN") Then
            startBound = FrameBound()
            Expect "AND"
            endBound = FrameBound()
        Else
            startBound = FrameBound()
        End If
        ' A range is of values rather than rows, and takes no distance.
        If kind = "RANGE" And (Not IsFrameEdge(startBound) Or _
                               Not IsFrameEdge(endBound)) Then
            RaiseQueryError ERR_NUMBER_BASE, "RANGE is only supported with " & _
                "UNBOUNDED and CURRENT ROW window frame delimiters."
        End If
    End If
    Expect ")"

    Set spec(0) = partitions
    Set spec(1) = orders
    spec(2) = kind
    spec(3) = startBound
    spec(4) = endBound
    ParseOver = NodeOf(N_WINDOW, callNode, spec, Empty)
End Function

' One end of a frame: UNBOUNDED PRECEDING or FOLLOWING, CURRENT ROW, or a
' number of rows either side, as an offset from the current row.
Private Function FrameBound() As Long
    If Accept("UNBOUNDED") Then
        If Accept("PRECEDING") Then
            FrameBound = WINDOW_UNBOUNDED_PRECEDING
        Else
            Expect "FOLLOWING"
            FrameBound = WINDOW_UNBOUNDED_FOLLOWING
        End If
    ElseIf Accept("CURRENT") Then
        Expect "ROW"
    Else
        FrameBound = CLng(Val(TakeText()))
        If Accept("PRECEDING") Then
            FrameBound = -FrameBound
        Else
            Expect "FOLLOWING"
        End If
    End If
End Function

Private Function IsFrameEdge(ByVal bound As Long) As Boolean
    IsFrameEdge = (bound = 0 Or bound = WINDOW_UNBOUNDED_PRECEDING Or _
                   bound = WINDOW_UNBOUNDED_FOLLOWING)
End Function

' Whether the select list or the ORDER BY holds a window.
Private Function HasWindow() As Boolean
    Dim item As Variant

    For Each item In mSelectItems
        If NodeHasWindow(item(0)) Then
            HasWindow = True
            Exit Function
        End If
    Next item
    For Each item In mOrderBy
        If NodeHasWindow(item(0)) Then
            HasWindow = True
            Exit Function
        End If
    Next item
End Function

Private Function NodeHasWindow(ByVal node As Variant) As Boolean
    Dim argument As Variant

    If Not IsArray(node) Then Exit Function

    Select Case node(0)
        Case N_WINDOW
            NodeHasWindow = True
        Case N_CALL
            For Each argument In node(2)
                If NodeHasWindow(argument) Then
                    NodeHasWindow = True
                    Exit Function
                End If
            Next argument
        Case N_BINARY
            NodeHasWindow = NodeHasWindow(node(2))
            If Not NodeHasWindow And Not IsObject(node(3)) Then
                NodeHasWindow = NodeHasWindow(node(3))
            End If
        Case N_UNARY
            NodeHasWindow = NodeHasWindow(node(2))
    End Select
End Function

' The same tree with each window replaced by the column its value will be
' read from, the windows collected in the order they were met.
Private Function RewriteWindows(ByVal node As Variant, _
                                ByVal found As Collection) As Variant
    Dim rebuilt As Collection
    Dim argument As Variant

    If Not IsArray(node) Then
        RewriteWindows = node
        Exit Function
    End If

    Select Case node(0)
        Case N_WINDOW
            found.Add node
            RewriteWindows = NodeOf(N_COLUMN, "$window" & found.count, _
                                    Empty, Empty)
        Case N_CALL
            Set rebuilt = New Collection
            For Each argument In node(2)
                rebuilt.Add RewriteWindows(argument, found)
            Next argument
            RewriteWindows = NodeOf(N_CALL, node(1), rebuilt, node(3))
        Case N_BINARY
            If IsObject(node(3)) Then
                Set rebuilt = New Collection
                For Each argument In node(3)
                    rebuilt.Add RewriteWindows(argument, found)
                Next argument
                RewriteWindows = NodeOf(N_BINARY, node(1), _
                    RewriteWindows(node(2), found), rebuilt)
            Else
                RewriteWindows = NodeOf(N_BINARY, node(1), _
                    RewriteWindows(node(2), found), _
                    RewriteWindows(node(3), found))
            End If
        Case N_UNARY
            RewriteWindows = NodeOf(N_UNARY, node(1), _
                RewriteWindows(node(2), found), Empty)
        Case Else
            RewriteWindows = node
    End Select
End Function

' Every window in the select list and the ORDER BY, worked out and put on
' the rows as columns of their own, and each replaced by its column. Copies
' of the rows are widened rather than the rows themselves, which may be the
' ones a source keeps for the next read of it.
Private Sub ComputeWindows(ByRef rows As Collection, _
                           ByRef columns As Collection)
    Dim found As Collection
    Dim rewritten As Collection
    Dim item As Variant
    Dim width As Long
    Dim count As Long
    Dim extended() As Variant
    Dim grown() As Variant
    Dim values As Variant
    Dim index As Long
    Dim position As Long
    Dim results As Variant
    Dim widened As Collection
    Dim column As Variant

    Set found = New Collection
    Set rewritten = New Collection
    For Each item In mSelectItems
        item(0) = RewriteWindows(item(0), found)
        rewritten.Add item
    Next item
    Set mSelectItems = rewritten
    Set rewritten = New Collection
    For Each item In mOrderBy
        item(0) = RewriteWindows(item(0), found)
        rewritten.Add item
    Next item
    Set mOrderBy = rewritten

    width = columns.count
    count = rows.count
    ReDim extended(0 To IIf(count > 0, count - 1, 0))
    For Each values In rows
        ReDim grown(0 To width + found.count - 1)
        For position = 0 To width - 1
            If LBound(values) + position <= UBound(values) Then
                grown(position) = values(LBound(values) + position)
            End If
        Next position
        extended(index) = grown
        index = index + 1
    Next values

    Set widened = New Collection
    For Each column In columns
        widened.Add column
    Next column

    For index = 1 To found.count
        results = WindowValues(found(index), extended, count, columns)
        For position = 0 To count - 1
            grown = extended(position)
            grown(width + index - 1) = results(position)
            extended(position) = grown
        Next position
        widened.Add WindowColumn(found(index), index, results, count, columns)
    Next index

    Set rows = New Collection
    For position = 0 To count - 1
        rows.Add extended(position)
    Next position
    Set columns = widened
End Sub

' One window's value for every row, in the rows' own order.
Private Function WindowValues(ByVal window As Variant, ByRef rows() As Variant, _
                              ByVal count As Long, _
                              ByVal columns As Collection) As Variant
    Dim fn As Variant
    Dim name As String
    Dim spec As Variant
    Dim partitions As Collection
    Dim orders As Collection
    Dim results() As Variant
    Dim groups As Collection
    Dim groupKeys As Collection
    Dim seen As Collection
    Dim key As Variant
    Dim index As Long

    fn = window(1)
    name = CStr(fn(1))
    spec = window(2)
    Set partitions = spec(0)
    Set orders = spec(1)

    If fn(3) = True Then
        RaiseQueryError 10759, "Use of DISTINCT is not allowed with the " & _
                               "OVER clause."
    End If
    Select Case name
        Case "ROW_NUMBER", "RANK", "DENSE_RANK", "NTILE", "LAG", "LEAD", _
             "FIRST_VALUE", "LAST_VALUE", "PERCENT_RANK", "CUME_DIST"
            If orders.count = 0 Then
                RaiseQueryError 4112, "The function '" & name & "' must " & _
                                      "have an OVER clause with ORDER BY."
            End If
        Case "COUNT", "SUM", "AVG", "MIN", "MAX", _
             "STDEV", "STDEVP", "VAR", "VARP"
        Case Else
            RaiseQueryError ERR_NUMBER_BASE, "'" & name & "' is not a " & _
                                             "windowed function answered here."
    End Select

    ReDim results(0 To IIf(count > 0, count - 1, 0))

    ' The rows by partition, each partition in the order its first row came.
    ' Which keys have been seen is kept apart from the partitions: HasKey
    ' reads what it finds into a Variant, which a Collection cannot go into
    ' without Set, so asked about a partition it always answered no.
    Set groups = New Collection
    Set groupKeys = New Collection
    Set seen = New Collection
    For index = 0 To count - 1
        key = PartitionKey(partitions, rows(index), columns)
        If Not HasKey(seen, CStr(key)) Then
            seen.Add True, CStr(key)
            groups.Add New Collection, CStr(key)
            groupKeys.Add key
        End If
        groups(CStr(key)).Add index
    Next index

    For Each key In groupKeys
        WindowOverPartition fn, spec, groups(CStr(key)), rows, columns, results
    Next key
    WindowValues = results
End Function

Private Function PartitionKey(ByVal partitions As Collection, _
                              ByRef values As Variant, _
                              ByVal columns As Collection) As String
    Dim node As Variant
    Dim value As Variant

    PartitionKey = "k"
    For Each node In partitions
        value = Evaluate(node, values, columns)
        If IsNull(value) Then
            PartitionKey = PartitionKey & ChrW$(1) & "null"
        Else
            PartitionKey = PartitionKey & ChrW$(1) & "v" & ValueKey(value)
        End If
    Next node
End Function

' One window over one partition: its rows put in the window's order, and
' each given its value.
Private Sub WindowOverPartition(ByVal fn As Variant, ByVal spec As Variant, _
                                ByVal members As Collection, _
                                ByRef rows() As Variant, _
                                ByVal columns As Collection, _
                                ByRef results() As Variant)
    Dim name As String
    Dim arguments As Collection
    Dim orders As Collection
    Dim kind As String
    Dim startBound As Long
    Dim endBound As Long
    Dim m As Long
    Dim sorted() As Long
    Dim positions() As Long
    Dim scratch() As Long
    Dim keys As Variant
    Dim descending() As Boolean
    Dim term As Variant
    Dim member As Variant
    Dim firstPeer() As Long
    Dim lastPeer() As Long
    Dim t As Long
    Dim p As Long
    Dim q As Long
    Dim s As Long
    Dim e As Long
    Dim value As Variant
    Dim inner As Variant
    Dim dense As Long
    Dim buckets As Long
    Dim bucket As Long
    Dim inBucket As Long
    Dim offset As Long
    Dim target As Long
    Dim growing As Boolean
    Dim upTo As Long
    Dim tally As Long
    Dim total As Double
    Dim best As Variant
    Dim frame As Collection
    Dim truncates As Boolean

    name = CStr(fn(1))
    Set arguments = fn(2)
    Set orders = spec(1)
    kind = CStr(spec(2))
    startBound = CLng(spec(3))
    endBound = CLng(spec(4))

    m = members.count
    ReDim sorted(0 To m - 1)
    ReDim positions(0 To m - 1)
    For Each member In members
        sorted(p) = member
        positions(p) = p
        p = p + 1
    Next member

    ' In the window's order, stably, so rows the order cannot tell apart
    ' keep the order they came in.
    If orders.count > 0 Then
        ReDim keys(0 To m - 1, 0 To orders.count - 1)
        ReDim descending(0 To orders.count - 1)
        t = 0
        For Each term In orders
            descending(t) = CBool(term(1))
            For p = 0 To m - 1
                keys(p, t) = NormalizeValue(Evaluate(term(0), rows(sorted(p)), _
                                                     columns))
            Next p
            t = t + 1
        Next term
        ReDim scratch(0 To m - 1)
        MergePositions positions, scratch, 0, m - 1, keys, descending
    End If

    ' Which rows the order cannot tell apart: a run of them shares a rank,
    ' and an ordered window with no frame runs to the last of the run.
    ReDim firstPeer(0 To m - 1)
    ReDim lastPeer(0 To m - 1)
    For p = 1 To m - 1
        If orders.count = 0 Then
            firstPeer(p) = 0
        ElseIf ComparePositions(positions(p), positions(p - 1), keys, _
                                descending) = 0 Then
            firstPeer(p) = firstPeer(p - 1)
        Else
            firstPeer(p) = p
        End If
    Next p
    lastPeer(m - 1) = m - 1
    For p = m - 2 To 0 Step -1
        If firstPeer(p + 1) = firstPeer(p) Then
            lastPeer(p) = lastPeer(p + 1)
        Else
            lastPeer(p) = p
        End If
    Next p

    If arguments.count > 0 Then inner = arguments(1)
    truncates = DeclaredWhole(inner, columns)
    growing = (Len(kind) = 0 Or startBound = WINDOW_UNBOUNDED_PRECEDING)
    ' A spread keeps no running figure here, so each frame is taken whole.
    Select Case name
        Case "STDEV", "STDEVP", "VAR", "VARP"
            growing = False
    End Select
    upTo = -1
    best = Null

    For p = 0 To m - 1
        FrameOf p, m, kind, startBound, endBound, orders.count > 0, _
                firstPeer, lastPeer, s, e
        Select Case name
            Case "ROW_NUMBER"
                value = CLngLng(p + 1)
            Case "RANK"
                value = CLngLng(firstPeer(p) + 1)
            Case "DENSE_RANK"
                If p = 0 Then
                    dense = 1
                ElseIf firstPeer(p) = p Then
                    dense = dense + 1
                End If
                value = CLngLng(dense)
            Case "NTILE"
                ' As many rows to a group as divide evenly, and the first
                ' groups one more each until the remainder is used up.
                If p = 0 Then
                    buckets = CLng(NullToZero(Evaluate(inner, _
                        rows(sorted(positions(0))), columns)))
                    If buckets < 1 Then
                        RaiseQueryError ERR_NUMBER_BASE, "The function " & _
                            "'NTILE' takes only a positive int or bigint " & _
                            "expression as its input."
                    End If
                    bucket = 1
                End If
                inBucket = inBucket + 1
                value = CLngLng(bucket)
                If inBucket >= (m \ buckets) + _
                               IIf(bucket <= (m Mod buckets), 1, 0) Then
                    bucket = bucket + 1
                    inBucket = 0
                End If
            Case "PERCENT_RANK"
                If m > 1 Then
                    value = CDbl(firstPeer(p)) / (m - 1)
                Else
                    value = 0#
                End If
            Case "CUME_DIST"
                value = CDbl(lastPeer(p) + 1) / m
            Case "LAG", "LEAD"
                offset = 1
                If arguments.count >= 2 Then
                    offset = CLng(NullToZero(Evaluate(arguments(2), _
                        rows(sorted(positions(p))), columns)))
                End If
                If name = "LAG" Then target = p - offset Else target = p + offset
                If target >= 0 And target < m Then
                    value = Evaluate(inner, rows(sorted(positions(target))), _
                                     columns)
                ElseIf arguments.count >= 3 Then
                    value = Evaluate(arguments(3), rows(sorted(positions(p))), _
                                     columns)
                Else
                    value = Null
                End If
            Case "FIRST_VALUE"
                value = Null
                If s <= e Then
                    value = Evaluate(inner, rows(sorted(positions(s))), columns)
                End If
            Case "LAST_VALUE"
                value = Null
                If s <= e Then
                    value = Evaluate(inner, rows(sorted(positions(e))), columns)
                End If
            Case Else
                ' An aggregate over the frame. A frame that starts at the
                ' first row only grows, so it is added to rather than summed
                ' again from the start for every row.
                If growing Then
                    Do While upTo < e
                        upTo = upTo + 1
                        Accumulate name, inner, rows(sorted(positions(upTo))), _
                                   columns, tally, total, best
                    Loop
                    value = Accumulated(name, tally, total, best, truncates)
                Else
                    Set frame = New Collection
                    For q = s To e
                        frame.Add rows(sorted(positions(q)))
                    Next q
                    value = ComputeAggregate(name, inner, frame, columns)
                End If
        End Select
        results(sorted(positions(p))) = value
    Next p
End Sub

' Which rows of a partition, by position, the frame around position p holds.
' With no frame written, an ordered window runs from the first row to the
' last the order cannot tell from this one, and an unordered one is the
' whole partition.
Private Sub FrameOf(ByVal p As Long, ByVal m As Long, ByVal kind As String, _
                    ByVal startBound As Long, ByVal endBound As Long, _
                    ByVal ordered As Boolean, ByRef firstPeer() As Long, _
                    ByRef lastPeer() As Long, ByRef s As Long, ByRef e As Long)
    If Len(kind) = 0 Then
        s = 0
        If ordered Then e = lastPeer(p) Else e = m - 1
        Exit Sub
    End If
    If kind = "RANGE" Then
        If startBound = WINDOW_UNBOUNDED_PRECEDING Then s = 0 Else s = firstPeer(p)
        If endBound = WINDOW_UNBOUNDED_FOLLOWING Then e = m - 1 Else e = lastPeer(p)
        Exit Sub
    End If
    s = FrameEdge(p, startBound, m)
    e = FrameEdge(p, endBound, m)
    If s < 0 Then s = 0
    If e > m - 1 Then e = m - 1
End Sub

Private Function FrameEdge(ByVal p As Long, ByVal bound As Long, _
                           ByVal m As Long) As Long
    If bound = WINDOW_UNBOUNDED_PRECEDING Then
        FrameEdge = 0
    ElseIf bound = WINDOW_UNBOUNDED_FOLLOWING Then
        FrameEdge = m - 1
    Else
        FrameEdge = p + bound
    End If
End Function

' A stable merge sort of positions by the keys each position has.
Private Sub MergePositions(ByRef positions() As Long, ByRef scratch() As Long, _
                           ByVal low As Long, ByVal high As Long, _
                           ByRef keys As Variant, ByRef descending() As Boolean)
    Dim middle As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    If low >= high Then Exit Sub
    middle = (low + high) \ 2
    MergePositions positions, scratch, low, middle, keys, descending
    MergePositions positions, scratch, middle + 1, high, keys, descending

    i = low
    j = middle + 1
    k = low
    Do While i <= middle And j <= high
        If ComparePositions(positions(j), positions(i), keys, descending) < 0 Then
            scratch(k) = positions(j)
            j = j + 1
        Else
            scratch(k) = positions(i)
            i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= middle
        scratch(k) = positions(i)
        i = i + 1
        k = k + 1
    Loop
    Do While j <= high
        scratch(k) = positions(j)
        j = j + 1
        k = k + 1
    Loop
    For k = low To high
        positions(k) = scratch(k)
    Next k
End Sub

Private Function ComparePositions(ByVal a As Long, ByVal b As Long, _
                                  ByRef keys As Variant, _
                                  ByRef descending() As Boolean) As Long
    Dim t As Long
    Dim outcome As Long

    For t = LBound(descending) To UBound(descending)
        outcome = CompareValues(keys(a, t), keys(b, t))
        If descending(t) Then outcome = -outcome
        If outcome <> 0 Then
            ComparePositions = outcome
            Exit Function
        End If
    Next t
End Function

' One more row into an aggregate that runs down a window, taken the way
' ComputeAggregate takes each of its rows.
Private Sub Accumulate(ByVal name As String, ByVal inner As Variant, _
                       ByRef values As Variant, ByVal columns As Collection, _
                       ByRef tally As Long, ByRef total As Double, _
                       ByRef best As Variant)
    Dim current As Variant
    Dim isStar As Boolean

    If IsEmpty(inner) Then
        isStar = True
    ElseIf IsArray(inner) Then
        isStar = (inner(0) = N_STAR)
    End If
    If isStar Then
        tally = tally + 1
        Exit Sub
    End If

    current = Evaluate(inner, values, columns)
    If IsNull(current) Then Exit Sub
    tally = tally + 1
    If IsNumeric(current) Then total = total + CDbl(current)
    If IsNull(best) Then
        best = current
    ElseIf name = "MIN" Then
        If CompareValues(current, best) < 0 Then best = current
    ElseIf name = "MAX" Then
        If CompareValues(current, best) > 0 Then best = current
    End If
End Sub

Private Function Accumulated(ByVal name As String, ByVal tally As Long, _
                             ByVal total As Double, ByVal best As Variant, _
                             ByVal truncates As Boolean) As Variant
    Select Case name
        Case "COUNT"
            Accumulated = tally
        Case "SUM"
            If tally = 0 Then Accumulated = Null Else Accumulated = total
        Case "AVG"
            If tally = 0 Then
                Accumulated = Null
            ElseIf truncates Then
                Accumulated = Fix(total / tally)
            Else
                Accumulated = total / tally
            End If
        Case Else
            Accumulated = best
    End Select
End Function

' The column a window's values ride on. The ranking functions are bigint on
' a real server and COUNT an int; an aggregate or a value taken from another
' row is whatever its argument is.
Private Function WindowColumn(ByVal window As Variant, ByVal index As Long, _
                              ByRef results As Variant, ByVal count As Long, _
                              ByVal columns As Collection) As Variant
    Dim fn As Variant
    Dim name As String
    Dim stated As Variant
    Dim position As Long

    fn = window(1)
    name = "$window" & index
    Select Case CStr(fn(1))
        Case "ROW_NUMBER", "RANK", "DENSE_RANK", "NTILE"
            WindowColumn = ColumnOf(name, TYPE_INTN, 8)
            Exit Function
        Case "COUNT"
            WindowColumn = ColumnOf(name, TYPE_INTN, 4)
            Exit Function
        Case "PERCENT_RANK", "CUME_DIST", "STDEV", "STDEVP", "VAR", "VARP"
            WindowColumn = ColumnOf(name, TYPE_FLTN, 8)
            Exit Function
        Case "SUM", "MIN", "MAX", "AVG"
            stated = StatedType(fn, columns)
        Case Else
            If fn(2).count > 0 Then stated = StatedType(fn(2)(1), columns)
    End Select

    If IsArray(stated) Then
        WindowColumn = ColumnOf(name, CLng(stated(0)), CLng(stated(1)), -1, _
                                vbNullString, CLng(NullToZero(stated(2))))
        Exit Function
    End If
    For position = 0 To count - 1
        If Not IsNull(results(position)) Then
            WindowColumn = ColumnOf(name, TypeOfValue(results(position)), _
                                    SizeOfValue(results(position)))
            Exit Function
        End If
    Next position
    WindowColumn = ColumnOf(name, TYPE_INTN, 4)
End Function


' Whether a star takes this column: a bare one takes every column, and one
' written with a source in front of it takes that source's.
Private Function StarWants(ByVal node As Variant, ByVal column As Variant) As Boolean
    ' A window's value rides on the row as a column of its own, which no
    ' star asked for.
    If Left$(CStr(column(COLUMN_NAME)), 7) = "$window" Then Exit Function
    If Len(CStr(node(1))) = 0 Then
        StarWants = True
    Else
        StarWants = (StrComp(CStr(column(COLUMN_QUALIFIER)), CStr(node(1)), _
                             vbTextCompare) = 0)
    End If
End Function


' Every token up to the bracket that closes the one already consumed.
Private Function BracketedTokens() As Collection
    Dim taken As Collection
    Dim depth As Long
    Dim text As String

    Set taken = New Collection
    depth = 1
    Do While Not AtEnd()
        text = PeekText()
        If PeekKind() = T_OPERATOR Then
            If text = "(" Then
                depth = depth + 1
            ElseIf text = ")" Then
                depth = depth - 1
                If depth = 0 Then
                    mCursor = mCursor + 1
                    Set BracketedTokens = taken
                    Exit Function
                End If
            End If
        End If
        taken.Add mTokens(mCursor)
        mCursor = mCursor + 1
    Loop
    Set BracketedTokens = taken
End Function

' A read run inside another read.
'
' The parser keeps one statement's worth of state, so running a second one
' inside the first means putting that state aside and giving it back. The
' row the outer read is on goes on a stack while the inner one runs, so a
' condition inside it may name a column of the outer source.
Private Sub RunSubquery(ByVal tokens As Collection, ByRef values As Variant, _
                        ByVal columns As Collection, _
                        ByRef outColumns As Collection, _
                        ByRef outRows As Collection)
    Dim kept As Variant
    Dim key As String
    Dim answered As Variant
    Dim before As Long

    ' A read inside another read is run once per row of the outer one. Most
    ' of them do not look at that row at all, and those have one answer:
    ' running them again for every row is a pass over the inner table per
    ' outer row, which is the whole cost of the query.
    key = CStr(ObjPtr(tokens))
    If Not mSubqueryAnswers Is Nothing Then
        On Error Resume Next
        answered = mSubqueryAnswers(key)
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            Set outColumns = answered(0)
            Set outRows = answered(1)
            Exit Sub
        End If
        Err.Clear
        On Error GoTo 0
    End If

    kept = SaveParseState()
    PushOuterRow values, columns
    before = mOuterReads
    On Error GoTo Restore

    Set mTokens = tokens
    mCursor = 1
    ParseSelect

    ExecuteWholeSelect outColumns, outRows

    PopOuterRow
    RestoreParseState kept
    On Error GoTo 0

    ' It looked at no outer row, so what it answered is what it will answer.
    ' Reading the outer row is the only way a read can be correlated, so
    ' counting those reads is how the two are told apart: no guessing at the
    ' shape of the statement, and no way for a correlated one to slip
    ' through as though it were not.
    If mOuterReads = before Then RememberSubquery key, outColumns, outRows
    Exit Sub

Restore:
    PopOuterRow
    RestoreParseState kept
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' A system rowset as it was last built, if it still stands.
Private Function BuiltRowset(ByVal name As String, _
                             ByRef outColumns As Collection, _
                             ByRef outRows As Collection) As Boolean
    Dim held As Variant

    If mSystemRowsets Is Nothing Then Exit Function
    On Error Resume Next
    held = mSystemRowsets(LCase$(name))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    Set outColumns = held(0)
    Set outRows = held(1)
    BuiltRowset = True
End Function

Private Sub RememberRowset(ByVal name As String, _
                           ByVal columns As Collection, _
                           ByVal rows As Collection)
    Dim held(0 To 1) As Variant

    If mSystemRowsets Is Nothing Then Set mSystemRowsets = New Collection
    Set held(0) = columns
    Set held(1) = rows
    On Error Resume Next
    mSystemRowsets.Add held, LCase$(name)
    Err.Clear
    On Error GoTo 0
End Sub

' Drop them, because what they describe has changed.
Private Sub ForgetSystemRowsets()
    Set mSystemRowsets = Nothing
End Sub

Private Sub RememberSource(ByVal key As String, _
                           ByVal columns As Collection, _
                           ByVal rows As Collection)
    Dim loaded(0 To 1) As Variant

    If mLoadedSources Is Nothing Then Set mLoadedSources = New Collection
    Set loaded(0) = columns
    Set loaded(1) = rows
    On Error Resume Next
    mLoadedSources.Add loaded, key
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub RememberSubquery(ByVal key As String, _
                             ByVal outColumns As Collection, _
                             ByVal outRows As Collection)
    Dim answered(0 To 1) As Variant

    If mSubqueryAnswers Is Nothing Then Set mSubqueryAnswers = New Collection
    Set answered(0) = outColumns
    Set answered(1) = outRows
    On Error Resume Next
    mSubqueryAnswers.Add answered, key
    Err.Clear
    On Error GoTo 0
End Sub

' What a read used as a value answers: the first value of the first row, or
' nothing at all.
Private Function ScalarSubquery(ByVal tokens As Collection, _
                                ByRef values As Variant, _
                                ByVal columns As Collection) As Variant
    Dim outColumns As Collection
    Dim outRows As Collection
    Dim first As Variant

    RunSubquery tokens, values, columns, outColumns, outRows
    If outRows.count = 0 Then
        ScalarSubquery = Null
        Exit Function
    End If
    first = outRows(1)
    ScalarSubquery = first(LBound(first))
End Function

' Whether a read found anything.
'
' One row is enough, and a select list nobody reads is a shape nobody needs,
' so the rows are walked until one passes the condition and neither the rest
' of them nor any of the columns are ever worked out. A statement that
' decides which rows there are after the condition has run is answered the
' long way, because stopping early would stop before the deciding.
Private Function ExistsSubquery(ByVal tokens As Collection, _
                                ByRef values As Variant, _
                                ByVal columns As Collection) As Boolean
    Dim kept As Variant
    Dim outColumns As Collection
    Dim outRows As Collection
    Dim answered As Boolean
    Dim key As String

    ' A read already turned into a lookup answers without being run.
    key = CStr(ObjPtr(tokens))
    If ExistsLookupReady(key) Then
        ExistsSubquery = ExistsLookupSays(key, values, columns)
        Exit Function
    End If

    kept = SaveParseState()
    PushOuterRow values, columns
    On Error GoTo Restore

    Set mTokens = tokens
    mCursor = 1
    ParseSelect

    If mGroupBy.count > 0 Or mTopCount >= 0 Or Not IsEmpty(mHavingNode) _
       Or FollowsSetOperator() Then
        ' Something here decides which rows there are after the condition
        ' has run, so the condition cannot be taken out of it. A set
        ' operator is one of those things, and it is also the rest of the
        ' read: stopping here would ask only its first branch.
        ExecuteWholeSelect outColumns, outRows
        answered = (outRows.count > 0)
    ElseIf UsesExistsLookup(key, columns) Then
        answered = ExistsLookupSays(key, values, columns)
    Else
        answered = AnyRowMatches()
    End If

    PopOuterRow
    RestoreParseState kept
    On Error GoTo 0
    ExistsSubquery = answered
    Exit Function

Restore:
    PopOuterRow
    RestoreParseState kept
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Turning a correlated EXISTS into a question about a set.
'
' WHERE EXISTS (SELECT 1 FROM o WHERE o.owner = p.id) asks the same thing of
' every outer row: is this value among the owners. Answering it a row at a
' time is a pass over the inner table per outer row, which is quadratic and
' was measured at five seconds over five hundred rows against five hundred.
' The owners are read once instead and each outer value looked up.
'
' The shortcut only fits some shapes. What has to hold: one outer column,
' matched by an equality that is ANDed with the rest of the condition, and
' nothing else in the subquery looking at the outer row. A GROUP BY, a
' HAVING or a TOP decides which rows there are after the condition, and
' those are declined further up. An OR, a comparison that is not equality,
' two outer columns matched at once, and an outer expression rather than an
' outer column all fall back to a row at a time.
'
' tests/test_shortcut.py runs every shape twice, once with this turned off,
' and the two have to agree; tests/test_differential.py puts the same shapes
' to a real SQL Server. A shortcut that never fired would pass a test that
' only checked the answers it did give, so which shapes take it is checked
' rather than assumed.
Private Function UsesExistsLookup(ByVal key As String, _
                                  ByVal outerColumns As Collection) As Boolean
    If mNoExistsShortcut Then Exit Function
    If ExistsLookupReady(key) Then
        UsesExistsLookup = True
        Exit Function
    End If
    UsesExistsLookup = BuildExistsLookup(key, outerColumns)
End Function

Private Function ExistsLookupReady(ByVal key As String) As Boolean
    Dim entry As Variant

    If mExistsLookups Is Nothing Then Exit Function
    On Error Resume Next
    entry = mExistsLookups(key)
    ExistsLookupReady = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function ExistsLookupSays(ByVal key As String, ByRef values As Variant, _
                                  ByVal columns As Collection) As Boolean
    Dim entry As Variant
    Dim probe As Variant

    entry = mExistsLookups(key)
    probe = Evaluate(entry(1), values, columns)

    ' NULL matches nothing, because NULL = anything is unknown and unknown
    ' is not a row that exists.
    If IsNull(probe) Or IsEmpty(probe) Then Exit Function
    ExistsLookupSays = HasValueKey(entry(0), probe)
End Function

' Read the inner column once, if the shape allows it. False means it does
' not, and the caller falls back to a row at a time.
Private Function BuildExistsLookup(ByVal key As String, _
                                   ByVal outerColumns As Collection) As Boolean
    Dim innerColumns As Collection
    Dim innerRows As Collection
    Dim parts As Collection
    Dim rest As Collection
    Dim part As Variant
    Dim outerNode As Variant
    Dim innerNode As Variant
    Dim matched As Boolean
    Dim values As Variant
    Dim value As Variant
    Dim lookup As Collection
    Dim keep As Variant
    Dim passes As Boolean
    Dim entry(0 To 1) As Variant

    If IsEmpty(mWhereNode) Then Exit Function

    CombinedSource innerColumns, innerRows

    Set parts = New Collection
    Conjuncts mWhereNode, parts

    Set rest = New Collection
    For Each part In parts
        If Not matched Then
            If IsCorrelatedEquality(part, innerColumns, outerColumns, _
                                    outerNode, innerNode) Then
                matched = True
                GoTo NextPart
            End If
        End If
        If ReadsOuter(part, innerColumns) Then Exit Function
        rest.Add part
NextPart:
    Next part
    If Not matched Then Exit Function

    Set lookup = New Collection
    For Each values In innerRows
        passes = True
        For Each part In rest
            keep = Evaluate(part, values, innerColumns)
            If IsNull(keep) Then
                passes = False
            ElseIf Not CBool(keep) Then
                passes = False
            End If
            If Not passes Then Exit For
        Next part

        If passes Then
            value = Evaluate(innerNode, values, innerColumns)

            ' A Boolean compares by converting whatever is beside it, which
            ' is a conversion that can fail rather than a key that can be
            ' written. A column holding one is answered a row at a time.
            If VarType(value) = vbBoolean Then Exit Function

            If Not (IsNull(value) Or IsEmpty(value)) Then
                On Error Resume Next
                lookup.Add True, ValueKey(value)
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next values

    If mExistsLookups Is Nothing Then Set mExistsLookups = New Collection
    Set entry(0) = lookup
    entry(1) = outerNode
    On Error Resume Next
    mExistsLookups.Add entry, key
    Err.Clear
    On Error GoTo 0
    mExistsLookupsBuilt = mExistsLookupsBuilt + 1
    BuildExistsLookup = True
End Function

' The conditions an AND joins, as a list. Anything else is one condition.
Private Sub Conjuncts(ByVal node As Variant, ByVal into As Collection)
    If Not IsArray(node) Then Exit Sub

    If node(0) = N_BINARY Then
        If CStr(node(1)) = "AND" Then
            Conjuncts node(2), into
            Conjuncts node(3), into
            Exit Sub
        End If
    End If
    into.Add node
End Sub

' Whether this condition is the one that correlates the read: an equality
' with a bare outer column on one side and nothing outer on the other.
Private Function IsCorrelatedEquality(ByVal node As Variant, _
                                      ByVal innerColumns As Collection, _
                                      ByVal outerColumns As Collection, _
                                      ByRef outerNode As Variant, _
                                      ByRef innerNode As Variant) As Boolean
    If Not IsArray(node) Then Exit Function
    If node(0) <> N_BINARY Then Exit Function
    If CStr(node(1)) <> "=" Then Exit Function

    If IsOuterColumn(node(2), innerColumns, outerColumns) Then
        If ReadsOuter(node(3), innerColumns) Then Exit Function
        outerNode = node(2)
        innerNode = node(3)
        IsCorrelatedEquality = True
        Exit Function
    End If

    If IsOuterColumn(node(3), innerColumns, outerColumns) Then
        If ReadsOuter(node(2), innerColumns) Then Exit Function
        outerNode = node(3)
        innerNode = node(2)
        IsCorrelatedEquality = True
    End If
End Function

' A bare column the inner source has not got and the outer one has.
Private Function IsOuterColumn(ByVal node As Variant, _
                               ByVal innerColumns As Collection, _
                               ByVal outerColumns As Collection) As Boolean
    If Not IsArray(node) Then Exit Function
    If node(0) <> N_COLUMN Then Exit Function
    If Left$(CStr(node(1)), 1) = "@" Then Exit Function
    If ColumnIndex(CStr(node(1)), innerColumns) > 0 Then Exit Function
    IsOuterColumn = (ColumnIndex(CStr(node(1)), outerColumns) > 0)
End Function

' Whether anything in this expression names a column the inner source has
' not got, which is the only way it can be reading the outer row.
Private Function ReadsOuter(ByVal node As Variant, _
                            ByVal innerColumns As Collection) As Boolean
    Dim item As Variant

    If Not IsArray(node) Then Exit Function

    Select Case node(0)
        Case N_LITERAL, N_STAR

        Case N_COLUMN
            If Left$(CStr(node(1)), 1) = "@" Then Exit Function
            ReadsOuter = (ColumnIndex(CStr(node(1)), innerColumns) = 0)

        Case N_SUBQUERY, N_EXISTS
            ' A read inside this one may name anything at all. Not taken
            ' apart, and so not allowed to hide an outer name.
            ReadsOuter = True

        Case N_UNARY
            ReadsOuter = ReadsOuter(node(2), innerColumns)

        Case N_BINARY
            Select Case CStr(node(1))
                Case "IN READ", "NOT IN READ"
                    ' The right side is a read's tokens rather than an
                    ' expression, and tokens are arrays too: walking them
                    ' as though they were nodes would read them wrong.
                    ReadsOuter = True
                Case "IN", "NOT IN"
                    If ReadsOuter(node(2), innerColumns) Then
                        ReadsOuter = True
                        Exit Function
                    End If
                    For Each item In node(3)
                        If ReadsOuter(item, innerColumns) Then
                            ReadsOuter = True
                            Exit Function
                        End If
                    Next item
                Case Else
                    If ReadsOuter(node(2), innerColumns) Then
                        ReadsOuter = True
                    Else
                        ReadsOuter = ReadsOuter(node(3), innerColumns)
                    End If
            End Select

        Case N_CALL
            For Each item In node(2)
                If ReadsOuter(item, innerColumns) Then
                    ReadsOuter = True
                    Exit Function
                End If
            Next item
    End Select
End Function

' The rows on the right, grouped by what an equality in the condition
' matches them on, or nothing when there is no such equality.
'
' A join tried every pair: five hundred rows against five hundred is a
' quarter of a million rows built and a quarter of a million conditions
' evaluated, and it was measured at three and a half seconds. An equality
' that is ANDed into the condition has to hold for a pair to be kept, so the
' pairs it rules out need never be built.
'
' It rules nothing else out. The whole condition is still evaluated on every
' pair the key kept, so a second condition beside the equality decides what
' it always decided.
Private Function JoinBuckets(ByVal onNode As Variant, _
                             ByVal leftColumns As Collection, _
                             ByVal rightColumns As Collection, _
                             ByVal rightRows As Collection, _
                             ByRef leftKey As Variant) As Collection
    Dim parts As Collection
    Dim part As Variant
    Dim rightKey As Variant
    Dim found As Boolean
    Dim buckets As Collection
    Dim values As Variant
    Dim value As Variant
    Dim key As String
    Dim bucket As Collection

    If mNoJoinShortcut Then Exit Function
    If IsEmpty(onNode) Then Exit Function
    If rightRows.count < 2 Then Exit Function

    ' An OR is not a condition every pair has to meet, so an equality inside
    ' one rules nothing out.
    Set parts = New Collection
    Conjuncts onNode, parts
    For Each part In parts
        If Not found Then
            If SidesOfEquality(part, leftColumns, rightColumns, _
                               leftKey, rightKey) Then
                found = True
            End If
        End If
    Next part
    If Not found Then Exit Function

    ' The key is worked out for every row on either side, so it is resolved
    ' against the columns once rather than once a row.
    rightKey = Prepared(rightKey, rightColumns)

    Set buckets = New Collection
    For Each values In rightRows
        value = Evaluate(rightKey, values, rightColumns)

        ' A Boolean has no key that can be written, and NULL matches
        ' nothing an equality would have kept.
        If VarType(value) = vbBoolean Then Exit Function
        If Not (IsNull(value) Or IsEmpty(value)) Then
            key = ValueKey(value)
            Set bucket = Nothing
            On Error Resume Next
            Set bucket = buckets(key)
            Err.Clear
            On Error GoTo 0
            If bucket Is Nothing Then
                Set bucket = New Collection
                buckets.Add bucket, key
            End If
            bucket.Add values
        End If
    Next values

    mJoinLookupsBuilt = mJoinLookupsBuilt + 1
    Set JoinBuckets = buckets
End Function

' The rows on the right that could match this one on the left.
Private Function JoinCandidates(ByVal buckets As Collection, _
                                ByVal leftKey As Variant, _
                                ByRef leftValues As Variant, _
                                ByVal leftColumns As Collection) As Collection
    Dim value As Variant
    Dim bucket As Collection

    value = Evaluate(leftKey, leftValues, leftColumns)
    If IsNull(value) Or IsEmpty(value) Then
        Set JoinCandidates = New Collection
        Exit Function
    End If
    If VarType(value) = vbBoolean Then
        Set JoinCandidates = New Collection
        Exit Function
    End If

    On Error Resume Next
    Set bucket = buckets(ValueKey(value))
    Err.Clear
    On Error GoTo 0
    If bucket Is Nothing Then Set bucket = New Collection
    Set JoinCandidates = bucket
End Function

' Whether this condition is an equality with one side wholly on the left and
' the other wholly on the right.
Private Function SidesOfEquality(ByVal node As Variant, _
                                 ByVal leftColumns As Collection, _
                                 ByVal rightColumns As Collection, _
                                 ByRef leftSide As Variant, _
                                 ByRef rightSide As Variant) As Boolean
    If Not IsArray(node) Then Exit Function
    If node(0) <> N_BINARY Then Exit Function
    If CStr(node(1)) <> "=" Then Exit Function

    If OnlyIn(node(2), leftColumns) And OnlyIn(node(3), rightColumns) Then
        leftSide = node(2)
        rightSide = node(3)
        SidesOfEquality = True
        Exit Function
    End If

    If OnlyIn(node(3), leftColumns) And OnlyIn(node(2), rightColumns) Then
        leftSide = node(3)
        rightSide = node(2)
        SidesOfEquality = True
    End If
End Function

' Whether every column this names is one of those, and it names at least
' one: a constant is true of both sides and keys nothing.
Private Function OnlyIn(ByVal node As Variant, _
                        ByVal columns As Collection) As Boolean
    If ReadsOuter(node, columns) Then Exit Function
    OnlyIn = NamesAColumn(node)
End Function

Private Function NamesAColumn(ByVal node As Variant) As Boolean
    Dim item As Variant

    If Not IsArray(node) Then Exit Function

    Select Case node(0)
        Case N_COLUMN
            NamesAColumn = (Left$(CStr(node(1)), 1) <> "@")
        Case N_UNARY
            NamesAColumn = NamesAColumn(node(2))
        Case N_BINARY
            If NamesAColumn(node(2)) Then
                NamesAColumn = True
            ElseIf Not IsObject(node(3)) Then
                NamesAColumn = NamesAColumn(node(3))
            End If
        Case N_CALL
            For Each item In node(2)
                If NamesAColumn(item) Then
                    NamesAColumn = True
                    Exit Function
                End If
            Next item
    End Select
End Function

' Whether any row of the source passes the condition, stopping at the first
' that does.
Private Function AnyRowMatches() As Boolean
    Dim columns As Collection
    Dim rows As Collection
    Dim values As Variant
    Dim keep As Variant

    CombinedSource columns, rows
    If IsEmpty(mWhereNode) Then
        AnyRowMatches = (rows.count > 0)
        Exit Function
    End If

    For Each values In rows
        keep = Evaluate(mWhereNode, values, columns)
        If Not IsNull(keep) Then
            If CBool(keep) Then
                AnyRowMatches = True
                Exit Function
            End If
        End If
    Next values
End Function

' The candidates of an IN, as something to look a value up in, or nothing
' when the walk has to be used instead.
'
' The keys follow the comparison this server does and no other: a number
' beside text converts the text, so 2 and '2' are one candidate, and text
' that is not a number matches without regard to case. A Boolean compares
' by converting whatever is beside it, which is a conversion that can fail
' rather than a key that can be written, so a list holding one declines the
' lookup and is walked.
Private Function SubqueryLookup(ByVal tokens As Collection, _
                                ByRef values As Variant, _
                                ByVal columns As Collection) As Collection
    Dim key As String
    Dim built As Collection
    Dim outColumns As Collection
    Dim outRows As Collection
    Dim row As Variant
    Dim value As Variant

    key = CStr(ObjPtr(tokens))
    If Not mSubqueryLookups Is Nothing Then
        On Error Resume Next
        Set built = mSubqueryLookups(key)
        Err.Clear
        On Error GoTo 0
        If Not built Is Nothing Then
            If built.count = 0 Then
                ' Nothing was worth keying, which is how a declined list is
                ' remembered: an empty one would be a wrong answer.
                Set SubqueryLookup = Nothing
            Else
                Set SubqueryLookup = built
            End If
            Exit Function
        End If
    End If

    RunSubquery tokens, values, columns, outColumns, outRows

    ' Only worth building for a read that answers the same for every row.
    ' A correlated one has different candidates each time, and keying them
    ' would be the pass this exists to avoid.
    If Not KnownSubquery(key) Then Exit Function

    Set built = New Collection
    For Each row In outRows
        value = row(LBound(row))
        If VarType(value) = vbBoolean Then
            Set built = New Collection
            Exit For
        End If
        On Error Resume Next
        If IsNull(value) Or IsEmpty(value) Then
            ' Kept rather than skipped: whether one of the candidates was
            ' NULL is the difference between no match and no answer.
            built.Add True, NULL_CANDIDATE
        Else
            built.Add True, ValueKey(value)
        End If
        Err.Clear
        On Error GoTo 0
    Next row

    If mSubqueryLookups Is Nothing Then Set mSubqueryLookups = New Collection
    On Error Resume Next
    mSubqueryLookups.Add built, key
    Err.Clear
    On Error GoTo 0

    If built.count > 0 Then Set SubqueryLookup = built
End Function

Private Function KnownSubquery(ByVal key As String) As Boolean
    Dim answered As Variant

    If mSubqueryAnswers Is Nothing Then Exit Function
    On Error Resume Next
    answered = mSubqueryAnswers(key)
    KnownSubquery = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' One value as the key it compares equal under.
Private Function ValueKey(ByVal value As Variant) As String
    ' A number first, because a join key usually is one and this runs for
    ' every row on both sides of the join. Text that reads as a number still
    ' goes the long way round, so '5' and 5 keep the one key: that is what
    ' makes a lookup find a number written as text.
    If IsNumberKind(VarType(value)) Then
        ValueKey = "#" & CStr(CDbl(value))
        Exit Function
    End If

    If IsNumeric(value) Then
        ValueKey = "#" & CStr(CDbl(value))
    Else
        ' Keyed the way text compares, trailing spaces and all: 'a ' and 'a'
        ' are one value, so a lookup has to find one under the other.
        ValueKey = "$" & LCase$(RTrim$(CStr(value)))
    End If
End Function

Private Function HasValueKey(ByVal lookup As Collection, _
                             ByVal value As Variant) As Boolean
    Dim found As Variant

    If VarType(value) = vbBoolean Then
        HasValueKey = HasKey(lookup, "#" & CStr(IIf(CBool(value), 1, 0)))
        Exit Function
    End If
    HasValueKey = HasKey(lookup, ValueKey(value))
End Function

Private Function HasKey(ByVal lookup As Collection, ByVal key As String) As Boolean
    Dim found As Variant

    On Error Resume Next
    found = lookup(key)
    HasKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' The first column of every row a read produced, for the right-hand side of
' IN.
Private Function SubqueryColumn(ByVal tokens As Collection, _
                                ByRef values As Variant, _
                                ByVal columns As Collection) As Collection
    Dim outColumns As Collection
    Dim outRows As Collection
    Dim row As Variant
    Dim taken As Collection

    RunSubquery tokens, values, columns, outColumns, outRows
    Set taken = New Collection
    For Each row In outRows
        taken.Add row(LBound(row))
    Next row
    Set SubqueryColumn = taken
End Function

' The outer rows a subquery can see, innermost last.
Private Sub PushOuterRow(ByRef values As Variant, ByVal columns As Collection)
    If mOuterValues Is Nothing Then
        Set mOuterValues = New Collection
        Set mOuterColumns = New Collection
        Set mOuterCaches = New Collection
    End If
    mOuterValues.Add values
    mOuterColumns.Add columns
    mOuterCaches.Add New Collection
End Sub

Private Sub PopOuterRow()
    If mOuterValues Is Nothing Then Exit Sub
    If mOuterValues.count = 0 Then Exit Sub
    mOuterValues.Remove mOuterValues.count
    mOuterColumns.Remove mOuterColumns.count
    mOuterCaches.Remove mOuterCaches.count
End Sub

' A name the running read has no column for may belong to one of the reads
' it is running inside. The innermost is asked first, as scoping requires.
'
' The answer is kept for as long as that row is the outer row, because it
' cannot change while it is: the name is resolved by walking the outer
' columns and comparing strings, and the outer read of a catalog query is a
' join of a dozen tables, so a subquery over three hundred rows walked two
' hundred columns three hundred times for each name it mentioned. Four and
' a half seconds of the Object Explorer's Tables node was two of those.
Private Function OuterColumnValue(ByVal name As String, _
                                  ByRef found As Boolean) As Variant
    Dim frame As Long
    Dim index As Long
    Dim outer As Variant
    Dim cached As Variant
    Dim answer(0 To 1) As Variant
    Dim held As Collection

    found = False
    If mOuterValues Is Nothing Then Exit Function
    If mOuterCaches.count > 0 Then
        Set held = mOuterCaches(mOuterCaches.count)
        On Error Resume Next
        cached = held(name)
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            found = CBool(cached(0))
            If found Then
                mOuterReads = mOuterReads + 1
                OuterColumnValue = cached(1)
            End If
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If

    For frame = mOuterColumns.count To 1 Step -1
        If Not mOuterColumns(frame) Is Nothing Then
            index = ColumnIndex(name, mOuterColumns(frame))
            If index > 0 Then
                outer = mOuterValues(frame)
                If index - 1 <= UBound(outer) Then
                    found = True
                    mOuterReads = mOuterReads + 1
                    OuterColumnValue = NormalizeValue(outer(index - 1))
                    If Not held Is Nothing Then
                        answer(0) = True
                        answer(1) = OuterColumnValue
                        held.Add answer, name
                    End If
                    Exit Function
                End If
            End If
        End If
    Next frame

    ' A name no outer read has either. Remembered as such, so the walk is
    ' not made again for every row of this one.
    If Not held Is Nothing Then
        answer(0) = False
        answer(1) = Empty
        held.Add answer, name
    End If
End Function

Private Function SaveParseState() As Variant
    Dim out(0 To 12) As Variant

    Set out(0) = mTokens
    out(1) = mCursor
    Set out(2) = mSelectItems
    Set out(3) = mFromItems
    out(4) = mFromName
    If IsObject(mWhereNode) Then Set out(5) = mWhereNode Else out(5) = mWhereNode
    Set out(6) = mGroupBy
    If IsObject(mHavingNode) Then Set out(7) = mHavingNode Else out(7) = mHavingNode
    Set out(8) = mOrderBy
    out(9) = mTopCount
    out(10) = mOffsetCount
    out(11) = mFetchCount
    out(12) = mTopPercent
    SaveParseState = out
End Function

Private Sub RestoreParseState(ByVal kept As Variant)
    Set mTokens = kept(0)
    mCursor = CLng(kept(1))
    Set mSelectItems = kept(2)
    Set mFromItems = kept(3)
    mFromName = CStr(kept(4))
    If IsObject(kept(5)) Then Set mWhereNode = kept(5) Else mWhereNode = kept(5)
    Set mGroupBy = kept(6)
    If IsObject(kept(7)) Then Set mHavingNode = kept(7) Else mHavingNode = kept(7)
    Set mOrderBy = kept(8)
    mTopCount = CLng(kept(9))
    mOffsetCount = CLng(kept(10))
    mFetchCount = CLng(kept(11))
    mTopPercent = CDbl(kept(12))
End Sub


' TRIM([LEADING | TRAILING | BOTH] [characters FROM] string), which reads its
' arguments as words rather than as a list: read as a list, the FROM between
' them was a syntax error. What comes out is the string, then the characters
' when there are any, then which end when that was said.
Private Function ParseTrim() As Variant
    Dim arguments As Collection
    Dim side As String
    Dim first As Variant

    Expect "("
    If Accept("LEADING") Then
        side = "LEADING"
    ElseIf Accept("TRAILING") Then
        side = "TRAILING"
    ElseIf Accept("BOTH") Then
        side = "BOTH"
    End If

    Set arguments = New Collection
    If Accept("FROM") Then
        ' TRIM(LEADING FROM x): an end and no characters, so spaces.
        arguments.Add ParseExpression()
        arguments.Add NodeOf(N_LITERAL, " ", Empty, Empty)
    Else
        first = ParseExpression()
        If Accept("FROM") Then
            arguments.Add ParseExpression()
            arguments.Add first
        Else
            arguments.Add first
        End If
    End If
    If Len(side) > 0 Then
        If arguments.count = 1 Then
            arguments.Add NodeOf(N_LITERAL, " ", Empty, Empty)
        End If
        arguments.Add NodeOf(N_LITERAL, side, Empty, Empty)
    End If
    Expect ")"
    ParseTrim = NodeOf(N_CALL, "TRIM", arguments, Empty)
End Function

' CONVERT(type, expression [, style]), which is CAST written the other way
' round. The style is read and ignored: nothing here formats a date by
' number.
Private Function ParseConvert() As Variant
    Dim typeName As String
    Dim node As Variant
    Dim arguments As Collection

    Expect "("
    typeName = TakeText()
    If PeekIs("(") Then typeName = typeName & BracketedText()
    Expect ","
    node = ParseExpression()
    If Accept(",") Then ParseExpression
    Expect ")"

    Set arguments = New Collection
    arguments.Add node
    arguments.Add NodeOf(N_LITERAL, typeName, Empty, Empty)
    ParseConvert = NodeOf(N_CALL, "$CAST", arguments, Empty)
End Function

' The text inside a bracketed string literal, with its doubled quotes put
' back to single ones. Tokenizing it is what unescapes them.
Private Function QuotedInside(ByVal written As String) As String
    Dim closing As Long
    Dim inner As String
    Dim tokens As Collection

    closing = MatchingBracket(written, 1)
    inner = TrimWhitespace(Mid$(written, 2, closing - 2))
    If Len(inner) = 0 Then Exit Function

    Set tokens = Tokenize(inner)
    If tokens.count = 0 Then Exit Function
    If tokens(1)(0) <> T_STRING Then Exit Function
    QuotedInside = CStr(tokens(1)(1))
End Function

' What CONNECTIONPROPERTY answers about the connection asking.
'
' A client reads these to find out how it got here and what authenticated
' it. Every connection this serves arrived over TCP and was let in by
' Windows, so the answers are the same for all of them.
Private Function ConnectionProperty(ByVal name As String) As Variant
    Select Case LCase$(Trim$(name))
        Case "net_transport": ConnectionProperty = "TCP"
        Case "protocol_type": ConnectionProperty = "TSQL"
        Case "auth_scheme": ConnectionProperty = "NTLM"
        Case "local_net_address": ConnectionProperty = "127.0.0.1"
        Case "client_net_address": ConnectionProperty = "127.0.0.1"
        Case "local_tcp_port": ConnectionProperty = 1433
        Case "physical_net_transport": ConnectionProperty = "TCP"
        Case Else: ConnectionProperty = Null
    End Select
End Function


' ============================================================================
' The functions a client asks with
'
' Ported wholesale from pySQLbridge rather than one at a time, because the
' list is not a matter of taste: it is what SSMS and the drivers actually
' call, and one missing name ends a connection with "is not a recognized
' built-in function name".
'
' The answers that are not simply computed are the ones worth reading. This
' server keeps no principals, no roles and no server state, and says so
' consistently: a yes here is followed by a question it cannot answer.
' ============================================================================

' What a name means when it depends on the connection rather than on its
' arguments. Answers Empty when the name is not one of these, so the caller
' can go on to the ordinary functions.
Private Function ContextFunction(ByVal name As String, _
                                 ByVal arguments As Collection, _
                                 ByRef values As Variant, _
                                 ByVal columns As Collection, _
                                 ByRef handled As Boolean) As Variant
    Dim first As Variant
    Dim second As Variant
    Dim third As Variant

    handled = True

    ' A date function names its part with a bare word, and a bare word is
    ' not an expression: working DATEDIFF(day, ...) out from the left looks
    ' for a column called day and fails there, before the case that knows
    ' what the word is for has been reached.
    If arguments.count >= 1 Then
        If Not NamesADatePart(name) Then
            first = Evaluate(arguments(1), values, columns)
        End If
    End If
    If arguments.count >= 2 Then second = Evaluate(arguments(2), values, columns)
    If arguments.count >= 3 Then third = Evaluate(arguments(3), values, columns)

    Select Case name
        ' What went wrong, inside the CATCH it sent a batch to, and NULL
        ' anywhere else. A CATCH that asked was refused as calling a
        ' function nothing had heard of, and the batch ended there.
        Case "ERROR_NUMBER"
            If mCaughtNumber = 0 Then ContextFunction = Null _
                Else ContextFunction = mCaughtNumber
        Case "ERROR_MESSAGE"
            If mCaughtNumber = 0 Then ContextFunction = Null _
                Else ContextFunction = mCaughtMessage
        Case "ERROR_SEVERITY"
            If mCaughtNumber = 0 Then ContextFunction = Null _
                Else ContextFunction = 16
        Case "ERROR_STATE", "ERROR_LINE"
            If mCaughtNumber = 0 Then ContextFunction = Null _
                Else ContextFunction = 1
        Case "ERROR_PROCEDURE"
            ContextFunction = Null

        ' With nothing to look up, whoever is asking. Given something to
        ' look up, nobody: this keeps no principals, so it cannot say who a
        ' sid belongs to, and answering with the caller would name the wrong
        ' person as the owner of everything.
        Case "SUSER_SNAME", "SUSER_NAME"
            If arguments.count = 0 Then
                ContextFunction = SessionValue("login")
            Else
                ContextFunction = Null
            End If

        Case "ORIGINAL_LOGIN"
            ContextFunction = SessionValue("login")

        Case "USER_NAME", "CURRENT_USER"
            ContextFunction = "dbo"

        Case "SCHEMA_NAME"
            ContextFunction = "dbo"

        Case "DB_NAME"
            ContextFunction = ConfiguredDatabase

        Case "DB_ID"
            ContextFunction = 1

        Case "HOST_NAME"
            ContextFunction = SessionValue("host")

        Case "APP_NAME"
            ContextFunction = SessionValue("app")

        ' Yes, whatever is named. This serves one catalog and a login may
        ' name any database to reach it, which the handshake already allows:
        ' SSMS connects naming msdb and is served. Saying no to a name the
        ' login accepts would be the inconsistent answer, and it is the one
        ' that makes a client report a feature unavailable for want of
        ' permission.
        Case "HAS_DBACCESS"
            ContextFunction = 1

        Case "OBJECT_ID"
            ContextFunction = ObjectIdOf(first)

        ' No roles are kept, so nobody is in one. Claiming otherwise would
        ' have a client offer what it cannot do.
        Case "IS_SRVROLEMEMBER", "IS_MEMBER", "IS_ROLEMEMBER"
            ContextFunction = 0

        ' No, whatever is asked. A real server says yes to a sysadmin here,
        ' and SSMS asks before reading the availability views. This keeps no
        ' server state to look at, so a yes would be followed by a question
        ' it cannot answer; no is both true and the answer that has the
        ' client skip the question.
        Case "HAS_PERMS_BY_NAME"
            ContextFunction = 0

        Case "GETDATE", "CURRENT_TIMESTAMP", "SYSDATETIME"
            ContextFunction = Now

        Case "GETUTCDATE", "SYSUTCDATETIME"
            ContextFunction = UtcNow()

        ' No sid for a name, because there are no principals to have one.
        Case "SID_BINARY", "SUSER_SID", "SUSER_ID"
            ContextFunction = Null

        Case "USER_ID"
            ContextFunction = 1

        ' Policy automation is off, and there are no diagrams because
        ' nothing here draws any.
        Case "FN_SYSPOLICY_IS_AUTOMATION_ENABLED", "FN_DIAGRAMOBJECTS"
            ContextFunction = 0

        Case "COLLATIONPROPERTY"
            ContextFunction = CollationProperty(CStr(NullToEmpty(second)))

        Case "DATABASEPROPERTYEX"
            ContextFunction = DatabaseProperty(CStr(NullToEmpty(second)))

        Case "OBJECTPROPERTY", "OBJECTPROPERTYEX"
            If IsNull(ObjectIdOf(first)) Then
                ContextFunction = Null
            Else
                ContextFunction = ObjectProperty(CStr(NullToEmpty(second)))
            End If

        Case "COLUMNPROPERTY"
            ' COLUMNPROPERTY(object, column, property): the column is named
            ' rather than numbered, and the property is the third argument.
            If IsNull(ObjectIdOf(first)) Then
                ContextFunction = Null
            Else
                ContextFunction = ColumnProperty(CStr(NullToEmpty(third)))
            End If

        Case "SERVERPROPERTY"
            ContextFunction = ServerProperty(CStr(NullToEmpty(first)))

        Case "CONNECTIONPROPERTY"
            ContextFunction = ConnectionProperty(CStr(NullToEmpty(first)))

        Case "FULLTEXTSERVICEPROPERTY"
            ContextFunction = 0

        Case Else
            handled = False
    End Select
End Function

' What the connection said about itself at login, or nothing when this is
' answering with no connection behind it.
Private Function SessionValue(ByVal name As String) As Variant
    Dim value As Variant

    If mSession Is Nothing Then
        SessionValue = Null
        Exit Function
    End If

    On Error Resume Next
    value = mSession(LCase$(name))
    If Err.Number <> 0 Then
        Err.Clear
        value = Null
    End If
    On Error GoTo 0
    SessionValue = value
End Function

' A number for a table this serves, and NULL for anything else.
'
' A client asks about objects a real server has and this one does not, and
' NULL is the answer that says so: SSMS reads it to decide whether a feature
' is installed. For a table that is served, the number is made from the name
' so it is the same number every time it is asked.
Private Function ObjectIdOf(ByVal name As Variant) As Variant
    Dim bare As String
    Dim parts() As String

    If IsNull(name) Or IsEmpty(name) Then
        ObjectIdOf = Null
        Exit Function
    End If

    bare = Replace(Replace(CStr(name), "[", ""), "]", "")
    parts = Split(bare, ".")
    bare = LCase$(parts(UBound(parts)))

    If Not HasTable(bare) Then
        ObjectIdOf = Null
        Exit Function
    End If
    ObjectIdOf = TableObjectId(bare)
End Function

' The number a table is known by, made from its name so it never moves.
Friend Function TableObjectId(ByVal name As String) As Long
    Dim total As Long
    Dim index As Long

    For index = 1 To Len(name)
        total = (total * 31 + Asc(Mid$(LCase$(name), index, 1))) And &HFFFFFF
    Next index
    ' Above the range a real server keeps for its own objects, so nothing
    ' this serves can be mistaken for one of them.
    TableObjectId = total + 1000000
End Function

' What a collation is, for the one collation this has. The numbers are what
' SQL Server reports for SQL_Latin1_General_CP1_CI_AS.
Private Function CollationProperty(ByVal wanted As String) As Variant
    Select Case UCase$(Trim$(wanted))
        Case "LCID": CollationProperty = 1033
        Case "CODEPAGE": CollationProperty = 1252
        Case "COMPARISONSTYLE": CollationProperty = 196609
        Case "VERSION": CollationProperty = 0
        Case Else: CollationProperty = Null
    End Select
End Function

' What a client asks about the database it is in. Read-only is the one that
' is not what a real server usually says, and it is what this is.
Private Function DatabaseProperty(ByVal wanted As String) As Variant
    Select Case UCase$(Trim$(wanted))
        Case "UPDATEABILITY": DatabaseProperty = "READ_ONLY"
        Case "STATUS": DatabaseProperty = "ONLINE"
        Case "COLLATION": DatabaseProperty = "SQL_Latin1_General_CP1_CI_AS"
        Case "RECOVERY": DatabaseProperty = "SIMPLE"
        Case "USERACCESS": DatabaseProperty = "MULTI_USER"
        Case "ISAUTOCLOSE", "ISINSTANDBY": DatabaseProperty = 0
        Case "ISREADONLY": DatabaseProperty = 1
        Case "VERSION": DatabaseProperty = 999
        Case "LASTGOODCHECKDBTIME": DatabaseProperty = Null
        Case Else: DatabaseProperty = Null
    End Select
End Function

' What a table this serves is: a plain one, not a view, not schema-bound,
' nothing generated.
Private Function ObjectProperty(ByVal wanted As String) As Variant
    Select Case UCase$(Trim$(wanted))
        Case "BASETYPE": ObjectProperty = "U"
        Case "ISTABLE", "ISUSERTABLE": ObjectProperty = 1
        Case "OWNERID", "SCHEMAID": ObjectProperty = 1
        Case Else: ObjectProperty = 0
    End Select
End Function

' What a column of a workbook is. Every value in one is stored rather than
' worked out, so nothing here is computed, indexed for full text, or given
' semantics; a cell that is empty is a null, so every column allows one.
Private Function ColumnProperty(ByVal wanted As String) As Variant
    Select Case UCase$(Trim$(wanted))
        Case "ALLOWSNULL": ColumnProperty = 1
        Case Else: ColumnProperty = 0
    End Select
End Function

' The functions that answer from their arguments alone.
Private Function PlainFunction(ByVal name As String, _
                               ByVal arguments As Collection, _
                               ByRef values As Variant, _
                               ByVal columns As Collection, _
                               ByRef handled As Boolean) As Variant
    Dim first As Variant
    Dim second As Variant
    Dim third As Variant
    Dim index As Long
    Dim text As String
    Dim part As Variant
    Dim started As Boolean
    Dim chosen As Variant

    handled = True

    ' A date function names its part with a bare word, and a bare word is
    ' not an expression: working DATEDIFF(day, ...) out from the left looks
    ' for a column called day and fails there, before the case that knows
    ' what the word is for has been reached.
    If arguments.count >= 1 Then
        If Not NamesADatePart(name) Then
            first = Evaluate(arguments(1), values, columns)
        End If
    End If
    If arguments.count >= 2 Then second = Evaluate(arguments(2), values, columns)
    If arguments.count >= 3 Then third = Evaluate(arguments(3), values, columns)

    If PropagatesNull(name) Then
        If IsNull(first) Or IsNull(second) Or IsNull(third) Then
            PlainFunction = Null
            Exit Function
        End If
    End If

    Select Case name
        Case "REVERSE"
            PlainFunction = StrReverse(CStr(first))

        Case "CHARINDEX"
            ' Where the first string appears inside the second, counting
            ' from one, and zero when it does not.
            If IsNull(first) Or IsNull(second) Then
                PlainFunction = Null
            ElseIf arguments.count >= 3 Then
                PlainFunction = InStr(CLng(third), CStr(second), CStr(first), _
                                      vbTextCompare)
            Else
                PlainFunction = InStr(1, CStr(second), CStr(first), vbTextCompare)
            End If

        Case "IIF"
            If IsNull(first) Then
                PlainFunction = third
            ElseIf CBool(first) Then
                PlainFunction = second
            Else
                PlainFunction = third
            End If

        Case "DATALENGTH"
            ' Bytes rather than characters, and text here is two bytes each.
            If IsNull(first) Then
                PlainFunction = Null
            ElseIf VarType(first) = vbString Then
                PlainFunction = Len(CStr(first)) * 2
            ElseIf IsNumeric(first) Then
                PlainFunction = 8
            Else
                PlainFunction = Len(CStr(first)) * 2
            End If

        Case "NULLIF"
            If IsNull(first) Or IsNull(second) Then
                PlainFunction = first
            ElseIf CompareValues(first, second) = 0 Then
                PlainFunction = Null
            Else
                PlainFunction = first
            End If

        Case "SIGN"
            If IsNull(first) Then
                PlainFunction = Null
            ElseIf CDbl(first) > 0 Then
                PlainFunction = 1
            ElseIf CDbl(first) < 0 Then
                PlainFunction = -1
            Else
                PlainFunction = 0
            End If

        Case "POWER"
            ' The answer is the type the base was declared, the way a
            ' division is: two to the minus one is a half, and a half in an
            ' integer is nought. By declaration and not by value, because a
            ' worksheet's whole numbers all arrive as reals.
            If IsWholeNumber(first) And IsWholeNumber(second) _
               And DividesWhole(arguments(1), columns) _
               And DividesWhole(arguments(2), columns) Then
                PlainFunction = CLngLng(Fix(CDbl(first) ^ CDbl(second)))
            Else
                PlainFunction = CDbl(first) ^ CDbl(second)
            End If
        Case "SQUARE"
            PlainFunction = CDbl(first) * CDbl(first)
        Case "EXP"
            PlainFunction = Exp(CDbl(first))
        Case "LOG"
            PlainFunction = Log(CDbl(first))
        Case "LOG10"
            PlainFunction = Log(CDbl(first)) / Log(10)
        Case "PI"
            PlainFunction = 3.14159265358979

        Case "SPACE"
            PlainFunction = Space$(CLng(first))
        Case "STR"
            PlainFunction = Trim$(Str$(CDbl(first)))
        Case "REPLICATE"
            If IsNull(first) Or IsNull(second) Then
                PlainFunction = Null
            ElseIf CLng(second) <= 0 Then
                PlainFunction = ""
            Else
                For index = 1 To CLng(second)
                    text = text & CStr(first)
                Next index
                PlainFunction = text
            End If

        Case "ASCII"
            If IsNull(first) Then
                PlainFunction = Null
            ElseIf Len(CStr(first)) = 0 Then
                PlainFunction = Null
            Else
                PlainFunction = Asc(Left$(CStr(first), 1))
            End If
        Case "UNICODE"
            If IsNull(first) Then
                PlainFunction = Null
            ElseIf Len(CStr(first)) = 0 Then
                PlainFunction = Null
            Else
                PlainFunction = AscW(Left$(CStr(first), 1))
            End If
        Case "CHAR"
            PlainFunction = Chr$(CLng(first) And &HFF)
        Case "NCHAR"
            PlainFunction = ChrW$(CLng(first))

        Case "PATINDEX"
            ' The pattern comes with its own wildcards, which are the ones
            ' LIKE uses rather than the ones VBA does.
            If IsNull(first) Or IsNull(second) Then
                PlainFunction = Null
            Else
                PlainFunction = PatternIndex(CStr(first), CStr(second))
            End If

        Case "STUFF"
            If IsNull(first) Then
                PlainFunction = Null
            Else
                PlainFunction = Left$(CStr(first), CLng(second) - 1) & _
                    CStr(NullToEmpty(Evaluate(arguments(4), values, columns))) & _
                    Mid$(CStr(first), CLng(second) + CLng(third))
            End If

        Case "CONCAT_WS"
            ' The separator is the first argument, and it goes between the
            ' values that are not NULL, empty ones included. The loop took
            ' the second argument for the separator and wrote each value
            ' over the real one, so a comma between 'a' and 'b' gave "aab".
            For index = 2 To arguments.count
                part = Evaluate(arguments(index), values, columns)
                If Not IsNull(part) Then
                    If started Then text = text & CStr(NullToEmpty(first))
                    text = text & CStr(part)
                    started = True
                End If
            Next index
            PlainFunction = text

        Case "GREATEST", "LEAST"
            ' The largest or the smallest, passing over NULLs: NULL only when
            ' every one of them is.
            chosen = Null
            For index = 1 To arguments.count
                part = Evaluate(arguments(index), values, columns)
                If Not IsNull(part) Then
                    If IsNull(chosen) Then
                        chosen = part
                    ElseIf name = "GREATEST" Then
                        If CompareValues(part, chosen) > 0 Then chosen = part
                    ElseIf CompareValues(part, chosen) < 0 Then
                        chosen = part
                    End If
                End If
            Next index
            PlainFunction = chosen

        Case "CHOOSE"
            index = CLng(first)
            If index < 1 Or index > arguments.count - 1 Then
                PlainFunction = Null
            Else
                PlainFunction = Evaluate(arguments(index + 1), values, columns)
            End If

        Case "TRANSLATE"
            If IsNull(first) Then
                PlainFunction = Null
            Else
                text = CStr(first)
                For index = 1 To Len(CStr(second))
                    text = Replace(text, Mid$(CStr(second), index, 1), _
                                   Mid$(CStr(third), index, 1))
                Next index
                PlainFunction = text
            End If

        Case "YEAR"
            PlainFunction = Year(CDate(first))
        Case "MONTH"
            PlainFunction = Month(CDate(first))
        Case "DAY"
            PlainFunction = Day(CDate(first))
        Case "EOMONTH"
            If IsNull(first) Then
                PlainFunction = Null
            Else
                PlainFunction = DateSerial(Year(CDate(first)), _
                                           Month(CDate(first)) + 1, 0)
            End If

        Case "DATEADD"
            If IsNull(second) Or IsNull(third) Then
                PlainFunction = Null
            Else
                PlainFunction = DateAdd(DatePartCode(CStr(NameOf(arguments(1)))), _
                                        CLng(second), CDate(third))
            End If
        Case "DATEDIFF"
            If IsNull(second) Or IsNull(third) Then
                PlainFunction = Null
            Else
                PlainFunction = DateDiff(DatePartCode(CStr(NameOf(arguments(1)))), _
                                         CDate(second), CDate(third))
            End If
        Case "DATEPART"
            If IsNull(second) Then
                PlainFunction = Null
            Else
                PlainFunction = DatePart(DatePartCode(CStr(NameOf(arguments(1)))), _
                                         CDate(second))
            End If
        Case "DATENAME"
            If IsNull(second) Then
                PlainFunction = Null
            Else
                PlainFunction = DateName(CStr(NameOf(arguments(1))), CDate(second))
            End If

        Case Else
            handled = False
    End Select
End Function

' Whether a whole number fits a column of this many bytes.
'
' One byte is a tinyint, which is unsigned and holds 0 to 255 rather than
' -128 to 127. Reading it as signed refused 170, which is a compatibility
' level the Object Explorer asks for on the way to drawing its first node.
Private Function FitsInteger(ByVal value As LongLong, _
                             ByVal width As Long) As Boolean
    Select Case width
        Case 1: FitsInteger = (value >= 0^ And value <= 255^)
        Case 2: FitsInteger = (value >= -32768^ And value <= 32767^)
        Case 4: FitsInteger = (value >= -2147483648^ And value <= 2147483647^)
        Case Else: FitsInteger = True
    End Select
End Function

' Half away from zero, which is what SQL Server does and VBA does not.
'
' VBA's Round is half to even: it answers 2 for 2.5 and 2 for 1.5, where a
' real server answers 3 and 2. A negative number of places rounds to the
' left of the point, which VBA's Round refuses outright.
Private Function RoundHalfUp(ByVal value As Double, _
                             ByVal places As Long) As Double
    Dim factor As Double
    Dim scaled As Double

    If places >= 0 Then
        factor = 10 ^ places
        scaled = value * factor
        RoundHalfUp = Sgn(scaled) * Int(Abs(scaled) + 0.5) / factor
    Else
        ' Scaled by dividing rather than by multiplying by a tenth, which
        ' no binary fraction holds exactly.
        factor = 10 ^ (-places)
        scaled = value / factor
        RoundHalfUp = Sgn(scaled) * Int(Abs(scaled) + 0.5) * factor
    End If
End Function

' SUBSTRING counts from one, and a start left of that is not an error: the
' string is read as though it went on to the left, and what comes back is
' the part of the span that lands on it. Mid$ raises instead.
Private Function Substring(ByVal text As String, ByVal start As Long, _
                           ByVal length As Long) As String
    Dim last As Long

    If length < 0 Then
        Err.Raise ERR_QUERY, "SqlBridge.Substring", _
                  "Invalid length parameter passed to the SUBSTRING function."
    End If

    last = start + length - 1
    If start < 1 Then start = 1
    If last < start Then Exit Function
    If start > Len(text) Then Exit Function
    Substring = Mid$(text, start, last - start + 1)
End Function

' The functions that answer with a whole number whatever they are given.
'
' A column is declared before any row is sent, and a function with nothing
' else to go on is declared from the value it produced. Where that value is
' NULL there is nothing to go on at all, and it came out as text one
' character wide: DATEDIFF over a NULL was declared nvarchar(1), so a client
' printing NULL against it printed N.
Private Function ReturnsWholeNumber(ByVal name As String) As Boolean
    Select Case UCase$(BareFunctionName(name))
        Case "LEN", "CHARINDEX", "ASCII", "UNICODE", _
             "DATEDIFF", "DATEPART", "YEAR", "MONTH", "DAY", _
             "ERROR_NUMBER", "ERROR_SEVERITY", "ERROR_STATE", "ERROR_LINE"
            ReturnsWholeNumber = True
    End Select
End Function

' The functions whose first argument is a part name rather than a value.
Private Function NamesADatePart(ByVal name As String) As Boolean
    Select Case name
        Case "DATEADD", "DATEDIFF", "DATEPART", "DATENAME"
            NamesADatePart = True
    End Select
End Function

' The word a date function names its part with, which is written as a bare
' word rather than as a string: DATEPART(year, x), not DATEPART('year', x).
Private Function NameOf(ByVal node As Variant) As String
    If Not IsArray(node) Then Exit Function
    If node(0) = N_COLUMN Then
        NameOf = CStr(node(1))
    ElseIf node(0) = N_LITERAL Then
        NameOf = CStr(NullToEmpty(node(1)))
    End If
End Function

Private Function DatePartCode(ByVal part As String) As String
    Select Case LCase$(Trim$(part))
        Case "year", "yy", "yyyy": DatePartCode = "yyyy"
        Case "quarter", "qq", "q": DatePartCode = "q"
        Case "month", "mm", "m": DatePartCode = "m"
        Case "dayofyear", "dy", "y": DatePartCode = "y"
        Case "day", "dd", "d": DatePartCode = "d"
        Case "week", "wk", "ww": DatePartCode = "ww"
        Case "weekday", "dw": DatePartCode = "w"
        Case "hour", "hh": DatePartCode = "h"
        Case "minute", "mi", "n": DatePartCode = "n"
        Case "second", "ss", "s": DatePartCode = "s"
        Case Else: DatePartCode = "d"
    End Select
End Function

Private Function DateName(ByVal part As String, ByVal value As Date) As String
    Select Case LCase$(Trim$(part))
        Case "month", "mm", "m": DateName = Format$(value, "mmmm")
        Case "weekday", "dw": DateName = Format$(value, "dddd")
        Case Else: DateName = CStr(DatePart(DatePartCode(part), value))
    End Select
End Function

' Where a LIKE pattern first matches, counting from one, and zero when it
' does not.
Private Function PatternIndex(ByVal pattern As String, ByVal value As String) As Long
    Dim at As Long
    Dim length As Long

    For at = 1 To Len(value)
        For length = 0 To Len(value) - at + 1
            If MatchesLike(Mid$(value, at, length), pattern) Then
                PatternIndex = at
                Exit Function
            End If
        Next length
    Next at
End Function


' A function's own name, without the database and schema a client may have
' written in front of it.
Private Function BareFunctionName(ByVal name As String) As String
    Dim parts() As String

    If Left$(name, 1) = "$" Then
        BareFunctionName = name
        Exit Function
    End If
    parts = Split(Replace(Replace(name, "[", ""), "]", ""), ".")
    BareFunctionName = UCase$(parts(UBound(parts)))
End Function


' The number sys.types gives a type, which a client reads to decide how to
' show a column.
Private Function SystemTypeId(ByVal typeCode As Long) As Long
    Select Case typeCode
        Case TYPE_INTN: SystemTypeId = 56           ' int
        Case TYPE_INT2: SystemTypeId = 52           ' smallint
        Case TYPE_DECIMALN: SystemTypeId = 106      ' decimal
        Case TYPE_FLTN: SystemTypeId = 62           ' float
        Case TYPE_BITN: SystemTypeId = 104          ' bit
        Case TYPE_DATETIMN: SystemTypeId = 61       ' datetime
        Case TYPE_GUIDN: SystemTypeId = 36          ' uniqueidentifier
        Case Else: SystemTypeId = 231               ' nvarchar
    End Select
End Function

Private Function NullToZero(ByVal value As Variant) As Variant
    If IsNull(value) Or IsEmpty(value) Then
        NullToZero = 0
    Else
        NullToZero = value
    End If
End Function


' A read used as a table: run it, and hand back what it produced under this
' source's own name.
Private Sub DerivedSource(ByVal item As Variant, ByRef columns As Collection, _
                          ByRef rows As Collection)
    Dim kept As Variant
    Dim inner As Collection
    Dim innerRows As Collection

    kept = SaveParseState()
    On Error GoTo Restore

    Set mTokens = item(6)
    mCursor = 1
    ParseSelect

    ExecuteWholeSelect inner, innerRows

    RestoreParseState kept
    On Error GoTo 0

    Set columns = QualifyColumns(inner, CStr(item(2)), item(4))
    Set rows = innerRows
    Exit Sub

Restore:
    RestoreParseState kept
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


' The first quoted string in a run of text, with its doubled quotes put back
' to single ones. Tokenizing is what unescapes them.
Private Function FirstQuotedIn(ByVal written As String) As String
    Dim tokens As Collection
    Dim index As Long

    If Len(Trim$(written)) = 0 Then Exit Function
    Set tokens = Tokenize(written)
    For index = 1 To tokens.count
        If tokens(index)(0) = T_STRING Then
            FirstQuotedIn = CStr(tokens(index)(1))
            Exit Function
        End If
    Next index
End Function


' One name, however it is written and however many parts it has.
'
' [sys].[dm_db_xtp_table_memory_stats] is one name, not three tokens and two
' dots: a client brackets whatever it quotes, and a parser that split there
' would read the schema as the table and the rest as an error.
Private Function ReadQualifiedName(ByRef sql As String, _
                                   ByRef cursor As Long) As String
    Dim length As Long
    Dim start As Long
    Dim ch As String
    Dim closing As String
    Dim name As String

    length = Len(sql)
    Do
        ch = Mid$(sql, cursor, 1)
        If ch = "[" Or ch = """" Then
            ' Brackets and double quotes both close a part, each with its
            ' own mark, and the mark doubled inside stands for itself. A
            ' worksheet may call a column [Bracketed], brackets and all, and
            ' a client asking for it writes [[Bracketed]]]: taking the first
            ' ] for the end of the name left ]] behind as an error.
            closing = IIf(ch = "[", "]", """")
            cursor = cursor + 1
            Do
                start = InStr(cursor, sql, closing)
                If start = 0 Then
                    RaiseQueryError ERR_NUMBER_BASE, _
                        "Unclosed quotation mark after the identifier."
                End If
                name = name & Mid$(sql, cursor, start - cursor)
                cursor = start + 1
                If Mid$(sql, cursor, 1) <> closing Then Exit Do
                name = name & closing
                cursor = cursor + 1
            Loop
        Else
            start = cursor
            Do While cursor <= length
                ch = Mid$(sql, cursor, 1)
                If IsNameChar(ch) Or IsDigitChar(ch) Or ch = "@" _
                   Or ch = "#" Or ch = "$" Then
                    cursor = cursor + 1
                Else
                    Exit Do
                End If
            Loop
            name = name & Mid$(sql, start, cursor - start)
        End If

        ' A dot joins the parts, and a bracket after one starts the next.
        If cursor > length Then Exit Do
        If Mid$(sql, cursor, 1) <> "." Then Exit Do
        If Right$(name, 1) <> "." Then name = name & "."
        cursor = cursor + 1
    Loop
    ReadQualifiedName = name
End Function
