# The Windows APIs this needs, and what VBA does to them

Measured on Windows 11 Pro 10.0.26200 with Microsoft 365 Excel x64, between
2026-09-08 and the same day's test run. Every claim here came from a failing
call rather than from documentation, because the documentation is right about
the API and silent about the language calling it.

## Schannel accepts two request flags, not six

The published sample for a TLS server passes six flags to
`AcceptSecurityContext`. Fed a real 1516-byte ClientHello from Python 3.14,
an inbound Schannel credential answers:

| Flags | Result |
| --- | --- |
| `ALLOCATE_MEMORY \| STREAM` | `SEC_I_CONTINUE_NEEDED`, 1167 bytes out |
| any of `REPLAY_DETECT`, `SEQUENCE_DETECT`, `CONFIDENTIALITY`, `EXTENDED_ERROR` added | `SEC_E_UNSUPPORTED_FUNCTION` |
| `STREAM` alone, or nothing | `SEC_E_INSUFFICIENT_MEMORY` |

The four that fail belong to the initiating side. The last row is not a flag
problem: without `ALLOCATE_MEMORY` the output buffer is null and Schannel has
nowhere to put the ServerHello.

The matrix was run against four protocol settings, TLS 1.2 server, system
default, 1.0 through 1.2, and 1.2 with 1.3, and the answer was the same in
every one, so the flags are the whole story.

Two further values matter and are easy to get wrong from memory.
`ASC_REQ_SEQUENCE_DETECT` is `0x8` and `ASC_REQ_REPLAY_DETECT` is `0x4`;
`0x2` is `ASC_REQ_MUTUAL_AUTH`, which asks Schannel for a client certificate
and is refused here.

## The certificate needs the AES provider

`CertCreateSelfSignCertificate` signs with SHA-1 when told nothing, and a
`PROV_RSA_FULL` container cannot sign with SHA-256 at all.
`PROV_RSA_AES` (24) can, so the key container is created there and the
signature algorithm is named outright as `1.2.840.113549.1.1.11`.

The key lives in a named container for as long as the server runs and is
deleted with `CRYPT_DELETEKEYSET` on shutdown. A container is a file in the
user's profile rather than something the handle owns, so a server that only
released its handles would leave one behind on every run.

Schannel finds the private key through the `CRYPT_KEY_PROV_INFO` stored in
the certificate, so the container name in that structure is load-bearing.

## What VBA does to the code calling all this

**A hex literal is typed by its width.** `&HFFFF` is an Integer holding -1,
and `&H8000` is -32768; both keep that value when assigned to a Long. The
maximum TDS packet size is written `65535` here for that reason, and every
hex constant in the class is either four digits and positive or eight digits
where the sign bit is the pattern the API wants.

**A `Declare` cannot be continued between `Lib` and `Alias`.** This is a
syntax error:

```vba
Private Declare PtrSafe Function sock_closesocket Lib "ws2_32.dll" _
    Alias "closesocket" (ByVal s As LongPtr) As Long
```

The compiler reports "Expected: line number or label or statement or end of
statement" and names no line, in a file of several thousand.

**A class module cannot expose a `Public Const`.** Constants, arrays,
fixed-length strings, user-defined types and `Declare` statements are all
refused as public members of an object module. Everything here is `Private`
and the tests assert against the wire values directly, which is the better
test anyway.

**A `Friend` member cannot expose a private type.** A connection therefore
reads the two halves of its server's credential handle as `LongPtr` and
rebuilds the `SecHandle` it passes to SSPI.

**A statically dimensioned array cannot be passed into a `Variant`
parameter.** `Dim node(0 To 3) As Variant` passed to `ByRef target As
Variant` raises type mismatch 13 at run time. `Collection.Add` takes one
happily, so the failure is specific to VBA's own procedures.

**Shadowing a VBA function with a local is a type mismatch at the call
site.** `Dim isNull As Boolean` turns `IsNull(value)` into an array index
into a Boolean, and the compiler says "Expected array". The same happened to
a `Column` factory shadowed by `Dim column As Variant`, which is why that
function is now called `ColumnOf`.

**`Trim` removes spaces and nothing else.** Not tabs, not carriage
returns, not line feeds. A batch split on its semicolons leaves each
statement after the first beginning with a newline, and a statement whose
first character is a newline has no first word: a batch of three reads was
answered with one, silently, because the other two were run by nothing.

**`Trim` removes spaces and nothing else, and it keeps costing.** The entry
above is one instance; there were five more, all of the same shape. A
statement runner takes the text after a keyword, trims it, and asks for its
first word: `INSERT INTO #t (a, b)` with the `SELECT` on the next line
leaves that text beginning with a newline, and a statement whose first
character is a newline has no first word. Everything here that splits
statement text now uses `TrimWhitespace`.

**A Collection asked for its nth item walks to it.** `things(i)` in a loop
over `things` is quadratic, and it reads as an ordinary indexed loop. `For
Each` is the fix wherever the position is not needed, and counting alongside
it is the fix where it is. This was in the sort, in the column lookup and in
the row encoder.

**An error handler is not free.** `On Error Resume Next`, a call, a check of
`Err.Number` and an `Err.Clear` cost more than the `UBound` they were
guarding. Asking a byte buffer for its capacity that way, once per field of
every row, was a measurable part of answering a large query.

**A statement splitter has to respect blocks.** Every keyword that can
start a statement can also appear inside one. `BEGIN TRY ... END TRY BEGIN
CATCH IF (...) BEGIN ... END END CATCH` split at the `IF`, which ran the
handler of a guarded block with nothing to handle. A block is one statement,
and so is an `IF` with its branches.

**A parameter cannot be called `scale`.** `Scale` is a drawing statement VB
kept from its form designer, so a procedure declaring one stops resolving
and the compiler reports every call to it as "Sub or Function not defined",
pointing at the call and not at the declaration. `precision`, `name`, `rows`
and `maxLength` are all accepted; `scale` alone is not.

**`Empty` is a reserved word.** `Dim empty() As Byte` is a compile error, and
under automation a compile error appears as a modal dialog that reads as a
harness failure rather than as the syntax error it is.

**A zero-length byte array is an unallocated one.** `ReDim` cannot express
it, so `EmptyBytes` is an array-valued function with an empty body: what it
returns is the only empty a `Byte()` has, and `ByteCount` reads it as zero
through an error trap around `UBound`.

**A guard cannot be an argument.** VBA works out a call's arguments before
the call, so this never guarded anything:

```vba
EvaluateCall = NullOr(first, UCase$(CStr(first)))
```

`CStr` has already run over the Null by the time `NullOr` is reached, and
`CStr` of a Null raises. Twenty-five functions were written that way and
every one of them answered a NULL argument with "Invalid use of Null". It
reads exactly like a guard, which is why it survived: the test has to happen
before the case does.

**`Left$` and `Right$` are shadowed by locals called `left` and `right`.**
A procedure holding `Dim left As Variant` cannot call `Left$`: VBA reads it
as that variable with a type-declaration character on it and reports
"Type-declaration character does not match declared data type", naming
neither the variable nor the call. `Mid$` does the same work and collides
with nothing. This is the `IsNull` entry above in another guise, and the
comparison evaluator is exactly where locals called left and right belong.

**`Mid$(s, 0)` is an error, not an empty string.** `Mid$` counts from one,
so asking a one-character operator for its last four characters asks for
position -2 and raises. Every comparison in every statement went through
that line, so the server answered nothing at all and the tests hung on a
call that never came back rather than failing with a message.

**A local cannot share a name with a parameter.** The compiler says
"Duplicate declaration in current scope" and names no line. `RunInsert`
already had a `Dim answers As Collection` for what its `SELECT` read when it
gained an `answers` parameter for what the client is told, and the two are
different things anyway: the rows going in, and the count going back.

**`ReDim a(2 To 1)` is a run-time error, not an empty array.** A table of
nothing but a header has no row to mark, and both `UPDATE` and `DELETE` say
so before they size anything.

**`Application.Union` costs more than linearly in what it already holds.**
Ten thousand single rows unioned one at a time took 249 seconds, of which
the delete at the end was 2. Deleting them in batches of a hundred takes
1.5. Measured at 25, 50, 100, 200, 400 and 800 areas per batch: flat from 25
to 200 and rising after 400, so the batch is a hundred.

**A `Range` follows the cells it named.** Taking `A4:B5`, inserting two rows
at row 4, and then assigning to that range writes at rows 6 and 7, over
whatever the insert pushed down there. The range has to be taken again after
the insert.

**Excel grows a table by itself, sometimes.** A value written in the row
directly under a `ListObject` is taken into the table, so a `Resize` to one
row more than it had swallows the row after that as well. What the table
should end up as is worked out before the write and only applied if the
table is not that size already.

**`WSAGetLastError` reads the thread's last error, and anything else on that
thread can have written it.** A non-blocking `accept` with nothing waiting
sets `WSAEWOULDBLOCK`, and the idle loop reads that and goes round again.
Under a harness that polls Excel over COM while a macro is running, it came
back 5 instead: a Windows access-denied left on the thread by something
else, which is not a code winsock produces at all, since its own start at
10004. The server raised on it and stopped.

Read at face value it looks like a socket failure and it is not one, so the
same read decided a receive was a closed connection and a send was a dead
one. Both dropped a live client. A code below 10000 now reads as nothing
waiting, which costs one more turn round the loop.

**A real server does not answer a session's SYN.** MARS runs over SMP, a
sixteen-byte header wrapping every message once a client and server have
agreed to it in PRELOGIN, and the linked-server provider will not connect
without it. The obvious reading of an open-acknowledge-close protocol is
that SYN is answered with SYN. It is not: a capture of a real server taking
a linked-server connection holds no server SYN at all, only DATA. Sending
one made the provider hang up with "the physical connection is not usable",
which names nothing.

Getting that capture needed the provider talked out of encryption, and
`Use Encryption for Data=false` does not do it for driver 19 -- every byte
still came through as TLS application data. `Encrypt=Optional` in the
provider string does.

**Another automation on the machine kills Excel by name.** A test file whose
Excel disappears partway through fails as a connection refused or an RPC
that is no longer there, which reads as a bug in whatever statement happened
to be running. It is not one. `Get-Process EXCEL | Stop-Process -Force` ends
every Excel on the machine, whoever started it and whatever it was in the
middle of, and a sibling repository's harness runs that line in three of its
scripts. There is no crash record for it because nothing crashed.

Measured: six runs lost out of sixty while one of those was up, none out of
forty once it had gone. What pointed at an outside cause rather than at this
code was that reads lost runs at the same rate as writes, that eight hundred
queries down one long-lived session never failed, and that one session died
at `Workbooks.Add`, before a line of this had been imported.

Two earlier entries stood here, blaming a `WorksheetFunction` call and then
Excel automation in general. Both were wrong, and both were written off
twenty clean runs. A rate of one in ten needs sixty runs before two changes
can be told apart, so twenty proves nothing either way, and a cause is worth
writing down only once the mechanism has been found rather than a
correlation.

## The timer, and where it will not run

`SetTimer` with a null window posts `WM_TIMER` to the calling thread, and
Excel's own message loop dispatches it into the `AddressOf` procedure
whenever nothing else is running. Measured at a 15 ms interval in an ordinary
Excel: 64 ticks in the first second, and a client logged in and queried
between them with no macro in flight.

The same pump takes Excel down on its first tick when the process is the
hidden one pyVBAharness supervises. That was measured with a callback whose
whole body was `gCount = gCount + 1`, so it is a property of that
environment rather than of anything the pump does. `tests/test_pump.py`
therefore drives Excel directly.
