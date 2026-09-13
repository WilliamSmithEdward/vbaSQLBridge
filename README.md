# vbaSQLBridge

Answer SQL Server's wire protocol from inside Excel, so Power BI, SSMS or
another Excel connects to a workbook and sees a database. Pure VBA: two files
you import, no reference, no add-in, no driver on either machine.

New to this? **[The guide](docs/guide.md)** is the same thing at walking
pace: what it is, five minutes to a working database, how to serve your own
sheets, what SQL you can write, and what to do when it does not work.

## Status

A client connects to a running workbook over TCP, negotiates a TLS tunnel,
authenticates through Windows, lists the tables, reads a worksheet and
writes back to it. The whole login sequence and the SQL below run inside
Excel with nothing installed: sockets come from ws2_32, the tunnel and the
login from secur32, and the certificate that tunnel presents from crypt32,
all through `Declare` statements.

```vba
Dim server As SqlBridge

Set server = SqlBridgeStart(1433)
server.AddTable "people", Sheet1.Range("A1:D5")        ' that rectangle
server.AddTable "orders", Sheet2.ListObjects("Orders")  ' that Excel table
server.AddTable "roster", Sheet3                        ' whatever it holds
```

That returns immediately. A timer serves connections while the workbook stays
usable, so nobody has to leave a macro running.

```
$ sqlcmd -S tcp:127.0.0.1,1433 -E
1> SELECT name, score FROM people WHERE score > 80 ORDER BY score DESC
2> GO
name             score
---------------- ------------------------
Ada Lovelace                         99.5
Barbara Liskov                      93.75
Grace Hopper                        87.25
```

| Piece | State |
| --- | --- |
| Packet framing, split and reassembly | done |
| PRELOGIN, byte for byte against a real server | done |
| TLS 1.2 handshake tunnelled in TDS packets | done |
| Self-signed certificate, generated at startup | done |
| Login state machine | done |
| LOGIN7 parse | done |
| Windows Authentication through SSPI | done |
| SQL Server logins, a name and a password the host allows | done |
| LOGINACK token stream | done |
| SQL batch parse | done |
| Result sets: int, float, bit, datetime, nvarchar, null | done |
| Worksheet ranges, tables and arrays as sources | done |
| Column types inferred from every value in a column | done |
| SELECT with a column list, TOP, WHERE, ORDER BY, DISTINCT | done |
| Aliases, expressions, CASE, LIKE, IN, BETWEEN, IS NULL | done |
| Whole-table aggregates: COUNT, SUM, MIN, MAX, AVG | done |
| STDEV, STDEVP, VAR and VARP, over a table, a group or a window | done |
| INFORMATION_SCHEMA tables, columns, schemata | done |
| A timer pump, so the workbook stays usable | done |
| Cancellation, acknowledged the way a client waits for | done |
| RPC, so parameterised queries work | done |
| GROUP BY and HAVING | done |
| The OLE DB catalog rowsets a schema browser calls | done |
| sp_tables, sp_columns, sp_databases | done |
| Multi-statement batches, variables, IF, temp tables | done |
| INNER and LEFT OUTER joins, CROSS APPLY over VALUES | done |
| CROSS and OUTER APPLY over a read that sees the row on its left | done |
| Scalar subqueries, CAST, bitwise operators, hex literals | done |
| SERVERPROPERTY, @@VERSION, xp_msver, the sys views | done |
| The batches SSMS 22 sends before it will connect | done |
| SMO reads the edition, version, platform and collation | done |
| The scalar functions a client asks with, all of them | done |
| sys.databases, sys.tables, sys.columns, sys.schemas | done |
| SMO enumerates the database and lists the tables | done |
| EXISTS, IN (SELECT ...), correlated subqueries | done |
| Derived tables, LIKE ... ESCAPE, comma cross joins | done |
| UNION, UNION ALL, EXCEPT, INTERSECT | done |
| BEGIN TRY, bitwise NOT, an INSERT with a column list | done |
| A CAST declares its column's type; decimal on the wire | done |
| The Object Explorer's Tables and Columns nodes | done |
| Indexes, keys, triggers, views, procedures, functions | done |
| Select Top 1000 Rows, three-part named | done |
| The policy store every node asks for its health state | done |
| A read that ignores the outer row run once, not once a row | done |
| Correlated EXISTS read as a lookup rather than a scan | done |
| A join that groups one side rather than trying every pair | done |
| Three-valued AND and OR, whichever side the unknown is | done |
| A string literal that spells an operator or a keyword | done |
| 7 / 2 against 7.0 / 2, and the sign of a remainder | done |
| A range in LIKE, and ORDER BY after a set operator | done |
| The same statements put to a real SQL Server and diffed | done |
| The Databases node's 139 properties, types included | done |
| An expression's type: CAST, ISNULL, CASE, aggregates | done |
| xp_instance_regread, so a client can read the settings | done |
| BEGIN TRY, and a block that stays one statement | done |
| A sheet read again per statement, so edits show up | done |
| An Excel table served by name, growing as rows are added | done |
| A worksheet read as a table: captions, gaps, blanks, errors | done |
| UPDATE, INSERT and DELETE, written back to the workbook | done |
| Rows affected, under the command field a client prints it for | done |
| Blocks rather than cells: runs, batched areas, one assignment | done |
| A column added, renamed, emptied or retyped, seen by the next query | done |
| Writes to a temporary table, whose rows are all there is of it | done |
| `ReadOnly`, refused the way a real read-only database refuses | done |
| A last-error value winsock did not write, read as nothing waiting | done |
| The date functions, whose first argument is a word and not a value | done |
| A NULL argument answered with NULL rather than raising | done |
| Text compared padded, so a trailing space decides nothing | done |
| A set operator inside a derived table, a subquery or an EXISTS | done |
| COUNT(DISTINCT x), and an integer too big for its column refused | done |
| NOT IN over a list holding NULL, which is never true | done |
| RIGHT and FULL OUTER joins | done |
| ANY, SOME and ALL against a read | done |
| Two strings compared as strings, whatever they spell | done |
| @@ROWCOUNT, which follows the statement before it | done |
| WHILE, which goes round rather than running once | done |
| OFFSET and FETCH, so a page is a page | done |
| An aggregate assigned to a variable | done |
| BEGIN TRANSACTION, COMMIT and a ROLLBACK that puts the rows back | done |
| A transaction per connection: @@TRANCOUNT, savepoints, held tables | done |
| Table hints, WITH (NOLOCK) and the older (NOLOCK), passed over | done |
| GRANT, REVOKE and DENY refused rather than completed | done |
| A read in brackets, and a UNION of reads in brackets | done |
| An Excel table's totals row left out, its names kept without a header | done |
| TRIM(... FROM ...), LTRIM and RTRIM of a set of characters | done |
| GREATEST and LEAST, passing over NULL | done |
| A date compared with text, or with another date, as a date | done |
| BeginTransaction, Save, Rollback and Commit from a client's API | done |
| OUTPUT parameters of sp_executesql, handed back as the batch left them | done |
| A batch carried on past the errors a real server carries on past | done |
| @@ERROR, and ERROR_NUMBER() and ERROR_MESSAGE() inside a CATCH | done |
| MARS, the session layer a linked server will not connect without | done |
| A real SQL Server lists this bridge's tables over a linked server | done |
| Four-part names, OPENQUERY and joins across both servers | done |
| A float written with an exponent, the way a pushed-down filter is | done |
| A prepared statement run again by its handle | done |
| INSERT, UPDATE and DELETE through a linked server | done |
| TOP (expression), and TOP ... PERCENT of the rows in the end | done |
| CONCAT_WS, its separator between whatever is not NULL | done |
| An UPDATE or DELETE refusing what it cannot read, before writing | done |
| Common table expressions, recursive ones and MAXRECURSION included | done |
| PIVOT, FOR XML and whatever else a read cannot take, refused | done |
| Window functions: ranking, LAG and LEAD, aggregates over a frame | done |
| An int's AVG a whole number, divided the way a real server divides | done |
| UPDATE ... FROM and DELETE ... FROM, the rows taken from a join | done |
| MERGE, with WHEN MATCHED, NOT MATCHED and NOT MATCHED BY SOURCE | done |
| TRUNCATE TABLE | done |
| A VALUES list joined to a table rather than read first | done |
| A Reload button, and every Excel table in the workbook served | done |
| Views, held as a named range pointing at the SQL in a cell | done |
| A view over a view, and a view of a join, read as one table | done |
| Views listed in INFORMATION_SCHEMA.VIEWS and sys.views | done |
| ORDER BY an aggregate, whether or not it is selected | done |
| A column name resolved to its position once a statement, not once a row | done |

## The workbook

`dist/vbaSQLBridge.xlsm` is built and ready to point something at. Open it,
enable macros, and press Start: the connection string appears on the sheet,
and the sample `people` and `orders` sheets are already served. Add a row to
the table list to serve another sheet of your own, and press Reload to serve
it without stopping. Every Excel table in the workbook is served too, under
its own name, so one added while the workbook serves is there after the next
Reload, and a client that refreshes its table list then shows it. The Writes
cell says whether a client may change the workbook; anything but `yes`
serves it read-only.

Clients log in with Windows authentication. Add login lets one in with a
name and a password as well: it asks for them in Windows' own dialog, so the
password is masked as it is typed, and puts the name and only the password's
hash in the login list beside the settings. A row there can name an
environment variable instead, `env:NAME`, for a password kept outside the
file. The workbook ships with the list empty, admitting Windows logins only.

```powershell
python scripts/build_workbook.py
```

That rebuilds it from `src/` and `demo/`, then opens what it made, starts it,
queries it with sqlcmd, writes to it, adds an Excel table and reloads to see
it served, and logs in with a name and a password from the login list and
has a wrong one refused, so a file that does not work does not ship.

## The demo workbook

`dist/vbaSQLBridge-demo.xlsm` is the same two modules with something to
chew on: 57,604 rows over six sheets, and eleven views over those.

| Sheet | Rows | What it is there for |
| --- | --- | --- |
| `orders` | 50,000 | a fact table: dates, customers, SKUs, prices, discounts, channels, a tenth of them unshipped |
| `Customers` | 5,000 | an Excel table, so the whole table path is exercised at size |
| `Products` | 1,000 | the third side of a three-way join |
| `metrics` | 1,096 | three years a day at a time, for running totals and moving averages |
| `wide` | 500 | 200 columns, because a column list is its own kind of load |
| `awkward` | 8 | spaces, brackets, accents, leading zeros, a column of two minds, gaps |

The views are stacked rather than parallel, which is the part worth looking
at. `order_lines` joins all three of orders, Customers and Products and
computes a line total. `monthly_revenue` groups that. `top_customers` ranks
it. `customer_health` joins `top_customers` back to Customers to find who is
over their credit limit. To a client that is four tables.

A dashboard sheet asks the same eight questions twice, once as an Excel
formula over the cells and once as SQL through the server, and prints both.
They agree on 1,838,690,979.93 of revenue, which is the point of it: the SQL
engine and Excel compute the same thing two different ways and land on the
same penny.

```powershell
python scripts/build_demo_workbook.py        # build, then verify
python scripts/build_demo_workbook.py verify # just the questions
```

The verify step puts fifteen statements through sqlcmd and prints what each
took, and the workbook does not ship unless all of them answer.

## Installing

Import two files into a macro-enabled workbook:

- `src/SqlBridge.cls`
- `src/SqlBridgeHost.bas`

`SqlBridgeHost` exists because Excel routes a `SetTimer` callback only to a
standard module. Everything else is in the class.

Serving on 1433 needs that port free, so a machine already running SQL Server
should pick another:

```vba
Set server = SqlBridgeStart(14330)
```

The listener binds to 127.0.0.1 by default. A bridge reachable from the
network is a database anyone on it can read and write, and that is a
decision for whoever runs it rather than a default:

```vba
Set server = SqlBridgeStart(14330, "0.0.0.0")
```

A client that can log in can change the workbook. To serve it the way this
served everything until now:

```vba
server.ReadOnly = True
```

A write can be taken back. A workbook has no log to undo from, so what
stands in for one is a copy: the first write to a table inside a transaction
keeps what that table looked like, and a rollback puts it back.

```
1> BEGIN TRANSACTION
2> DELETE FROM people
3> ROLLBACK
4> GO
```

Values, not formatting and not formulas, which is the part of a real
rollback this cannot do. A rollback unwinds to the outermost BEGIN, the way
a real server's does, and `SAVE TRANSACTION` marks a point that `ROLLBACK
TRANSACTION name` goes back to instead.

A transaction is its connection's own. `@@TRANCOUNT` counts that
connection's, and one client's `ROLLBACK` leaves another's writes alone. A
table written inside an open transaction is held by that connection until it
commits or rolls back. A real server would make another connection's write
wait; this one refuses it with 1222, the error a real server gives when a
lock is not granted in time. Reads are not held back, so another connection
sees what an open transaction has written. A connection that closes with a
transaction open has it rolled back, as a real server does. A client that
begins, saves, rolls back and commits through its API, as .NET's
`SqlTransaction` and ADO's `BeginTrans` do, sends requests of their own in
place of those statements, and they count, save and roll back the same
transaction. Joining a distributed transaction is refused: there is no
coordinator here to join.

When a statement fails, the batch goes on or stops the way a real server's
does, measured one error at a time against the one on this machine. After a
divide by zero, an arithmetic overflow, a COMMIT or ROLLBACK with nothing
open, a procedure that is not there, or a table another connection holds,
the rest of the batch runs and the error arrives among its results. After a
table that is not there, text that is not a date, or a SAVE with nothing
open, the batch stops, and what ran before it is kept. A column that is not
there is answered by itself, as on a real server, with one difference: a
real server finds it before running any of the batch, where here the
statements before it have already run, so a write earlier in the batch stays
written. `@@ERROR` reads the last statement's error until the next statement
finishes, and inside a CATCH, `ERROR_NUMBER()` and `ERROR_MESSAGE()` say
what sent it there.

## The database it serves

One database, called `vbaSQLBridge` unless you say otherwise:

```vba
server.Database = "Sales"
```

Deliberately not `master`. A client files a database away by its name, and
SSMS puts anything called master under **System Databases**, where the
Databases node reads as empty and nobody looking for their worksheet thinks
to open it. Connecting still works whatever catalog a connection string
names, because there is only one to reach.

## Serving data

`AddTable` takes a worksheet, a range, a list object, or a two-dimensional
array. The first row is the header.

```vba
server.AddTable "people", Sheet1.Range("A1:D5")   ' that rectangle
server.AddTable "orders", Sheet2.ListObjects("Orders")  ' that table
server.AddTable "roster", Sheet3                  ' whatever it holds
server.AddTable "grid", someArray                 ' a copy, once
```

The workbook's control sheet takes the same three. Leave its last column
empty to serve the whole sheet, or put an Excel table's name or a range
address in it; a table is looked for first, because Excel's `Range` answers
to a table's name as well and hands back the rows without the header, which
would make the first row of data into the column names.

A worksheet, a range and an Excel table are read again for every statement,
so a client asking twice about a workbook that was edited in between gets two
different answers, which is the point of pointing a client at a workbook.
What changes is what was named. A range is the rectangle it names, and a row
typed under it is outside the table. A sheet is whatever it holds, so rows
appear and go as they are typed and cleared. An Excel table is whatever the
table holds: its own range grows as rows are added to it, so it grows here
too, and it keeps its shape when something else is typed beside it on the
same sheet. An array is a copy and stays one; there is nothing left to read
it from.

The re-read is one pass per statement rather than one per mention of a table,
so a join naming the same sheet twice is one question about one moment, and a
sheet nobody has touched costs a comparison rather than a retyping of every
column.

A column's type is inferred from every value in it rather than per row,
because COLMETADATA declares a column once and every row is encoded against
that declaration. One value that is not a whole number drops the column to
float, and one that is not a number at all drops it to text.

A text column produced by an expression is declared at the width the answer
actually needs rather than at the widest a column can be, because a client
believes the declaration: sqlcmd pads every value out to it, and one grouped
name comes back four thousand characters wide.

Tables can be added and removed while a client is connected. The next query
sees them.

## Views, kept in the workbook

A view is a read with a name on it. `AddView` takes the name and the SQL:

```vba
server.AddView "big", "SELECT * FROM orders WHERE total > 1000"
server.AddView "by_month", "SELECT MONTH(ordered) AS m, SUM(total) AS n " & _
                           "FROM big GROUP BY MONTH(ordered)"
```

A client selects from a view, joins to it, groups and sorts it as though it
were a sheet, and never finds out it is a query. That is what makes a
report tool usable against a workbook: the awkward join lives here, and the
tool points at a table.

Views read views. `by_month` reads `big`, which reads a sheet; nesting stops
at 32 deep, so a view that reads itself is refused rather than followed until
Excel gives up. A view runs where it is read rather than where it was
defined, so it follows the cells underneath it, and its columns are worked
out per statement and remembered, so `SELECT *` and `SELECT one_column` do
not both pay for the whole thing.

`AddView` refuses anything that is not a read. A view that deletes rows when
you look at it is not a view.

The workbook keeps them as **defined names**. Any name beginning `sql_`
points at the cell holding the statement, and `sql_by_month` is served as
`by_month`:

```
views!C6  =  SELECT MONTH(ordered) AS m, SUM(total) AS n FROM big ...
name      =  sql_by_month  ->  =views!$C$6
```

The name points at the cell rather than holding a copy, so editing the cell
edits the view, and a name spanning several cells is joined top to bottom,
which is how a forty-line statement fits somewhere readable. The demo
workbook's views sheet is a name and a statement per row; Sync views writes
the names, Reload serves them. Nothing stops you naming a cell yourself in
the Name Box instead, which is the whole mechanism.

A view that does not parse is passed over by name and the rest are served,
because one bad statement should not take the workbook down with it. They
appear in `INFORMATION_SCHEMA.VIEWS`, in `sys.views` and in a client's object
list, marked as views, which is how a client finds them.

## A worksheet is not a table

It has a title over the top, a blank row somebody liked the look of, a
column that was emptied but not deleted, two headers spelt the same, a
totals row at the bottom, and a number typed as text. Serving one means
deciding what all of that means, so the decisions are written down and
tested rather than discovered a sheet at a time. `tests/corpus.py` holds
twenty-two of them, one worksheet each, and says what serving each should
produce.

* The table need not start at A1. Blank rows and columns around it are not
  part of it.
* A row with one cell on it, over a table wider than one column, is a
  caption rather than a header of one column. Any number of them can sit
  above the table.
* A column with nothing anywhere in it, header included, is not a column. A
  column with data and no name is a column, served under its position.
* Two columns with one name become `team` and `team2`, because a client
  asking for `team` has to get one of them.
* A blank row is not a row of nothings.
* An Excel error is not a value that can be sent, and goes as NULL. A
  formula goes as what it works out to.
* A number typed as text stays text, leading zeros and all. A column of
  numbers with one word in it is a column of text, because a column is
  declared once and every row is sent against that declaration.
* A totals row is served as the row it is. Guessing that a row is a summary
  would drop somebody's data.
* A sheet with nothing on it is a table with no columns rather than an
  error.

Whether the first row is the header is worked out: a row of nothing but
numbers and dates is data, and anything with a word in it is a header. That
is the only kind of missing header there is any telling about, because a
first row of `1` and `Ada` looks exactly like the header of a table whose
first column is called `1`. `AddTable` takes a third argument to say which
it is when the sheet cannot.

```vba
server.AddTable "readings", Sheet1, False   ' no header; name them by position
server.AddTable "counts", Sheet2, True      ' the first row is the header
```

All of it happens over the array Excel already handed across, so an untidy
sheet costs no more trips to Excel than a tidy one: reading twenty thousand
rows off a sheet and making a table of them is about a hundred milliseconds,
against eighty-five for the reading alone.

## Writing back

`UPDATE`, `INSERT` and `DELETE` change the workbook.

```
1> UPDATE people SET score = score + 1 WHERE team = 'red'
2> GO

(2 rows affected)
```

What can be written to follows from what was served. A sheet and an Excel
table both say how many rows they have by how many they have, so both can
gain and lose them. A range is the rectangle it names, and a row added to it
or taken out of it would make it a different rectangle, so it takes an
`UPDATE` and refuses the other two.

A temporary table and a table served from an array have no workbook behind
them: their rows are all there is of them, so their rows are what changes.
Writing to one changes the copy this holds and not the array you passed in,
and its columns stay as they were declared, because a temporary table's came
from its `CREATE TABLE`.

`server.ReadOnly = True` refuses all three, with the error a real server
sends for a read-only database, and `sys.databases.is_read_only` follows it
so a client shows the database that way.

`UPDATE ... FROM` and `DELETE ... FROM` take the rows to change from a
join. The join runs the way a read runs it, over a copy of the table whose
rows each carry their place in it, so every row the join keeps says which
row of the sheet it came from. A row the join keeps more than once is
changed once, from the first row it joined to, which is as much as a real
server promises. A `FROM` that leaves the table out has it joined on, the
way a real server reads one.

`MERGE` joins its target and its source in full, over copies that say for
each row which side it came from. Every joined row is then matched, a
source row the target has not got, or a target row the source has not got,
and takes the first `WHEN` clause of its kind whose condition holds. The
updates land where the rows are, the deletes go next and the inserts last.
A target row that two source rows match is refused before anything is
written, as a real server refuses it, and the count covers every row the
statement touched, under 0x117, the command a real server closes a `MERGE`
with. `TRUNCATE TABLE` empties a table and, like a real server, gives no
count.

An `UPDATE` or a `DELETE` that says anything this does not read, such as an
`OUTPUT` clause, is refused before a row is touched. Read past, the `WHERE`
behind it went unread as well, and the write went to every row.

A client only prints the count if the DONE token names the right command.
Real SQL Server sends 0xC3 after an INSERT, 0xC5 after an UPDATE and 0xC4
after a DELETE. This sent 0xC1, the SELECT, at first, and sqlcmd printed
nothing at all: what it had was a SELECT that answered with no rows. The
three values were read off the wire between sqlcmd and a real server on this
machine rather than guessed at.

The schema follows the sheet, because the sheet is the schema. A column
added, renamed, emptied, or given one value that changes its type is a table
of a different shape, and the next statement sees it: `sys.columns`,
`INFORMATION_SCHEMA.COLUMNS` and the COLMETADATA of the next answer all come
out of the same re-read.

The rows are worked out first and handed to Excel afterwards, in blocks. A
cell assigned on its own is a round trip, and a sheet of twenty thousand rows
is twenty thousand of them. An `UPDATE` groups the rows that changed into
runs that sit under each other on the sheet and assigns one array per run per
column. An `INSERT` writes the whole new block in one assignment, making room
first if something is already there, and afterwards tells an Excel table how
much bigger it is. A `DELETE` groups into runs the same way and hands them
over a hundred areas at a time.

## Over a linked server

A real SQL Server can take the workbook as a linked server, and a query
written in SSMS against that server can then join its own tables to a
worksheet:

```sql
EXEC sp_addlinkedserver @server = 'EXCEL', @srvproduct = '',
     @provider = 'MSOLEDBSQL', @datasrc = 'tcp:127.0.0.1,1433',
     @provstr = 'TrustServerCertificate=yes';
EXEC sp_addlinkedsrvlogin @rmtsrvname = 'EXCEL', @useself = 'true';

SELECT o.id, o.total, p.name
FROM   dbo.orders AS o
JOIN   [EXCEL].[vbaSQLBridge].[dbo].[people] AS p ON p.id = o.person_id;
```

Four-part names, `OPENQUERY`, `sp_tables_ex`, the filters the real server
pushes down and joins between the two servers all answer. So do `INSERT`,
`UPDATE` and `DELETE` on a four-part name, which land on the sheet.
`@useself = 'true'` is the mapping for a caller logged in with Windows
authentication. For callers Windows cannot vouch for, add a SQL Server login
to the bridge and map it with `@useself = 'false'`, `@rmtuser` and
`@rmtpassword`; the link then logs in as that login.

A linked server asks a good deal before it reads anything. Every question
below was read off the wire between the provider and a real server on this
machine, through a relay that recorded both directions:

* It will not connect without MARS, and a real server does not answer a
  session's SYN. `docs/windows-apis-from-vba.md` has the details.
* `@@SPID` is a smallint. Answered as an int, the provider hangs up saying
  the physical connection is not usable.
* It reads the catalog through rowset procedures a browser never calls:
  `sp_tables_info_90_rowset_64`, `sp_columns_100_rowset`,
  `sp_indexes_100_rowset`, `sp_check_constbytable_rowset` and
  `sp_table_statistics2_rowset`. A text column that comes back from
  `sp_columns_100_rowset` without a collation is refused with Msg 7368.
* It opens a transaction around every read with transaction-manager
  requests, which are packets of type 0x0E rather than statements, and
  rolls it back afterwards.
* It takes a schema lock with `sp_getschemalock`, which hands back a handle
  and a version as output parameters, and releases it after the read.
* It sends the read as `sp_prepexec`, with every name double-quoted,
  `"Tbl1002"."score"`, and any float it pushes down written with an
  exponent, `9.0000000000000000e+001`.
* `OPENQUERY` asks `sp_prepare` for the query's columns before it runs it.
  Without them the real server reports that the object has no columns. A
  read is run to find them and its rows are thrown away, so a pass-through
  query reads the sheet twice.
* It asks each catalog rowset about one table, by name, and takes more
  than one row back as more than one table of that name: Msg 7315. The
  rowsets here answered every table whatever was asked, which nobody
  noticed while only one table was served.
* An `UPDATE` or a `DELETE` is pushed down as a statement inside
  `sp_prepexec`, and the call closes on the write's own command, 0xC5 or
  0xC4, with the number of rows it changed.
* An `INSERT` comes as a server-side cursor: `sp_cursoropen` over
  `select * from` the table, answered with its columns, a hidden `ROWSTAT`
  and no rows, then `sp_cursor` once per row with the values named after
  their columns, then `sp_cursorclose`. A column flagged as computed, which
  is how a select list's columns go out here, is refused as one nobody may
  write: Msg 7344.

## What it costs

Measured on twenty thousand rows of four columns, answering offline so the
socket and the tunnel are a separate question. Milliseconds, best of three:

| Question | Was | Is |
| --- | --- | --- |
| `SELECT *` | 576 | 194 |
| `SELECT id, name` | 414 | 115 |
| `SELECT id, score ... ORDER BY score DESC` | 3960 | 184 |
| `WHERE score > 90` | 200 | 156 |
| `GROUP BY team` | 121 | 107 |

Five things were in the way, each found by measuring rather than by reading.
A Collection asked for its nth item walks to it, so numbering rows one at a
time read the whole list once per row. An ORDER BY term was resolved inside
the comparison, which is a quarter of a million resolutions for twenty
thousand rows. Every value was encoded into an array of its own, then
trimmed to length and copied in. The capacity of the answer's buffer was
asked for through an error-trapped `UBound` once per field. And a column
name was taken apart with two `Replace` calls and a search for a dot before
every lookup, on every row.

A third, larger than either: what each select item takes from a row was
worked out per row. A name is resolved by walking the columns and comparing
strings, so `SELECT id, name, team, score` cost 418ms over twenty thousand
rows where `SELECT *` over the same four cost 254, the star being faster
only because it reads a position it already knows. Settling both once per
statement took four named columns to 204, one column from 100 to 62, and a
qualified `p.id` from 158 to 62. `SELECT *` itself went to 194, because
which columns a star wants was also being decided per row.

Two more, over the same twenty thousand rows. A string value was
encoded by allocating a byte array, copying into it, and then asking its
length twice through an error-trapped `UBound`: a VBA string is already
UTF-16, so it goes into the buffer as it stands. `SELECT *` went from 281 to
255 and a text column from 161 to 145. The value each row sorts on was read
twice per comparison, which for twenty thousand rows is half a million
readings of twenty thousand values; read once per row instead, `ORDER BY`
went from 305 to 281.

Measured beside those: an append through a helper costs 109ns against 10.5
for the same bytes written where they go, of which 27 is the call itself and
17 the error handler inside `EnsureCapacity`. That is the floor on what any
further inlining could win, and it is not where the rest of the time goes.

A read inside another read is run once per row of the outer one, which is a
pass over the inner table per outer row. Most of them do not look at the
outer row at all and so have one answer; those are run once now. Reading the
outer row is the only way a read can be correlated, so counting those reads
is what tells the two apart, rather than guessing at the shape of the
statement. Over five hundred rows against five hundred:

| Question | Was | Is |
| --- | --- | --- |
| `WHERE id > (SELECT MIN(owner) FROM orders)` | 776 | 9 |
| `WHERE id IN (SELECT owner FROM orders)` | 989 | 15 |
| `WHERE EXISTS (SELECT 1 FROM orders WHERE owner = p.id)` | 4981 | 94 |
| `FROM people p JOIN orders o ON o.owner = p.id` | 3596 | 24 |

All four are linear from there: doubling the rows to a thousand takes them
to 16, 22, 101 and 37 milliseconds.

A join tried every pair, which for five hundred rows against five hundred is
a quarter of a million rows built and a quarter of a million conditions
evaluated. An equality ANDed into the condition has to hold for a pair to be
kept, so the rows on one side are grouped by it and the pairs it rules out
are never built. It rules nothing else out: the whole condition is still
evaluated on every pair the key kept, so a second condition beside the
equality decides what it always decided. A join on an inequality, or on an
OR, has nothing to group by and tries every pair as before.

`IN` looks its value up rather than walking the candidates, keyed the way
this server compares: a number beside text converts the text, so 2 and '2'
are one candidate, and text matches without regard to case. A list holding a
Boolean declines the lookup and is walked, because a Boolean compares by
converting whatever is beside it and that is a conversion that can fail
rather than a key that can be written.

A correlated `EXISTS` asks the same thing of every outer row: is this value
among the owners. The owners are read once and each outer value looked up,
which is what turns the pass-per-row into a pass. It also stops at the first
row that passes and never works out a column, which is what the shapes that
cannot take the lookup gained.

The lookup only fits some shapes. What has to hold: one outer column,
matched by an equality that is ANDed with the rest of the condition, and
nothing else in the subquery looking at the outer row. A `GROUP BY`, a
`HAVING` or a `TOP` decides which rows there are after the condition has
run, so dropping the condition would decide it over different rows. An `OR`,
a comparison that is not equality, two outer columns matched at once, and an
outer expression rather than an outer column all fall back to a row at a
time.

Twenty-six shapes are checked three ways. `tests/test_differential.py` puts
each of them to a real SQL Server on the same machine and to this one and
compares; `tests/test_shortcut.py` runs each twice, once with the lookup
turned off, and the two have to agree. Agreement is cheap for a shortcut
that never fires, so which shapes take it is counted as well, and with the
shortcuts turned off none is built at all.

A table served from a worksheet is read again before every statement, which
at twenty thousand rows costs about eighty milliseconds on top of the
numbers above. Most of that is Excel handing over the cells; the comparison
that decides whether anything changed compares them as they are rather than
as text.

Writing, over the same twenty thousand rows, each statement measured on a
freshly filled and freshly served sheet. Reading the sheet is most of every
number here: `SELECT *` over the same table is 387.

| Statement | ms |
| --- | --- |
| `UPDATE ... WHERE id = 7` | 275 |
| `UPDATE`, every row | 235 |
| `UPDATE`, two columns, every row | 301 |
| `UPDATE ... WHERE id % 2 = 0`, ten thousand runs of one | 521 |
| `INSERT`, a thousand rows | 410 |
| `DELETE ... WHERE id <= 1000`, one run | 268 |
| `DELETE ... WHERE id % 2 = 0`, ten thousand runs of one | 1680 |

That last one was 254 seconds before the areas went over in batches, and
`Application.Union` was 249 of them: the delete at the end of it was 2. Ten
thousand cells assigned one at a time is 294, which is what a write that did
not group into runs would have cost at best.

At fifty thousand rows the shape of the cost changes, and measuring the demo
workbook says where it went. Starting a statement is a few hundredths of a
second whether the workbook holds six rows or sixty thousand, and re-reading
the sheet is not it either: what is left is per row and per comparison.
Adding a row to a Collection is 1.4 microseconds, so fifty thousand of them
is 0.07 of a second and not worth a thought. A join probe was ten.

Four things, measured by serving the same 57,604-row workbook twice with
only the library swapped, so the data and the machine are held still:

| Statement over 50,000 rows | Was | Is |
| --- | --- | --- |
| `JOIN Customers c ON c.id = o.customer_id` | 1.30s | 0.76s |
| the same, grouped by country | 2.81s | 1.32s |
| `order_lines`, the three-way join | 5.05s | 2.47s |
| that view grouped by category | 6.17s | 3.92s |
| `COUNT(*)`, `WHERE`, `GROUP BY` | 0.19-0.50s | 0.20-0.50s |

A column name in an expression was resolved against the column list every
time the expression was evaluated, which for a join condition is twice per
pair tried. The statement knows its columns before the first row, so the
expression is walked once and every column node replaced by the position it
resolved to; what is left at run time is an array index. The rewritten tree
is thrown away with the statement, so nothing has to be invalidated when a
sheet changes shape.

`EvaluateBinary` reached a comparison through a dozen string comparisons
against the other operators, once per row asked about, before it had even
fetched the two values. Comparisons are most of what a `WHERE`, a join and a
sort ask for, so they are answered first. `CompareValues` asked `VarType`
six times a comparison and now asks twice, taking two numbers, which most
comparisons are, out at the top. And a join key went through `IsNumeric`,
which is a conversion attempt with an error trapped around it, for every row
on both sides; a number is now recognised by its type and the trap is only
reached by text that reads as a number, which is the case that needs it so
that `5` and `'5'` keep the one key.

The scans did not move, which is the useful half of the result: they were
already down to the cost of walking rows and putting them somewhere, and
that is the floor until rows stop being Collections.

## Without a network

The catalog and the SQL work without a socket, which is how the query surface
is tested and how you can check a statement from a worksheet:

```vba
Dim server As SqlBridge

Set server = SqlBridge.Offline()
server.AddTable "people", Sheet1.Range("A1:D5")
```

## What a client asks before it will show anything

A driver does not open with a query. SSMS sends a batch that makes a temp
table, fills it from `xp_msver` and from a `CROSS APPLY` over `VALUES`,
declares variables, skips a block guarded by `SERVERPROPERTY('EngineEdition')`,
and only then reads a row built out of `@@microsoftversion` taken apart with
integer division and a bitwise and, with a scalar subquery in the middle of
the select list. Answering that batch with nothing is what "Cannot find table
0" means.

Those batches are in `tests/probes/`, and `tests/test_ssms.py` runs them
after every change. SSMS 22 asks a different set again, so `tests/test_smo.py`
drives SMO itself, which is what the Object Explorer is built on: connect,
read the server's properties, ask whether msdb is reachable, enumerate the
databases, list the tables, then open the Tables and Columns nodes. Each
step reports separately, because one unanswered question otherwise hides the
next one.

Opening a node is not asking for names. The Databases node reads a hundred
and thirty-nine properties of every database it draws, several of them into
enum members and three into a `Guid`. An enum only accepts its own
underlying type, so `tinyint` read as `int` throws rather than rounds; a
`Guid` cannot hold a null at all. A database that fails any of the hundred
and thirty-nine is a database that does not appear in the tree.

That is why an expression's type is worked out rather than guessed from the
value it produced. A `CAST` says what it is, `ISNULL` takes its first
argument's, an aggregate keeps what it summed, a `CASE` is as wide as the
widest branch it can produce, arithmetic is as wide as its operands, a read
used as a value is asked what it selects, and a bare `NULL` is an integer.
Only what none of those covers is read back off the values, which is wrong
whenever the values are all null.

The Tables node sends one statement
of some seventeen thousand characters that reads every property of every
table it will draw: twenty-six system views, five temporary tables, a dozen
outer joins, correlated `EXISTS` subqueries and a `LIKE` with an `ESCAPE`.
The Columns node sends another that reads fifty-seven fields per column,
each with a getter for the type the column declared, so a column typed
merely close to a real server's is a failed read rather than a rounded
value. The second was settled by running the same statement against a real
SQL Server and diffing the fifty-seven declared types until none differed.

The nodes under a table, and the Views, Programmability and Security nodes
beside it, are answered too. A workbook has no indexes, keys or triggers, so
each of them answers zero rather than failing, which is the difference
between an empty node and a red cross. The lists of columns those views
declare while empty were not guessed either: the server logs every statement
it is asked, so the list is every column an SMO run actually named, read
back out of that log.

`SELECT TOP (1000) ... FROM [vbaSQLBridge].[dbo].[people]`, which is what
the menu on a table generates, reads the worksheet. Writing that statement
costs more than reading it: the menu asks for one property SMO has not got,
which initialises the whole server, which reads a dozen of its own settings
through `xp_instance_regread`. A procedure a batch cannot find ends the
batch, and the menu reports that as a failure to select from the table. The
settings this can answer it answers, and the query is written to expect
nothing for the rest.

One more thing reads before the tree is drawn at all. Every node asks the
policy store for a health state, and the store reads whether policy
management is enabled before it will answer. Answering with no rows is not
the same as answering no: the store reads that flag with a getter for a
bool, and a missing row reaches it as a null it cannot cast, which surfaces
at login as "Specified cast is not valid (Microsoft.SqlServer.Dmf)" with no
mention of what it was reading. `msdb.dbo.syspolicy_configuration` is served
with policy management off, which is both true and the answer that stops
SSMS offering to evaluate policies against a worksheet.
`tests/test_dmf.py` drives that path, up to the `GetAggregatedHealthState`
call the error came out of.

The functions come from the same place. Porting them one at a time as a
client complained was the slow way round: the list in pySQLbridge's
`predicate.py` is not a matter of taste, it is what SSMS and the drivers
actually call, and one missing name ends a connection with "is not a
recognized built-in function name".

## SQL Server logins

A client can log in with a name and a password as well as through Windows,
the way it logs in to a SQL Server login:

```vba
server.AddLogin "reporter", Environ$("REPORTER_PASSWORD")
server.AddLoginHash "analyst", "pbkdf2_sha256$210000$...$..."
```

Nothing is admitted that way until a login is added. With none, a client
asking for SQL authentication is refused, as a real server refuses a login
it does not have. A password is kept only as a salted PBKDF2-SHA256 hash,
worked out by Windows' CNG, and it is never logged. `HashPassword` writes the
form `AddLoginHash` takes, so a module or a workbook can hold that rather
than the password; pySQLbridge writes the same form, and each reads the
other's. A name matches in any case and a password only exactly.
`RemoveLogin`, `ClearLogins` and `LoginNames` do what they say, and the names
never come with what they log in with.

A login that fails gets what a real server sends, measured against the one
on this machine: error 18456 at severity 14 and state 1, `Login failed for
user 'reporter'.`, and then the connection closes. A name that is not there
and a password that is wrong get the same answer and the same work: the
unknown name is checked against a decoy hash, so the time a refusal takes
does not say whether the name exists. The password arrives masked with a
constant the protocol publishes, which is not encryption. What keeps it
private is the TLS tunnel every login here goes through.

## What the login actually does

Everything in `docs/` was measured rather than read from a specification.
Three things are worth knowing before reading the code.

**Encryption is negotiated off and happens anyway.** Both sides advertise
"encryption is available but off", a TLS handshake follows regardless, the
login goes through it, and the connection reverts to cleartext immediately
afterwards. A server without a working TLS handshake cannot log anybody in,
including in the configuration that sounds like it does not need one.

**The handshake rides inside TDS packets.** The TLS records travel as the
payload of packets typed 0x12, so the framing has to be added and removed
underneath them for the duration of the handshake and then stopped. The
application records that carry the login are bare.

**Windows does the authentication** of a Windows login. The client's blob is
SPNEGO wrapping NTLM, and it goes to `AcceptSecurityContext` unopened, so
nothing here handles that credential. A SQL Server login is the one
credential checked here, and only once the host has added one; the rules that
check keeps are in the section above.

## Tests

```powershell
python -m pytest tests
```

VBA procedures in `tests/vba/test_*.bas` are collected as pytest items
through [pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness),
and the Python tests in `tests/` drive a real TDS client against a server
running inside Excel: PRELOGIN, a TLS handshake framed in TDS, a LOGIN7
carrying a token Windows produced, then queries. Nothing is mocked, because
the thing worth proving is that a real client's sequence gets a real answer.

`tests/test_ssms.py` runs the two batches SSMS sends while connecting,
through sqlcmd, against the live server.

`tests/test_oledb.py` drives MSOLEDBSQL, which is the provider Excel and
Power BI use and the one the reference capture came from. It sends a
parameterised query, which arrives as a call to sp_executesql rather than as
a batch. It runs a prepared command three times, which arrives as
sp_prepexec and then as sp_execute with the handle the first call handed
back. It sets two OUTPUT parameters, a number and a string, and reads them
back after the call. It also calls `OpenSchema`, which does not send SQL at
all.

`tests/test_linked.py` has the real SQL Server on the machine take the
bridge as a linked server. It asks that server for the catalog, a four-part
read, an `OPENQUERY`, a filter it pushes down, a count, and a join between
rows that exist only on the real server and rows that exist only in the
workbook. It also inserts, updates and deletes through the link and reads
the cells back through Excel, and makes a second link that logs in with a
name and a password rather than as its caller. It is skipped where there is
no SQL Server, and the links are dropped afterwards.

`tests/test_transactions.py` holds transactions open on real connections,
ADO and sqlcmd side by side. One connection's ROLLBACK leaves another's
writes alone, a table one connection's open transaction has written is
refused to the others, and a bare client that closes its socket in the
middle of a transaction has it rolled back. It also sends the requests a
client's transaction API sends in place of SQL, begin, save, roll back to a
savepoint and commit, both as bare packets and through .NET's
`SqlTransaction`, and checks the count and the cells after each one.

`tests/test_sql_login.py` logs in with a name and a password through
sqlcmd, ADO, .NET's SqlClient and a bare client, has a wrong password and an
unknown name refused, reads the refusal's tokens against what a real server
sends, and checks the password never reaches the log.
`tests/vba/test_logins.bas` checks the unmasking against a byte worked out by
hand and the hashing against hashes Python's `hashlib` wrote.

`tests/test_dmf.py` drives the policy store, which the Object Explorer reads
once per node.

`tests/test_corpus.py` builds every sheet in `tests/corpus.py` and checks
what a client sees of it. A surface is declared once and built two ways, as
VBA for the harness and over COM for a script holding its own Excel, so a
test and a probe cannot drift apart.

`tests/test_write.py` writes to a worksheet and to an Excel table, and reads
the cells back through Excel as well as asking the server: a write nobody can
see in the workbook is not a write. It covers what each kind of source will
and will not take, what a row count says, what happens to something sitting
where a new row is going, and the catalog following a column that was added,
renamed, emptied or retyped.

`tests/test_differential.py` puts the same statements to a real SQL Server
and to this one and compares what came back, across four hundred and
twenty-five shapes: every operator beside NULL, the string and number
functions, joins, grouping, ordering, set operations and subqueries. It is
skipped where there is no SQL Server to compare against, which is most
machines.

`tests/test_views.py` reads a view as a client does: from the catalog, from
the schema rowsets, joined to a sheet, stacked on another view, and
following the cells underneath it when one of them is typed over.

It has found thirty-six bugs so far, and all but one of them nobody had
thought to write a test for. The first ten came from the operators and the
NULLs:

* `NULL OR TRUE` read as unknown rather than true, and `NULL AND FALSE` the
  same, so a row that should have been kept was dropped. The short-circuit
  only settled it when the unknown was on the right.
* A token was matched by its text, so the string `'-'` was read as a minus
  sign and `ISNULL(x, '-')` failed to parse. Any string spelling an operator
  or a keyword did the same.
* `7.0 / 2` was 3 rather than 3.5. Whether a division is a whole one depends
  on how the number was written, and the value alone cannot say.
* `-7 % 3` was 2 rather than -1: a remainder takes the sign of what was
  divided, and the quotient truncates toward zero rather than downward.
* `ISNULL(NULL, 'x')` failed. A bare NULL says nothing about the type, so
  taking its own sent a string down a column declared an integer.
* `ORDER BY` after a `UNION`, `EXCEPT` or `INTERSECT` sorted inside the last
  side by a name that side had not got, both when it named a column and when
  it gave a position.
* `ORDER BY p.team` over a grouped answer could not find `team`, because a
  grouped answer's columns carry no source name.
* `LIKE '[AB]%'` matched nothing: a set of characters was escaped into a
  literal bracket rather than passed through.

The next eleven came from asking about the parts nobody had asked about:
the date functions, which had no case here at all, the writes, and the
edges of the string and number ones.

* Every function given NULL raised "Invalid use of Null" instead of
  answering NULL. Twenty-five of them were guarded as
  `NullOr(first, UCase$(CStr(first)))`, and VBA works out a call's arguments
  before the call, so `CStr` had already run over the Null by the time the
  guard was reached. It reads exactly like a guard and never was one.
* `DATEADD`, `DATEDIFF`, `DATEPART` and `DATENAME` were unusable. Their
  first argument is a bare word rather than an expression, and working the
  arguments out from the left looked for a column called `day`.
* A NULL from a function that answers with a number was declared
  `nvarchar(1)`, because a NULL value says nothing about its own type. A
  client printing NULL against it printed `N`.
* `ROUND` rounded half to even, which is what VBA does: 2.5 came back 2
  where a real server says 3. A negative number of places was refused
  outright.
* `SUBSTRING` refused a start position below one. A real server reads the
  string as though it went on to the left and answers with the part of the
  span that lands on it.
* `POWER(2, -1)` answered 0.5. The answer takes the type of its base, the
  way a division does, so two whole numbers give a whole number.
* An integer too big for the column it was declared in was sent as its low
  four bytes: `100000 * 100000` came back 1410065408. It is refused now,
  with the number and the words a real server uses.
* Text compared unpadded, so `'a' = 'a '` was false. SQL Server makes the
  shorter side up to the longer with spaces, in `=`, in `IN` and in `LIKE`,
  and a leading space still counts.
* A set operator inside a derived table, a subquery or an `EXISTS` ran only
  its first branch and dropped the rest without saying so. `COUNT(*)` over a
  derived table of two rows answered one. The read carries on past where the
  SELECT stopped, and only the outermost one was reading on.
* `COUNT(DISTINCT team)` was a syntax error: `DISTINCT` was read as the
  start of an expression, which made it a column name and the column after
  it the error.
* A one-byte integer was range-checked as though it were signed, which
  refused 170. That one is the exception above: the check that found it was
  the overflow check three bullets up, and the Object Explorer asks for a
  compatibility level of 170 on the way to drawing its first node.

A third round asked about the classics, the joins, and the conversions,
and found six more.

* `NOT IN` over anything holding a NULL returned rows. A NULL among the
  candidates is a candidate the value might be equal to, so nothing matching
  is unknown rather than false, and `NOT IN` over it is never true. Both the
  list form and the subquery form were wrong, and the subquery form is the
  one a client writes without thinking about it.
* Two strings compared as numbers whenever both spelled numbers, so `'10'`
  was greater than `'9'`. A column of numbers typed as text is a column of
  text, which is the whole point of the corpus rule about them.
* `CAST(-2.7 AS int)` was -3. A cast truncates toward zero rather than
  downward, the same way a remainder does.
* `CAST('abcdef' AS nvarchar(3))` sent all six characters down a column
  declared for three. sqlcmd stopped with an internal error rather than
  showing a wrong value, which is what an overlong value on the wire does.
* `RIGHT JOIN` and `FULL JOIN` were not implemented and failed as an invalid
  column name.
* `ANY`, `SOME` and `ALL` were a syntax error.

A fourth round asked what a batch remembers, and found four more.

* `@@ROWCOUNT` answered nought whatever had just happened. A client asks how
  many rows the last statement came to far more often than it asks for the
  rows, and every answer was that nothing had.
* `WHILE` was not a statement here at all. The splitter took its body for a
  statement of its own and ran it once, which is a loop that does not loop
  and no error to say so.
* `OFFSET` and `FETCH` were read as the end of a clause and then dropped, so
  a client asking for the second page of an answer was sent all of it.
* `SELECT @n = COUNT(*)` refused the count: an aggregate assigned to a
  variable was handed to the expression evaluator a row at a time, and that
  has no COUNT in it. A count over no rows is nought rather than nothing,
  too.

A fifth round asked about window functions, and found one more, which was
not in the windows at all.

* The average of whole numbers was rounded rather than cut: `AVG` over 1
  and 2 answered 2 where a real server answers 1. The total was divided in
  doubles and the result rounded into the int column, and a plain `AVG` and
  one under `GROUP BY` did the same. The averages already in the surface
  came to 3.2, which rounds and cuts to the same 3; a running average put
  1.5 in front of it.

A sixth round asked about dates, which pySQLbridge had just found nine
faults in, and found four here.

* A date compared with text compared as text. A date is not a number to
  VBA's `IsNumeric`, so `WHERE hired > '2024-9-1'` set the date's printed
  form, 10/1/2024, against the string.
* Two dates compared the same way, so the first of October sorted before
  the ninth of September, in an `ORDER BY` and in `MIN` and `MAX`.
* `MAX` over a `VALUES` list of dates was declared as text, and went out as
  the text a date prints as.
* `'2024-10-01'` inserted into a temporary table's `datetime` column stayed
  text, and compared as text from then on.

Sixteen differences are left, listed with their reasons in
`tests/surface.py`, across 422 cases. Most agree on the value and differ on
how it is declared, which sqlcmd then renders differently: `SELECT 7.0 / 2`
is 3.5 either way and a real server prints 3.500000. Two are answered here
where a real server refuses, and one is refused with the same number and
the same words after a different share of the batch has run. An error is
compared on its number, its level, its words and where it falls among the
results; its state, its line and the way a server spells its own name are
not compared. The differences are asserted to differ rather than skipped,
so one quietly starting to agree is noticed too.

The types were settled by running each captured statement against a real SQL
Server on the same machine and diffing the declared types column by column
until none differed, and the empty views' column lists come from the server's
own log of an SMO run rather than from guesswork.

`tests/test_pump.py` drives Excel directly rather than through the harness.
The harness runs Excel hidden under a supervisor, and dispatching WM_TIMER
into VBA there takes the process down on the first tick, measured with a
callback that did nothing but increment a counter.

A test procedure's name has to be 31 characters or fewer; the harness will
not run a longer one.

The suite drives a real Excel, so anything else on the machine that ends
Excel ends the run. `Get-Process EXCEL | Stop-Process -Force` is a common
line in a test harness and it does not stop at its own instance: six runs in
sixty were lost while another project's harness was up, and none in forty
after it had gone. A session that fails to come up is started again, which
covers a kill at the beginning and not one in the middle, so a run that
fails with a connection refused or an RPC that is no longer there is worth
repeating on a quiet machine before it is read as a bug.

`scripts/compile_check.py` compiles the library and reports what the VBA
editor says, which turns a modal dialog at run time into a message. It takes
a source directory, so a candidate copy can be compiled without touching the
one in the repository.
`scripts/find_syntax_error.py` binary-searches for a compile error the editor
will not give a line for.

## Credit

The wire format comes from
[pySQLbridge](https://github.com/WilliamSmithEdward/pySQLbridge), which
measured it against SQL Server 2025 rather than reading the specification,
and the pump follows the one in
[ReDim](https://github.com/WilliamSmithEdward/ReDim).

## License

MIT.
