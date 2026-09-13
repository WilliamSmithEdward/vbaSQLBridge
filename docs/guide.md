# A guide to vbaSQLBridge

This is the long, gentle version of the README. It assumes you can open
Excel and have maybe written a macro once, and it assumes nothing else. If
you already know what TDS is, the README is faster.

- [What this actually is](#what-this-actually-is)
- [Five minutes to a working database](#five-minutes-to-a-working-database)
- [Connecting something to it](#connecting-something-to-it)
- [Serving your own data](#serving-your-own-data)
- [Views: a saved query that behaves like a table](#views-a-saved-query-that-behaves-like-a-table)
- [Writing back to the workbook](#writing-back-to-the-workbook)
- [Who is allowed in](#who-is-allowed-in)
- [The SQL you can write](#the-sql-you-can-write)
- [The demo workbook](#the-demo-workbook)
- [Putting it in a workbook of your own](#putting-it-in-a-workbook-of-your-own)
- [The API, one page](#the-api-one-page)
- [How fast it is](#how-fast-it-is)
- [When something goes wrong](#when-something-goes-wrong)
- [What it is not](#what-it-is-not)

## What this actually is

Every database speaks a language over the network. SQL Server's is called
TDS, and every tool that talks to SQL Server, SSMS, Power BI, Excel's own
Get Data, Python, a reporting tool nobody has touched since 2011, knows how
to speak it.

vbaSQLBridge is that language, written in VBA, running inside a workbook.
Press Start and the workbook opens a network port and waits. When a tool
connects, the workbook does everything a SQL Server would do: it negotiates
encryption, checks who you are against Windows, lists its tables, parses the
SQL you send, and sends back rows. The rows come from worksheets.

So the shape of it is:

```
   SSMS / Power BI / Python / another Excel
                  |
                  |  TDS over TCP, the same protocol SQL Server uses
                  v
        a workbook with two modules in it
                  |
                  v
            your worksheets
```

Three things follow from that, and they are the whole point:

**Nothing is installed.** No driver, no add-in, no reference to tick, no
admin rights. The sockets come from `ws2_32.dll`, the encryption from
`secur32.dll`, the certificate from `crypt32.dll`, all of which are already
on the machine because Windows is. You import two text files into a
workbook.

**The client does not know.** It connects to what it believes is SQL Server
and behaves accordingly. That is why a tool with no Excel connector can read
your spreadsheet: it never finds out it is reading a spreadsheet.

**The data stays live.** The worksheet is read again for every statement, so
if you type a number into a cell and the client asks again, it gets the new
number. Nothing is copied, imported or synchronised.

It is not a client. It does not connect to SQL Server and fetch data into
Excel; every tool in the world already does that. It is the other direction,
which nothing much does.

## Five minutes to a working database

1. Download `vbaSQLBridge.xlsm` from the
   [releases page](https://github.com/WilliamSmithEdward/vbaSQLBridge/releases).
2. Right-click the file, choose Properties, and tick **Unblock** if it is
   there. Windows marks anything downloaded, and a marked workbook will not
   run macros.
3. Open it and click **Enable Content**.
4. Press **Start** on the Server sheet.

The status line now says something like `listening on 127.0.0.1:14330,
serving 2 table(s)`, and the cell below it holds a connection string you can
copy.

The workbook is still usable. Type in cells, switch sheets, open another
file; the server runs on a Windows timer that wakes up a few times a second,
does whatever the network needs, and gets out of the way. You do not have a
macro stuck in a loop.

To check it from a command prompt, if you have `sqlcmd` (it comes with SSMS,
and Microsoft ships it on its own):

```bash
sqlcmd -S tcp:127.0.0.1,14330 -E -C -Q "SELECT * FROM people"
```

`-E` means log in as the Windows user you already are, and `-C` means trust
the certificate the workbook generated for itself when it started. You
should get the sample rows back.

Press **Stop** when you are done, or just close the workbook.

## Connecting something to it

The address is always the same shape: the machine, a comma, the port. Not a
colon. SQL Server uses a comma and so does this.

**SSMS.** Server name `tcp:127.0.0.1,14330`, Windows Authentication, and on
the Connection Properties tab tick **Trust server certificate**. Your
worksheets appear under Databases > vbaSQLBridge > Tables.

**Power BI or Excel's Get Data.** Choose SQL Server database, server
`127.0.0.1,14330`, database `vbaSQLBridge`. Pick DirectQuery if you want the
report to follow the workbook as it is edited rather than taking a copy.

**A connection string**, which is what the workbook puts in cell C8:

```
Provider=MSOLEDBSQL;Data Source=tcp:127.0.0.1,14330;
Initial Catalog=vbaSQLBridge;Integrated Security=SSPI;
TrustServerCertificate=yes;
```

**Python**, with `pyodbc`:

```python
import pyodbc

connection = pyodbc.connect(
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=tcp:127.0.0.1,14330;Database=vbaSQLBridge;"
    "Trusted_Connection=yes;TrustServerCertificate=yes;")
for row in connection.execute("SELECT name, score FROM people"):
    print(row.name, row.score)
```

**A real SQL Server**, as a linked server, which lets a stored procedure
join a workbook to a proper table. The README has the `sp_addlinkedserver`
call and the settings that matter.

Two things to know about the certificate. The workbook generates a
self-signed one at startup, so every client must be told to trust it, hence
`-C`, `TrustServerCertificate=yes`, or the tickbox. And the traffic really is
encrypted: the login is tunnelled through TLS the way a real server's is.

## Serving your own data

The Server sheet has a table list starting at row 13, three columns wide:

| Name | Sheet | Range or table |
| --- | --- | --- |
| `people` | `People` | *(empty)* |
| `q3` | `Sales` | `B4:H219` |
| `orders` | `Sales` | `Orders` |

The first column is the name the client will see. The second is the
worksheet. The third decides how much of it:

- **Empty** serves the whole sheet. The table grows as rows are typed under
  it, which is usually what you want.
- **A range address** like `B4:H219` serves exactly that rectangle. A row
  typed underneath is outside it and will not appear.
- **An Excel table's name** serves that table. It grows with the table, and
  it keeps its shape if somebody types something else on the same sheet.

Add a row, press **Reload**, and it is served. You do not have to stop
first, and a connected client sees the new table on its next query.

Every Excel table in the workbook is served anyway, under its own name,
without being listed. That is the fastest way to get something served: make
it an Excel table (select it, Ctrl+T) and press Reload.

The first row is the header. What if your sheet does not look like that?
It probably does not, and the project has opinions about all of it, tested
against twenty-two deliberately awful worksheets:

- The table need not start at A1. Blank rows and columns around it are
  ignored.
- A single cell sitting alone above a wide table is a title, not a
  one-column header. Any number of those can stack up.
- A column with a header and nothing under it is still a column.
- A column with no header gets called `Column3` rather than vanishing.
- Two columns with the same header get told apart.
- An Excel table's totals row is left out of the data, because it is a
  summary and not a row.
- A number typed as text is text. The whole column decides its own type: one
  value that is not a whole number makes the column a float, and one value
  that is not a number at all makes the column text. This is not
  negotiable at the row level, because the protocol declares a column's type
  once and then encodes every row against that declaration.

## Views: a saved query that behaves like a table

A view is a query with a name. Once it exists, a client can select from it,
join to it, group it and sort it as though it were a worksheet, and it never
has to know the query exists.

This matters more here than in a real database, because the client on the
other end is often something like Power BI that would rather point at a
table than hold SQL. A view lets the workbook keep the complicated part.

The demo workbook has a **views** sheet. Column B is the name, column C is
the statement:

| View | SQL |
| --- | --- |
| `order_lines` | `SELECT o.order_id, c.name AS customer, ... FROM orders o JOIN Customers c ON ...` |
| `monthly_revenue` | `SELECT YEAR(ordered) AS year, SUM(line_total) AS revenue FROM order_lines GROUP BY ...` |

Type a name and a statement, press **Sync views**, and it is served. From
then on `SELECT * FROM monthly_revenue` works from any client.

### How it works, and why it is worth knowing

Sync views does not copy your SQL anywhere. It creates an Excel **defined
name** called `sql_monthly_revenue` that points at the cell the statement is
in. Start and Reload walk the workbook's defined names, and any name
beginning `sql_` becomes a view called whatever follows the prefix.

So the convention is the whole mechanism, and you can use it without the
sheet. Put a statement in any cell anywhere, select the cell, type
`sql_whatever` into the Name Box to the left of the formula bar, press
Enter, press Reload. That is a view.

Because the name points at the cell rather than holding a copy, editing the
cell edits the view. And because a name can point at several cells, a
statement can be written down a column, one line per cell, and it is joined
back together top to bottom, which is how you write a forty-line query
without a cell you cannot read.

A few rules the server enforces:

- A view must be a read. `AddView` refuses anything that is not a SELECT,
  because a view that deletes rows when you look at it is not a view.
- Views can be built on views. `monthly_revenue` reads `order_lines`, which
  joins three sheets. Nesting is capped at 32 deep, so a view that reads
  itself is refused rather than running until Excel gives up.
- A view is run when it is read, not when it is defined. It follows the
  worksheets underneath it, so a view over live data stays live.
- Views appear in `INFORMATION_SCHEMA.VIEWS`, in `sys.views`, and in the
  table list a client shows, marked as views. That is how Power BI finds
  them.
- A bad view does not stop the server. Start reports `passed over the
  view(s) x, y` and serves everything else.

## Writing back to the workbook

A client can change the workbook. `INSERT`, `UPDATE`, `DELETE`, `MERGE` and
`TRUNCATE` all land in cells, and the sheet updates while you watch.

This is off unless you turn it on. The **Writes** cell on the Server sheet
must say `yes`; anything else serves everything read-only, and a client
trying to write gets the error a real read-only database gives.

What lands where:

- An `INSERT` adds rows under the last one. A sheet grows; a named range
  does not, and refuses instead of overwriting whatever is below it.
- An `UPDATE` writes into the cells that hold those values. Formatting
  survives; formulas do not, because a value written over a formula replaces
  it, the same as typing.
- A `DELETE` removes the rows and closes the gap.
- Writes go in blocks rather than cell by cell: a run of adjacent cells is
  one assignment, which is the difference between a second and a minute for
  a few thousand rows.
- `BEGIN TRANSACTION` works, and `ROLLBACK` really does put the rows back.

Excel's own undo does not know about any of this. A client that deletes a
thousand rows has deleted them. Keep a copy of anything you would miss, and
leave Writes at `no` until you need it.

## Who is allowed in

By default, only you: whoever is logged into Windows on this machine can
connect, through the same Windows authentication a real SQL Server uses, and
nobody else can.

The listener also binds to `127.0.0.1`, which means connections from this
machine only. That is a deliberate default. A bridge reachable from the
network is a database anyone on that network can read, and possibly write,
and that should be a decision somebody makes rather than one that happens.
Change the Address cell to `0.0.0.0` if you mean it.

**Add login** lets somebody in with a name and a password instead. It asks
for them in Windows' own credential dialog, so the password is masked as you
type it, and writes the name and a hash of the password into the login list
beside the settings. The password itself is not stored.

A row in that list can also say `env:SOMENAME`, which means the password is
whatever the `SOMENAME` environment variable holds. That keeps the secret
out of the file entirely, which matters because a workbook is a thing people
email to each other. A password typed into that column in plain text is
ignored and reported, on purpose.

## The SQL you can write

More than you would expect. The parser and the evaluator are written from
scratch in VBA, and 452 differential tests run each statement against a real
SQL Server and compare the answers, so where it disagrees with SQL Server
that is a bug with a test attached.

**Reading**

```sql
SELECT TOP 10 name, score FROM people WHERE score > 80 ORDER BY score DESC
SELECT DISTINCT country FROM Customers
SELECT * FROM orders ORDER BY ordered OFFSET 100 ROWS FETCH NEXT 25 ROWS ONLY
```

**Expressions, conditions and types**

`CASE`, `LIKE`, `IN`, `BETWEEN`, `IS NULL`, `ISNULL`, `COALESCE`, `NULLIF`,
`CAST`, `CONVERT`, `IIF`, `CHOOSE`, arithmetic, string concatenation with
`+`, `CONCAT` and `CONCAT_WS`, `GREATEST` and `LEAST`.

**Functions**

String: `LEFT`, `RIGHT`, `SUBSTRING`, `LEN`, `DATALENGTH`, `TRIM` (including
`TRIM(x FROM y)`), `LTRIM`, `RTRIM`, `REPLACE`, `TRANSLATE`, `UPPER`,
`LOWER`, `CHARINDEX`, `PATINDEX`, `STUFF`, `REVERSE`, `REPLICATE`, `SPACE`,
`STR`, `QUOTENAME`, `ASCII`, `UNICODE`, `CHAR`, `NCHAR`.

Date: `GETDATE`, `CURRENT_TIMESTAMP`, `SYSDATETIME`, `DATEADD`, `DATEDIFF`,
`DATEPART`, `DATENAME`, `YEAR`, `MONTH`, `DAY`, `EOMONTH`.

Maths: `ABS`, `ROUND`, `FLOOR`, `CEILING`, `POWER`, `SQUARE`, `SQRT`, `LOG`,
`LOG10`, `EXP`, `SIGN`, `PI`.

**Aggregates and grouping**

```sql
SELECT country, COUNT(*) AS n, SUM(total) AS revenue, AVG(total) AS mean
FROM orders
GROUP BY country
HAVING COUNT(*) > 5
ORDER BY SUM(total) DESC
```

`COUNT`, `COUNT(DISTINCT x)`, `SUM`, `MIN`, `MAX`, `AVG`, `STDEV`, `STDEVP`,
`VAR` and `VARP`, over the whole table, over a group or over a window. An
integer column's `AVG` comes back a whole number, because that is what SQL
Server does.

**Joins**

`INNER`, `LEFT`, `RIGHT`, `FULL OUTER`, `CROSS`, `CROSS APPLY` and `OUTER
APPLY`, on any condition, across as many tables as you like, including
joining a worksheet to a view to another worksheet.

**Subqueries and set operators**

Scalar subqueries, `IN (SELECT ...)`, `EXISTS`, `ANY`, `SOME`, `ALL`,
derived tables, `UNION`, `UNION ALL`, `INTERSECT`, `EXCEPT`, and common
table expressions including recursive ones with `MAXRECURSION`.

**Window functions**

```sql
SELECT order_id, channel, ordered,
       ROW_NUMBER() OVER (PARTITION BY channel ORDER BY ordered) AS n,
       SUM(quantity) OVER (ORDER BY ordered ROWS UNBOUNDED PRECEDING) AS running,
       AVG(revenue) OVER (ORDER BY day ROWS BETWEEN 6 PRECEDING AND CURRENT ROW)
           AS week_average
FROM orders
```

`ROW_NUMBER`, `RANK`, `DENSE_RANK`, `PERCENT_RANK`, `NTILE`, `LAG`, `LEAD`,
`FIRST_VALUE`, `LAST_VALUE`, and the aggregates over a frame.

**Procedural bits**

Variables with `DECLARE` and `SET`, `IF`, `WHILE`, `BEGIN TRY ... BEGIN
CATCH` with `ERROR_NUMBER()` and `ERROR_MESSAGE()`, `@@ROWCOUNT`, `@@ERROR`,
`@@TRANCOUNT`, `sp_executesql` with output parameters, and prepared
statements run again by handle.

**The catalog**, which is how a client fills its object list without being
told anything: `INFORMATION_SCHEMA.TABLES`, `.COLUMNS`, `.VIEWS`,
`.SCHEMATA`, `sys.tables`, `sys.views`, `sys.objects`, `sys.columns`,
`sp_tables`, `sp_columns`, and the OLE DB rowsets.

What is refused, clearly rather than silently: `PIVOT`, `FOR XML`, `GRANT`,
`REVOKE`, `DENY`, and anything else that would need a real engine underneath.
You get an error message saying so.

## The demo workbook

`vbaSQLBridge-demo.xlsm` on the releases page is the same engine with
something substantial to chew on: just under 60,000 rows over six sheets,
and eleven views on top of them.

| Sheet | What is in it |
| --- | --- |
| `orders` | 50,000 order rows: dates, customers, SKUs, quantities, prices, discounts, channels, a tenth unshipped |
| `Customers` | a 5,000-row Excel table: names, countries, segments, credit limits |
| `Products` | 1,000 products with SKUs, categories and prices |
| `metrics` | 1,096 daily rows, three years, for running totals and moving averages |
| `wide` | 500 rows of 200 columns, because a column list is its own kind of load |
| `awkward` | 8 rows of names with spaces, brackets, accents, leading zeros and blanks |

The views build on each other, which is the part worth looking at:
`order_lines` joins all three of orders, Customers and Products and computes
a line total; `monthly_revenue` groups that view by month; `top_customers`
ranks it; and `customer_health` joins `top_customers` back to Customers to
work out who is over their credit limit. That is a view over a view over a
join, and to a client it is four tables.

There is a dashboard sheet that asks the server the same questions Excel
formulas answer on the sheet, and prints both. They agree to the penny on
1,838,690,979.93 of revenue, which is the point: the SQL engine and Excel
are computing the same thing two different ways.

Some times from a run on an ordinary laptop, taken with the client's own
startup subtracted out so what is left is the server's work:

| Statement | Time |
| --- | --- |
| `SELECT COUNT(*) FROM orders` (50,000 rows) | 0.20s |
| `GROUP BY` over 50,000 rows | 0.50s |
| A join, 50,000 rows to 5,000 | 0.76s |
| `order_lines`, the three-way join | 2.5s |
| A window function over 50,000 rows | 2.0s |

Slower than SQL Server, obviously. Fast enough that a dashboard refreshes
while you wait rather than while you get coffee.

## Putting it in a workbook of your own

Two files, imported into any macro-enabled workbook:

- `src/SqlBridge.cls`
- `src/SqlBridgeHost.bas`

In the VBA editor (Alt+F11), right-click the project, File > Import File,
and pick each one. There is nothing else: no reference to add, no library to
register.

`SqlBridgeHost` exists for one reason. Excel will only deliver a `SetTimer`
callback to a standard module, and the timer is what keeps the server
running without a loop. Everything else lives in the class.

Then, from anywhere, a button or the Immediate window:

```vba
Dim server As SqlBridge

Set server = SqlBridgeStart(14330)
server.AddTable "people", Sheet1
server.AddTable "orders", Sheet2.ListObjects("Orders")
server.AddView "big", "SELECT * FROM orders WHERE total > 1000"
```

`SqlBridgeStart` returns immediately, having started the timer. To stop:

```vba
SqlBridgeStop
```

Port 1433 is the default and is what a client assumes when you give it no
port. If SQL Server is installed on the same machine it already has 1433, so
pick something else and say so in the connection string. `Auto_Close` stops
the server when the workbook closes, so a forgotten server does not keep a
port.

Copying the demo's Server sheet into your own workbook is often quicker than
writing the VBA: it is only the same calls behind buttons.

## The API, one page

Starting and stopping, from `SqlBridgeHost`:

| Call | What it does |
| --- | --- |
| `SqlBridgeStart(port, address)` | Starts the server and the timer, returns the `SqlBridge` |
| `SqlBridgeServer()` | The running server, or `Nothing` |
| `SqlBridgeStop()` | Stops it and frees the port |
| `SqlBridgeIsRunning()` | True while it is listening |

On the server object:

| Member | What it does |
| --- | --- |
| `AddTable name, source, [hasHeader]` | Serves a sheet, range, list object or array |
| `RemoveTable name` | Stops serving one |
| `TableNames` | What is served, comma separated |
| `AddView name, sql` | Serves a SELECT under a name |
| `RemoveView name`, `ClearViews` | Removes one, or all |
| `ViewNames` | The views, comma separated |
| `AddLogin name, password` | Lets a name and password in |
| `AddLoginHash name, hash` | The same, from a stored hash |
| `RemoveLogin`, `ClearLogins`, `LoginNames` | The rest of the login list |
| `HashPassword(password)` | The hash to store |
| `Database` | The catalog name a client sees, `vbaSQLBridge` by default |
| `ReadOnly` | True refuses every write |
| `RequireEncryption` | True refuses a client that will not use TLS |
| `IsListening`, `Port`, `ConnectionCount` | Where it stands |
| `LogText`, `ClearLog` | What it has been doing |
| `Shutdown` | Stops it |

And one for testing, which is how most of this project's own tests run:

```vba
Dim factory As SqlBridge
Dim engine As SqlBridge

Set factory = New SqlBridge
Set engine = factory.Offline()
```

That is the whole SQL surface with no socket, no certificate and no login,
so you can call it from a macro and get rows back without a network in the
way. The first object is only there to call the method on; the one it hands
back is the engine.

## How fast it is

Worth knowing where the time goes, because the shape is not obvious.

Starting a statement costs almost nothing: a few hundredths of a second
between the packet arriving and the answer going out, for a statement that
reads nothing, and the same whether the workbook holds six rows or sixty
thousand. Refreshing the sheet is not where the time goes either. The
cost is per row and per comparison. A 50,000-row scan is 0.2s. Each extra
condition in a `WHERE` costs about as much again. A join costs most, because
every row on one side has to find its partners on the other.

Three things make a query fast:

**Let a view do the joining once.** Reading `order_lines` is one pass. Two
clients each writing their own three-way join is two.

**Filter before joining.** A `WHERE` that removes 90% of the rows before a
join removes 90% of the join.

**Serve the range you mean.** Serving a whole sheet when you want 200 rows
means reading the whole sheet.

Recent work roughly halved join time by resolving column names to positions
once per statement instead of once per row, and by taking the commonest
comparison out from behind a dozen string tests. Scans and `WHERE` clauses
were already at the floor.

## When something goes wrong

| What you see | What it means |
| --- | --- |
| Macros disabled, nothing happens | Unblock the file in Properties, then Enable Content |
| `could not start: ...` in the status | Usually the port is taken. Try another one |
| Client hangs, then a timeout | Wrong port, or the address is `127.0.0.1` and you are connecting from another machine |
| A certificate error | The client was not told to trust it: `-C`, `TrustServerCertificate=yes`, or the tickbox |
| `Login failed for user` | Windows auth only, unless you added a login. Check the login list |
| `Invalid object name 'x'` | The name is not in the table list and is not an Excel table. Press Reload |
| A view is missing | Start said `passed over the view(s) x`: the statement did not parse, or is not a SELECT |
| Every write refused | The Writes cell is not `yes` |
| A column arrives as text | One value in it is not a number. The whole column follows its worst value |
| Numbers where a date should be | The cell holds a number formatted as a date rather than a date |
| Excel goes unresponsive | A long statement runs on Excel's own thread. It comes back |

The Server sheet keeps a log of the last 24 lines, updated by the **Refresh
log** button. The **Try statement** button runs a statement from the sheet
against the server without a client, which is the quickest way to find out
whether the problem is your SQL or your connection.

## What it is not

It is one workbook, answering one statement at a time, on Excel's thread. It
is not a database server, and where that matters it shows: no indexes, no
query planner, no concurrent execution, no durability beyond what saving the
workbook gives you. A hundred users pointed at it will queue.

What it is good for is the case where the data genuinely lives in a
worksheet and something else needs to read it as though it did not: a report
that must follow a workbook somebody edits daily, a tool with no Excel
connector, a linked server that joins a spreadsheet to a real table, a
prototype where the data is not in a database yet and may never be.

For the protocol details, the differential test setup, and the Windows API
work underneath, see the [README](../README.md) and
[windows-apis-from-vba.md](windows-apis-from-vba.md).
