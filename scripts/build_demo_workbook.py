"""Build dist/vbaSQLBridge-demo.xlsm, the same server under a real load.

    python scripts/build_demo_workbook.py

Where dist/vbaSQLBridge.xlsm is small enough to read at a glance, this one is
the workbook to point a client at when the question is whether any of this
holds up: fifty thousand order lines on a worksheet that grows, five
thousand customers and a thousand products as Excel tables, three years of
daily figures, a two-hundred-column sheet, and a sheet of the awkward names
and mixed types a real workbook grows on its own.

It also carries its views. A defined name beginning with sql_ whose cells
hold a SELECT is served as a view of that name, so the eleven statements on
the views sheet arrive at a client as views it can read, join and nest.

The build ends by opening what it made, starting it, and putting every view
and a battery of joins, groupings and window functions through sqlcmd, with
the time each took written onto the workbook's own notes sheet. A workbook
that cannot answer does not ship.
"""
import datetime
import random
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_workbook import (  # noqa: E402
    INK,
    MUTED,
    PANEL,
    REPO,
    RULE,
    VB_EXT_CLASS,
    VB_EXT_STD,
    XL_OPEN_XML_WORKBOOK_MACRO_ENABLED,
    build_control,
    label,
    strip_header,
)

OUT = REPO / "dist" / "vbaSQLBridge-demo.xlsm"

# Its own port, so this workbook and the small one can both be running.
PORT = 14331

# Everything below is drawn from one seed, so two builds of this file differ
# only where the code did.
SEED = 20260912

ORDER_LINES = 50_000
CUSTOMER_ROWS = 5_000
PRODUCT_ROWS = 1_000
METRIC_DAYS = 1_096            # three years, the last one a leap year
WIDE_ROWS = 500
WIDE_COLUMNS = 200

FIRST_DAY = datetime.date(2023, 1, 1)
LAST_DAY = datetime.date(2025, 12, 31)

FIRST_NAMES = [
    "Ada", "Grace", "Edsger", "Barbara", "Alan", "Katherine", "Donald",
    "Margaret", "Tony", "Radia", "Leslie", "Frances", "Ken", "Shafi",
    "Vint", "Anita", "Linus", "Jean", "Guido", "Karen",
]
LAST_NAMES = [
    "Lovelace", "Hopper", "Dijkstra", "Liskov", "Turing", "Johnson", "Knuth",
    "Hamilton", "Hoare", "Perlman", "Lamport", "Allen", "Thompson",
    "Goldwasser", "Cerf", "Borg", "Torvalds", "Bartik", "van Rossum",
    "Sparck Jones",
]
PLACES = [
    ("United Kingdom", ["London", "Manchester", "Bristol"]),
    ("United States", ["Seattle", "Austin", "Boston"]),
    ("Germany", ["Berlin", "Hamburg", "Munich"]),
    ("France", ["Paris", "Lyon", "Toulouse"]),
    ("Japan", ["Tokyo", "Osaka", "Sapporo"]),
    ("Brazil", ["Sao Paulo", "Recife", "Curitiba"]),
    ("India", ["Bengaluru", "Pune", "Chennai"]),
    ("Canada", ["Toronto", "Montreal", "Vancouver"]),
    ("Australia", ["Sydney", "Perth", "Hobart"]),
    ("Poland", ["Warsaw", "Krakow", "Gdansk"]),
    ("Kenya", ["Nairobi", "Mombasa", "Kisumu"]),
    ("Norway", ["Oslo", "Bergen", "Tromso"]),
]
SEGMENTS = ["enterprise", "mid-market", "small business"]
CATEGORIES = [
    "engines", "compilers", "storage", "networking", "instruments",
    "printing", "power", "spares",
]
PARTS = [
    "analytical engine", "difference engine", "compiler", "assembler",
    "punch card", "tape reel", "core store", "drum memory", "relay",
    "vacuum tube", "transistor", "capacitor", "oscilloscope", "plotter",
    "line printer", "ribbon", "power supply", "rectifier", "cable loom",
    "bearing",
]
CHANNELS = ["web", "phone", "partner", "store"]
NOTES = [
    None, None, None, None, None, None, None, None,
    "rush order",
    "customer asked for the invoice in euros",
    "cafe order, paid in cash",           # plain ASCII beside the ones below
    "naive estimate, revised twice",
    "\u5317\u4eac shipment, customs cleared",
    "caf\u00e9 order, r\u00e9sum\u00e9 attached",
    "part shipped short and the rest followed a week later, which is why "
    "the line reads twice on the statement and once here; the second "
    "delivery went out on the same order number rather than a new one, so "
    "the quantity below is the whole of it",
]

# name, what it is for, the statement itself
VIEWS = [
    ("order_lines",
     "every order line, with its customer and product",
     "SELECT o.order_id, o.ordered, o.customer_id, c.name AS customer,\n"
     "       c.country, c.segment, o.sku, p.name AS product, p.category,\n"
     "       o.quantity, o.unit_price, o.discount, o.channel,\n"
     "       o.quantity * o.unit_price * (1 - o.discount) AS line_total\n"
     "FROM orders o\n"
     "JOIN Customers c ON c.id = o.customer_id\n"
     "JOIN Products p ON p.sku = o.sku"),
    ("monthly_revenue",
     "a view over a view: what each month came to",
     "SELECT YEAR(ordered) AS year, MONTH(ordered) AS month,\n"
     "       COUNT(*) AS lines, SUM(line_total) AS revenue\n"
     "FROM order_lines\n"
     "GROUP BY YEAR(ordered), MONTH(ordered)"),
    ("top_customers",
     "the twenty-five biggest, by what they spent",
     "SELECT TOP 25 customer_id, customer, country,\n"
     "       COUNT(*) AS lines, SUM(line_total) AS revenue\n"
     "FROM order_lines\n"
     "GROUP BY customer_id, customer, country\n"
     "ORDER BY SUM(line_total) DESC"),
    ("customer_health",
     "the biggest customers against what they may owe",
     "SELECT t.customer, t.revenue, c.segment, c.credit_limit,\n"
     "       CASE WHEN t.revenue > c.credit_limit THEN 'over'\n"
     "            ELSE 'within' END AS standing\n"
     "FROM top_customers t\n"
     "JOIN Customers c ON c.id = t.customer_id"),
    ("category_spread",
     "how wide each category's line totals are",
     "SELECT category, COUNT(*) AS lines, AVG(line_total) AS average,\n"
     "       STDEV(line_total) AS spread, MIN(line_total) AS smallest,\n"
     "       MAX(line_total) AS largest\n"
     "FROM order_lines\n"
     "GROUP BY category"),
    ("running_revenue",
     "a running total and a seven-day average, over a window",
     "SELECT day, revenue,\n"
     "       SUM(revenue) OVER (ORDER BY day ROWS UNBOUNDED PRECEDING)\n"
     "           AS to_date,\n"
     "       AVG(revenue) OVER (ORDER BY day ROWS BETWEEN 6 PRECEDING\n"
     "           AND CURRENT ROW) AS week_average\n"
     "FROM metrics"),
    ("never_ordered",
     "products nobody has ordered",
     "SELECT p.sku, p.name, p.category, p.unit_price\n"
     "FROM Products p\n"
     "WHERE NOT EXISTS (SELECT 1 FROM orders o WHERE o.sku = p.sku)"),
    ("late_shipments",
     "what shipped late, what has not shipped at all",
     "SELECT order_id, ordered, shipped, channel,\n"
     "       CASE WHEN shipped IS NULL THEN 'not shipped'\n"
     "            WHEN DATEDIFF(day, ordered, shipped) > 7 THEN 'late'\n"
     "            ELSE 'on time' END AS state\n"
     "FROM orders"),
    ("wide_sample",
     "twenty rows of the two-hundred-column sheet",
     "SELECT TOP 20 id, m001, m002, m100, m199\n"
     "FROM wide\n"
     "ORDER BY m001 DESC"),
    ("awkward_read",
     "names with spaces, brackets and accents, read as they are",
     "SELECT [Order ID], [Unicode name], [Leading Zeros], [Mixed]\n"
     "FROM awkward\n"
     "WHERE [Order ID] IS NOT NULL"),
    ("served_objects",
     "the workbook describing itself",
     "SELECT TABLE_NAME, TABLE_TYPE\n"
     "FROM INFORMATION_SCHEMA.TABLES"),
]

# What the build puts through sqlcmd, and what it writes the time of onto
# the notes sheet. Each one has to answer for the workbook to ship.
CHECKS = [
    ("count the order lines", "SELECT COUNT(*) AS n FROM orders"),
    ("read a view", "SELECT COUNT(*) AS n FROM order_lines"),
    ("group a view by month",
     "SELECT TOP 3 year, month, revenue FROM monthly_revenue "
     "ORDER BY year, month"),
    ("the biggest customers",
     "SELECT TOP 3 customer, revenue FROM top_customers"),
    ("a view over a view",
     "SELECT TOP 3 customer, standing FROM customer_health"),
    ("spread by category", "SELECT category, spread FROM category_spread"),
    ("a running total",
     "SELECT TOP 3 day, to_date FROM running_revenue ORDER BY day"),
    ("what nobody ordered", "SELECT COUNT(*) AS n FROM never_ordered"),
    ("late and unshipped",
     "SELECT state, COUNT(*) AS n FROM late_shipments GROUP BY state"),
    ("two hundred columns", "SELECT COUNT(*) AS n FROM wide_sample"),
    ("awkward names", "SELECT COUNT(*) AS n FROM awkward_read"),
    ("the catalog", "SELECT COUNT(*) AS n FROM served_objects"),
    ("a join written by hand",
     "SELECT c.country, COUNT(*) AS lines FROM orders o "
     "JOIN Customers c ON c.id = o.customer_id GROUP BY c.country"),
    ("a window over the orders",
     "SELECT TOP 5 order_id, ordered, "
     "SUM(quantity) OVER (PARTITION BY channel ORDER BY ordered "
     "ROWS UNBOUNDED PRECEDING) AS running FROM orders ORDER BY ordered"),
    ("the views the catalog lists",
     "SELECT COUNT(*) AS n FROM INFORMATION_SCHEMA.VIEWS"),
]


def put(sheet, rows, first_row=1, chunk=5_000) -> None:
    """Write rows in blocks. A cell at a time is minutes; a block is seconds."""
    width = len(rows[0])
    for start in range(0, len(rows), chunk):
        block = rows[start:start + chunk]
        top = first_row + start
        target = sheet.Range(sheet.Cells(top, 1),
                             sheet.Cells(top + len(block) - 1, width))
        target.Value = tuple(tuple(row) for row in block)


def add_sheet(book, name):
    sheet = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
    sheet.Name = name
    return sheet


def a_day(rnd) -> datetime.date:
    span = (LAST_DAY - FIRST_DAY).days
    return FIRST_DAY + datetime.timedelta(days=rnd.randint(0, span))


# Excel counts days from here, and the cells are formatted as dates, so
# Excel hands a client a date and the server reads one.
EXCEL_EPOCH = datetime.date(1899, 12, 30)


def as_serial(day: datetime.date) -> int:
    """Excel's own day number for a day.

    Written as a datetime instead, COM converts it out of the local
    timezone on the way in, and every date in the workbook lands eight
    hours past midnight.
    """
    return (day - EXCEL_EPOCH).days


def customer_rows(rnd) -> list:
    rows = [["id", "name", "country", "city", "segment", "signed",
             "credit_limit", "active"]]
    for number in range(1, CUSTOMER_ROWS + 1):
        country, cities = rnd.choice(PLACES)
        rows.append([
            number,
            f"{rnd.choice(FIRST_NAMES)} {rnd.choice(LAST_NAMES)}",
            country,
            rnd.choice(cities),
            rnd.choice(SEGMENTS),
            as_serial(FIRST_DAY - datetime.timedelta(
                days=rnd.randint(0, 3_000))),
            round(rnd.choice([2_500, 10_000, 50_000, 250_000]) *
                  rnd.uniform(0.5, 1.5), 2),
            rnd.random() > 0.12,
        ])
    return rows


def product_rows(rnd) -> list:
    rows = [["sku", "name", "category", "unit_price", "in_stock",
             "reorder_level"]]
    for number in range(1, PRODUCT_ROWS + 1):
        rows.append([
            # Text with its zeros kept, which is the case a column typed as
            # a number gets wrong.
            f"P{number:05d}",
            f"{rnd.choice(PARTS)} mk{rnd.randint(1, 9)}",
            rnd.choice(CATEGORIES),
            round(rnd.uniform(0.5, 4_000), 2),
            rnd.random() > 0.2,
            rnd.choice([0, 5, 10, 25, 100]),
        ])
    return rows


def order_rows(rnd, products) -> list:
    rows = [["order_id", "customer_id", "sku", "ordered", "quantity",
             "unit_price", "discount", "shipped", "channel", "note"]]
    for number in range(1, ORDER_LINES + 1):
        ordered = a_day(rnd)
        # A tenth of the lines never shipped, which is what makes the NULLs
        # in this sheet worth having.
        shipped = None
        if rnd.random() > 0.1:
            shipped = as_serial(
                ordered + datetime.timedelta(days=rnd.randint(1, 21)))
        # Only the front of the catalogue is ever ordered, so the view that
        # asks what nobody has ordered has something to answer.
        product = products[rnd.randint(1, int((len(products) - 1) * 0.85))]
        rows.append([
            100_000 + number,
            rnd.randint(1, CUSTOMER_ROWS),
            product[0],
            as_serial(ordered),
            rnd.randint(1, 40),
            product[3],
            rnd.choice([0, 0, 0, 0.05, 0.1, 0.15, 0.2]),
            shipped,
            rnd.choice(CHANNELS),
            rnd.choice(NOTES),
        ])
    return rows


def metric_rows(rnd) -> list:
    rows = [["day", "visitors", "signups", "revenue", "refunds", "uptime"]]
    visitors = 4_000
    for offset in range(METRIC_DAYS):
        day = FIRST_DAY + datetime.timedelta(days=offset)
        # A walk rather than noise, so a running total and a seven-day
        # average have a shape to follow.
        visitors = max(500, int(visitors * rnd.uniform(0.96, 1.05)))
        weekend = day.weekday() >= 5
        signups = int(visitors * rnd.uniform(0.01, 0.04) * (0.6 if weekend else 1))
        revenue = round(signups * rnd.uniform(40, 260), 2)
        rows.append([
            as_serial(day),
            visitors,
            signups,
            revenue,
            round(revenue * rnd.uniform(0, 0.08), 2),
            round(rnd.uniform(0.978, 1.0), 4),
        ])
    return rows


def wide_rows(rnd) -> list:
    header = ["id"] + [f"m{number:03d}" for number in range(1, WIDE_COLUMNS)]
    rows = [header]
    for number in range(1, WIDE_ROWS + 1):
        rows.append([number] + [round(rnd.uniform(-100, 100), 3)
                                for _ in range(WIDE_COLUMNS - 1)])
    return rows


def awkward_rows() -> list:
    """The shapes a workbook grows on its own: names a client has to bracket,
    codes that are text however they read, a column of two minds, and gaps."""
    long_text = ("the note somebody pasted in from an email, which runs past "
                 "what any column was made for and keeps going for a while "
                 "yet, as these do; " * 4)
    rows = [["Order ID", "[Bracketed]", "Unicode name", "Leading Zeros",
             "Mixed", "Long Text", "Blank"]]
    samples = [
        (100_001, "square", "Ada Lovelace", "007", 42, long_text, None),
        (100_002, "square", "\u00c5sa Lindqvist", "0042", "forty-two",
         "short", None),
        (100_003, "curly", "\u5317\u4eac branch", "000", 0, "", None),
        (100_004, "curly", "Jos\u00e9 Mart\u00ednez", "0100", -7.5,
         long_text, None),
        (None, "none", "row with no id", "0007", None, None, None),
        (100_006, "square", "\u041c\u043e\u0441\u043a\u0432\u0430 office",
         "0808", True, "yes", None),
        (100_007, "curly", "caf\u00e9 crawl", "0909", "n/a", "maybe", None),
        (100_008, "square", "na\u00efve plan", "1000", 3.14159, "pi", None),
    ]
    rows.extend(list(sample) for sample in samples)
    return rows


def build_views_sheet(book, sheet) -> None:
    """The views, and a defined name for each one pointing at its statement.

    The sheet is the index and the names are the convention: a name that
    begins with sql_ is served as a view of what follows it. Sync views
    writes these names from this sheet, and this build writes them here so
    the workbook serves its views the first time it is started.
    """
    sheet.Columns("A").ColumnWidth = 2
    sheet.Columns("B").ColumnWidth = 22
    sheet.Columns("C").ColumnWidth = 96
    sheet.Columns("D").ColumnWidth = 44

    label(sheet, "B1", "Views", size=18)
    label(sheet, "B2",
          "A defined name beginning with sql_ is served as a view of what "
          "follows it, and this sheet is where those names point. Edit a "
          "statement and press Reload; add a row and press Sync views, "
          "which writes the name for it.",
          bold=False, colour=MUTED)

    label(sheet, "B4", "View")
    label(sheet, "C4", "Statement")
    label(sheet, "D4", "What it answers")
    sheet.Range("B4:D4").Interior.Color = PANEL
    sheet.Range("B4:D4").Borders.Color = RULE

    row = 5
    for name, about, sql in VIEWS:
        sheet.Cells(row, 2).Value = name
        sheet.Cells(row, 3).Value = sql
        sheet.Cells(row, 4).Value = about
        sheet.Cells(row, 3).WrapText = True
        sheet.Cells(row, 3).VerticalAlignment = -4160        # xlTop
        sheet.Cells(row, 4).VerticalAlignment = -4160
        sheet.Cells(row, 4).Font.Color = MUTED
        sheet.Rows(row).RowHeight = 14 * (sql.count("\n") + 1)
        book.Names.Add(
            Name="sql_" + name,
            RefersTo="=" + sheet.Name + "!" + sheet.Cells(row, 3).Address,
        )
        row += 1
    sheet.Range(sheet.Cells(5, 2), sheet.Cells(row - 1, 4)).Borders.Color = RULE


def build_notes_sheet(sheet) -> None:
    sheet.Columns("A").ColumnWidth = 2
    sheet.Columns("B").ColumnWidth = 26
    sheet.Columns("C").ColumnWidth = 104

    label(sheet, "B1", "What to try", size=18)
    label(sheet, "B2",
          "Press Start on the Server sheet, then point a client at the "
          "connection string it writes there.", bold=False, colour=MUTED)

    lines = [
        ("Served as sheets", "orders (50,000 lines), metrics (three years "
                             "of days), wide (200 columns), awkward"),
        ("Served as Excel tables", "Customers (5,000), Products (1,000, with "
                                   "a totals row Excel keeps and SQL does not "
                                   "see)"),
        ("Served as views", "eleven, on the views sheet; each is a defined "
                            "name beginning with sql_"),
        ("", ""),
        ("SSMS or sqlcmd", "sqlcmd -S tcp:127.0.0.1,%d -E -C -Q \"SELECT TOP 5 "
                           "* FROM top_customers\"" % PORT),
        ("ADO or Excel", "Provider=MSOLEDBSQL;Data Source=tcp:127.0.0.1,%d;"
                         "Initial Catalog=vbaSQLBridge;Integrated "
                         "Security=SSPI;TrustServerCertificate=yes;" % PORT),
        (".NET", "Server=tcp:127.0.0.1,%d;Database=vbaSQLBridge;Integrated "
                 "Security=SSPI;Encrypt=False;TrustServerCertificate=True"
                 % PORT),
        ("Power BI", "Get Data, SQL Server, 127.0.0.1,%d, database "
                     "vbaSQLBridge" % PORT),
        ("From a real SQL Server", "EXEC sp_addlinkedserver @server = "
                                   "'WORKBOOK', @srvproduct = '', @provider = "
                                   "'MSOLEDBSQL', @datasrc = "
                                   "'tcp:127.0.0.1,%d', @provstr = "
                                   "'TrustServerCertificate=yes'" % PORT),
        ("", ""),
        ("Read a view", "SELECT TOP 10 * FROM monthly_revenue ORDER BY year, "
                        "month"),
        ("Join across sources", "SELECT c.country, COUNT(*) AS lines FROM "
                                "orders o JOIN Customers c ON c.id = "
                                "o.customer_id GROUP BY c.country"),
        ("Window functions", "SELECT TOP 10 day, to_date, week_average FROM "
                             "running_revenue ORDER BY day DESC"),
        ("Spread", "SELECT * FROM category_spread ORDER BY spread DESC"),
        ("Awkward names", "SELECT TOP 5 [Order ID], [Unicode name] FROM "
                          "awkward"),
        ("What is served", "SELECT TABLE_NAME, TABLE_TYPE FROM "
                           "INFORMATION_SCHEMA.TABLES ORDER BY TABLE_TYPE"),
        ("Write to the sheet", "UPDATE orders SET channel = 'store' WHERE "
                               "order_id = 100001"),
        ("", ""),
        ("A view of your own", "Type a name and a SELECT on the views sheet, "
                               "press Sync views, and it is served."),
        ("A login of your own", "Press Add login on the Server sheet: only "
                                "the password's hash is kept in the file."),
    ]
    row = 4
    for left, right in lines:
        if left:
            label(sheet, f"B{row}", left, size=10)
            sheet.Cells(row, 3).Value = right
            sheet.Cells(row, 3).Font.Size = 10
        row += 1

    label(sheet, f"B{row + 1}", "How long it took", size=14)
    label(sheet, f"B{row + 2}",
          "Measured by the build, on the machine that made this file.",
          bold=False, size=9, colour=MUTED)
    return row + 3


def build_dashboard(sheet) -> None:
    """Excel's own answers beside the server's, over the same cells."""
    sheet.Columns("A").ColumnWidth = 2
    sheet.Columns("B").ColumnWidth = 34
    sheet.Columns("C").ColumnWidth = 22
    sheet.Columns("D").ColumnWidth = 60

    label(sheet, "B1", "The same numbers, both ways", size=18)
    label(sheet, "B2",
          "Excel works these out from the cells; a client works them out "
          "from the server. They are the same numbers.",
          bold=False, colour=MUTED)

    rows = [
        ["Order lines", f"=COUNT(orders!A2:A{ORDER_LINES + 1})",
         "SELECT COUNT(*) FROM orders"],
        ["Customers", f"=COUNTA(customers!B2:B{CUSTOMER_ROWS + 1})",
         "SELECT COUNT(*) FROM Customers"],
        ["Products", f"=COUNTA(products!A2:A{PRODUCT_ROWS + 1})",
         "SELECT COUNT(*) FROM Products"],
        ["Revenue",
         f"=ROUND(SUMPRODUCT(orders!E2:E{ORDER_LINES + 1},"
         f"orders!F2:F{ORDER_LINES + 1},"
         f"1-orders!G2:G{ORDER_LINES + 1}),2)",
         "SELECT SUM(line_total) FROM order_lines"],
        ["Unshipped lines",
         f"=COUNTBLANK(orders!H2:H{ORDER_LINES + 1})",
         "SELECT COUNT(*) FROM orders WHERE shipped IS NULL"],
        ["Busiest channel",
         f"=INDEX(orders!I2:I{ORDER_LINES + 1},MATCH(MAX(COUNTIF("
         f"orders!I2:I{ORDER_LINES + 1},orders!I2:I{ORDER_LINES + 1})),"
         f"COUNTIF(orders!I2:I{ORDER_LINES + 1},"
         f"orders!I2:I{ORDER_LINES + 1}),0))",
         "SELECT TOP 1 channel, COUNT(*) AS n FROM orders "
         "GROUP BY channel ORDER BY COUNT(*) DESC"],
        ["Days of figures", f"=COUNT(metrics!A2:A{METRIC_DAYS + 1})",
         "SELECT COUNT(*) FROM metrics"],
        ["Visitors, all days", f"=SUM(metrics!B2:B{METRIC_DAYS + 1})",
         "SELECT SUM(visitors) FROM metrics"],
    ]

    label(sheet, "B4", "What")
    label(sheet, "C4", "Excel")
    label(sheet, "D4", "The same question in SQL")
    sheet.Range("B4:D4").Interior.Color = PANEL
    sheet.Range("B4:D4").Borders.Color = RULE

    row = 5
    for what, formula, sql in rows:
        sheet.Cells(row, 2).Value = what
        sheet.Cells(row, 3).Formula = formula
        sheet.Cells(row, 4).Value = sql
        sheet.Cells(row, 4).Font.Color = MUTED
        row += 1
    sheet.Range(sheet.Cells(5, 2), sheet.Cells(row - 1, 4)).Borders.Color = RULE
    sheet.Range(f"C5:C{row - 1}").NumberFormat = "#,##0.00"


def build_data(book, rnd) -> None:
    customers = customer_rows(rnd)
    products = product_rows(rnd)

    sheet = add_sheet(book, "customers")
    put(sheet, customers)
    sheet.Range(f"F2:F{CUSTOMER_ROWS + 1}").NumberFormat = "yyyy-mm-dd"
    sheet.Range(f"G2:G{CUSTOMER_ROWS + 1}").NumberFormat = "#,##0.00"
    listed = sheet.ListObjects.Add(
        1, sheet.Range(f"A1:H{CUSTOMER_ROWS + 1}"), None, 1)
    listed.Name = "Customers"

    sheet = add_sheet(book, "products")
    put(sheet, products)
    sheet.Range(f"D2:D{PRODUCT_ROWS + 1}").NumberFormat = "#,##0.00"
    listed = sheet.ListObjects.Add(
        1, sheet.Range(f"A1:F{PRODUCT_ROWS + 1}"), None, 1)
    listed.Name = "Products"
    # A totals row, which Excel keeps under the table and a client never
    # sees as a record.
    listed.ShowTotals = True
    listed.ListColumns("unit_price").TotalsCalculation = 2      # average

    sheet = add_sheet(book, "orders")
    put(sheet, order_rows(rnd, products))
    sheet.Range(f"D2:D{ORDER_LINES + 1}").NumberFormat = "yyyy-mm-dd"
    sheet.Range(f"H2:H{ORDER_LINES + 1}").NumberFormat = "yyyy-mm-dd"
    sheet.Range(f"F2:F{ORDER_LINES + 1}").NumberFormat = "#,##0.00"
    sheet.Range(f"G2:G{ORDER_LINES + 1}").NumberFormat = "0%"

    sheet = add_sheet(book, "metrics")
    put(sheet, metric_rows(rnd))
    sheet.Range(f"A2:A{METRIC_DAYS + 1}").NumberFormat = "yyyy-mm-dd"
    sheet.Range(f"D2:E{METRIC_DAYS + 1}").NumberFormat = "#,##0.00"
    sheet.Range(f"F2:F{METRIC_DAYS + 1}").NumberFormat = "0.0000"

    sheet = add_sheet(book, "wide")
    put(sheet, wide_rows(rnd), chunk=250)

    sheet = add_sheet(book, "awkward")
    put(sheet, awkward_rows())
    sheet.Columns("A:G").AutoFit()
    sheet.Columns("F").ColumnWidth = 60


def build(excel) -> Path:
    rnd = random.Random(SEED)
    book = excel.Workbooks.Add()
    while book.Worksheets.Count > 1:
        book.Worksheets(book.Worksheets.Count).Delete()

    control = book.Worksheets(1)
    notes = add_sheet(book, "notes")
    views = add_sheet(book, "views")
    dashboard = add_sheet(book, "dashboard")
    build_data(book, rnd)

    build_notes_sheet(notes)
    build_views_sheet(book, views)
    build_dashboard(dashboard)
    build_control(
        control,
        tables=[
            ["orders", "orders", ""],
            ["metrics", "metrics", ""],
            ["wide", "wide", ""],
            ["awkward", "awkward", ""],
        ],
        port=PORT,
        statement="SELECT TOP 10 * FROM monthly_revenue ORDER BY year, month",
        hint="The four sheets above are served whole, so a row typed under "
             "the last one is served too. Customers and Products are Excel "
             "tables and are served under their own names without being "
             "listed. The views sheet holds eleven more, each a defined "
             "name beginning with sql_.",
    )

    project = book.VBProject
    for path, kind in (
        (REPO / "src" / "SqlBridge.cls", VB_EXT_CLASS),
        (REPO / "src" / "SqlBridgeHost.bas", VB_EXT_STD),
        (REPO / "demo" / "SqlBridgeApp.bas", VB_EXT_STD),
    ):
        component = project.VBComponents.Add(kind)
        component.Name = path.stem
        component.CodeModule.AddFromString(
            strip_header(path.read_text(encoding="utf-8")))

    control.Activate()
    excel.ActiveWindow.DisplayGridlines = False
    control.Range("B1").Select()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    target = OUT
    if OUT.exists():
        try:
            OUT.unlink()
        except PermissionError:
            target = OUT.with_name(OUT.stem + "-new" + OUT.suffix)
            print(f"{OUT.name} is open; building {target.name} instead")
            if target.exists():
                target.unlink()

    book.SaveAs(str(target), FileFormat=XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)
    book.Close(SaveChanges=False)
    return target


def verify(excel, built: Path) -> None:
    """Open what was built, start it, and put the whole battery through
    sqlcmd, writing what each one took onto the workbook's notes sheet."""
    sqlcmd = shutil.which("sqlcmd") or (
        r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\180"
        r"\Tools\Binn\SQLCMD.EXE")
    book = excel.Workbooks.Open(str(built))
    try:
        excel.Run("SqlBridgeApp.StartServer")
        time.sleep(0.5)
        status = book.Worksheets("Server").Range("C7").Value
        print("status:", status)
        if "listening" not in str(status):
            raise SystemExit(f"the workbook did not start: {status}")
        if "11 view(s)" not in str(status):
            raise SystemExit(f"the views were not served: {status}")

        timings = []
        if Path(sqlcmd).exists():
            for what, sql in CHECKS:
                started = time.perf_counter()
                done = subprocess.run(
                    [sqlcmd, "-S", f"tcp:127.0.0.1,{PORT}", "-E", "-C",
                     "-l", "30", "-h", "-1", "-W", "-Q", sql],
                    capture_output=True, text=True, timeout=600, check=False)
                took = time.perf_counter() - started
                answer = " ".join(line.strip() for line in
                                  done.stdout.splitlines()
                                  if line.strip() and
                                  not line.startswith("("))[:70]
                if done.returncode != 0 or "Msg " in done.stdout:
                    raise SystemExit(
                        f"{what} failed: {done.stdout}{done.stderr}")
                print(f"{took:6.2f}s  {what}: {answer}")
                timings.append((what, round(took, 2), answer))

            # Excel's own answers against the server's, over the same cells.
            # The dashboard sheet asks both ways and this checks they agree.
            def ask_one(sql: str) -> str:
                answered = subprocess.run(
                    [sqlcmd, "-S", f"tcp:127.0.0.1,{PORT}", "-E", "-C",
                     "-l", "30", "-h", "-1", "-W", "-Q", sql],
                    capture_output=True, text=True, timeout=600, check=False)
                for line in answered.stdout.splitlines():
                    line = line.strip()
                    if line and not line.startswith("("):
                        return line
                raise SystemExit(f"nothing came back for {sql}")

            excel.Calculate()
            dashboard = book.Worksheets("dashboard")
            by_excel = float(dashboard.Range("C5").Value)
            if int(ask_one("SELECT COUNT(*) AS n FROM orders")) != by_excel:
                raise SystemExit("Excel and the server disagree on the lines")
            by_excel = float(dashboard.Range("C8").Value)
            by_server = float(ask_one(
                "SELECT SUM(line_total) AS total FROM order_lines"))
            if abs(by_server - by_excel) > 0.5:
                raise SystemExit(
                    f"Excel says {by_excel} and the server {by_server}")
            print(f"both ways agree on {by_excel:,.2f} of revenue")

            # A write, and the same write undone: what ships has to be a
            # workbook a client can change.
            before = book.Worksheets("orders").Range("I2").Value
            subprocess.run(
                [sqlcmd, "-S", f"tcp:127.0.0.1,{PORT}", "-E", "-C", "-l", "30",
                 "-Q", "UPDATE orders SET channel = 'store' "
                       "WHERE order_id = 100001"],
                capture_output=True, text=True, timeout=600, check=False)
            if book.Worksheets("orders").Range("I2").Value != "store":
                raise SystemExit("the built workbook did not take a write")
            book.Worksheets("orders").Range("I2").Value = before

        notes = book.Worksheets("notes")
        row = notes.Cells(notes.Rows.Count, 2).End(-4162).Row + 2   # xlUp
        for what, took, answer in timings:
            notes.Cells(row, 2).Value = what
            notes.Cells(row, 3).Value = f"{took:.2f} seconds   {answer}"
            notes.Cells(row, 3).Font.Size = 10
            row += 1

        excel.Run("SqlBridgeApp.RefreshLog")
        excel.Run("SqlBridgeApp.StopServer")

        sheet = book.Worksheets("Server")
        sheet.Range(sheet.Cells(20, 2), sheet.Cells(44, 2)).ClearContents()
        sheet.Range("C7").Value = "stopped"
        sheet.Range("C8").Value = ""
        book.Worksheets("notes").Activate()
        book.Worksheets("Server").Activate()
        sheet.Range("B1").Select()
        book.Save()
    finally:
        book.Close(SaveChanges=False)


def main() -> int:
    """Both halves by default. Either on its own, as

        python scripts/build_demo_workbook.py build
        python scripts/build_demo_workbook.py verify

    because building fifty thousand lines takes long enough that a change
    to what is asked of them should not have to wait for it again.
    """
    step = sys.argv[1] if len(sys.argv) > 1 else "both"
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    try:
        built = OUT
        if step in ("both", "build"):
            started = time.perf_counter()
            built = build(excel)
            print(f"built {built} in {time.perf_counter() - started:.0f}s")
        excel.ScreenUpdating = True
        if step in ("both", "verify"):
            verify(excel, built)
            print("verified", built)
    finally:
        excel.Quit()
    return 0


if __name__ == "__main__":
    sys.exit(main())
