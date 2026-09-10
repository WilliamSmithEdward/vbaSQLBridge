"""Statements to put to both servers, across the SQL this claims to answer.

One list, run against a real SQL Server and against this one, and the two
answers compared character for character. The point is the cases nobody
thought to write a test for: NULL beside every operator, an aggregate over
nothing, an ordering with ties, a join that keeps rows on one side.

The data is the same as shapes.py so both can share the setup.
"""
from shapes import SETUP  # noqa: F401  (re-exported for the tests)

# What the two servers say differently, and why. Every one of these agrees
# on the value and differs on how it is declared, which sqlcmd then renders
# differently. Listed rather than quietly skipped: a case that starts
# agreeing, or a new one that stops, is something to look at either way.
DIFFERENT = {
    "float divide":
        "both 3.5. A real server types a decimal literal as decimal and "
        "renders 3.500000; this types it float.",
    "round":
        "both 3.46, rendered 3.460 against a decimal(n,3) declaration.",
    "sqrt":
        "both 4, rendered 4.0 by a float column a digit wider.",
    "concat null":
        "both NULL. 'a' + NULL + 'c' is nvarchar(3) on a real server, so "
        "sqlcmd prints the word NULL cut to three characters.",
    "case no else":
        "both none. The CASE holds only 'a', so a real server calls it "
        "nvarchar(1) and sqlcmd cuts the answer to one character.",
    "a rollback outside a transaction":
        "both refuse it, with the same number and the same words. A real "
        "server then carries on with the rest of the batch, where this "
        "stops at the statement that failed; and the header names the line "
        "from the top of the batch and spells the host as Windows does.",
    "datalength of text":
        "both the size of 'abc' as this server holds it. Every string here "
        "is nvarchar, so three characters are six bytes; a real server "
        "types a bare literal varchar and answers three.",
    "nullif then isnull":
        "both the word gone. NULLIF('a', 'a') is varchar(1) on a real "
        "server, so ISNULL takes that type and sqlcmd prints one character "
        "of it.",
    "case with no match and no else":
        "both NULL. The CASE holds only 'yes', so a real server calls it "
        "nvarchar(3) and sqlcmd cuts the word NULL to three characters.",
    "null concatenated":
        "both NULL. 'a' + NULL is nvarchar(1) on a real server, so sqlcmd "
        "prints the word NULL cut to one character. This declares it wide "
        "enough to say NULL.",
    "order by two of the same":
        "a real server refuses a column named twice in an ORDER BY. This "
        "sorts by it and then by it again, which decides nothing the first "
        "term had not decided already, so the rows come back in the order "
        "both would agree on.",
    "round half up":
        "both 3. A decimal literal is decimal on a real server, so it "
        "prints 3.0 against a decimal(2,1); it is float here.",
    "round half down":
        "both -3, and the same decimal declaration as round half up.",
    "round negative length":
        "both 1200, printed 1200.0000 against the decimal a real server "
        "keeps its four places in.",
    "big product":
        "both refuse it, with the same number and the same words. The "
        "header differs: a real server counts the line from the top of the "
        "batch, where this counts from the statement, and it spells its own "
        "host name in the case Windows gave it.",
    "eomonth":
        "both the last instant of February 2020. A real server types "
        "EOMONTH as date and prints 2020-02-29; there is no date type here "
        "that is not a datetime, so it prints the midnight as well.",
    "nullif same":
        "a real server types NULLIF(3, 3) tinyint and then refuses -1 as an "
        "overflow of it. This answers -1, which is the useful answer and "
        "not the one a real server gives.",
}

CASES = [
    # ---------------------------------------------------------- arithmetic
    ("plus", "SELECT 2 + 3 AS n"),
    ("minus", "SELECT 2 - 3 AS n"),
    ("times", "SELECT 2 * 3 AS n"),
    ("integer divide", "SELECT 7 / 2 AS n"),
    ("float divide", "SELECT 7.0 / 2 AS n"),
    ("modulo", "SELECT 7 % 2 AS n"),
    ("precedence", "SELECT 2 + 3 * 4 AS n"),
    ("brackets", "SELECT (2 + 3) * 4 AS n"),
    ("negative", "SELECT -5 + 2 AS n"),
    ("null plus", "SELECT 1 + NULL AS n"),
    ("bitwise and", "SELECT 12 & 10 AS n"),
    ("bitwise or", "SELECT 12 | 10 AS n"),
    ("bitwise xor", "SELECT 12 ^ 10 AS n"),

    # ---------------------------------------------------------- comparison
    ("equal", "SELECT id FROM #people WHERE id = 3"),
    ("not equal", "SELECT id FROM #people WHERE id <> 3 ORDER BY id"),
    ("greater", "SELECT id FROM #people WHERE id > 3 ORDER BY id"),
    ("at least", "SELECT id FROM #people WHERE id >= 3 ORDER BY id"),
    ("between", "SELECT id FROM #people WHERE id BETWEEN 2 AND 4 ORDER BY id"),
    ("not between",
     "SELECT id FROM #people WHERE id NOT BETWEEN 2 AND 4 ORDER BY id"),
    ("in list", "SELECT id FROM #people WHERE id IN (1, 3, 5) ORDER BY id"),
    ("not in list",
     "SELECT id FROM #people WHERE id NOT IN (1, 3, 5) ORDER BY id"),
    ("is null", "SELECT item FROM #orders WHERE owner IS NULL"),
    ("is not null",
     "SELECT owner FROM #orders WHERE owner IS NOT NULL ORDER BY owner"),
    ("equals null", "SELECT COUNT(*) AS n FROM #orders WHERE owner = NULL"),
    ("not equals null",
     "SELECT COUNT(*) AS n FROM #orders WHERE owner <> NULL"),
    ("in with null",
     "SELECT COUNT(*) AS n FROM #people WHERE id IN "
     "(SELECT owner FROM #orders)"),

    # ---------------------------------------------------------------- like
    ("like start", "SELECT name FROM #people WHERE name LIKE 'A%' ORDER BY id"),
    ("like end", "SELECT name FROM #people WHERE name LIKE '%a' ORDER BY id"),
    ("like middle",
     "SELECT name FROM #people WHERE name LIKE '%r%' ORDER BY id"),
    ("like one", "SELECT name FROM #people WHERE name LIKE '_da'"),
    ("not like",
     "SELECT name FROM #people WHERE name NOT LIKE 'A%' ORDER BY id"),
    ("like case", "SELECT name FROM #people WHERE name LIKE 'a%' ORDER BY id"),
    ("like escape",
     "SELECT CASE WHEN 'a%b' LIKE 'a!%b' ESCAPE '!' THEN 1 ELSE 0 END AS n"),
    ("like escape misses",
     "SELECT CASE WHEN 'axb' LIKE 'a!%b' ESCAPE '!' THEN 1 ELSE 0 END AS n"),

    # ----------------------------------------------------------- the logic
    ("null or true",
     "SELECT COUNT(*) AS n FROM #orders WHERE owner = 99 OR qty = 9"),
    ("true or null",
     "SELECT COUNT(*) AS n FROM #orders WHERE qty = 9 OR owner = 99"),
    ("null and false",
     "SELECT COUNT(*) AS n FROM #orders WHERE owner = 99 AND qty = 99"),
    ("false and null",
     "SELECT COUNT(*) AS n FROM #orders WHERE qty = 99 AND owner = 99"),
    ("not null",
     "SELECT COUNT(*) AS n FROM #orders WHERE NOT (owner = 99)"),
    ("nested logic",
     "SELECT name FROM #people WHERE (team = 'red' AND id > 1) "
     "OR name = 'Grace' ORDER BY id"),

    # ------------------------------------------------------------- strings
    ("length", "SELECT LEN('hello') AS n"),
    ("length trailing", "SELECT LEN('hello  ') AS n"),
    ("left", "SELECT LEFT('hello', 3) AS s"),
    ("right", "SELECT RIGHT('hello', 3) AS s"),
    ("substring", "SELECT SUBSTRING('hello', 2, 3) AS s"),
    ("upper", "SELECT UPPER('hello') AS s"),
    ("lower", "SELECT LOWER('HeLLo') AS s"),
    ("ltrim", "SELECT '[' + LTRIM('  x  ') + ']' AS s"),
    ("rtrim", "SELECT '[' + RTRIM('  x  ') + ']' AS s"),
    ("replace", "SELECT REPLACE('banana', 'an', 'X') AS s"),
    ("charindex", "SELECT CHARINDEX('an', 'banana') AS n"),
    ("charindex from", "SELECT CHARINDEX('an', 'banana', 3) AS n"),
    ("charindex missing", "SELECT CHARINDEX('z', 'banana') AS n"),
    ("reverse", "SELECT REVERSE('hello') AS s"),
    ("concat", "SELECT 'a' + 'b' + 'c' AS s"),
    ("concat null", "SELECT 'a' + NULL + 'c' AS s"),
    ("replicate", "SELECT REPLICATE('ab', 3) AS s"),
    ("string of number", "SELECT 'n=' + CAST(5 AS nvarchar(10)) AS s"),

    # ------------------------------------------------------------- numbers
    ("abs", "SELECT ABS(-7) AS n"),
    ("round", "SELECT ROUND(3.456, 2) AS n"),
    ("floor", "SELECT FLOOR(3.7) AS n"),
    ("ceiling", "SELECT CEILING(3.2) AS n"),
    ("power", "SELECT POWER(2, 10) AS n"),
    ("sqrt", "SELECT SQRT(16) AS n"),
    ("sign", "SELECT SIGN(-3) AS n"),

    # --------------------------------------------------------- the null ops
    ("isnull", "SELECT ISNULL(NULL, 'x') AS s"),
    ("isnull passes", "SELECT ISNULL('a', 'x') AS s"),
    ("coalesce", "SELECT COALESCE(NULL, NULL, 'third') AS s"),
    ("nullif same", "SELECT ISNULL(NULLIF(3, 3), -1) AS n"),
    ("nullif different", "SELECT NULLIF(3, 4) AS n"),

    # ---------------------------------------------------------------- case
    ("case searched",
     "SELECT CASE WHEN 1 = 1 THEN 'yes' ELSE 'no' END AS s"),
    ("case simple", "SELECT CASE 2 WHEN 1 THEN 'a' WHEN 2 THEN 'b' END AS s"),
    ("case no else", "SELECT ISNULL(CASE WHEN 1 = 0 THEN 'a' END, 'none') AS s"),
    ("case over rows",
     "SELECT name, CASE WHEN qty > 1 THEN 'many' ELSE 'one' END AS how "
     "FROM #orders o JOIN #people p ON p.id = o.owner ORDER BY o.item"),
    ("case with null",
     "SELECT CASE WHEN owner IS NULL THEN 'none' ELSE 'some' END AS s "
     "FROM #orders ORDER BY item"),

    # ------------------------------------------------------------- casting
    ("cast to int", "SELECT CAST('42' AS int) AS n"),
    ("cast float to int", "SELECT CAST(3.9 AS int) AS n"),
    ("cast to bit", "SELECT CAST(5 AS bit) AS n"),
    ("cast zero to bit", "SELECT CAST(0 AS bit) AS n"),
    ("cast to text", "SELECT CAST(42 AS nvarchar(10)) + '!' AS s"),
    ("convert", "SELECT CONVERT(int, '42') AS n"),

    # ---------------------------------------------------------- aggregates
    ("count star", "SELECT COUNT(*) AS n FROM #orders"),
    ("count column", "SELECT COUNT(owner) AS n FROM #orders"),
    ("sum", "SELECT SUM(qty) AS n FROM #orders"),
    ("min", "SELECT MIN(qty) AS n FROM #orders"),
    ("max", "SELECT MAX(qty) AS n FROM #orders"),
    ("avg integer", "SELECT AVG(qty) AS n FROM #orders"),
    ("sum of nothing",
     "SELECT ISNULL(SUM(qty), -1) AS n FROM #orders WHERE qty > 99"),
    ("count of nothing",
     "SELECT COUNT(*) AS n FROM #orders WHERE qty > 99"),
    ("min of nothing",
     "SELECT ISNULL(MIN(qty), -1) AS n FROM #orders WHERE qty > 99"),
    ("expression in aggregate", "SELECT SUM(qty * 2) AS n FROM #orders"),
    ("aggregate in expression", "SELECT SUM(qty) + 1 AS n FROM #orders"),

    # ------------------------------------------------------------- grouping
    ("group", "SELECT team, COUNT(*) AS n FROM #people GROUP BY team "
              "ORDER BY team"),
    ("group sum", "SELECT owner, SUM(qty) AS n FROM #orders "
                  "WHERE owner IS NOT NULL GROUP BY owner ORDER BY owner"),
    ("group having",
     "SELECT owner, COUNT(*) AS n FROM #orders WHERE owner IS NOT NULL "
     "GROUP BY owner HAVING COUNT(*) > 1 ORDER BY owner"),
    ("group by two",
     "SELECT team, id, COUNT(*) AS n FROM #people GROUP BY team, id "
     "ORDER BY team, id"),
    ("group with null",
     "SELECT owner, COUNT(*) AS n FROM #orders GROUP BY owner ORDER BY owner"),

    # ------------------------------------------------------------ ordering
    ("order asc", "SELECT name FROM #people ORDER BY name"),
    ("order desc", "SELECT name FROM #people ORDER BY name DESC"),
    ("order two keys",
     "SELECT team, name FROM #people ORDER BY team, name"),
    ("order by position", "SELECT name, id FROM #people ORDER BY 2"),
    ("order by alias", "SELECT id AS who FROM #people ORDER BY who"),
    ("order with nulls", "SELECT owner FROM #orders ORDER BY owner"),
    ("order with nulls desc",
     "SELECT owner FROM #orders ORDER BY owner DESC"),
    ("distinct", "SELECT DISTINCT team FROM #people ORDER BY team"),
    ("distinct with null",
     "SELECT DISTINCT owner FROM #orders ORDER BY owner"),
    ("top", "SELECT TOP 2 name FROM #people ORDER BY id"),
    ("top all", "SELECT TOP 99 name FROM #people ORDER BY id"),

    # --------------------------------------------------------------- joins
    ("inner join",
     "SELECT p.name, o.item FROM #people p JOIN #orders o "
     "ON o.owner = p.id ORDER BY p.id, o.item"),
    ("left join",
     "SELECT p.name, ISNULL(o.item, '-') AS item FROM #people p "
     "LEFT OUTER JOIN #orders o ON o.owner = p.id ORDER BY p.id, o.item"),
    ("join with condition",
     "SELECT p.name, o.item FROM #people p JOIN #orders o "
     "ON o.owner = p.id AND o.qty > 1 ORDER BY p.id, o.item"),
    ("comma join",
     "SELECT p.name, o.item FROM #people p, #orders o "
     "WHERE o.owner = p.id ORDER BY p.id, o.item"),
    ("join then group",
     "SELECT p.team, COUNT(*) AS n FROM #people p JOIN #orders o "
     "ON o.owner = p.id GROUP BY p.team ORDER BY p.team"),

    # ------------------------------------------------------------ set ops
    ("union", "SELECT team FROM #people UNION SELECT team FROM #people "
              "ORDER BY team"),
    ("union all",
     "SELECT id FROM #people UNION ALL SELECT owner FROM #orders "
     "WHERE owner IS NOT NULL ORDER BY id"),
    ("except",
     "SELECT id FROM #people EXCEPT SELECT owner FROM #orders ORDER BY id"),
    ("intersect",
     "SELECT id FROM #people INTERSECT SELECT owner FROM #orders ORDER BY id"),

    # --------------------------------------------------------- subqueries
    ("scalar subquery",
     "SELECT name FROM #people WHERE id = (SELECT MIN(owner) FROM #orders) "
     "ORDER BY id"),
    ("subquery in select",
     "SELECT name, (SELECT COUNT(*) FROM #orders o WHERE o.owner = p.id) AS n "
     "FROM #people p ORDER BY id"),
    ("derived table",
     "SELECT x.team, x.n FROM (SELECT team, COUNT(*) AS n FROM #people "
     "GROUP BY team) AS x ORDER BY x.team"),

    # ------------------------------------------------------ a second look
    ("nested subquery",
     "SELECT name FROM #people WHERE id IN (SELECT owner FROM #orders "
     "WHERE qty IN (SELECT MAX(qty) FROM #orders)) ORDER BY id"),
    ("subquery in having",
     "SELECT owner, COUNT(*) AS n FROM #orders WHERE owner IS NOT NULL "
     "GROUP BY owner HAVING COUNT(*) > (SELECT 1) ORDER BY owner"),
    ("exists in select",
     "SELECT name, CASE WHEN EXISTS (SELECT 1 FROM #orders o "
     "WHERE o.owner = p.id) THEN 1 ELSE 0 END AS has FROM #people p "
     "ORDER BY id"),
    ("distinct two columns",
     "SELECT DISTINCT team, id FROM #people ORDER BY team, id"),
    ("distinct then top",
     "SELECT DISTINCT TOP 1 team FROM #people ORDER BY team"),
    ("top with order",
     "SELECT TOP 2 name FROM #people ORDER BY name DESC"),
    ("order by expression",
     "SELECT name FROM #people ORDER BY LEN(name), name"),
    ("order by two directions",
     "SELECT team, id FROM #people ORDER BY team ASC, id DESC"),
    ("group by expression",
     "SELECT LEN(name) AS n, COUNT(*) AS c FROM #people "
     "GROUP BY LEN(name) ORDER BY n"),
    ("aggregate of text", "SELECT MIN(name) AS s, MAX(name) AS t FROM #people"),
    ("count distinct-ish",
     "SELECT COUNT(*) AS n FROM (SELECT DISTINCT team FROM #people) AS x"),

    # --------------------------------------------------------- more nulls
    ("null in aggregate", "SELECT SUM(owner) AS n FROM #orders"),
    ("null in group key",
     "SELECT ISNULL(CAST(owner AS nvarchar(10)), 'none') AS who, "
     "COUNT(*) AS n FROM #orders GROUP BY owner ORDER BY who"),
    ("null ordering two keys",
     "SELECT owner, item FROM #orders ORDER BY owner, item"),
    ("null in like", "SELECT COUNT(*) AS n FROM #orders WHERE item LIKE NULL"),
    ("null compared both ways",
     "SELECT COUNT(*) AS n FROM #orders WHERE NULL = NULL"),
    ("null in case",
     "SELECT COUNT(*) AS n FROM #orders WHERE "
     "CASE WHEN owner IS NULL THEN 1 ELSE 0 END = 1"),
    ("coalesce over columns",
     "SELECT COALESCE(CAST(owner AS nvarchar(10)), item) AS s FROM #orders "
     "ORDER BY item"),

    # --------------------------------------------------- more comparisons
    ("negative modulo", "SELECT -7 % 3 AS n"),
    ("divide negative", "SELECT -7 / 2 AS n"),
    ("compare text to number",
     "SELECT COUNT(*) AS n FROM #people WHERE name > 'B'"),
    ("compare with case",
     "SELECT COUNT(*) AS n FROM #people WHERE team = 'RED'"),
    ("string ordering", "SELECT name FROM #people ORDER BY name"),
    ("like a range",
     "SELECT name FROM #people WHERE name LIKE '[AB]%' ORDER BY id"),
    ("like not a range",
     "SELECT name FROM #people WHERE name LIKE '[^AB]%' ORDER BY id"),

    # -------------------------------------------------------- more joins
    ("self join",
     "SELECT a.name, b.name AS other FROM #people a JOIN #people b "
     "ON a.team = b.team AND a.id < b.id ORDER BY a.id, b.id"),
    ("three tables",
     "SELECT p.name, o.item FROM #people p JOIN #orders o ON o.owner = p.id "
     "JOIN #people q ON q.id = p.id ORDER BY p.id, o.item"),
    ("left join keeps all",
     "SELECT COUNT(*) AS n FROM #people p LEFT OUTER JOIN #orders o "
     "ON o.owner = p.id"),
    ("left join misses",
     "SELECT p.name FROM #people p LEFT OUTER JOIN #orders o "
     "ON o.owner = p.id WHERE o.owner IS NULL ORDER BY p.id"),
    ("join on an expression",
     "SELECT COUNT(*) AS n FROM #people p JOIN #orders o "
     "ON o.owner = p.id + 0"),

    # ------------------------------------------------------- more set ops
    ("union three",
     "SELECT 1 AS n UNION SELECT 2 UNION SELECT 3 ORDER BY n"),
    ("union orders by position",
     "SELECT id FROM #people UNION SELECT owner FROM #orders ORDER BY 1"),
    ("union of expressions",
     "SELECT id * 10 AS n FROM #people UNION SELECT qty FROM #orders "
     "ORDER BY n"),
    ("union inside a derived table",
     "SELECT name FROM (SELECT 'x' AS name UNION ALL SELECT 'y') t"),
    ("union inside a derived table counted",
     "SELECT COUNT(*) AS n FROM "
     "(SELECT 'x' AS name UNION ALL SELECT 'x') t"),
    ("union inside a subquery",
     "SELECT id FROM #people WHERE id IN "
     "(SELECT 1 UNION ALL SELECT 2) ORDER BY id"),
    ("union inside an exists",
     "SELECT id FROM #people WHERE EXISTS "
     "(SELECT 1 WHERE 1 = 0 UNION ALL SELECT 1) ORDER BY id"),
    ("except inside a derived table",
     "SELECT id FROM (SELECT id FROM #people EXCEPT SELECT 3) t ORDER BY id"),

    ("except twice",
     "SELECT id FROM #people EXCEPT SELECT owner FROM #orders "
     "EXCEPT SELECT 5 ORDER BY id"),

    # ------------------------------------------------------ transactions
    ("transaction commits",
     "BEGIN TRANSACTION; UPDATE #people SET team = 'x'; COMMIT; "
     "SELECT team FROM #people ORDER BY id"),
    ("transaction rolls back",
     "BEGIN TRANSACTION; UPDATE #people SET team = 'x'; ROLLBACK; "
     "SELECT team FROM #people ORDER BY id"),
    ("rollback puts deleted rows back",
     "BEGIN TRANSACTION; DELETE FROM #people; ROLLBACK; "
     "SELECT id FROM #people ORDER BY id"),
    ("rollback takes an insert away",
     "BEGIN TRANSACTION; INSERT INTO #people (id, name) VALUES (9, 'K'); "
     "ROLLBACK; SELECT id FROM #people ORDER BY id"),
    ("rollback over two tables",
     "BEGIN TRANSACTION; UPDATE #people SET team = 'x'; "
     "DELETE FROM #orders; ROLLBACK; "
     "SELECT team FROM #people ORDER BY id; "
     "SELECT COUNT(*) AS n FROM #orders"),
    ("rollback then more work",
     "BEGIN TRANSACTION; UPDATE #people SET team = 'x'; ROLLBACK; "
     "UPDATE #people SET team = 'y' WHERE id = 1; "
     "SELECT team FROM #people ORDER BY id"),
    ("commit keeps it",
     "BEGIN TRANSACTION; DELETE FROM #people WHERE id > 3; COMMIT; "
     "SELECT id FROM #people ORDER BY id"),
    ("a rollback outside a transaction",
     "UPDATE #people SET team = 'x'; ROLLBACK; "
     "SELECT team FROM #people ORDER BY id"),

    # ------------------------------------------- what a batch remembers
    # A client asks how many rows the last statement touched far more often
    # than it asks for the rows.
    ("rowcount after a read",
     "SELECT id FROM #people WHERE id < 3; SELECT @@ROWCOUNT AS n"),
    ("rowcount after an update",
     "UPDATE #people SET team = 'x'; SELECT @@ROWCOUNT AS n"),
    ("rowcount after a delete",
     "DELETE FROM #people WHERE team = 'red'; SELECT @@ROWCOUNT AS n"),
    ("rowcount after an insert",
     "INSERT INTO #people (id, name) VALUES (9, 'K'); "
     "SELECT @@ROWCOUNT AS n"),
    ("rowcount after nothing matched",
     "UPDATE #people SET team = 'x' WHERE id = 99; SELECT @@ROWCOUNT AS n"),
    ("rowcount is reset by the read of it",
     "SELECT id FROM #people; SELECT @@ROWCOUNT AS a; SELECT @@ROWCOUNT AS b"),

    # ------------------------------------------------------- variables
    ("declare and set",
     "DECLARE @n int; SET @n = 7; SELECT @n AS n"),
    ("declare with a value",
     "DECLARE @n int = 7; SELECT @n AS n"),
    ("a variable in a where",
     "DECLARE @n int; SET @n = 2; "
     "SELECT id FROM #people WHERE id > @n ORDER BY id"),
    ("a variable from a read",
     "DECLARE @n int; SELECT @n = COUNT(*) FROM #people; SELECT @n AS n"),
    ("arithmetic on a variable",
     "DECLARE @n int; SET @n = 2; SET @n = @n * 3; SELECT @n AS n"),
    ("a text variable",
     "DECLARE @s nvarchar(10); SET @s = 'ab'; SELECT @s + 'c' AS n"),
    ("an unset variable is null",
     "DECLARE @n int; SELECT ISNULL(@n, -1) AS n"),

    # ------------------------------------------------------ branching
    ("if that runs",
     "IF 1 = 1 SELECT 'yes' AS n"),
    ("if that does not run",
     "IF 1 = 0 SELECT 'yes' AS n ELSE SELECT 'no' AS n"),
    ("if over a count",
     "IF (SELECT COUNT(*) FROM #people) > 3 SELECT 'many' AS n "
     "ELSE SELECT 'few' AS n"),
    ("a while loop",
     "DECLARE @n int; SET @n = 0; WHILE @n < 3 SET @n = @n + 1; "
     "SELECT @n AS n"),

    # ------------------------------------------------ more of the writes
    ("update from a read",
     "UPDATE #people SET team = (SELECT MIN(item) FROM #orders) "
     "WHERE id = 1; SELECT id, team FROM #people ORDER BY id"),
    ("delete by a read",
     "DELETE FROM #people WHERE id IN (SELECT owner FROM #orders); "
     "SELECT id FROM #people ORDER BY id"),
    ("update by a correlated read",
     "UPDATE #people SET team = 'busy' WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = #people.id); "
     "SELECT id, team FROM #people ORDER BY id"),
    ("insert an expression",
     "INSERT INTO #people (id, name) SELECT id + 100, UPPER(name) "
     "FROM #people WHERE id = 1; SELECT id, name FROM #people ORDER BY id"),
    ("insert nothing",
     "INSERT INTO #people (id, name) SELECT id, name FROM #people "
     "WHERE id = 99; SELECT COUNT(*) AS n FROM #people"),
    ("update every row twice over",
     "UPDATE #orders SET qty = qty * 2; UPDATE #orders SET qty = qty + 1; "
     "SELECT item, qty FROM #orders ORDER BY item"),

    # --------------------------------------------------- a few functions
    ("iif", "SELECT IIF(1 = 1, 'yes', 'no') AS n"),
    ("iif over a null", "SELECT IIF(NULL = 1, 'yes', 'no') AS n"),
    ("choose", "SELECT CHOOSE(2, 'a', 'b', 'c') AS n"),
    ("nullif then isnull", "SELECT ISNULL(NULLIF('a', 'a'), 'gone') AS n"),
    ("concat of three", "SELECT CONCAT('a', NULL, 'c') AS n"),
    ("stuff", "SELECT STUFF('abcdef', 2, 3, 'XY') AS n"),
    ("a unicode literal", "SELECT N'caf' + NCHAR(233) AS n"),
    ("datalength of text", "SELECT DATALENGTH('abc') AS n"),
    ("a long string", "SELECT LEN(REPLICATE('ab', 200)) AS n"),

    # --------------------------------------------------------- ordering
    ("offset and fetch",
     "SELECT id FROM #people ORDER BY id OFFSET 1 ROWS FETCH NEXT 2 ROWS ONLY"),
    ("order by a case",
     "SELECT name FROM #people "
     "ORDER BY CASE WHEN team = 'red' THEN 0 ELSE 1 END, name"),
    ("aggregate over a join",
     "SELECT p.team, COUNT(*) AS n FROM #people p JOIN #orders o "
     "ON o.owner = p.id GROUP BY p.team ORDER BY p.team"),

    # ----------------------------------------------------- NULL and the set
    # NOT IN over a list holding NULL is never true, because the value might
    # be the NULL. The classic way to lose every row and not notice.
    ("not in with a null", "SELECT id FROM #people WHERE id NOT IN (1, NULL) "
                           "ORDER BY id"),
    ("in with a null present", "SELECT id FROM #people WHERE id IN (1, NULL) "
                              "ORDER BY id"),
    ("not in a read with nulls",
     "SELECT id FROM #people WHERE id NOT IN (SELECT owner FROM #orders) "
     "ORDER BY id"),
    ("in a read with nulls",
     "SELECT id FROM #people WHERE id IN (SELECT owner FROM #orders) "
     "ORDER BY id"),
    ("not exists over nulls",
     "SELECT p.id FROM #people p WHERE NOT EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = p.id) ORDER BY p.id"),

    # ------------------------------------------------------ ANY, ALL, SOME
    ("greater than any",
     "SELECT id FROM #people WHERE id > ANY (SELECT owner FROM #orders) "
     "ORDER BY id"),
    ("greater than all",
     "SELECT id FROM #people WHERE id > ALL (SELECT qty FROM #orders) "
     "ORDER BY id"),
    ("equal to some",
     "SELECT id FROM #people WHERE id = SOME (SELECT owner FROM #orders) "
     "ORDER BY id"),

    # ---------------------------------------------- comparing across types
    ("text equals number",
     "SELECT CASE WHEN '10' = 10 THEN 'same' ELSE 'not' END AS n"),
    ("text orders as number",
     "SELECT CASE WHEN '10' > 9 THEN 'more' ELSE 'less' END AS n"),
    ("text orders as text",
     "SELECT CASE WHEN '10' > '9' THEN 'more' ELSE 'less' END AS n"),
    ("bit equals number",
     "SELECT CASE WHEN CAST(1 AS bit) = 1 THEN 'same' ELSE 'not' END AS n"),
    ("number in a text list",
     "SELECT CASE WHEN 2 IN ('1', '2') THEN 'in' ELSE 'out' END AS n"),

    # ------------------------------------------------------- conversions
    ("cast text with spaces", "SELECT CAST('  7  ' AS int) AS n"),
    ("cast text with a sign", "SELECT CAST('-7' AS int) AS n"),
    ("cast float rounds", "SELECT CAST(2.7 AS int) AS n"),
    ("cast negative float", "SELECT CAST(-2.7 AS int) AS n"),
    ("cast to bit of two", "SELECT CAST(2 AS bit) AS n"),
    ("cast to bit of text", "SELECT CAST('1' AS bit) AS n"),
    ("cast int to text", "SELECT '[' + CAST(7 AS nvarchar(10)) + ']' AS n"),
    ("cast to a narrow text", "SELECT CAST('abcdef' AS nvarchar(3)) AS n"),
    ("convert with a style", "SELECT CONVERT(int, '7') AS n"),

    # ------------------------------------------------------------- joins
    ("right join",
     "SELECT p.name, o.item FROM #orders o RIGHT JOIN #people p "
     "ON o.owner = p.id ORDER BY p.id, o.item"),
    ("full join",
     "SELECT p.id, o.item FROM #people p FULL JOIN #orders o "
     "ON o.owner = p.id ORDER BY p.id, o.item"),
    ("join on two columns",
     "SELECT p.name FROM #people p JOIN #orders o ON o.owner = p.id "
     "AND o.item = p.name ORDER BY p.name"),
    ("join on an or",
     "SELECT DISTINCT p.name FROM #people p JOIN #orders o "
     "ON o.owner = p.id OR o.qty = p.id ORDER BY p.name"),
    ("left join then where",
     "SELECT p.id FROM #people p LEFT JOIN #orders o ON o.owner = p.id "
     "WHERE o.item IS NULL ORDER BY p.id"),
    ("join a derived table",
     "SELECT p.name FROM #people p JOIN "
     "(SELECT owner, COUNT(*) AS n FROM #orders GROUP BY owner) o "
     "ON o.owner = p.id ORDER BY p.name"),

    # ------------------------------------------------- grouping and having
    ("group by an expression",
     "SELECT LEN(name) AS n, COUNT(*) AS c FROM #people GROUP BY LEN(name) "
     "ORDER BY n"),
    ("having without a group",
     "SELECT COUNT(*) AS n FROM #people HAVING COUNT(*) > 1"),
    ("having that keeps nothing",
     "SELECT team, COUNT(*) AS n FROM #people GROUP BY team "
     "HAVING COUNT(*) > 90 ORDER BY team"),
    ("group by two with a null",
     "SELECT owner, qty, COUNT(*) AS n FROM #orders GROUP BY owner, qty "
     "ORDER BY owner, qty"),
    ("count over a group of nulls",
     "SELECT owner, COUNT(owner) AS n FROM #orders GROUP BY owner "
     "ORDER BY owner"),

    # ----------------------------------------------------------- ordering
    ("order by a column not selected",
     "SELECT name FROM #people ORDER BY id DESC"),
    ("order by two of the same",
     "SELECT team FROM #people ORDER BY team, team"),
    ("top with an order",
     "SELECT TOP 2 name FROM #people ORDER BY name"),
    ("top more than there are",
     "SELECT TOP 99 id FROM #people ORDER BY id"),
    ("distinct then order",
     "SELECT DISTINCT team FROM #people ORDER BY team DESC"),

    # ------------------------------------------------------------- names
    ("a bracketed name", "SELECT [id] FROM #people ORDER BY [id]"),
    ("a quoted alias", "SELECT id AS [the id] FROM #people ORDER BY id"),
    ("keywords in any case", "select Id from #People order by ID"),
    ("a table alias twice",
     "SELECT a.id FROM #people a JOIN #people b ON a.id = b.id "
     "ORDER BY a.id"),

    # ------------------------------------------------------ CASE and LIKE
    ("case with a null test",
     "SELECT CASE WHEN NULL = 1 THEN 'yes' ELSE 'no' END AS n"),
    ("simple case over a null",
     "SELECT CASE owner WHEN 1 THEN 'one' ELSE 'other' END AS n "
     "FROM #orders ORDER BY item"),
    ("case with no match and no else",
     "SELECT CASE WHEN 1 = 0 THEN 'yes' END AS n"),
    ("like an underscore", "SELECT name FROM #people WHERE name LIKE '_da' "
                           "ORDER BY name"),
    ("like a null pattern",
     "SELECT CASE WHEN 'a' LIKE NULL THEN 'yes' ELSE 'no' END AS n"),
    ("like an escaped wildcard",
     "SELECT CASE WHEN 'a%b' LIKE 'a!%b' ESCAPE '!' THEN 'yes' ELSE 'no' "
     "END AS n"),

    # --------------------------------------------------- arithmetic on null
    ("null plus one", "SELECT NULL + 1 AS n"),
    ("null times zero", "SELECT NULL * 0 AS n"),
    ("null concatenated", "SELECT 'a' + NULL AS n"),
    ("sum with a null in it", "SELECT SUM(qty) AS n FROM #orders"),

    # ---------------------------------------------------------------- dates
    # A whole family of functions with no case here until now. Literals
    # throughout: GETDATE cannot be compared between two servers.
    ("year", "SELECT YEAR('2020-03-04') AS n"),
    ("month", "SELECT MONTH('2020-03-04') AS n"),
    ("day", "SELECT DAY('2020-03-04') AS n"),
    ("datepart quarter", "SELECT DATEPART(quarter, '2020-08-04') AS n"),
    ("datepart dayofyear", "SELECT DATEPART(dayofyear, '2020-03-01') AS n"),
    ("datediff day", "SELECT DATEDIFF(day, '2020-01-01', '2020-03-01') AS n"),
    ("datediff month", "SELECT DATEDIFF(month, '2020-01-31', '2020-02-01') AS n"),
    ("datediff year", "SELECT DATEDIFF(year, '2019-12-31', '2020-01-01') AS n"),
    ("datediff backwards", "SELECT DATEDIFF(day, '2020-03-01', '2020-01-01') AS n"),
    ("dateadd day", "SELECT DATEADD(day, 30, '2020-02-01') AS n"),
    ("dateadd month over end", "SELECT DATEADD(month, 1, '2020-01-31') AS n"),
    ("dateadd negative", "SELECT DATEADD(day, -1, '2020-03-01') AS n"),
    ("eomonth", "SELECT EOMONTH('2020-02-07') AS n"),
    ("leap day", "SELECT DAY(EOMONTH('2020-02-07')) AS n"),
    ("year of null", "SELECT YEAR(NULL) AS n"),
    ("datediff with null", "SELECT DATEDIFF(day, NULL, '2020-01-01') AS n"),

    # ------------------------------------------------- padded comparison
    # A char comparison pads the shorter side, so a trailing space does not
    # make two strings different. LIKE does not pad, which is the one place
    # the two rules can be told apart.
    ("trailing space equal",
     "SELECT CASE WHEN 'a' = 'a ' THEN 'same' ELSE 'not' END AS n"),
    ("trailing space in",
     "SELECT CASE WHEN 'a ' IN ('a') THEN 'in' ELSE 'out' END AS n"),
    ("trailing space like",
     "SELECT CASE WHEN 'a ' LIKE 'a' THEN 'like' ELSE 'not' END AS n"),
    ("leading space equal",
     "SELECT CASE WHEN 'a' = ' a' THEN 'same' ELSE 'not' END AS n"),
    ("trailing space order",
     "SELECT name FROM (SELECT 'a ' AS name UNION ALL SELECT 'a') t "
     "ORDER BY name, LEN(name)"),

    # ------------------------------------------------------ number edges
    ("round negative length", "SELECT ROUND(1234.5678, -2) AS n"),
    ("round half up", "SELECT ROUND(2.5, 0) AS n"),
    ("round half down", "SELECT ROUND(-2.5, 0) AS n"),
    ("ceiling negative", "SELECT CEILING(-2.5) AS n"),
    ("floor negative", "SELECT FLOOR(-2.5) AS n"),
    ("power fraction", "SELECT POWER(9, 0.5) AS n"),
    ("power negative", "SELECT POWER(2, -1) AS n"),
    ("int and float compare",
     "SELECT CASE WHEN 1 = 1.0 THEN 'same' ELSE 'not' END AS n"),
    ("int plus float", "SELECT 1 + 0.5 AS n"),
    ("big product", "SELECT 100000 * 100000 AS n"),

    # ------------------------------------------------------ string edges
    ("substring past end", "SELECT SUBSTRING('abc', 2, 10) AS n"),
    ("substring zero start", "SELECT SUBSTRING('abc', 0, 2) AS n"),
    ("substring negative start", "SELECT SUBSTRING('abcdef', -1, 4) AS n"),
    ("substring zero length", "SELECT '[' + SUBSTRING('abc', 2, 0) + ']' AS n"),
    ("left zero", "SELECT '[' + LEFT('abc', 0) + ']' AS n"),
    ("right longer", "SELECT RIGHT('abc', 10) AS n"),
    ("charindex from", "SELECT CHARINDEX('a', 'banana', 4) AS n"),
    ("replace with nothing", "SELECT REPLACE('banana', 'a', '') AS n"),
    ("replace longer", "SELECT REPLACE('aaa', 'a', 'bb') AS n"),
    ("replicate zero", "SELECT '[' + REPLICATE('ab', 0) + ']' AS n"),
    ("upper of null", "SELECT UPPER(NULL) AS n"),
    ("len of spaces", "SELECT LEN('   ') AS n"),
    ("ltrim of spaces", "SELECT '[' + LTRIM('   ') + ']' AS n"),

    # --------------------------------------------------- aggregate edges
    ("count of a column with nulls", "SELECT COUNT(owner) AS n FROM #orders"),
    ("sum over nulls", "SELECT SUM(owner) AS n FROM #orders"),
    ("avg of integers", "SELECT AVG(qty) AS n FROM #orders"),
    ("max of text", "SELECT MAX(name) AS n FROM #people"),
    ("min of text", "SELECT MIN(name) AS n FROM #people"),
    ("count distinct", "SELECT COUNT(DISTINCT team) AS n FROM #people"),
    ("sum of nothing at all",
     "SELECT SUM(qty) AS n FROM #orders WHERE qty > 1000"),
    ("top zero", "SELECT TOP 0 name FROM #people"),
    ("top percent", "SELECT TOP 50 PERCENT id FROM #people ORDER BY id"),
    ("top percent rounds up",
     "SELECT TOP 10 PERCENT id FROM #people ORDER BY id"),
    ("top percent of groups",
     "SELECT TOP 50 PERCENT team, COUNT(*) AS n FROM #people "
     "GROUP BY team ORDER BY team"),
    ("top from a variable",
     "DECLARE @n int = 2; SELECT TOP (@n) id FROM #people ORDER BY id"),
    ("top from arithmetic", "SELECT TOP (1 + 2) id FROM #people ORDER BY id"),
    ("concat_ws", "SELECT CONCAT_WS(',', 'a', NULL, 'b', '') AS s"),
    ("concat_ws of nothing", "SELECT CONCAT_WS(',', NULL, NULL) AS s"),
    ("concat_ws over rows",
     "SELECT CONCAT_WS('-', id, name) AS s FROM #people WHERE id < 3 "
     "ORDER BY id"),
    ("distinct over nulls", "SELECT DISTINCT owner FROM #orders ORDER BY owner"),

    # --------------------------------------------------------- the writes
    # Nothing here asked a write before. The count each server reports is
    # dropped by the reader, so each one reads back what it changed.
    ("update every row",
     "UPDATE #people SET team = 'grey'; "
     "SELECT id, team FROM #people ORDER BY id"),
    ("update where",
     "UPDATE #people SET name = 'X' WHERE id = 2; "
     "SELECT id, name FROM #people ORDER BY id"),
    ("update from the row",
     "UPDATE #orders SET qty = qty + 1; "
     "SELECT item, qty FROM #orders ORDER BY item"),
    ("update to null",
     "UPDATE #people SET team = NULL WHERE id = 1; "
     "SELECT id, team FROM #people ORDER BY id"),
    ("update where null",
     "UPDATE #orders SET qty = 0 WHERE owner = NULL; "
     "SELECT item, qty FROM #orders ORDER BY item"),
    ("update where is null",
     "UPDATE #orders SET qty = 0 WHERE owner IS NULL; "
     "SELECT item, qty FROM #orders ORDER BY item"),
    ("update two columns",
     "UPDATE #people SET name = 'X', team = 'Y' WHERE id > 3; "
     "SELECT id, name, team FROM #people ORDER BY id"),
    ("update twice",
     "UPDATE #people SET id = id + 10; UPDATE #people SET id = id + 100; "
     "SELECT id FROM #people ORDER BY id"),
    ("update swaps nothing",
     "UPDATE #people SET name = team, team = name WHERE id = 1; "
     "SELECT id, name, team FROM #people ORDER BY id"),
    ("delete where",
     "DELETE FROM #people WHERE team = 'red'; "
     "SELECT id FROM #people ORDER BY id"),
    ("delete where null",
     "DELETE FROM #orders WHERE owner = NULL; "
     "SELECT item FROM #orders ORDER BY item"),
    ("delete is null",
     "DELETE FROM #orders WHERE owner IS NULL; "
     "SELECT item FROM #orders ORDER BY item"),
    ("delete every row",
     "DELETE FROM #orders; SELECT COUNT(*) AS n FROM #orders"),
    ("delete then insert",
     "DELETE FROM #people WHERE id > 2; "
     "INSERT INTO #people (id, name, team) VALUES (9, 'Katherine', 'grey'); "
     "SELECT id, name FROM #people ORDER BY id"),
    ("insert then update",
     "INSERT INTO #people (id, name, team) VALUES (6, 'Katherine', 'grey'); "
     "UPDATE #people SET team = 'blue' WHERE id = 6; "
     "SELECT id, team FROM #people ORDER BY id"),
    ("update an aggregate's source",
     "UPDATE #orders SET qty = 2; SELECT SUM(qty) AS n FROM #orders"),
]
