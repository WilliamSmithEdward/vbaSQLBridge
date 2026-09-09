"""The correlated shapes, and the data they are asked about.

Kept apart from the tests that run them so the same list can be put to a
real SQL Server and to this one and the two answers compared. A shortcut
that only fires on some shapes has to be checked on the ones it declines as
well, or a shortcut that never fires would pass every test.
"""

# Two small tables, built the same way on either server. Temporary, so
# nothing is left behind on a machine that only lent its SQL Server for an
# afternoon. The nulls and the repeated owner are the interesting part: a
# lookup that loses them answers differently from a scan.
SETUP = """
CREATE TABLE #people (id int, name nvarchar(40), team nvarchar(20));
INSERT INTO #people (id, name, team) VALUES (1, 'Ada', 'red');
INSERT INTO #people (id, name, team) VALUES (2, 'Grace', 'blue');
INSERT INTO #people (id, name, team) VALUES (3, 'Edsger', 'red');
INSERT INTO #people (id, name, team) VALUES (4, 'Barbara', 'blue');
INSERT INTO #people (id, name, team) VALUES (5, 'Alan', 'red');
CREATE TABLE #orders (owner int, item nvarchar(40), qty int);
INSERT INTO #orders (owner, item, qty) VALUES (1, 'Ada', 3);
INSERT INTO #orders (owner, item, qty) VALUES (1, 'engine', 1);
INSERT INTO #orders (owner, item, qty) VALUES (2, 'compiler', 2);
INSERT INTO #orders (owner, item, qty) VALUES (4, 'Barbara', 1);
INSERT INTO #orders (owner, item, qty) VALUES (NULL, 'nobody', 9);
"""

# name, whether the EXISTS lookup is expected to be built, the statement.
SHAPES = [
    ("plain", True,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = p.id) ORDER BY p.id"),
    ("reversed", True,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE p.id = o.owner) ORDER BY p.id"),
    ("star", True,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT * FROM #orders o WHERE o.owner = p.id) ORDER BY p.id"),
    ("negated", True,
     "SELECT p.name FROM #people p WHERE NOT EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = p.id) ORDER BY p.id"),
    ("extra inner condition", True,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = p.id AND o.qty > 1) "
     "ORDER BY p.id"),
    ("inner condition first", True,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.qty > 1 AND o.owner = p.id) "
     "ORDER BY p.id"),
    ("text key", True,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.item = p.name) ORDER BY p.id"),
    ("outer condition beside it", True,
     "SELECT p.name FROM #people p WHERE p.team = 'red' AND EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = p.id) ORDER BY p.id"),

    # The shapes the shortcut has to decline. Dropping the correlated
    # condition and matching on a key would answer these differently.
    ("or", False,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = p.id OR o.qty = 9) "
     "ORDER BY p.id"),
    ("not equality", False,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner > p.id) ORDER BY p.id"),
    ("two outer columns", False,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = p.id AND o.item = p.name) "
     "ORDER BY p.id"),
    ("outer expression", False,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT 1 FROM #orders o WHERE o.owner = p.id + 0) ORDER BY p.id"),
    ("grouped", False,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT o.owner FROM #orders o WHERE o.owner = p.id "
     "GROUP BY o.owner HAVING COUNT(*) > 1) ORDER BY p.id"),
    ("top none", False,
     "SELECT p.name FROM #people p WHERE EXISTS "
     "(SELECT TOP 0 1 FROM #orders o WHERE o.owner = p.id) ORDER BY p.id"),

    # Not an EXISTS at all, so no EXISTS lookup is built for them either.
    # They are here because the answers have to keep agreeing while the
    # EXISTS path is changed underneath them.
    ("in", False,
     "SELECT p.name FROM #people p WHERE p.id IN "
     "(SELECT o.owner FROM #orders o) ORDER BY p.id"),
    ("correlated value", False,
     "SELECT p.name, (SELECT COUNT(*) FROM #orders o WHERE o.owner = p.id) "
     "AS n FROM #people p ORDER BY p.id"),
]


# The join shapes, and whether the rows on one side are expected to be
# grouped by a key rather than every pair tried.
JOIN_SHAPES = [
    ("join on equality", True,
     "SELECT p.name, o.item FROM #people p JOIN #orders o "
     "ON o.owner = p.id ORDER BY p.id, o.item"),
    ("join reversed", True,
     "SELECT p.name, o.item FROM #people p JOIN #orders o "
     "ON p.id = o.owner ORDER BY p.id, o.item"),
    ("join with a second condition", True,
     "SELECT p.name, o.item FROM #people p JOIN #orders o "
     "ON o.owner = p.id AND o.qty > 1 ORDER BY p.id, o.item"),
    ("left join on equality", True,
     "SELECT p.name, ISNULL(o.item, '-') AS item FROM #people p "
     "LEFT OUTER JOIN #orders o ON o.owner = p.id ORDER BY p.id, o.item"),
    ("join on text", True,
     "SELECT p.name, o.item FROM #people p JOIN #orders o "
     "ON o.item = p.name ORDER BY p.id"),
    ("self join", True,
     "SELECT a.name, b.name AS other FROM #people a JOIN #people b "
     "ON a.team = b.team AND a.id < b.id ORDER BY a.id, b.id"),

    # The shapes with nothing to group the rows by.
    ("join on an inequality", False,
     "SELECT COUNT(*) AS n FROM #people p JOIN #orders o "
     "ON o.owner < p.id"),
    ("join on an or", False,
     "SELECT COUNT(*) AS n FROM #people p JOIN #orders o "
     "ON o.owner = p.id OR o.qty = 9"),
    ("join on an expression", True,
     "SELECT COUNT(*) AS n FROM #people p JOIN #orders o "
     "ON o.owner = p.id + 0"),
    ("cross join", False,
     "SELECT COUNT(*) AS n FROM #people p, #orders o"),
]
