"""Shared pieces for the Python-driven tests."""
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src"

# The rows the end-to-end tests query, written into a worksheet so the client
# is reading a real sheet rather than an array built for the occasion.
PEOPLE = [
    ["id", "name", "score", "retired"],
    [1, "Ada Lovelace", 99.5, True],
    [2, "Grace Hopper", 87.25, True],
    [3, "Edsger Dijkstra", 78.0, True],
    [4, "Barbara Liskov", 93.75, False],
]

# A server that serves that sheet for a while and hands back what it logged.
# Serve is a factory, so a caller with the class imported writes
# SqlBridge.Serve(1433); the harness strips the header that makes SqlBridge
# itself an instance, so the tests take the same route through one of theirs.
SERVE_SOURCE = """
Option Explicit

Public Function Main(ByVal port As Long, ByVal budgetMs As Long) As String
    Dim factory As SqlBridge
    Dim server As SqlBridge

    Set factory = New SqlBridge
    Set server = factory.Serve(port)
    server.AddTable "people", ThisWorkbook.Worksheets("Sheet1").Range("A1:D5")
    server.RunFor budgetMs
    Main = server.LogText
    server.Shutdown
End Function
"""
