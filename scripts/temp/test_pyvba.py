from pyvbaanalysis import analyze_loose_file, VbaDiagnostic, DiagnosticSeverity

with open("scripts/temp/__broken.bas", "w", encoding="utf-8") as f:
    f.write(
        'Attribute VB_Name = "Broken"\r\n'
        "Option Explicit\r\n"
        "Sub Foo()\r\n"
        "    x = UndeclaredVar + 1\r\n"
        "End Sub\r\n"
    )

count = 0
for d in analyze_loose_file("scripts/temp/__broken.bas"):
    print("SEV:", d.severity, "CODE:", d.code, "MSG:", d.message, "LINE:", getattr(d, "line", None))
    count += 1
print("diagnostics:", count)
