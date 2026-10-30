"""Run the VBA tests headless in LibreOffice Calc (Linux, no Excel needed).

    python3 tests/run_in_libreoffice.py [--trace] [test name filter]

--trace numbers the statements of both modules as it loads them and, when a
test errors, reports the last source line each module reached
(LibreOffice gives no line numbers for runtime errors).

Loads src/IRExCollection.bas and tests/IRExCollectionTests.bas into a hidden
Calc document in VBA compatibility mode and runs each test through
RunOneHeadless. Exits 1 if any test fails.

LibreOffice's VBA support has gaps: no ListObjects, SaveAs always writes the
old .xls format whatever the file name, and Workbook.Close leaves documents
open. So the harness closes documents between tests, and the tests mark the
few checks that need real Excel as skipped here. Desktop Excel on Windows
stays the reference run: see tests/README.md.
"""

import os
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODULES = [
    ("IRExCollection", os.path.join(ROOT, "src", "IRExCollection.bas")),
    ("IRExCollectionTests", os.path.join(ROOT, "tests", "IRExCollectionTests.bas")),
]


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def basic_source(path):
    """Excel's exported .bas, made loadable by LibreOffice Basic."""
    with open(path, encoding="cp1252") as f:
        lines = [l.rstrip("\r\n") for l in f]
    # Attribute lines become blank so that line numbers match the file.
    lines = ["" if l.startswith("Attribute VB_") else l for l in lines]
    return "Option VBASupport 1\n" + "\n".join(lines[1:]) + "\n"


NOT_STATEMENTS = ("case ", "case\t", "else", "elseif ", "end ", "loop", "next", "wend",
                  "private ", "public ", "dim ", "const ", "option ", "type ", "'", "#")


def traced(source, var):
    """Puts "<var> = n" before each statement (n = line in the .bas file)."""
    out, continued, in_type = [], False, False
    for n, line in enumerate(source.split("\n"), start=1):
        bare = line.strip().lower()
        if bare.startswith(("private type", "public type", "type ")):
            in_type = True
        if in_type or continued or not bare or bare.startswith(NOT_STATEMENTS) \
                or bare.endswith(":") or bare.startswith("select case"):
            out.append(line)
        else:
            out.append(f"{var} = {n}: " + line.lstrip())
        if bare.startswith("end type"):
            in_type = False
        continued = line.rstrip().endswith(" _")
    text = "\n".join(out)
    head, _, rest = text.partition("Option Explicit\n")
    extra = "Public TraceIRExCollectionTestsCopy As Long\n" if var == "TraceIRExCollection" else ""
    return (head + f"Option Explicit\nPublic {var} As Long\n" + extra + rest +
            f"\nPublic Function {var}Headless() As Long\n{var}Headless = {var}\nEnd Function\n")


def close_others(desktop, keep):
    comps = desktop.Components.createEnumeration()
    docs = []
    while comps.hasMoreElements():
        docs.append(comps.nextElement())
    for d in docs:
        if d == keep:
            continue
        try:
            d.setModified(False)
        except Exception:
            pass
        try:
            d.close(True)
        except Exception:
            try:
                d.dispose()
            except Exception:
                pass
    left = 0
    comps = desktop.Components.createEnumeration()
    while comps.hasMoreElements():
        if comps.nextElement() != keep:
            left += 1
    if left:
        print(f"warning: {left} document(s) would not close", file=sys.stderr)


def free_port():
    with socket.socket() as s:
        s.bind(("localhost", 0))
        return s.getsockname()[1]


def main():
    args = sys.argv[1:]
    trace = "--trace" in args
    args = [a for a in args if a != "--trace"]
    name_filter = args[0] if args else ""
    port = free_port()
    profile = tempfile.mkdtemp(prefix="lo-profile-")
    out_dir = os.path.join(ROOT, "tests", "out")
    os.makedirs(out_dir, exist_ok=True)
    proc = subprocess.Popen(
        ["soffice", "--headless", "--invisible", "--norestore", "--nologo",
         f"-env:UserInstallation=file://{profile}",
         f"--accept=socket,host=localhost,port={port};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(120):
            try:
                ctx = resolver.resolve(
                    f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        if ctx is None:
            sys.exit("LibreOffice did not start")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0,
                                           (prop("Hidden", True),))
        libs = doc.BasicLibraries
        libs.VBACompatibilityMode = True
        lib = libs.getByName("Standard")
        for name, path in MODULES:
            source = basic_source(path)
            if trace:
                source = traced(source, "Trace" + name)
                # Record where both modules were when an error escaped a test.
                source = re.sub(
                    r'TraceIRExCollectionTests = \d+: (Fail "ERROR " & Err.Number & ": " & Err.Description)',
                    r'\1 & vbLf & "last lines reached: src/IRExCollection.bas " & '
                    r'IRExCollection.TraceIRExCollection & ", tests/IRExCollectionTests.bas " & '
                    r'IRExCollection.TraceIRExCollectionTestsCopy', source)
                source = re.sub(r"(TraceIRExCollectionTests = (\d+): )",
                                r"\1IRExCollection.TraceIRExCollectionTestsCopy = \2: ", source)
            lib.insertByName(name, source)
        provider = doc.getScriptProvider()

        def call(name, *args):
            return provider.getScript(
                f"vnd.sun.star.script:Standard.IRExCollectionTests.{name}"
                "?language=Basic&location=document").invoke(args, (), ())[0]

        fixtures = os.path.join(ROOT, "tests", "fixtures") + os.sep
        lines, passed, failed = [], 0, 0
        try:
            names = call("TestNamesHeadless")
        except Exception as e:
            sys.exit(f"The VBA modules don't load: {e}")
        if not names:
            sys.exit("The VBA modules don't compile (LibreOffice gives no details: "
                     "check the last edit, or open the modules in LibreOffice's Basic IDE)")
        for test in names.split("|"):
            if name_filter and name_filter.lower().replace(" ", "_") not in test.lower():
                continue
            # One test per call, closing every other document afterwards:
            # LibreOffice's Workbook.Close leaves documents open.
            try:
                line = call("RunOneHeadless", fixtures, out_dir + os.sep, test)
            except Exception as e:  # a Basic runtime error escaped the test
                line = f"ERROR  {test.replace('_', ' ')}\n      {e}"
            close_others(desktop, doc)
            lines.append(line)
            if line.startswith("PASS"):
                passed += 1
            else:
                failed += 1
        result = "\n".join(lines) + f"\n\n{passed} passed, {failed} failed"
        print(result)
        try:
            doc.close(True)
            desktop.terminate()
        except Exception:
            pass
        sys.exit(1 if failed else 0)
    finally:
        try:
            proc.wait(timeout=20)
        except subprocess.TimeoutExpired:
            proc.kill()
        shutil.rmtree(profile, ignore_errors=True)


if __name__ == "__main__":
    main()
