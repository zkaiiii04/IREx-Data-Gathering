"""Find which procedure (and line) stops LibreOffice Basic compiling a .bas file.

    python3 tests/find_compile_error.py src/IRExCollection.bas

LibreOffice reports no detail when a module fails to compile, so this removes
one procedure at a time until the rest compiles. Known gap it has found:
LibreOffice can't ReDim an array that is a member of a Type.
"""
import re
import socket
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

def prop(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

src = open(sys.argv[1], encoding="cp1252").read().replace("\r\n", "\n")
lines = [l for l in src.split("\n") if not l.startswith("Attribute VB_")]
# split into header + procedures
starts = [i for i, l in enumerate(lines) if re.match(r"^(Public |Private )?(Sub|Function) ", l)]
header = lines[:starts[0]]
procs = []
for k, s in enumerate(starts):
    e = starts[k + 1] if k + 1 < len(starts) else len(lines)
    procs.append(lines[s:e])

with socket.socket() as s:
    s.bind(("localhost", 0)); port = s.getsockname()[1]
profile = tempfile.mkdtemp()
proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", f"-env:UserInstallation=file://{profile}", f"--accept=socket,host=localhost,port={port};urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
local = uno.getComponentContext()
res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(120):
    try:
        ctx = res.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(0.5)
desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
n = [0]
def compiles(text):
    n[0] += 1
    doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, (prop("Hidden", True),))
    doc.BasicLibraries.VBACompatibilityMode = True
    lib = doc.BasicLibraries.getByName("Standard")
    lib.insertByName("M", "Option VBASupport 1\n" + text + "\nFunction ProbeOK() As String\nProbeOK = \"ok\"\nEnd Function\n")
    try:
        r = doc.getScriptProvider().getScript("vnd.sun.star.script:Standard.M.ProbeOK?language=Basic&location=document").invoke((), (), ())[0]
    except Exception as e:
        r = str(e)
    doc.close(True)
    return r == "ok"

full = "\n".join(header + sum(procs, []))
print("full compiles:", compiles(full))
print("header only compiles:", compiles("\n".join(header)))
for i, p in enumerate(procs):
    rest = header + sum(procs[:i] + procs[i + 1:], [])
    if compiles("\n".join(rest)):
        print("culprit:", p[0])
        # line bisect inside procedure: try blanking single lines
        for j in range(1, len(p) - 1):
            q = p[:j] + ["' removed"] + p[j + 1:]
            if compiles("\n".join(header + sum(procs[:i], []) + q + sum(procs[i + 1:], []))):
                print("   line:", p[j])
try: desk.terminate()
except Exception: pass
proc.wait(timeout=20)
