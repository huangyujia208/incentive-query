import os, zipfile

ROOT = r"D:\wordbuddy\2026-10-20-17-05-44\incentive-query"
OUT = r"D:\wordbuddy\2026-10-20-17-05-44\incentive-query\render-deploy-package.zip"

EXCLUDE_DIRS = {"__pycache__", ".git", ".workbuddy"}
EXCLUDE_FILES = {
    ".log", ".pyc", ".DS_Store", "local_preview.html",
    "server.log", "app.log", "render-deploy-package.zip",
    "query_log.csv",
}

def keep(path):
    name = os.path.basename(path)
    if name in EXCLUDE_FILES:
        return False
    if name.endswith(".pyc"):
        return False
    return True

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for dp, dns, fns in os.walk(ROOT):
        dns[:] = [d for d in dns if d not in EXCLUDE_DIRS]
        for fn in fns:
            full = os.path.join(dp, fn)
            if not keep(full):
                continue
            rel = os.path.relpath(full, ROOT)
            z.write(full, rel)

print("written:", OUT)
