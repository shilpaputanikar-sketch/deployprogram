import json
import sys
import urllib.request
import urllib.error

# Usage:
#   python test_endpoints.py https://deployprogram-xxxx.onrender.com
# Defaults to local dev server if no URL is given.
BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:5000"


def call(method, path, body=None):
    url = BASE.rstrip("/") + path
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, r.read().decode()
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode()
    except Exception as e:
        return "ERR", str(e)


tests = [
    ("GET", "/", None),
    ("GET", "/users", None),
    ("GET", "/tasks", None),
    ("POST", "/tasks", {"title": "test task", "done": False}),
    ("GET", "/tasks", None),
    ("POST", "/tasks", {"title": "second task"}),
    ("DELETE", "/tasks/1", None),
    ("GET", "/tasks", None),
]

print("Target:", BASE)
print("-" * 60)
for method, path, body in tests:
    status, resp = call(method, path, body)
    print(f"{method:6} {path:15} -> {status}")
    print(f"       {resp}")
print("-" * 60)
print("Done. Every endpoint should return 200/201 (or 404 only for a missing id).")
