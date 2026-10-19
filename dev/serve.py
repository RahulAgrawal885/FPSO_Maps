"""Local preview server for the FPSO map that tells the browser never to cache, so edits show on reload.

Usage (from the repo root):  python3 dev/serve.py [port]      then open http://127.0.0.1:8766/
"""
import functools, http.server, pathlib, sys

ROOT = pathlib.Path(__file__).resolve().parent.parent   # repo root, where index.html lives
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8766


class NoCache(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


print(f"Serving {ROOT} at http://127.0.0.1:{PORT}/")
http.server.ThreadingHTTPServer(("127.0.0.1", PORT), functools.partial(NoCache, directory=ROOT)).serve_forever()
