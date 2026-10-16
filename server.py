#!/usr/bin/env python3
"""Control remoto de PowerPoint desde el celular (solo biblioteca estándar).

Uso:  python server.py [--port 8765] [--pin 4821]
Los celulares abren http://<IP-del-portátil>:<puerto>/?k=<PIN> (el QR de la consola ya lo trae).
PowerPoint se controla con un puente de PowerShell (bridge.ps1): no hay que instalar nada.
"""
import argparse
import json
import queue
import random
import socket
import subprocess
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

HERE = Path(__file__).resolve().parent


class Bridge:
    """Un proceso PowerShell de larga vida; una orden a la vez."""

    def __init__(self):
        self.lock = threading.Lock()
        self.proc = None
        self.out = None
        self._start()

    def _start(self):
        self.proc = subprocess.Popen(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", str(HERE / "bridge.ps1")],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
            text=True, encoding="utf-8", bufsize=1)
        self.out = queue.Queue()
        threading.Thread(target=self._pump, args=(self.proc, self.out), daemon=True).start()

    @staticmethod
    def _pump(proc, out):
        for line in proc.stdout:
            out.put(line.lstrip("﻿").strip())
        out.put(None)

    def ask(self, line, timeout=8.0):
        with self.lock:
            try:
                self.proc.stdin.write(line + "\n")
                self.proc.stdin.flush()
                resp = self.out.get(timeout=timeout)
                if resp is None:
                    raise RuntimeError("el puente se cerró")
                return json.loads(resp)
            except Exception as exc:  # puente colgado o muerto: reiniciarlo
                try:
                    self.proc.kill()
                except Exception:
                    pass
                self._start()
                return {"ok": False, "error": f"puente reiniciado ({exc})"}


class State:
    show_started = None  # instante en que se vio la presentación en curso por primera vez


def local_ips():
    ips = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ips.add(info[4][0])
    except OSError:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ips.add(s.getsockname()[0])
        s.close()
    except OSError:
        pass
    return sorted(i for i in ips if not i.startswith(("127.", "169.254.")))


def make_handler(bridge, pin):
    html = (HERE / "remote.html").read_bytes()

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *a):  # silencio: llega una petición cada 700 ms por celular
            pass

        def _send(self, code, body, ctype):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _json(self, code, obj):
            self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                       "application/json; charset=utf-8")

        def do_GET(self):
            u = urlparse(self.path)
            q = parse_qs(u.query)
            if u.path in ("/", "/index.html"):
                return self._send(200, html, "text/html; charset=utf-8")
            if not u.path.startswith("/api/"):
                return self._send(404, b"no encontrado", "text/plain")
            if q.get("k", [""])[0] != pin:
                return self._json(403, {"ok": False, "error": "PIN incorrecto"})

            if u.path == "/api/cmd":
                c = q.get("c", [""])[0]
                if c == "reset":
                    State.show_started = time.time() if State.show_started else None
                    c = "state"
                at = q.get("at", [""])[0]
                n = q.get("n", [""])[0]
                if c not in ("next", "prev", "goto", "start", "state"):
                    return self._json(400, {"ok": False, "error": "orden desconocida"})
                st = bridge.ask(f"{c}|{n}|{at if at.isdigit() else ''}")
            else:
                st = bridge.ask("state")

            if st.get("ok") and st.get("mode") == "show":
                State.show_started = State.show_started or time.time()
            else:
                State.show_started = None
            st["elapsed"] = int(time.time() - State.show_started) if State.show_started else 0
            return self._json(200, st)

    return Handler


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--pin", default=str(random.randint(1000, 9999)))
    args = ap.parse_args()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    bridge = Bridge()
    srv = ThreadingHTTPServer(("0.0.0.0", args.port), make_handler(bridge, args.pin))

    print("\n=== Control remoto de la charla ===")
    print(f"PIN: {args.pin}   (puerto {args.port})")
    urls = [f"http://{ip}:{args.port}/?k={args.pin}" for ip in local_ips()]
    for u in urls:
        print("  ", u)
    if not urls:
        print("  (sin red: conecta el portátil al hotspot del celular y reinicia)")
    else:
        try:
            import segno
            qr = segno.make(urls[-1], error="l")
            qr.save(str(HERE / "qr.png"), scale=8, border=2)
            print(f"\nQR del último enlace guardado en {HERE / 'qr.png'}\n")
            qr.terminal(compact=True)
        except Exception:
            pass
        print("Si hay varias IP, usa la del hotspot (p. ej. 172.20.10.x en iPhone, 192.168.43.x en Android).")
    print("\nCtrl+C para salir.\n")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        try:
            bridge.proc.kill()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(main())
