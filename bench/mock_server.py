"""Servidor falso compatível com OpenAI para medir o tempo do lote sem gastar tokens.

Simula a latência de uma API rápida: 0,6 s até o primeiro token + N tokens a
--tps tokens/s. Não reescreve de verdade (só reorganiza), serve só para tempo.
"""
import argparse, json, re, threading, time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

ap = argparse.ArgumentParser(); ap.add_argument("--port", type=int, default=8765)
ap.add_argument("--tps", type=float, default=90); ap.add_argument("--ttft", type=float, default=0.6)
A = ap.parse_args()
lock = threading.Lock(); live = [0, 0]

class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        sysm, user = body["messages"][0]["content"], body["messages"][-1]["content"]
        text = user.split("\n\n", 1)[-1] if user.startswith(("Seção", "Article")) else user
        if sysm.startswith(("Reduza", "Reduce")):
            out = "\n".join("- " + s.strip() for s in re.split(r"(?<=[.!?])\s+", text) if s.strip())
            stop = {"o","a","os","as","de","da","do","das","dos","que","e","em","um","uma","para","com","por","no","na",
                    "the","of","and","to","in","a","an","that","for","with","is","are","which","such","as","by"}
            out = "\n".join(" ".join(w for w in l.split() if w.lower() not in stop) for l in out.splitlines())
        elif sysm.startswith(("Escreva", "Write")):
            out = " ".join(l.lstrip("- ").strip() for l in text.splitlines() if l.startswith("-"))
        else:
            out = text
        m = re.search(r"um a cada (\d+) palavras|once every (\d+) words", sysm)
        if m:  # imita o redator inserindo marcadores
            k = int(m.group(1) or m.group(2)); w = out.split(" ")
            out = " ".join(x + (" ¦" if (i + 1) % k == 0 else "") for i, x in enumerate(w))
        ntok = int(len(out.split()) * 1.4) + 5
        if sysm.startswith(("Escreva", "Write")):  # um redator de verdade devolve a prosa inteira,
            ntok = int(len(out.split()) * 1.8 * 1.4) + 5  # não só as palavras de conteúdo
        with lock: live[0] += 1; live[1] = max(live[1], live[0])
        time.sleep(A.ttft + ntok / A.tps)
        with lock: live[0] -= 1
        res = json.dumps({"choices": [{"message": {"content": out}}],
                          "usage": {"prompt_tokens": int(len(user.split()) * 1.4 + len(sysm.split()) * 1.4),
                                    "completion_tokens": ntok}}).encode()
        self.send_response(200); self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(res))); self.end_headers(); self.wfile.write(res)
    def do_GET(self):
        res = json.dumps({"max_concorrentes": live[1]}).encode()
        self.send_response(200); self.end_headers(); self.wfile.write(res)

class S(ThreadingHTTPServer):
    request_queue_size = 1024  # API real aceita centenas de conexões; o padrão (5) recusaria
    daemon_threads = True
S(("127.0.0.1", A.port), H).serve_forever()
