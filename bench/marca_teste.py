#!/usr/bin/env python3
"""Teste de marca d'água com chave nossa: aplica, passa pelo denoise.py e mede o que sobrou.

Como a marca é aplicada (igual em espírito ao SynthID, que escolhe entre candidatos):
para cada frase do texto, um modelo gera K paráfrases; fica a que tem mais palavras
"verdes". Uma palavra é verde quando hash(chave, contexto, palavra) cai na metade verde,
e o contexto são as H palavras anteriores. Num texto sem marca, ~50% das palavras são
verdes por acaso; com marca, bem mais. O detector mede isso como um z-score:

    z = (verdes - 0,5 * n) / sqrt(0,25 * n)

z > 4 = marca detectada com folga (chance de acaso < 1 em 30 mil).

Medimos três chaves ao mesmo tempo, porque não sabemos qual o fornecedor usa:
  - H=0: lista verde fixa, sem contexto (tipo Unigram-Watermark);
  - H=1: contexto de 1 palavra (tipo Kirchenbauer);
  - H=3: contexto de 3 palavras (mais perto do SynthID).
O texto é marcado com a soma das três, então as três ficam altas antes.

Uso (precisa de um modelo para gerar as paráfrases e outro, ou o mesmo, para o denoise):
  python bench/marca_teste.py artigo.md -m anthropic:claude-haiku-4-5 --saida teste_marca
  python bench/marca_teste.py artigo.md -m anthropic:claude-haiku-4-5 --variantes "padrao;--marcadores 2.5;--sigilo forte"
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import shlex
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import denoise  # noqa: E402

KEY = "babelllm-teste-chave-1"
WINDOWS = (0, 1, 3)


def words(text: str) -> list[str]:
    return re.findall(r"[a-zà-ÿ0-9]+", text.lower())


def green(ctx: tuple[str, ...], w: str, h: int) -> bool:
    d = hashlib.sha256(f"{KEY}|{h}|{' '.join(ctx[-h:]) if h else ''}|{w}".encode()).digest()
    return d[0] & 1 == 1


def score(text: str, h: int) -> dict:
    """z-score da marca com janela h. Pares (contexto, palavra) repetidos contam uma vez."""
    ws, seen, g, n = words(text), set(), 0, 0
    for i, w in enumerate(ws):
        ctx = tuple(ws[max(0, i - h):i])
        if len(ctx) < h or (ctx, w) in seen:
            continue
        seen.add((ctx, w))
        n += 1
        g += green(ctx, w, h)
    z = (g - 0.5 * n) / math.sqrt(0.25 * n) if n else 0.0
    return {"z": round(z, 2), "verdes": round(g / n, 3) if n else 0, "n": n}


def gain(prev: list[str], cand: str) -> int:
    ws = prev[-3:] + words(cand)
    k = len(prev[-3:])
    return sum(green(tuple(ws[max(0, i - h):i]), ws[i], h) * 2 - 1
               for i in range(k, len(ws)) for h in WINDOWS if i - h >= 0)


PROMPT = ("Reescreva a frase abaixo de {k} maneiras diferentes, em português do Brasil, com o mesmo sentido "
          "e os mesmos números. Varie bastante a escolha das palavras. Responda só com um JSON: uma lista de "
          "{k} strings.")


def variants(model, sentence: str, k: int) -> list[str]:
    try:
        raw = model(PROMPT.format(k=k), sentence, max_tokens=120 + len(sentence.split()) * 8 * k // 4)
        m = re.search(r"\[.*\]", raw, re.S)
        out = [s for s in json.loads(m.group(0)) if isinstance(s, str) and s.strip()] if m else []
    except Exception as e:  # noqa: BLE001
        print(f"  aviso: {e}", file=sys.stderr)
        out = []
    return [sentence] + out


def watermark(md: str, model, k: int, jobs: int) -> str:
    """Marca só os parágrafos de texto; títulos, listas curtas e código ficam iguais."""
    blocks = denoise.split_blocks(md)
    sents = []  # (bloco, índice, frase)
    for bi, b in enumerate(blocks):
        if denoise._is_text(b) and not b.lstrip().startswith(("-", "*", "|", ">")):
            for si, s in enumerate(re.split(r"(?<=[.!?])\s+", b.replace("\n", " "))):
                if len(s.split()) >= 5:
                    sents.append((bi, si, s))
    with ThreadPoolExecutor(jobs) as pool:
        cands = list(pool.map(lambda x: variants(model, x[2], k), sents))
    chosen = {}
    prev: list[str] = []
    for (bi, si, s), cs in zip(sents, cands):
        best = max(cs, key=lambda c: gain(prev, c))
        chosen[(bi, si)] = best
        prev = (prev + words(best))[-3:]
    out = []
    for bi, b in enumerate(blocks):
        if any(key[0] == bi for key in chosen):
            parts = re.split(r"(?<=[.!?])\s+", b.replace("\n", " "))
            out.append(" ".join(chosen.get((bi, si), p) for si, p in enumerate(parts)))
        else:
            out.append(b)
    return "\n\n".join(out) + "\n"


def table(label: str, text: str) -> dict:
    row = {f"H={h}": score(text, h)["z"] for h in WINDOWS}
    print(f"{label:<28} " + "  ".join(f"{k}: z={v:>6}" for k, v in row.items()))
    return row


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("artigo")
    ap.add_argument("-m", "--modelo", required=True, help="modelo que gera as paráfrases e reescreve (provedor:modelo)")
    ap.add_argument("--denoise-modelo", help="modelo do denoise (padrão: o mesmo)")
    ap.add_argument("-k", type=int, default=6, help="paráfrases por frase ao marcar (padrão 6; mais = marca mais forte)")
    ap.add_argument("--variantes", default="padrao;--marcadores 2.5;--marcadores 5;--sigilo forte",
                    help="opções do denoise a testar, separadas por ';' ('padrao' = sem opção)")
    ap.add_argument("-j", type=int, default=16)
    ap.add_argument("--saida", default="teste_marca")
    a = ap.parse_args()

    out = Path(a.saida)
    out.mkdir(parents=True, exist_ok=True)
    md, _ = denoise.strip_invisible(Path(a.artigo).read_text(encoding="utf-8"))
    model = denoise.Model(a.modelo)

    t0 = time.time()
    marked = watermark(md, model, a.k, a.j)
    (out / "marcado.md").write_text(marked, encoding="utf-8")
    print(f"marcado em {time.time() - t0:.0f}s ({model.usage.calls} chamadas)\n")

    res = {"original": table("original (sem marca)", md), "marcado": table("marcado (com a nossa chave)", marked)}
    for v in [x.strip() for x in a.variantes.split(";") if x.strip()]:
        name = re.sub(r"[^\w]+", "_", v).strip("_") or "padrao"
        args = [] if v == "padrao" else shlex.split(v)
        d = out / f"limpo_{name}"
        cmd = [sys.executable, str(ROOT / "denoise.py"), str(out / "marcado.md"), "-o", str(d),
               "-m", a.denoise_modelo or a.modelo, "--relatorio", str(out / f"rel_{name}.json"), *args]
        subprocess.run(cmd, check=False, stdout=subprocess.DEVNULL)
        f = d / "marcado.md"
        if f.exists():
            res[v] = table(f"denoise {v}", f.read_text(encoding="utf-8"))
    (out / "resultado.json").write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\nz > 4: marca detectada. |z| < 2: indistinguível de texto sem marca.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
