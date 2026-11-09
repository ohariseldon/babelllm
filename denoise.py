#!/usr/bin/env python3
"""denoise.py: tira a marca d'água de LLM de artigos, em lote e rápido.

Para quem usa IA para escrever e assume isso, mas não quer o texto carimbado
pela marca d'água estatística do modelo que o gerou.

Por trecho (até ~150 palavras, sem cruzar títulos), em uma chamada só:
  1. limpeza grátis: caracteres invisíveis e metadados do arquivo saem;
  2. notas locais, sem LLM: cada frase vira só as palavras de conteúdo, com
     números, nomes e citações intactos (sai a costura: artigos, preposições,
     muletas como "vale ressaltar");
  3. REDATOR: um modelo escreve o trecho do zero a partir das notas, sem ver a
     prosa original. Se faltar número, citação ou URL, tenta de novo uma vez.
O redator também reorganiza a lógica do parágrafo (ordem das ideias, junta e
divide frases), contra marcas de nível de frase ou de sentido que ainda não
conhecemos (--manter-logica desliga).
Sigilo (padrão): números, nomes, siglas, links e citações vão ao modelo como
códigos com chave ({Nxy}, {Exy}...) e são destrocados no fim; a API não vê os dados.
Opcional: --notas provedor:modelo (notas feitas por LLM, reescrita mais funda),
--revisao provedor:modelo (mais um passe de correção) e --marcadores N: o último
modelo escreve com N marcadores "¦" a cada 10 palavras e o script os apaga. Cada
token escrito logo depois de um marcador foi escolhido num contexto que deixa de
existir, então a marca do PRÓPRIO redator perde o sinal ali. Quanto mais
marcadores, melhor a remoção (e menos fluido o texto).
Todos os trechos de todos os artigos vão ao mesmo tempo, em conexões persistentes.

Por que tira a marca: a marca d'água (tipo SynthID) vive na sequência de
tokens que o modelo de origem escolheu. Aqui o redator escolhe de novo toda a
costura das frases. Por isso o redator NÃO deve ser o modelo que marcou o texto
(nem outro que também marque): o último modelo a escrever deixa a marca dele.

Uso:
  python denoise.py artigos/ -o limpos/ -m openai:gpt-4o-mini
  python denoise.py artigos/ -o limpos/ -m ollama:qwen2.5:7b-instruct
  python denoise.py artigo.docx -o limpos/ --so-limpar     # só invisíveis + metadados

Só biblioteca padrão. .docx precisa de python-docx; .pdf de pypdf.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import struct
import sys
import threading
import time
import unicodedata
import http.client
import urllib.parse
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path

__version__ = "0.6.0"

# ---------------------------------------------------------------- Unicode

_INVISIBLE = re.compile(
    "[­͏؜ᅟᅠ឴឵᠋-᠏​-‏‪-‮"
    "⁠-⁤⁦-⁯ㅤ︀-️﻿ﾠ￹-￻"
    "\U0001d173-\U0001d17a\U000e0000-\U000e007f\U000e0100-\U000e01ef]"
)
_ODD_SPACE = re.compile("[   -   　]")


def strip_invisible(text: str) -> tuple[str, int]:
    """Remove invisíveis (largura zero, tags, bidi, seletores...) e espaços exóticos."""
    n = len(_INVISIBLE.findall(text)) + len(_ODD_SPACE.findall(text))
    text = _ODD_SPACE.sub(" ", _INVISIBLE.sub("", text))
    text = "".join(c for c in text if unicodedata.category(c) != "Cf" or c in "\n\t")
    return unicodedata.normalize("NFC", text), n


# ---------------------------------------------------------------- arquivos

_HEAD = re.compile(r"^#{1,6}\s")
_FENCE = re.compile(r"^\s*```")


def read_any(path: Path) -> str:
    """Devolve o conteúdo como markdown simples."""
    ext = path.suffix.lower()
    if ext == ".docx":
        import docx
        out = []
        for p in docx.Document(str(path)).paragraphs:
            t, st = p.text.strip(), (p.style.name if p.style is not None else "").lower()
            if not t:
                continue
            if st.startswith(("heading", "título", "title")):
                lv = re.findall(r"\d", st)
                out.append("#" * (int(lv[0]) if lv else 1) + " " + t)
            elif "list" in st or "lista" in st:
                out.append("- " + t)
            else:
                out.append(t)
        return "\n\n".join(out)
    if ext == ".pdf":
        from pypdf import PdfReader
        raw = "\n".join((pg.extract_text() or "") for pg in PdfReader(str(path)).pages)
        raw = re.sub(r"(\w)-\n(\w)", r"\1\2", raw)
        return re.sub(r"(?<![.!?:;])\n(?!\n)", " ", raw)
    return path.read_text(encoding="utf-8")


_CORE_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties '
             'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
             'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"/>')
_APP_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties '
            'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>')


def clean_docx_bytes(data: bytes) -> bytes:
    """Zera autor/datas/app, tira customXml e partes C2PA, datas do zip em 1980."""
    zin, buf = zipfile.ZipFile(io.BytesIO(data)), io.BytesIO()
    drop = {i.filename for i in zin.infolist() if "c2pa" in i.filename.lower()
            or i.filename.lower().startswith("customxml/") or i.filename.lower() == "docprops/custom.xml"}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            name, low = info.filename, info.filename.lower()
            if name in drop:
                continue
            body = zin.read(name)
            if low == "docprops/core.xml":
                body = _CORE_XML.encode()
            elif low == "docprops/app.xml":
                body = _APP_XML.encode()
            elif low.endswith(".rels") or low == "[content_types].xml":
                txt = body.decode("utf-8")
                for d in drop:
                    txt = re.sub(r'<Relationship[^>]*Target="[^"]*' + re.escape(d.split("/")[-1]) + r'"[^>]*/>', "", txt)
                    txt = re.sub(r'<Override[^>]*PartName="/' + re.escape(d) + r'"[^>]*/>', "", txt)
                body = txt.encode("utf-8")
            zi = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, body)
    return buf.getvalue()


def clean_pdf_bytes(data: bytes) -> bytes:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import NameObject
    w = PdfWriter(clone_from=PdfReader(io.BytesIO(data)))
    for key in ("/Metadata", "/AF"):
        if key in w._root_object:
            del w._root_object[NameObject(key)]
    for p in w.pages:
        if "/Metadata" in p:
            del p[NameObject("/Metadata")]
    w._info = None
    out = io.BytesIO()
    w.write(out)
    return out.getvalue()


def clean_image_bytes(data: bytes) -> bytes:
    """PNG: só chunks de imagem/cor. JPEG: tira EXIF/XMP/C2PA (APP1, APP11-15, COM)."""
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        keep = {b"IHDR", b"PLTE", b"IDAT", b"IEND", b"tRNS", b"gAMA", b"cHRM", b"sRGB", b"iCCP", b"sBIT", b"pHYs"}
        out, i = [data[:8]], 8
        while i + 8 <= len(data):
            (ln,) = struct.unpack(">I", data[i:i + 4])
            if data[i + 4:i + 8] in keep:
                out.append(data[i:i + 12 + ln])
            if data[i + 4:i + 8] == b"IEND":
                break
            i += 12 + ln
        return b"".join(out)
    if data[:2] == b"\xff\xd8":
        out, i = [data[:2]], 2
        while i + 4 <= len(data) and data[i] == 0xFF:
            m = data[i + 1]
            if m == 0xDA:
                out.append(data[i:])
                break
            (ln,) = struct.unpack(">H", data[i + 2:i + 4])
            if not (m == 0xE1 or 0xEB <= m <= 0xEF or m == 0xFE) or (m == 0xEE and b"Adobe" in data[i:i + 16]):
                out.append(data[i:i + 2 + ln])
            i += 2 + ln
        return b"".join(out)
    raise ValueError("imagem não suportada")


def write_any(md: str, path: Path) -> None:
    if path.suffix.lower() != ".docx":
        path.write_text(md, encoding="utf-8")
        return
    import docx
    d = docx.Document()
    for blk in split_blocks(md):
        if _HEAD.match(blk):
            d.add_heading(blk.lstrip("#").strip(), level=min(9, len(blk) - len(blk.lstrip("#"))))
        elif re.match(r"^\s*(?:[-*+]|\d+[.)])\s", blk):
            for line in blk.splitlines():
                d.add_paragraph(re.sub(r"^\s*(?:[-*+]|\d+[.)])\s+", "", line), style="List Bullet")
        else:
            d.add_paragraph(blk)
    buf = io.BytesIO()
    d.save(buf)
    path.write_bytes(clean_docx_bytes(buf.getvalue()))


# ---------------------------------------------------------------- trechos

def split_blocks(md: str) -> list[str]:
    """Blocos separados por linha em branco; código cercado fica num bloco só."""
    blocks, cur, in_code = [], [], False
    for line in md.replace("\r\n", "\n").split("\n"):
        if _FENCE.match(line):
            in_code = not in_code
            cur.append(line)
            continue
        if not in_code and _HEAD.match(line):
            if cur:
                blocks.append("\n".join(cur))
            blocks.append(line)
            cur = []
            continue
        if not in_code and not line.strip():
            if cur:
                blocks.append("\n".join(cur))
            cur = []
            continue
        cur.append(line)
    if cur:
        blocks.append("\n".join(cur))
    return blocks


def _is_text(blk: str) -> bool:
    s = blk.lstrip()
    return bool(s) and not _HEAD.match(s) and not s.startswith(("```", "|", "<", "![")) and len(s.split()) >= 4


@dataclass
class Piece:
    text: str
    rewrite: bool
    heading: str = ""
    out: str = ""
    lost: list[str] = field(default_factory=list)
    error: str = ""
    marks: int = 0
    coverage: float = 0.0


def make_pieces(md: str, max_words: int) -> list[Piece]:
    """Junta blocos de texto seguidos (sem cruzar título) até ~max_words."""
    pieces, buf, heading = [], [], ""

    def flush():
        if buf:
            pieces.append(Piece("\n\n".join(buf), True, heading))
            buf.clear()

    for blk in split_blocks(md):
        if _is_text(blk):
            if buf and sum(len(b.split()) for b in buf) + len(blk.split()) > max_words:
                flush()
            buf.append(blk)
        else:
            flush()
            if _HEAD.match(blk):
                heading = blk.lstrip("#").strip()
            pieces.append(Piece(blk, False))
    flush()
    return pieces


_MUST = re.compile(
    r"https?://\S+|\[\d+(?:[,–-]\s*\d+)*\]|\([A-ZÀ-Ú][^()]{0,60}?\d{4}[a-z]?\)|\d+(?:[.,]\d+)*\s?%?"
)


def must_keep(text: str) -> list[str]:
    """Números, citações (Autor, 2020) / [3] e URLs que a saída precisa conter."""
    items = []
    for m in _MUST.findall(text):
        m = m.strip()
        if not m.startswith("("):
            m = m.rstrip(".,;:)")
        if m and m not in items:
            items.append(m)
    return items


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", s).replace(",", ".").lower()


def missing(items: list[str], out: str) -> list[str]:
    o = _norm(out)
    lost = []
    for it in items:
        if _norm(it) in o:
            continue
        # citação reescrita: basta o sobrenome e o ano estarem lá
        parts = re.findall(r"[A-Za-zÀ-ú]{3,}|\d{4}", it)
        if it.startswith("(") and parts and all(_norm(p) in o for p in parts[:1] + parts[-1:]):
            continue
        lost.append(it)
    return lost


def rho3(src: str, out: str) -> float:
    """Fração dos trigramas de palavras da saída que já estavam no original."""
    tok = lambda s: re.findall(r"\w+", s.lower())
    a, b = tok(src), tok(out)
    A = {tuple(a[i:i + 3]) for i in range(len(a) - 2)}
    B = [tuple(b[i:i + 3]) for i in range(len(b) - 2)]
    return round(sum(g in A for g in B) / len(B), 3) if B else 0.0


def detect_lang(text: str) -> str:
    w = re.findall(r"[a-zà-ú]+", text.lower())
    pt = sum(x in {"de", "que", "não", "uma", "para", "com", "os", "das", "são", "é", "em", "da", "do"} for x in w)
    en = sum(x in {"the", "and", "of", "to", "is", "that", "with", "for", "are", "it", "this", "in"} for x in w)
    return "pt" if pt >= en else "en"


# ---------------------------------------------------------------- prompts

PROMPTS = {
    "pt": {
        "notas": ("Reduza o texto a notas telegráficas, uma por linha com '- ', sem conectivos nem artigos. "
                  "Copie números, nomes, citações e URLs exatamente. Linha em branco entre parágrafos. Só as notas."),
        "redator": ("Escreva em português do Brasil, como prosa de artigo, o que estas notas dizem, sem acrescentar "
                    "nada. Palavras e ordem suas, frases de tamanhos variados, sem clichês de IA. Bloco de notas = "
                    "parágrafo; '[lista]' = item com '- '. Números, nomes, citações e URLs exatos. Só o texto."),
        "revisao": ("Revise este trecho em português do Brasil: corrija gramática e frases estranhas, pode trocar "
                    "palavras. Sem acrescentar ou tirar informação; números, citações e listas intactos. Só o trecho."),
        "secao": "Seção: {h}\n\n", "faltou": "\n\nInclua também: {x}",
        "logica": (" Reorganize a lógica do parágrafo: mude a ordem das ideias, junte ou divida frases, inverta "
                   "causa e efeito ou geral e exemplo quando o sentido permitir."),
        "codigos": (" Os códigos entre chaves ({N..} número, {E..} nome ou sigla, {Em..} masculino, {Ef..} feminino, "
                    "{L..} link, {C..} citação) são opacos: use cada um exatamente uma vez, copiado igual, "
                    "sem traduzir, explicar ou inventar códigos."),
        "marcas": (" Enquanto escreve, insira o símbolo ¦ entre palavras, em média um a cada {k} palavras, "
                   "espalhado também no meio das frases. Ele será apagado depois; não o explique."),
    },
    "en": {
        "notas": ("Reduce the text to terse notes, one per line with '- ', no connectives or articles. Copy numbers, "
                  "names, citations and URLs exactly. Blank line between paragraphs. Notes only."),
        "redator": ("Write, as English article prose, what these notes say, adding nothing. Your own words and order, "
                    "varied sentence length, no AI stock phrases. Note block = paragraph; '[list]' = '- ' item. "
                    "Numbers, names, citations and URLs exact. Text only."),
        "revisao": ("Copy-edit this English passage: fix grammar and awkward sentences, you may change words. Add or "
                    "remove nothing; keep numbers, citations and lists. Passage only."),
        "secao": "Section: {h}\n\n", "faltou": "\n\nAlso include: {x}",
        "logica": (" Reorganize the paragraph's logic: change the order of ideas, merge or split sentences, swap "
                   "cause and effect or general and example where the meaning allows."),
        "codigos": (" The codes in braces ({N..} number, {E..} name or acronym, {L..} link, {C..} citation) are "
                    "opaque: use each exactly once, copied verbatim, never translate, explain or invent codes."),
        "marcas": (" While writing, insert the symbol ¦ between words, on average once every {k} words, spread "
                   "inside sentences too. It will be deleted afterwards; do not explain it."),
    },
}

# Palavras de função e muletas que as notas locais descartam. Negação, quantidade
# e conectivos lógicos (mas, porém, portanto, however...) ficam, porque mudam o sentido.
# O redator recebe só isso, na ordem, e refaz toda a costura da frase.
_DROP = {
    "pt": set("""o a os as um uma uns umas de do da dos das no na nos nas ao aos à às em por pelo pela pelos pelas
        para pra com e que isso isto esse essa este esta
        esses essas estes estas aquele aquela seu sua seus suas lhe lhes ele ela eles elas é são foi foram ser
        está estão há tem têm vem além disso
        ressaltar destacar importante vale suma ademais cenário atual crucial fundamental""".split()),
    "en": set("""the a an of to in on at by for with from and that which who this these those it its is are
        was were be been being as such moreover furthermore additionally indeed worth noting important crucial pivotal landscape overall
        conclusion notably""".split()),
}
_KEEP_TOKEN = re.compile(r"https?://\S+|\[[\d,–\- ]+\]|\([^()]{0,60}?\d{4}[a-z]?\)|\S+")


def local_notes(text: str, lang: str) -> str:
    """Notas sem LLM: uma linha por frase, só palavras de conteúdo, números e citações intactos."""
    drop, blocks = _DROP[lang], []
    for para in text.split("\n\n"):
        lines = []
        for raw in para.splitlines():
            is_list = bool(re.match(r"^\s*(?:[-*+]|\d+[.)])\s", raw))
            raw = re.sub(r"^\s*(?:[-*+]|\d+[.)])\s+", "", raw)
            for sent in re.split(r"(?<=[.!?;:])\s+", raw):
                words = [w for w in _KEEP_TOKEN.findall(sent)
                         if w.startswith(("(", "[", "http")) or w.lower().strip(".,;:!?—–-\"“”'") not in drop]
                words = [w.strip(".,;:!?—–\"“”") if not w.startswith(("(", "[")) else w for w in words]
                words = [w for w in words if w]
                if words:
                    lines.append(("[lista] " if is_list and lang == "pt" else "[list] " if is_list else "") + "- " + " ".join(words))
        if lines:
            blocks.append("\n".join(lines))
    return "\n\n".join(blocks)


# ---------------------------------------------------------------- modelos


class Usage:
    def __init__(self):
        self.lock = threading.Lock()
        self.calls = self.tin = self.tout = 0

    def add(self, i: int, o: int):
        with self.lock:
            self.calls += 1
            self.tin += i
            self.tout += o


@dataclass
class Model:
    """``spec`` = provedor:modelo, ex.: openai:gpt-4o-mini, ollama:qwen2.5:7b-instruct,
    anthropic:claude-3-haiku-20240307, gemini:gemini-2.5-flash, compat:meu-modelo (com --base-url)."""
    spec: str
    base_url: str | None = None
    temperature: float | None = None
    timeout: float = 120
    usage: Usage = field(default_factory=Usage)

    def __post_init__(self):
        self.provider, _, self.name = self.spec.partition(":")
        self.provider = self.provider.lower()
        if self.provider not in {"openai", "ollama", "anthropic", "gemini", "compat"} or not self.name:
            raise SystemExit(f"modelo inválido: {self.spec!r} (use provedor:modelo)")
        if self.provider == "anthropic":
            self.url = (self.base_url or os.environ.get("ANTHROPIC_BASE_URL") or "https://api.anthropic.com") + "/v1/messages"
            self.key = os.environ.get("ANTHROPIC_API_KEY", "")
        else:
            default = {"openai": "https://api.openai.com/v1", "ollama": "http://localhost:11434/v1",
                       "gemini": "https://generativelanguage.googleapis.com/v1beta/openai"}.get(self.provider)
            base = self.base_url or (os.environ.get("OPENAI_BASE_URL") if self.provider in {"openai", "compat"} else None) or default
            if not base:
                raise SystemExit("compat precisa de --base-url")
            self.url = base.rstrip("/") + "/chat/completions"
            if self.provider == "gemini":
                self.key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or ""
            else:
                self.key = os.environ.get("LLM_API_KEY") or os.environ.get("OPENAI_API_KEY") or "local"

    def __call__(self, system: str, user: str, max_tokens: int = 2048) -> str:
        if self.provider == "anthropic":
            body = {"model": self.name, "max_tokens": max_tokens, "system": system,
                    "messages": [{"role": "user", "content": user}]}
            headers = {"x-api-key": self.key, "anthropic-version": "2023-06-01"}
        else:
            body = {"model": self.name, "max_tokens": max_tokens,
                    "messages": [{"role": "system", "content": system}, {"role": "user", "content": user}]}
            headers = {"Authorization": f"Bearer {self.key}"}
        if self.temperature is not None:  # vários modelos novos recusam temperature: só manda se pedido
            body["temperature"] = self.temperature
        data = self._post(body, headers)
        if self.provider == "anthropic":
            u = data.get("usage", {})
            self.usage.add(u.get("input_tokens", 0), u.get("output_tokens", 0))
            return "".join(b.get("text", "") for b in data.get("content", [])).strip()
        u = data.get("usage") or {}
        self.usage.add(u.get("prompt_tokens", 0), u.get("completion_tokens", 0))
        return (data["choices"][0]["message"]["content"] or "").strip()

    def _conn(self):
        """Uma conexão HTTP persistente por thread: evita refazer TCP/TLS a cada chamada."""
        c = getattr(_local, self.url, None)
        if c is None:
            u = urllib.parse.urlsplit(self.url)
            cls = http.client.HTTPSConnection if u.scheme == "https" else http.client.HTTPConnection
            c = cls(u.netloc, timeout=self.timeout)
            setattr(_local, self.url, c)
        return c

    def _post(self, body: dict, headers: dict) -> dict:
        path = urllib.parse.urlsplit(self.url).path
        for attempt in range(5):
            c = self._conn()
            try:
                c.request("POST", path, json.dumps(body).encode(), {"Content-Type": "application/json", **headers})
                resp = c.getresponse()
                raw = resp.read()
            except (OSError, http.client.HTTPException) as e:
                c.close()
                setattr(_local, self.url, None)
                if attempt < 4:
                    time.sleep(0.5 * 2 ** attempt)
                    continue
                raise RuntimeError(f"{self.spec}: {e}") from None
            if resp.status == 200:
                return json.loads(raw)
            msg = raw.decode(errors="replace")[:300]
            if resp.status == 400 and "max_tokens" in msg and "max_tokens" in body:  # o-series etc.
                body["max_completion_tokens"] = body.pop("max_tokens")
                continue
            if resp.status in (429, 500, 502, 503, 504, 529) and attempt < 4:
                time.sleep(min(float(resp.getheader("retry-after") or 0) or 1.5 * 2 ** attempt, 30))
                continue
            raise RuntimeError(f"{self.spec}: HTTP {resp.status}: {msg}")
        raise RuntimeError(f"{self.spec}: sem resposta")


_local = threading.local()
_LEAD = re.compile(r"^(?:aqui está|segue|here is|here's)[^\n]*:\s*\n+", re.I)


def _tidy(s: str) -> str:
    s = _LEAD.sub("", s.strip())
    s = re.sub(r"^```\w*\n|\n```$", "", s)
    return s.strip().strip('"“”').strip()


# ---------------------------------------------------------------- marcadores

MARK = "¦"
WINDOW = 4  # janela de contexto suposta da marca (SynthID usa poucos tokens; a da Anthropic não é pública)


def strip_marks(text: str) -> tuple[str, int, float]:
    """Apaga os marcadores. Devolve (texto, nº de marcadores, cobertura).

    Cobertura = fração das palavras que tinham um marcador entre as WINDOW posições
    anteriores. Para essas palavras, o contexto que o redator viu ao escolhê-las não
    existe mais no texto final, e uma marca d'água baseada em contexto perde o sinal nelas."""
    n = text.count(MARK)
    covered = total = 0
    since = WINDOW + 1
    for tok in re.findall(MARK + r"|[^\s" + MARK + r"]+", text):
        if tok == MARK:
            since = 0
            continue
        total += 1
        covered += since < WINDOW
        since += 1
    clean = text.replace(MARK, " ")
    clean = re.sub(r"[ \t]+", " ", clean)
    clean = re.sub(r" +([,.;:!?)\]%])", r"\1", clean)
    clean = re.sub(r"([(\[]) +", r"\1", clean)
    clean = "\n".join(line.strip() for line in clean.split("\n"))
    return clean, n, (covered / total if total else 0.0)


# ---------------------------------------------------------------- sigilo

_ART_M = {"o", "os", "do", "dos", "no", "nos", "ao", "aos", "pelo", "pelos", "um", "the"}
_ART_F = {"a", "as", "da", "das", "na", "nas", "à", "às", "pela", "pelas", "uma"}
_LINK = re.compile(r"\[[^\]\n]{1,200}\]\([^)\s]+\)")
_URL = re.compile(r"https?://[^\s)\]]+")
_CITE = re.compile(r"\([A-ZÀ-Ú][^()\n]{0,60}?\d{4}[a-z]?\)|\[\d+(?:[,–-]\s*\d+)*\]")
_NUM = re.compile(r"(?<![{\w])(?:R\$\s?|US\$\s?)?\d+(?:[.,]\d+)*(?:\s?%)?")
_CAP = r"[A-ZÀ-Ý][\wÀ-ÿ'’-]*"
_NAME = re.compile(r"(?<![{\w])" + _CAP + r"(?: (?:de |da |do |dos |das |of |the )?" + _CAP + r")*")
_CODE = re.compile(r"\{[NELC][mf]?[a-z]{2}\}")
_SENT_START = re.compile(r"(?:^|[.!?:;]\s+|\n\s*(?:[-*+>]\s+|\d+[.)]\s+|\*+)?|\*+|[(\"«“—–|·]\s*)$")


_COMMON_CAPS = set("""antes depois quem quando onde como hoje ainda nas nos essa esse isso ela ele eles elas não sim
    mas porém próximo próxima quantidade página mapa painel análises acontecimentos benefícios diagnóstico observado
    dado dados taxa fonte nota tabela gráfico figura when where who however after before then there this these""".split())


def known_names(md: str) -> set[str]:
    """Nomes próprios e siglas do documento: palavras com maiúscula no meio de frase que
    nunca aparecem em minúscula no texto (isso descarta "Em", "A", "Ela" etc.)."""
    lower = {w for w in re.findall(r"[\wÀ-ÿ'’-]+", md) if w.islower()}
    names = set()
    for m in _NAME.finditer(md):
        words = m.group(0).split(" ")
        names.update(x for x in words if len(x) > 1 and x.isupper())
        if not _SENT_START.search(md[max(0, m.start() - 6):m.start()]):
            names.add(words[0])
    common = _DROP["pt"] | _DROP["en"] | _COMMON_CAPS
    return {n for n in names if len(n) > 1 and (n.isupper() or (n.lower() not in lower and n.lower() not in common))}


class Vault:
    """Troca números, nomes, siglas, links e citações por códigos opacos com chave aleatória.

    O modelo só vê {Nxy}, {Exy}, {Lxy}, {Cxy}: não sabe os valores, não faz conta nem
    "corrige" fatos, e a API nunca recebe os dados reais. O script destroca no fim."""

    def __init__(self, rng):
        self.rng, self.fwd, self.back = rng, {}, {}

    def _code(self, kind: str, value: str) -> str:
        if value in self.fwd:
            return self.fwd[value]
        while True:
            c = "{" + kind + "".join(self.rng.choice("abcdefghijkmnpqrstuvwxyz") for _ in range(2)) + "}"
            if c not in self.back:
                break
        self.fwd[value], self.back[c] = c, value
        return c

    def encode(self, text: str, names: set[str]) -> str:
        text = _LINK.sub(lambda m: self._code("L", m.group(0)), text)
        text = _URL.sub(lambda m: self._code("L", m.group(0)), text)
        text = _CITE.sub(lambda m: self._code("C", m.group(0)), text)
        text = _NUM.sub(lambda m: self._code("N", m.group(0)), text)

        def name(m):
            words = m.group(0).split(" ")
            ok = lambda x: x in names or (len(x) > 1 and x.isupper())
            i = next((k for k, x in enumerate(words) if ok(x)), None)
            if i is None:
                return m.group(0)
            lead, w = " ".join(words[:i]), " ".join(words[i:])
            prev = (text[:m.start()] + " " + lead).split()[-1:] or [""]
            return (lead + " " if lead else "") + self._one(w, prev[0])

        return _NAME.sub(name, text)

    def _one(self, w: str, prev: str) -> str:
        if w in self.fwd:
            return self.fwd[w]
        g = "m" if prev.lower() in _ART_M else "f" if prev.lower() in _ART_F else ""
        return self._code("E" + g, w)


    def decode(self, text: str) -> tuple[str, list[str]]:
        """Destroca. Devolve (texto, códigos que o modelo perdeu)."""
        seen = set()

        def back(m):
            c = m.group(0)
            if c in self.back:
                seen.add(c)
                return self.back[c]
            # gênero trocado pelo modelo ({Em..} escrito como {E..}): procura pelas letras
            for k, v in self.back.items():
                if k[-3:] == c[-3:] and k[1] == c[1]:
                    seen.add(k)
                    return v
            return ""
        text = _CODE.sub(back, text)
        return re.sub(r"[ \t]{2,}", " ", text), [c for c in self.back if c not in seen]


# ---------------------------------------------------------------- motor

def process_piece(p: Piece, lang: str, notas: Model | None, redator: Model, revisao: Model | None,
                  marks: float = 0.0, logic: bool = True, secret: bool = True, names: set[str] | None = None,
                  rng=None) -> Piece:
    """``marks`` = marcadores por 10 palavras (0 desliga). O último modelo a escrever os insere.
    ``secret``: números, nomes, links e citações vão ao modelo como códigos e voltam no fim."""
    P = PROMPTS[lang]
    head = P["secao"].format(h=p.heading) if p.heading else ""
    need = must_keep(p.text)
    vault = Vault(rng or _rng) if secret else None
    src = vault.encode(p.text, names or set()) if vault else p.text
    codes_rule = P["codigos"] if vault and vault.back else ""
    mark_rule = P["marcas"].format(k=max(1, round(10 / marks))) if marks > 0 else ""
    w_sys = P["redator"] + (P["logica"] if logic else "") + codes_rule + ("" if revisao else mark_rule)
    r_sys = P["revisao"] + codes_rule + mark_rule

    def clean(s: str) -> tuple[str, str, list[str]]:
        """(texto ainda codificado, texto final, o que faltou na forma que pode ir à API)."""
        s, n, cov = strip_marks(_tidy(s))
        p.marks, p.coverage = n, cov
        if vault is None:
            return s, s, missing(need, s)
        final, lost_codes = vault.decode(s)
        return s, final, lost_codes

    try:
        if notas is None:
            notes = local_notes(src, lang)
        else:
            notes = notas(P["notas"] + codes_rule, head + src, max_tokens=64 + len(p.text.split()) * 2)
        cap = int((200 + len(p.text.split()) * 3) * (1 + marks / 10))
        coded, draft, lost = clean(redator(w_sys, head + notes, max_tokens=cap))
        if lost:
            st = (p.marks, p.coverage)
            c2, retry, lost2 = clean(redator(w_sys, head + notes + P["faltou"].format(x="; ".join(lost)), max_tokens=cap))
            if len(lost2) < len(lost):
                coded, draft, lost = c2, retry, lost2
            else:
                p.marks, p.coverage = st
        out = draft
        if revisao is not None:
            p.marks, p.coverage = 0, 0.0
            _, fixed, lost3 = clean(revisao(r_sys, coded, max_tokens=cap))
            if fixed and len(lost3) <= len(lost) and len(fixed.split()) > 0.6 * len(draft.split()):
                out, lost = fixed, lost3
            else:  # ficou a versão do redator, que não tinha marcadores
                p.marks, p.coverage = 0, 0.0
        p.out = out
        p.lost = missing(need, out) + ([vault.back[c] for c in lost if vault.back[c] not in need] if vault else [])
    except Exception as e:  # noqa: BLE001 - o lote não para por um trecho
        p.error = str(e)
        p.out = p.text
    return p


_rng = __import__("random").SystemRandom()


def run(files: list[Path], outdir: Path, notas: Model | None, redator: Model | None, revisao: Model | None,
        jobs: int, max_words: int, fmt: str | None, only_clean: bool, lang_opt: str,
        marks: float = 0.0, logic: bool = True, secret: bool = True) -> dict:
    outdir.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    docs = []
    for f in files:
        md, nzw = strip_invisible(read_any(f))
        lang = lang_opt if lang_opt != "auto" else detect_lang(md)
        docs.append({"src": f, "md": md, "lang": lang, "invisiveis": nzw, "pieces": make_pieces(md, max_words),
                     "names": known_names(md) if secret else set()})

    if not only_clean:
        todo = [(p, d["lang"], d["names"]) for d in docs for p in d["pieces"] if p.rewrite]
        if not jobs:  # API: tudo de uma vez. Ollama atende poucos pedidos por vez (OLLAMA_NUM_PARALLEL).
            local = any(m is not None and m.provider == "ollama" for m in (notas, redator, revisao))
            jobs = 8 if local else max(1, min(256, len(todo)))
        with ThreadPoolExecutor(max_workers=jobs) as pool:
            futs = [pool.submit(process_piece, p, lang, notas, redator, revisao, marks, logic, secret, names)
                    for p, lang, names in todo]
            done = 0
            for fu in futs:
                fu.result()
                done += 1
                print(f"\r  trechos: {done}/{len(futs)}  ({time.time() - t0:.1f}s)", end="", file=sys.stderr)
            print(file=sys.stderr)

    report = {"versao": __version__, "arquivos": []}
    for d in docs:
        ext = fmt or (d["src"].suffix.lower() if d["src"].suffix.lower() in {".md", ".txt", ".docx"} else ".md")
        dst = outdir / (d["src"].stem + ext)
        if only_clean and d["src"].suffix.lower() in {".docx", ".pdf", ".png", ".jpg", ".jpeg"} and fmt is None:
            data = d["src"].read_bytes()
            cleaner = {".docx": clean_docx_bytes, ".pdf": clean_pdf_bytes}.get(d["src"].suffix.lower(), clean_image_bytes)
            dst = outdir / d["src"].name
            dst.write_bytes(cleaner(data))
            out_md = d["md"]
        else:
            out_md = "\n\n".join(p.out if (p.rewrite and p.out) else p.text for p in d["pieces"]).strip() + "\n"
            write_any(out_md, dst)
        rw = [p for p in d["pieces"] if p.rewrite]
        report["arquivos"].append({
            "arquivo": str(d["src"]), "saida": str(dst), "idioma": d["lang"], "invisiveis_removidos": d["invisiveis"],
            "trechos": len(rw), "trechos_com_erro": sum(bool(p.error) for p in rw),
            "rho_3gramas": rho3("\n".join(p.text for p in rw), "\n".join(p.out or p.text for p in rw)) if rw and not only_clean else None,
            "itens_perdidos": [x for p in rw for x in p.lost],
            "marcadores": sum(p.marks for p in rw) if marks else None,
            "cobertura_marcadores": round(sum(p.coverage * len(p.out.split()) for p in rw)
                                          / max(1, sum(len(p.out.split()) for p in rw)), 3) if marks else None,
            "erros": [p.error for p in rw if p.error][:3],
        })
    report["segundos"] = round(time.time() - t0, 2)
    seen = {}
    for m in (notas, redator, revisao):
        if m is not None and id(m.usage) not in seen:
            seen[id(m.usage)] = (m.spec, m.usage)
    report["uso"] = [{"modelo": s, "chamadas": u.calls, "tokens_entrada": u.tin, "tokens_saida": u.tout}
                     for s, u in seen.values()]
    return report


def expand(paths: list[str]) -> list[Path]:
    out = []
    for s in paths:
        p = Path(s)
        if p.is_dir():
            out += sorted(x for x in p.iterdir() if x.suffix.lower() in {".md", ".txt", ".docx", ".pdf"})
        elif p.exists():
            out.append(p)
        else:
            raise SystemExit(f"não achei: {s}")
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Tira a marca d'água de LLM de artigos, em lote.")
    ap.add_argument("entradas", nargs="+", help="arquivos ou pastas (.md .txt .docx .pdf)")
    ap.add_argument("-o", "--saida", default="limpos", help="pasta de saída (padrão: limpos/)")
    ap.add_argument("-m", "--modelo", help="modelo redator (provedor:modelo); é o último a escrever")
    ap.add_argument("--redator", help="sinônimo de -m")
    ap.add_argument("--notas", default="local",
                    help="quem faz as notas: 'local' (padrão, sem LLM, grátis) ou provedor:modelo")
    ap.add_argument("--revisao", help="modelo de uma revisão extra depois do redator (padrão: nenhuma)")
    ap.add_argument("--marcadores", type=float, default=0, metavar="N",
                    help="marcadores por 10 palavras que o último modelo insere e o script apaga (0 a 10; "
                         "quanto mais, melhor a remoção da marca do próprio redator, e pior a fluidez). Padrão: 0")
    ap.add_argument("--sem-sigilo", action="store_true",
                    help="manda números, nomes, links e citações em claro (padrão: vão como códigos com chave)")
    ap.add_argument("--manter-logica", action="store_true",
                    help="não pede ao redator para reorganizar a ordem das ideias e das frases")
    ap.add_argument("--base-url", help="endpoint para provedor compat/openai (OpenRouter, Groq, LM Studio...)")
    ap.add_argument("--temperatura", type=float, help="só se o modelo aceitar")
    ap.add_argument("-j", "--paralelo", type=int, default=0,
                    help="chamadas simultâneas (padrão: todos os trechos de uma vez, até 256; baixe se der 429)")
    ap.add_argument("--palavras", type=int, default=150, help="tamanho máximo do trecho em palavras (padrão: 150)")
    ap.add_argument("--formato", choices=["md", "txt", "docx"], help="formato de saída (padrão: o da entrada)")
    ap.add_argument("--idioma", choices=["auto", "pt", "en"], default="auto")
    ap.add_argument("--so-limpar", action="store_true", help="só tira invisíveis e metadados, sem LLM")
    ap.add_argument("--relatorio", help="grava o relatório JSON neste arquivo")
    a = ap.parse_args(argv)

    notas = redator = revisao = None
    if not a.so_limpar:
        spec = a.redator or a.modelo
        if not spec:
            raise SystemExit("informe -m provedor:modelo, ou use --so-limpar")
        pool: dict[str, Model] = {}
        mk = lambda s: pool.setdefault(s, Model(s, a.base_url, a.temperatura))
        redator = mk(spec)
        notas = None if a.notas.lower() == "local" else mk(a.notas)
        revisao = mk(a.revisao) if a.revisao and a.revisao.lower() not in {"nao", "não", "no", "off"} else None

    files = expand(a.entradas)
    rep = run(files, Path(a.saida), notas, redator, revisao, a.paralelo, a.palavras,
              "." + a.formato if a.formato else None, a.so_limpar, a.idioma,
              max(0.0, min(10.0, a.marcadores)), not a.manter_logica, not a.sem_sigilo)
    for f in rep["arquivos"]:
        flag = " ERRO" if f["trechos_com_erro"] else ""
        lost = f" perdeu: {', '.join(f['itens_perdidos'][:5])}" if f["itens_perdidos"] else ""
        rho = f" rho={f['rho_3gramas']}" if f["rho_3gramas"] is not None else ""
        if f.get("cobertura_marcadores") is not None:
            rho += f" marcadores={f['marcadores']} cobertura={f['cobertura_marcadores']:.0%}"
        print(f"{Path(f['saida']).name}: {f['trechos']} trechos{rho}{flag}{lost}")
    for u in rep["uso"]:
        print(f"{u['modelo']}: {u['chamadas']} chamadas, {u['tokens_entrada']} tokens in, {u['tokens_saida']} out")
    print(f"{len(files)} arquivos em {rep['segundos']}s")
    if a.relatorio:
        Path(a.relatorio).write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
    return 1 if any(f["trechos_com_erro"] for f in rep["arquivos"]) else 0


if __name__ == "__main__":
    sys.exit(main())
