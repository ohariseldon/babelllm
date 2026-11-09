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

__version__ = "0.7.0"

# ---------------------------------------------------------------- Unicode

_INVISIBLE = re.compile(
    "[­͏؜ᅟᅠ឴឵᠋-᠏​-‏‪-‮"
    "⁠-⁤⁦-⁯ㅤ︀-️﻿ﾠ￹-￻"
    "\U0001d173-\U0001d17a\U000e0000-\U000e007f\U000e0100-\U000e01ef]"
)
_ODD_SPACE = re.compile("[   -   　]")


_HOMO = {"а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "х": "x", "у": "y", "і": "i", "ј": "j", "һ": "h",
         "ѕ": "s", "ԁ": "d", "ɡ": "g", "ο": "o", "α": "a", "ν": "v", "ρ": "p",
         "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T",
         "Х": "X", "Ѕ": "S", "І": "I", "Α": "A", "Β": "B", "Ε": "E", "Η": "H", "Κ": "K", "Μ": "M", "Ν": "N",
         "Ο": "O", "Ρ": "P", "Τ": "T", "Χ": "X"}
_HOMO_WORD = re.compile(r"\w*[" + "".join(_HOMO) + r"]\w*")


def _fold_homoglyphs(text: str) -> tuple[str, int]:
    """Troca letras cirílicas/gregas parecidas com latinas, só em palavras que misturam alfabetos
    (texto grego ou russo de verdade fica como está)."""
    n = 0

    def fix(m):
        nonlocal n
        w = m.group(0)
        if not re.search(r"[A-Za-zÀ-ÿ]", w):
            return w
        n += sum(c in _HOMO for c in w)
        return "".join(_HOMO.get(c, c) for c in w)
    return _HOMO_WORD.sub(fix, text), n


def strip_invisible(text: str) -> tuple[str, int]:
    """Remove invisíveis (largura zero, tags, bidi, seletores...), espaços exóticos e letras disfarçadas."""
    n = len(_INVISIBLE.findall(text)) + len(_ODD_SPACE.findall(text))
    text = _ODD_SPACE.sub(" ", _INVISIBLE.sub("", text))
    text, h = _fold_homoglyphs(text)
    n += h
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
    if "\n" not in s.strip() and len(s.split()) < 14 and (" · " in s or not re.search(r"[.!?:…]\W*$", s)):
        return False  # linha de assinatura, data ou crédito: fica como está
    if re.match(r"\W*(Fontes?|Notas?|Nota técnica|Metodologia|Sources?|Notes?|Methodology)\b\W", s, re.I) \
            or len(_LINK.findall(s) + _URL.findall(s)) * 12 > len(s.split()):
        return False  # bloco de fontes ou nota técnica: só links e códigos de série, nada a reescrever
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
    wrap: str = ""
    patterns: int = 0
    meaning: list[str] = field(default_factory=list)


_WRAP = re.compile(r"^(\*{1,2}|_)(?![*_\s])([^*\n]+?)(?<![*_\s])\1$")


def make_pieces(md: str, max_words: int) -> list[Piece]:
    """Junta blocos de texto seguidos (sem cruzar título) até ~max_words."""
    pieces, buf, heading = [], [], ""

    def flush():
        if buf:
            pieces.append(Piece("\n\n".join(buf), True, heading))
            buf.clear()

    for blk in split_blocks(md):
        w = _WRAP.match(blk.strip())
        if w and len(w.group(2).split()) < 15:  # nota curta em itálico ("*Estimado a partir de...*"): intacta
            flush()
            pieces.append(Piece(blk, False))
            continue
        if w and _is_text(w.group(2)):  # linha fina em itálico/negrito: trecho próprio, volta com a mesma ênfase
            flush()
            pieces.append(Piece(w.group(2), True, heading, wrap=w.group(1)))
            continue
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
                    "parágrafo; '[lista]' = item com '- '. Números, nomes, citações e URLs exatos. Não acrescente fatos, "
                    "causas, exemplos nem opiniões; tom sóbrio, sem palavras pomposas; tamanho parecido com o pedido. "
                    "Sem títulos. Só o texto."),
        "redator_prosa": ("Reescreva em português do Brasil, como prosa de artigo, o texto abaixo com outras palavras "
                          "e outra construção de frase, dizendo exatamente o mesmo: mesmos fatos, mesmas relações entre "
                          "eles (quem, quanto, quando, a mais ou a menos, causa), sem acrescentar nada. Frases de "
                          "tamanhos variados, sem clichês de IA, tom sóbrio, sem títulos, sem alongar e sem omitir nenhum item, nome ou número. "
                          "Só o texto."),
        "revisao": ("Revise este trecho em português do Brasil: corrija gramática e frases estranhas, pode trocar "
                    "palavras. Sem acrescentar ou tirar informação; números, citações e listas intactos. Só o trecho."),
        "secao": "(contexto, não escreva: seção «{h}»)\n\n", "faltou": "\n\nInclua também: {x}",
        "tamanho": "\n\n(cerca de {n} palavras)",
        "sentido": "\n\nAtenção, numa versão anterior estas frases mudaram de sentido; mantenha a direção (sobe/desce, "
                   "sim/não, antes/depois) exatamente como no original: {x}",
        "juiz": ("Você compara um ORIGINAL com uma REESCRITA. Os códigos entre chaves são opacos. Liste, uma por linha, "
                 "cada frase do original que a reescrita contradiz (sentido invertido, número ou ano trocado de papel, "
                 "quem fez o quê trocado, causa inventada) ou omite. Paráfrase com outras palavras e outra ordem não é "
                 "problema. Antes de apontar, pergunte: as duas podem ser verdade ao mesmo tempo? Se nada estiver "
                 "errado, responda só OK."),
        "juiz_user": "ORIGINAL:\n{a}\n\nREESCRITA:\n{b}",
        "pares": "\n\nAtenção: cada número fica na mesma frase do seu ano, sem outro número entre eles: {x}.",
        "reparo": ("Neste texto faltam alguns itens. Insira cada um deles uma vez, no lugar que a nota indica, "
                   "mudando o mínimo possível do resto. Mantenha todos os códigos entre chaves já presentes. "
                   "Responda só com o texto completo corrigido."),
        "reparo_user": "Itens que faltam: {x}\n\nNotas onde eles aparecem:\n{n}\n\nTexto:\n{t}",
        "logica": (" Reorganize a lógica do parágrafo: mude a ordem das ideias, junte ou divida frases, "
                   "sem mudar quem fez o quê nem a que dado cada número se refere. Troque por sinônimos ou outra formulação "
                   "pelo menos metade das palavras de conteúdo das notas (substantivos, verbos, adjetivos), "
                   "menos os termos técnicos."),
        "estrutura": {
            "base": (" Estrutura deste trecho, sorteada: {a}; ritmo das frases: {r}; {e}. Não use a fórmula "
                     "'não é X, é Y' nem abra com uma negação para depois corrigir, a não ser que a negação seja o fato; "
                     "sem pergunta retórica, sem frase de efeito com dois-pontos, sem frase final de resumo."),
            "abertura": ["comece pelo dado mais concreto", "comece pela consequência", "comece pela causa",
                         "comece pelo contexto de tempo ou de lugar", "comece por uma constatação curta",
                         "comece pelo ator principal (quem faz ou mede)"],
            "ritmo": ["curta, longa, média", "longa, curta, longa", "média, média, curta, longa",
                      "curta, curta, longa", "longa, média, curta"],
            "enumeracao": ["onde houver três itens seguidos, separe um deles em outra frase",
                           "evite listas de três; use dois itens ou quatro", "prefira juntar itens numa frase só, sem série"],
        },
        "codigos": (" Os códigos entre chaves ({k}) são opacos: use cada um exatamente uma vez, copiado igual, "
                    "sem traduzir, explicar ou inventar códigos. Mantenha as preposições da nota em volta de cada "
                    "código ('de', 'em', 'desde'): 'em' + ano continua sendo o ano daquele dado. Não repita depois "
                    "do código a unidade que ele já contém."),
        "tipos": {"N": "N número, com a unidade se houver", "A": "A ano", "E": "E nome ou sigla",
                  "Em": "Em nome masculino: leva artigo (o, do, no)", "Ef": "Ef nome feminino: leva artigo (a, da, na)",
                  "L": "L link", "C": "C citação",
                  "P": "P palavra de sentido oculto: ocupa o lugar de uma palavra só e mantém a função dela"},
        "marcas": (" Enquanto escreve, insira o símbolo ¦ entre palavras, em média um a cada {k} palavras, "
                   "espalhado também no meio das frases. Ele será apagado depois; não o explique."),
    },
    "en": {
        "notas": ("Reduce the text to terse notes, one per line with '- ', no connectives or articles. Copy numbers, "
                  "names, citations and URLs exactly. Blank line between paragraphs. Notes only."),
        "redator": ("Write, as English article prose, what these notes say, adding nothing. Your own words and order, "
                    "varied sentence length, no AI stock phrases. Note block = paragraph; '[list]' = '- ' item. "
                    "Numbers, names, citations and URLs exact. Add no facts, causes, examples or opinions; plain tone; "
                    "length close to the one asked. No headings. Text only."),
        "redator_prosa": ("Rewrite the English text below in other words and other sentence structure, saying exactly "
                          "the same: same facts, same relations between them (who, how much, when, more or less, cause), "
                          "adding nothing. Varied sentence length, no AI stock phrases, plain tone, no headings, no padding and no item, name or number left out "
                          "(keep every item). Text only."),
        "revisao": ("Copy-edit this English passage: fix grammar and awkward sentences, you may change words. Add or "
                    "remove nothing; keep numbers, citations and lists. Passage only."),
        "secao": "(context, do not write: section «{h}»)\n\n", "faltou": "\n\nAlso include: {x}",
        "tamanho": "\n\n(about {n} words)",
        "sentido": "\n\nCareful, in an earlier version these sentences changed meaning; keep the direction (up/down, "
                   "yes/no, before/after) exactly as in the original: {x}",
        "juiz": ("You compare an ORIGINAL with a REWRITE. Brace codes are opaque. List, one per line, each original "
                 "sentence the rewrite contradicts (meaning reversed, number or year in the wrong role, who did what "
                 "swapped, invented cause) or leaves out. Paraphrase in other words and order is fine. Before flagging, "
                 "ask: could both be true at the same time? If nothing is wrong, reply only OK."),
        "juiz_user": "ORIGINAL:\n{a}\n\nREWRITE:\n{b}",
        "pares": "\n\nCareful: each number stays in the same sentence as its year, no other number between them: {x}.",
        "reparo": ("Some items are missing from this text. Insert each one once, where its note places it, changing "
                   "as little else as possible. Keep every brace code already present. Reply with the full corrected text only."),
        "reparo_user": "Missing items: {x}\n\nNotes where they appear:\n{n}\n\nText:\n{t}",
        "logica": (" Reorganize the paragraph's logic: change the order of ideas, merge or split sentences, "
                   "without changing who did what or which fact each number belongs to. Replace at least half of the notes' "
                   "content words (nouns, verbs, adjectives) with synonyms or other phrasing, except technical terms."),
        "estrutura": {
            "base": (" Structure for this passage, drawn at random: {a}; sentence rhythm: {r}; {e}. Do not use the "
                     "'not X, but Y' formula or open with a negation to correct it, unless the negation is the fact; no "
                     "rhetorical questions, no colon punchlines, no closing summary sentence."),
            "abertura": ["open with the most concrete fact", "open with the consequence", "open with the cause",
                         "open with the time or place", "open with a short plain statement",
                         "open with the main actor (who does or measures)"],
            "ritmo": ["short, long, medium", "long, short, long", "medium, medium, short, long", "short, short, long",
                      "long, medium, short"],
            "enumeracao": ["where three items run together, move one to another sentence",
                           "avoid lists of three; use two items or four", "prefer merging items into one sentence, no series"],
        },
        "codigos": (" The codes in braces ({k}) are opaque: use each exactly once, copied verbatim, never translate, "
                    "explain or invent codes. Keep the note's prepositions around each code: 'in' + year is still the "
                    "year of that fact. Do not repeat after a code the unit it already holds."),
        "tipos": {"N": "N number, with its unit if any", "A": "A year", "E": "E name or acronym", "Em": "E name",
                  "Ef": "E name", "L": "L link", "C": "C citation", "P": "P hidden word, one word, same role"},
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


_NOT_NOUN = {"mais", "menos", "apenas", "depois", "antes", "pois", "após", "através", "nós", "vós", "jamais",
             "quase", "talvez", "ambos", "ambas", "todos", "todas", "outros", "outras", "eles", "elas", "seus", "suas",
             "os", "as", "dos", "das", "nos", "nas", "aos", "às", "pelos", "pelas", "es", "seus", "estes", "estas", "esses",
             "essas", "anos", "meses", "dias", "pontos", "vezes"}
_GLUE = {"de", "do", "da", "dos", "das", "em", "no", "na", "of", "in"}


def local_notes(text: str, lang: str) -> str:
    """Notas sem LLM: uma linha por frase, só palavras de conteúdo, números e citações intactos."""
    drop, blocks = _DROP[lang], []
    for para in text.split("\n\n"):
        lines = []
        for raw in para.splitlines():
            is_list = bool(re.match(r"^\s*(?:[-*+]|\d+[.)])\s", raw))
            raw = re.sub(r"^\s*(?:[-*+]|\d+[.)])\s+", "", raw)
            for sent in re.split(r"(?<=[.!?;:])\s+", raw):
                toks = _KEEP_TOKEN.findall(sent)
                words = [w for i, w in enumerate(toks)
                         if w.startswith(("(", "[", "http", "{")) or w.lower().strip(".,;:!?—–-\"“”'") not in drop
                         or (i + 1 < len(toks) and toks[i + 1].startswith("{"))
                         or (i > 0 and toks[i - 1].rstrip(".,;:").endswith("}") and w.lower() in _GLUE)]
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
    rpm: float = 0

    def __post_init__(self):
        self.limiter = Limiter(self.rpm)
        self.gate = Gate()
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
        tries = 10
        for attempt in range(tries):
            self.limiter.wait()
            c = self._conn()
            try:
                with self.gate:
                    c.request("POST", path, json.dumps(body).encode(), {"Content-Type": "application/json", **headers})
                    resp = c.getresponse()
                    raw = resp.read()
            except (OSError, http.client.HTTPException) as e:
                c.close()
                setattr(_local, self.url, None)
                if attempt < tries - 1:
                    time.sleep(min(0.5 * 2 ** attempt, 20))
                    continue
                raise RuntimeError(f"{self.spec}: {e}") from None
            if resp.status == 200:
                self.gate.success()
                return json.loads(raw)
            msg = raw.decode(errors="replace")[:600]
            if resp.status == 400 and "max_tokens" in msg and "max_tokens" in body:  # o-series etc.
                body["max_completion_tokens"] = body.pop("max_tokens")
                continue
            if resp.status in (429, 500, 502, 503, 504, 529) and attempt < tries - 1:
                # espera o que a API pede (cabeçalho retry-after ou "retryDelay": "17s" do Gemini)
                hint = resp.getheader("retry-after") or (re.findall(r'retryDelay"?\s*:\s*"?(\d+(?:\.\d+)?)s', msg) or [0])[0]
                wait = float(hint or 0) or min(1.5 * 2 ** attempt, 60)
                if resp.status == 429 and "concurrent" in msg.lower():
                    self.gate.too_many()  # limite de simultâneos: menos threads, não menos ritmo
                    time.sleep(min(1.0 + attempt, 10))
                elif resp.status == 429:
                    self.limiter.pause(min(wait, 90))
                else:
                    time.sleep(min(wait + 0.5, 90))
                continue
            raise RuntimeError(f"{self.spec}: HTTP {resp.status}: {msg[:300]}")
        raise RuntimeError(f"{self.spec}: sem resposta")


class Gate:
    """Limite adaptativo de pedidos simultâneos de um modelo (AIMD).

    Começa sem teto prático; quando a API recusa por excesso de pedidos simultâneos
    (429 "concurrent"), corta o teto pela metade; a cada 4 respostas boas, sobe 1."""

    def __init__(self, start: int = 256):
        self.cv, self.limit, self.busy, self.ok, self.cut_at = threading.Condition(), start, 0, 0, 0.0

    def __enter__(self):
        with self.cv:
            while self.busy >= self.limit:
                self.cv.wait()
            self.busy += 1

    def __exit__(self, *exc):
        with self.cv:
            self.busy -= 1
            self.cv.notify_all()

    def success(self):
        with self.cv:
            self.ok += 1
            if self.ok % 4 == 0 and self.limit < 256:
                self.limit += 1
                self.cv.notify_all()

    def too_many(self):
        """Corta pela metade no máximo uma vez a cada 3 s: uma rajada de 429 conta como um corte só."""
        with self.cv:
            now = time.monotonic()
            if now - self.cut_at >= 3:
                self.limit = max(4, min(self.limit, self.busy + 1) // 2)
                self.cut_at = now


class Limiter:
    """Ritmo dos pedidos de um modelo, compartilhado entre as threads.

    ``rpm`` > 0 espaça os pedidos (ex.: 15 no plano grátis do Gemini). Ao receber 429,
    todas as threads param pelo tempo que a API pediu e depois voltam ao ritmo normal:
    não há redução permanente."""

    def __init__(self, rpm: float = 0):
        self.lock, self.next, self.rpm, self.until = threading.Lock(), 0.0, rpm, 0.0

    def wait(self):
        with self.lock:
            now = time.monotonic()
            at = max(now, self.until)
            if self.rpm:
                at = max(at, self.next)
                self.next = at + 60.0 / self.rpm
        if at > now:
            time.sleep(at - now)

    def pause(self, seconds: float):
        with self.lock:
            self.until = max(self.until, time.monotonic() + seconds)


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
_NUM = re.compile(r"(?<![{\w])(?:R\$\s?|US\$\s?)?\d+(?:[.,]\d+)*(?:\s?%|\s(?:mil|milh(?:ão|ões)|bilh(?:ão|ões)|trilh(?:ão|ões)|thousand|million|billion|trillion|p\.p\.|pontos percentuais|pontos?|anos?|mes(?:es)?|mês|dias?|semanas?|décadas?|séculos?|horas?|vezes|years?|months?|days?|weeks?|decades?|centuries|times|points?)(?![\wÀ-ÿ]))?")
_YEAR = re.compile(r"(?:1[5-9]|20)\d\d")
_CAP = r"[A-ZÀ-Ý][\wÀ-ÿ'’-]*"
_NAME = re.compile(r"(?<![{\w])" + _CAP + r"(?: (?:de |da |do |dos |das |of |the )?" + _CAP + r")*")
_CODE = re.compile(r"\{[NAELCP][mf]?[a-z]{2}\}")
_FAKE_CODE = re.compile(r"\{[NAELCPQ][mf]?(?:[a-z]{0,3}|\.\.|…)\}")  # código inventado ou copiado do prompt
# Sigilo forte: palavras que dão direção ou polaridade ao fato (subiu/caiu, mais/menos, não...).
_POL = re.compile(r"(?<![{\w])(?:" + "|".join([
    r"n[ãa]o", r"nunca", r"jamais", r"nenhuma?", r"nada",
    r"sub(?:iu|iram|ia|iam|indo|ir|ido|ida|ir[áa]|e|em)", r"ca(?:iu|[íi]ram|[íi]a|[íi]am|indo|ir|[íi]do|i|em)",
    r"cresc\w*", r"aument\w*", r"diminu\w*", r"redu[zç]\w*", r"recu\w*", r"avan[çc]\w*", r"encolh\w*",
    r"dobr\w*", r"triplic\w*", r"despenc\w*", r"dispar\w*", r"piora\w*", r"piore\w*", r"melhora\w*", r"melhore\w*",
    r"mais", r"menos", r"maior(?:es)?", r"menor(?:es)?", r"melhor(?:es)?", r"pior(?:es)?", r"acima", r"abaixo",
    r"alta", r"altas", r"queda", r"quedas", r"recorde",
    r"not", r"never", r"no", r"none", r"rose", r"rise[sn]?", r"rising", r"fell", r"fall(?:s|en|ing)?",
    r"increase[sd]?", r"decrease[sd]?", r"grew", r"grow(?:s|n|ing)?", r"declin\w*", r"more", r"less", r"fewer",
    r"higher", r"lower", r"better", r"worse", r"above", r"below", r"record",
]) + r")(?![\wÀ-ÿ])", re.I)
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

    def __init__(self, rng, strong: bool = False, numbers: bool = True):
        self.rng, self.fwd, self.back, self.strong, self.numbers = rng, {}, {}, strong, numbers
        self.lang, self.fake, self.de_after = "pt", 0, set()

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
        self.braces = "{" in text or "}" in text
        text = _LINK.sub(lambda m: self._code("L", m.group(0)), text)
        text = _URL.sub(lambda m: self._code("L", m.group(0)), text)
        text = _CITE.sub(lambda m: self._code("C", m.group(0)), text)
        self.src_text = text
        if self.numbers:
            text = _NUM.sub(lambda m: self._code("A" if _YEAR.fullmatch(m.group(0)) else "N", m.group(0)), text)

        def name(m):
            words = m.group(0).split(" ")
            ok = lambda x: x in names or (len(x) > 1 and x.isupper())
            i = next((k for k, x in enumerate(words) if ok(x)), None)
            if i is None:
                return m.group(0)
            lead, w = " ".join(words[:i]), " ".join(words[i:])
            prev = (text[:m.start()] + " " + lead).split()[-1:] or [""]
            return (lead + " " if lead else "") + self._one(w, prev[0])

        text = _NAME.sub(name, text)
        for c in self.back:  # números seguidos de "de" no original ("2,39 milhões de brasileiros")
            if c[1] == "N" and re.search(re.escape(c) + r" de ", text):
                self.de_after.add(c)
        if self.strong:  # cada forma ("subiu", "Subiu") vira um código próprio: a volta é exata
            text = _POL.sub(lambda m: self._code("P", m.group(0)), text)
        return text

    def _one(self, w: str, prev: str) -> str:
        if w in self.fwd:
            return self.fwd[w]
        g = "m" if prev.lower() in _ART_M else "f" if prev.lower() in _ART_F else ""
        return self._code("E" + g, w)


    _PREP = {"em", "de", "desde", "até", "a", "e", "ou", "entre", "para", "após", "antes", "partir", "ano", "anos",
             "no", "do", "dos", "da", "in", "of", "since", "until", "by", "from", "to", "and", "or", "year", "como"}

    def _prev_ok(self, text, m) -> bool:
        prev = re.findall(r"[\wÀ-ÿ{}%]+", text[:m.start()])[-1:]
        return not prev or prev[0].lower() in self._PREP or self.lang != "pt"

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
        def contract(m):
            w, c = m.group(1), m.group(2)
            g = c[2] if c[2] in "mf" else ""
            new = {"de": {"m": "do", "f": "da"}, "em": {"m": "no", "f": "na"}, "por": {"m": "pelo", "f": "pela"}}[w.lower()][g]
            return (new.capitalize() if w[0].isupper() else new) + " " + c
        text = re.sub(r"(?<![\wÀ-ÿ])([Dd]e|[Ee]m|[Pp]or) (\{E[mf][a-z]{2}\})", contract, text)
        def unit(m):  # "{N..} anos" quando o código já é "29 anos"
            c, after = m.group(1), m.group(2)
            v = self.back.get(c, "")
            return c if v.split()[-1:] == [after] else m.group(0)
        text = re.sub(r"(\{N[a-z]{2}\})\s+([\wÀ-ÿ%]+)(?![\wÀ-ÿ])", unit, text)
        text = re.sub(r"(?<=[\wÀ-ÿ}%]) (\{A[a-z]{2}\})(?=[.,;:)]|$)",
                      lambda m: m.group(0) if self._prev_ok(text, m) else " em " + m.group(1), text)
        def de(m):
            c, nxt = m.group(1), m.group(2)
            v = self.back.get(c, "")
            if self.lang == "pt" and c in self.de_after and nxt.lower() not in _NOT_NOUN \
                    and not re.search(r"(?:ad|id)[oa]s$", nxt.lower()):
                return f"{c} de {nxt}"
            return m.group(0)
        text = re.sub(r"(\{N[a-z]{2}\}) ([a-zà-ú]+s)(?![\wÀ-ÿ])", de, text)
        text = _CODE.sub(back, text)
        self.fake = len(_FAKE_CODE.findall(text))
        text = _FAKE_CODE.sub("", text)
        text = re.sub(r"(?<![\wÀ-ÿ])[NA](?=(?:1[5-9]|20)\d\d(?!\d))", "", text)  # "A1970": prefixo de código imitado
        text = re.sub(r"\{([^{}\n]{1,60})\}", r"\1", text)  # {Palavra} que o modelo inventou: tira as chaves
        if not getattr(self, "braces", True):  # o original não tinha chaves: as que sobraram são do modelo
            text = text.replace("{", "").replace("}", "")
        text = re.sub(r" +([,.;:!?)])", r"\1", re.sub(r"[ \t]{2,}", " ", text))
        return text, [c for c in self.back if c not in seen]


# ---------------------------------------------------------------- motor

_PAIR_GAP = 3


def code_pairs(coded: str) -> list[tuple[str, str]]:
    """Pares número→ano do original: '{N..} em {A..}' com até 3 palavras entre eles e nenhum outro código."""
    toks = re.findall(r"\{[NAELCP][mf]?[a-z]{2}\}|[\wÀ-ÿ%]+|[.!?;]", coded)
    pairs = []
    for i, t in enumerate(toks):
        if t.startswith("{N"):
            for j in range(i + 1, min(len(toks), i + 2 + _PAIR_GAP)):
                if toks[j] in ".!?;" or (toks[j].startswith("{") and not toks[j].startswith("{A")):
                    break
                if toks[j].startswith("{A"):
                    pairs.append((t, toks[j]))
                    break
    return pairs


def broken_pairs(pairs: list[tuple[str, str]], coded: str) -> list[tuple[str, str]]:
    """Pares que a saída separou: os dois códigos precisam estar na mesma frase, sem outro número ou ano entre eles."""
    bad = []
    for n, a in pairs:
        ok = False
        for sent in re.split(r"(?<=[.!?;])\s+", coded):
            if n in sent and a in sent:
                i, j = sorted((sent.index(n), sent.index(a)))
                ok = not re.search(r"\{[NA][a-z]{2}\}", sent[i + 5:j])
                break
        if not ok:
            bad.append((n, a))
    return bad


_AI_PAT = {
    "pt": re.compile(r"\bnão (?:é|são|foi|se trata de)\b[^.;]{1,60}?[,;—–-]\s*(?:mas|e sim|e)?\s*(?:é|são|foi)\b"
                     r"|\b\w+, \w+ e \w+\b(?=[.;])|^[^:\n]{3,40}: [A-ZÀ-Ú]|\?\s|\b(?:Em suma|Em resumo|No fim das contas)\b",
                     re.I | re.M),
    "en": re.compile(r"\b(?:is|are|was) not\b[^.;]{1,60}?[,;—–-]\s*(?:but|it)?\s*(?:is|are|was)\b"
                     r"|\b\w+, \w+,? and \w+\b(?=[.;])|^[^:\n]{3,40}: [A-Z]|\?\s|\b(?:In short|Ultimately|In the end)\b",
                     re.I | re.M),
}


def ai_patterns(text: str, lang: str) -> int:
    """Conta fórmulas retóricas típicas de IA: 'não é X, é Y', série de três, dois-pontos de efeito, pergunta, resumo."""
    return len(_AI_PAT[lang].findall(text))


# Checagem de sentido sem modelo (ideia do NegEx: pseudo-negação vem antes da negação).
_NEG = {"pt": {"não", "nao", "nunca", "jamais", "nenhum", "nenhuma", "nada", "ninguém", "sem", "nem"},
        "en": {"not", "no", "never", "none", "nothing", "nobody", "without", "nor", "cannot"}}
_PSEUDO = {"pt": [("não", "só"), ("não", "apenas"), ("não", "somente"), ("não", "necessariamente"),
                  ("não", "mais", "só"), ("nem", "sempre"), ("sem", "dúvida")],
           "en": [("not", "only"), ("not", "just"), ("not", "merely"), ("not", "necessarily"), ("no", "longer", "just")]}
_SIDES = {"pt": [
    ({"subiu", "subiram", "sobe", "sobem", "subir", "aumentou", "aumentaram", "aumenta", "aumento", "cresceu",
      "cresceram", "cresce", "crescem", "crescimento", "alta", "avançou", "avanço", "mais", "maior", "maiores", "acima",
      "superou", "ganhou", "recorde"},
     {"caiu", "caíram", "cai", "caem", "cair", "diminuiu", "diminuíram", "diminui", "queda", "recuou", "recuo",
      "encolheu", "baixa", "menos", "menor", "menores", "abaixo", "perdeu", "redução", "reduziu", "desce", "descer",
      "despencou", "declínio"}),
    ({"antes", "anterior", "precedeu"}, {"depois", "posterior", "após", "seguiu"}),
    ({"todos", "todas", "sempre", "tudo"}, {"nenhum", "nenhuma", "nunca", "nada"}),
    ({"aprovou", "aprovado", "permitido", "venceu"}, {"rejeitou", "rejeitado", "proibido", "perdeu"})],
    "en": [
    ({"rose", "rise", "rises", "increased", "increase", "grew", "growth", "up", "higher", "more", "above", "gained",
      "record", "climbed", "surged"},
     {"fell", "fall", "falls", "decreased", "decrease", "shrank", "decline", "declined", "down", "lower", "less",
      "fewer", "below", "lost", "dropped", "sank"}),
    ({"before", "earlier", "prior"}, {"after", "later", "subsequent"}),
    ({"all", "always", "everything"}, {"none", "never", "nothing"}),
    ({"approved", "allowed", "won"}, {"rejected", "forbidden", "lost"})]}


def _polarity(sent: str, lang: str) -> tuple[bool, tuple[int, ...]]:
    toks = re.findall(r"[a-zà-ÿ]+", sent.lower())
    masked = [False] * len(toks)
    for i in range(len(toks)):
        for ph in _PSEUDO[lang]:
            if toks[i:i + len(ph)] == list(ph):
                for j in range(i, i + len(ph)):
                    masked[j] = True
    neg = sum(1 for i, t in enumerate(toks) if not masked[i] and t in _NEG[lang]) % 2 == 1
    bag = {t for i, t in enumerate(toks) if not masked[i]}
    sides = tuple(3 if (bag & a and bag & b) else 1 if bag & a else 2 if bag & b else 0 for a, b in _SIDES[lang])
    return neg, sides


def flipped_claims(src: str, out: str, lang: str) -> list[str]:
    """Frases do original cuja direção mudou na saída. O alinhamento é pelos códigos: a frase da saída que
    contém todos os códigos da frase original é a mesma afirmação. Sem código, a frase não é checada."""
    split = lambda t: [x for x in re.split(r"(?<=[.!?;])\s+", t) if x.strip()]
    outs, bad = split(out), []
    for s in split(src):
        codes = set(_CODE.findall(s))
        if not codes:
            continue
        hits = [o for o in outs if codes <= set(_CODE.findall(o))]
        if len(hits) != 1:
            continue
        (sn, ss), (on, os_) = _polarity(s, lang), _polarity(hits[0], lang)
        if sn != on or any(a in (1, 2) and b in (1, 2) and a != b for a, b in zip(ss, os_)):
            bad.append(s)
    return bad


def _no_heads(s: str, heading: str) -> str:
    """Tira títulos que o modelo ecoou: linhas com #, «Seção:» e cópias do título da seção."""
    key = lambda x: re.sub(r"[\W_]+", " ", x).strip().lower()
    h = key(heading)
    keep = [l for l in s.split("\n") if not (_HEAD.match(l.lstrip()) or re.match(r"\s*\**(Seção|Section)\s*:", l)
                                           or re.match(r"\s*(?:-{3,}|\*{3,}|_{3,})\s*$", l)
                                           or re.match(r"\s*[*_(\[]*(Nota|Observação|Obs|Note|N\.B)\b\W", l, re.I)
                                           or (h and key(l) == h))]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(keep)).strip()


def process_piece(p: Piece, lang: str, notas: Model | None, redator: Model, revisao: Model | None,
                  marks: float = 0.0, logic: bool = True, secret: bool = True, names: set[str] | None = None,
                  rng=None, verify=None) -> Piece:
    """``marks`` = marcadores por 10 palavras (0 desliga). O último modelo a escrever os insere.
    ``secret``: números, nomes, links e citações vão ao modelo como códigos e voltam no fim."""
    P = PROMPTS[lang]
    head = P["secao"].format(h=p.heading) if p.heading else ""
    need = must_keep(p.text)
    vault = Vault(rng or _rng, strong=(secret == "forte"), numbers=(secret != "nomes")) if secret else None
    if vault:
        vault.lang = lang
    src = vault.encode(p.text, names or set()) if vault else p.text
    codes_rule = ""
    if vault and vault.back:
        kinds = []
        for c in vault.back:
            k = c[1:3] if c[1] == "E" and len(c) == 6 else c[1]
            if P["tipos"][k] not in kinds:
                kinds.append(P["tipos"][k])
        codes_rule = P["codigos"].format(k="; ".join(kinds))
    mark_rule = P["marcas"].format(k=max(1, round(10 / marks))) if marks > 0 else ""
    struct = ""
    if logic:  # a estrutura vem do nosso sorteio, não do hábito do modelo
        E, r = P["estrutura"], rng or _rng
        struct = E["base"].format(a=r.choice(E["abertura"]), r=r.choice(E["ritmo"]), e=r.choice(E["enumeracao"]))
    w_sys = P["redator"] + (P["logica"] + struct if logic else "") + codes_rule + ("" if revisao else mark_rule)
    r_sys = P["revisao"] + codes_rule + mark_rule

    def clean(s: str) -> tuple[str, str, list[str]]:
        """(texto ainda codificado, texto final, o que faltou na forma que pode ir à API)."""
        s, n, cov = strip_marks(_no_heads(_tidy(s), p.heading))
        p.marks, p.coverage = n, cov
        if vault is None:
            return s, s, missing(need, s)
        final, lost_codes = vault.decode(s)
        return s, final, lost_codes

    try:
        if notas == "prosa":  # o redator vê o texto codificado inteiro: mais fiel, menos distante do original
            notes = src
            w_sys = P["redator_prosa"] + w_sys[len(P["redator"]):]
        elif notas is None:
            notes = local_notes(src, lang)
        else:
            notes = notas(P["notas"] + codes_rule, head + src, max_tokens=64 + len(p.text.split()) * 2)
        cap = int((200 + len(p.text.split()) * 3) * (1 + marks / 10))
        size = P["tamanho"].format(n=len(p.text.split()))
        coded, draft, lost = clean(redator(w_sys, head + notes + size, max_tokens=cap))
        pairs = code_pairs(src) if vault else []
        bad = broken_pairs(pairs, coded)
        if bad:  # número trocado de ano: refaz uma vez, avisando quais pares andam juntos
            hint = P["pares"].format(x="; ".join(f"{n} com {a}" if lang == "pt" else f"{n} with {a}" for n, a in bad))
            st = (p.marks, p.coverage)
            c2, d2, l2 = clean(redator(w_sys, head + notes + size + hint, max_tokens=cap))
            b2 = broken_pairs(pairs, c2)
            if len(b2) < len(bad) and len(l2) <= len(lost):
                coded, draft, lost, bad = c2, d2, l2, b2
            else:
                p.marks, p.coverage = st
        for _ in range(2 if lost else 0):  # reparo: o redator insere só o que faltou, no texto que já escreveu
            if not lost:
                break
            st = (p.marks, p.coverage)
            where = "\n".join(l for l in notes.splitlines() if any(x in l for x in lost)) or notes
            c2, retry, lost2 = clean(redator(P["reparo"] + codes_rule, P["reparo_user"].format(
                x="; ".join(lost), n=where, t=coded), max_tokens=cap))
            grew = len(c2.split()) - len(coded.split())
            glued = any(re.search(r",(?:\s+[\wÀ-ÿ]+){0,2}\s+" + re.escape(c) + r"\s*[.!?]", c2)
                        for c in lost if c.startswith("{"))
            if not glued and len(lost2) < len(lost) and -0.2 * len(coded.split()) < grew <= 6 * len(lost) + 8:
                coded, draft, lost = c2, retry, lost2
            else:
                p.marks, p.coverage = st
        flips = flipped_claims(src, coded, lang) if vault else []
        issues = list(flips)
        if verify is not None and vault:  # juiz opcional: compara original e saída, ambos ainda codificados
            issues += judge(verify, src, coded, lang)
        if issues:  # sentido invertido ou contradição: refaz uma vez, dizendo o que conferir
            st = (p.marks, p.coverage)
            hint = P["sentido"].format(x=" | ".join(issues[:4]))
            c2, d2, l2 = clean(redator(w_sys, head + notes + size + hint, max_tokens=cap))
            f2 = flipped_claims(src, c2, lang)
            i2 = f2 + (judge(verify, src, c2, lang) if verify is not None else [])
            if len(i2) < len(issues) and len(l2) <= len(lost) and len(broken_pairs(pairs, c2)) <= len(bad):
                coded, draft, lost, flips, issues = c2, d2, l2, f2, i2
            else:
                p.marks, p.coverage = st
        p.meaning = [vault.decode(x)[0] for x in issues] if vault else []
        out = draft
        if revisao is not None:
            p.marks, p.coverage = 0, 0.0
            _, fixed, lost3 = clean(revisao(r_sys, coded, max_tokens=cap))
            if fixed and len(lost3) <= len(lost) and len(fixed.split()) > 0.6 * len(draft.split()):
                out, lost = fixed, lost3
            else:  # ficou a versão do redator, que não tinha marcadores
                p.marks, p.coverage = 0, 0.0
        if p.wrap:
            out = p.wrap + " ".join(out.split()) + p.wrap
        p.out = out
        if vault and vault.fake:
            lost = lost + ["(palavra trocada por um código inventado)"]
        miss = missing(need, out)
        extra = [vault.back.get(c, c) for c in lost] if vault else []
        extra = [v for v in extra if v.startswith("(") and "código inventado" in v or not (
            _norm(v) in _norm(out) or any(m in v for m in miss))]
        if vault and pairs:
            bad = broken_pairs(pairs, coded)
            extra += [f"(par trocado: {vault.back[n]} / {vault.back[a]})" for n, a in bad]
        p.lost = miss + extra + [f"(sentido a conferir: {x[:90]})" for x in p.meaning]
        p.patterns = ai_patterns(out, lang)
    except Exception as e:  # noqa: BLE001 - o lote não para por um trecho
        p.error = str(e)
        p.out = p.wrap + p.text + p.wrap
    return p


def judge(model, src: str, out: str, lang: str) -> list[str]:
    """Pergunta a um modelo se a reescrita contradiz ou omite algo do original. Devolve as frases do
    original com problema (vazio = ok). Os dois textos vão codificados: o juiz não vê os dados."""
    P = PROMPTS[lang]
    try:
        r = model(P["juiz"], P["juiz_user"].format(a=src, b=out), max_tokens=200)
    except Exception:  # noqa: BLE001 - o juiz é opcional; se falhar, segue
        return []
    lines = [l.strip(" -*•\t") for l in r.strip().splitlines() if l.strip()]
    if not lines or lines[0].upper().startswith("OK"):
        return []
    return [l for l in lines if l.upper() not in {"PROBLEMAS:", "PROBLEMS:"}][:4]


_rng = __import__("random").SystemRandom()


def run(files: list[Path], outdir: Path, notas: Model | None, redator: Model | None, revisao: Model | None,
        jobs: int, max_words: int, fmt: str | None, only_clean: bool, lang_opt: str,
        marks: float = 0.0, logic: bool = True, secret: bool | str = True, verify=None) -> dict:
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
            local = any(m is not None and not isinstance(m, str) and m.provider == "ollama" for m in (notas, redator, revisao))
            jobs = 8 if local else max(1, min(256, len(todo)))
        with ThreadPoolExecutor(max_workers=jobs) as pool:
            futs = [pool.submit(process_piece, p, lang, notas, redator, revisao, marks, logic, secret, names,
                                None, verify)
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
            "formulas_ia": sum(p.patterns for p in rw),
            "marcadores": sum(p.marks for p in rw) if marks else None,
            "cobertura_marcadores": round(sum(p.coverage * len(p.out.split()) for p in rw)
                                          / max(1, sum(len(p.out.split()) for p in rw)), 3) if marks else None,
            "erros": [p.error for p in rw if p.error][:3],
        })
    report["segundos"] = round(time.time() - t0, 2)
    seen = {}
    for m in (notas, redator, revisao, verify):
        if m is not None and not isinstance(m, str) and id(m.usage) not in seen:
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
    ap.add_argument("--notas", default="prosa",
                    help="o que o redator recebe: 'prosa' (padrão: o texto codificado inteiro, o mais fiel), "
                         "'local' (notas sem LLM: reescrita mais distante, erra mais a relação entre fatos) "
                         "ou provedor:modelo")
    ap.add_argument("--verificar", metavar="provedor:modelo",
                    help="juiz opcional que compara original e reescrita (ambos codificados) e manda refazer o "
                         "trecho que mudou de sentido; ~1 chamada a mais por trecho")
    ap.add_argument("--revisao", help="modelo de uma revisão extra depois do redator (padrão: nenhuma)")
    ap.add_argument("--marcadores", type=float, default=0, metavar="N",
                    help="marcadores por 10 palavras que o último modelo insere e o script apaga (0 a 10; "
                         "quanto mais, melhor a remoção da marca do próprio redator, e pior a fluidez). Padrão: 0")
    ap.add_argument("--sigilo", choices=["basico", "nomes", "forte", "nao"], default="basico",
                    help="basico (padrão): números, nomes, siglas, links e citações vão como códigos com chave; "
                         "nomes: só nomes, siglas, links e citações (números em claro; melhor para artigos de dados); "
                         "forte: também as palavras de direção e polaridade (subiu, caiu, mais, menos, não...); "
                         "nao: tudo em claro")
    ap.add_argument("--sem-sigilo", action="store_true", help="o mesmo que --sigilo nao")
    ap.add_argument("--manter-logica", action="store_true",
                    help="não pede ao redator para reorganizar a ordem das ideias e das frases")
    ap.add_argument("--base-url", help="endpoint para provedor compat/openai (OpenRouter, Groq, LM Studio...)")
    ap.add_argument("--temperatura", type=float, help="só se o modelo aceitar")
    ap.add_argument("-j", "--paralelo", type=int, default=0,
                    help="chamadas simultâneas (padrão: todos os trechos de uma vez, até 256; baixe se der 429)")
    ap.add_argument("--palavras", type=int, default=150, help="tamanho máximo do trecho em palavras (padrão: 150)")
    ap.add_argument("--rpm", type=float, default=0,
                    help="teto de pedidos por minuto por modelo (ex.: 15 no plano grátis do Gemini). "
                         "Padrão: sem teto; ao receber 429 o script reduz o ritmo sozinho")
    ap.add_argument("--formato", choices=["md", "txt", "docx"], help="formato de saída (padrão: o da entrada)")
    ap.add_argument("--idioma", choices=["auto", "pt", "en"], default="auto")
    ap.add_argument("--so-limpar", action="store_true", help="só tira invisíveis e metadados, sem LLM")
    ap.add_argument("--relatorio", help="grava o relatório JSON neste arquivo")
    a = ap.parse_args(argv)

    notas = redator = revisao = verify = None
    if not a.so_limpar:
        spec = a.redator or a.modelo
        if not spec:
            raise SystemExit("informe -m provedor:modelo, ou use --so-limpar")
        pool: dict[str, Model] = {}
        mk = lambda s: pool.setdefault(s, Model(s, a.base_url, a.temperatura, rpm=a.rpm))
        redator = mk(spec)
        notas = None if a.notas.lower() == "local" else "prosa" if a.notas.lower() == "prosa" else mk(a.notas)
        revisao = mk(a.revisao) if a.revisao and a.revisao.lower() not in {"nao", "não", "no", "off"} else None
        verify = mk(a.verificar) if a.verificar else None
        last = revisao or redator
        if last.provider in ("anthropic", "gemini") or (last.provider == "openai" and "gpt" in last.spec):
            print(f"aviso: {last.spec} é de um fornecedor que pode marcar o texto que escreve. A marca de origem sai, "
                  "mas a do redator entra. Prefira um modelo local (ollama:...) ou use --marcadores.", file=sys.stderr)

    files = expand(a.entradas)
    rep = run(files, Path(a.saida), notas, redator, revisao, a.paralelo, a.palavras,
              "." + a.formato if a.formato else None, a.so_limpar, a.idioma,
              max(0.0, min(10.0, a.marcadores)), not a.manter_logica,
              False if (a.sem_sigilo or a.sigilo == "nao") else (a.sigilo if a.sigilo in ("forte", "nomes") else True),
              verify)
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
