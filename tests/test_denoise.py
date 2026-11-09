import io, sys, zipfile
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import denoise as d


class Fake:
    """Modelo falso: responde por etapa com funções dadas."""
    def __init__(self, notas=None, redator=None, revisao=None):
        self.f = {"notas": notas, "redator": redator, "revisao": revisao}
        self.usage = d.Usage(); self.spec = "fake:x"; self.provider = "fake"; self.calls = []
    def __call__(self, system, user, max_tokens=0):
        stage = "notas" if system.startswith(("Reduza", "Reduce")) else "redator" if system.startswith(("Escreva", "Write", "Neste texto", "Some items")) else "revisao"
        self.calls.append(stage); self.usage.add(1, 1)
        return self.f[stage](user)


def test_strip_invisible():
    s, n = d.strip_invisible("a​b\U000E0041c d﻿")
    assert s == "abc d" and n == 4


def test_pieces_keep_headings_code_and_split():
    md = "# T\n\nPrimeiro parágrafo com várias palavras aqui.\n\n```\ncode x y z\n```\n\n## S\n\n" + \
         "\n\n".join("Parágrafo número %d com bastante texto para contar." % i for i in range(30))
    ps = d.make_pieces(md, max_words=40)
    assert ps[0].text == "# T" and not ps[0].rewrite
    assert any(p.text.startswith("```") and not p.rewrite for p in ps)
    rw = [p for p in ps if p.rewrite]
    assert all(len(p.text.split()) <= 40 or p.text.count("\n\n") == 0 for p in rw)
    assert rw[-1].heading == "S" and rw[0].heading == "T"


def test_must_keep_and_missing():
    t = "Em 2023, 67% das escolas (Silva, 2023) usaram [3]; veja https://x.org/a."
    items = d.must_keep(t)
    assert {"2023", "67%", "(Silva, 2023)", "[3]", "https://x.org/a"} <= set(items)
    assert d.missing(items, "Segundo Silva (2023), 67 % das escolas em 2023 [3] https://x.org/a") == []
    assert "67%" in d.missing(items, "Em 2023 a maioria (Silva, 2023) [3] https://x.org/a")


def test_retry_when_number_lost():
    tries = []
    def red(u):
        tries.append(u)
        return "Texto sem o número." if len(tries) == 1 else "Texto com 42 itens."
    m = Fake(notas=lambda u: "- 42 itens", redator=red, revisao=lambda u: u)
    p = d.process_piece(d.Piece("Foram contados 42 itens no total.", True), "pt", m, m, m, secret=False)
    assert p.out == "Texto com 42 itens." and p.lost == [] and "Itens que faltam: 42" in tries[1]


def test_revision_dropped_if_it_loses_facts():
    m = Fake(notas=lambda u: "- 42", redator=lambda u: "São 42 itens.", revisao=lambda u: "São muitos itens.")
    p = d.process_piece(d.Piece("Há 42 itens na caixa.", True), "pt", m, m, m, secret=False)
    assert p.out == "São 42 itens."


def test_writer_never_sees_original_prose():
    seen = []
    m = Fake(notas=lambda u: "- fato um", redator=lambda u: (seen.append(u), "Novo texto.")[1], revisao=lambda u: u)
    d.process_piece(d.Piece("A frase original muito específica aqui.", True), "pt", m, m, None, secret=False)
    assert "original muito específica" not in seen[0]


def test_error_keeps_original_and_is_reported():
    def boom(u): raise RuntimeError("HTTP 401")
    m = Fake(notas=boom)
    p = d.process_piece(d.Piece("Texto qualquer com palavras.", True), "pt", m, m, None, secret=False)
    assert p.error and p.out == "Texto qualquer com palavras."


def test_end_to_end(tmp_path):
    src = tmp_path / "in"; src.mkdir()
    (src / "a.md").write_text("# Título\n\nVale ressal​tar que 10 alunos (Silva, 2020) vieram.\n", encoding="utf-8")
    m = Fake(notas=lambda u: "- 10 alunos vieram (Silva, 2020)", redator=lambda u: "Vieram 10 alunos (Silva, 2020).", revisao=lambda u: u)
    rep = d.run([src / "a.md"], tmp_path / "out", m, m, m, 4, 350, None, False, "auto", secret=False)
    out = (tmp_path / "out" / "a.md").read_text(encoding="utf-8")
    assert out == "# Título\n\nVieram 10 alunos (Silva, 2020).\n"
    f = rep["arquivos"][0]
    assert f["invisiveis_removidos"] == 1 and f["itens_perdidos"] == [] and f["idioma"] == "pt"
    assert rep["uso"][0]["chamadas"] == 3


def test_docx_out_has_no_metadata(tmp_path):
    import pytest
    pytest.importorskip("docx")
    p = tmp_path / "x.docx"
    d.write_any("# T\n\nUm parágrafo.\n\n- item um\n- item dois", p)
    z = zipfile.ZipFile(p)
    core = z.read("docProps/core.xml").decode()
    assert "python-docx" not in core and "creator" not in core
    assert all(i.date_time[0] == 1980 for i in z.infolist())
    assert "Um parágrafo." in d.read_any(p)


def test_png_metadata_stripped():
    import struct, zlib
    def chunk(t, b): return struct.pack(">I", len(b)) + t + b + struct.pack(">I", zlib.crc32(t + b))
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", b"\0" * 13) + chunk(b"tEXt", b"c2pa\0x") + chunk(b"caBX", b"m") + chunk(b"IEND", b"")
    out = d.clean_image_bytes(png)
    assert b"tEXt" not in out and b"caBX" not in out and out.endswith(chunk(b"IEND", b""))


def test_model_spec_parsing():
    m = d.Model("ollama:qwen2.5:7b-instruct")
    assert m.name == "qwen2.5:7b-instruct" and m.url.startswith("http://localhost:11434")
    a = d.Model("anthropic:claude-3-haiku-20240307")
    assert a.url.endswith("/v1/messages")
    g = d.Model("gemini:gemini-2.5-flash")
    assert g.url == "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"


def test_local_notes_drop_glue_keep_facts():
    n = d.local_notes("Vale ressaltar que 67% das escolas (Silva, 2023) já usam IA.\n\n- Um item da lista", "pt")
    assert "(Silva, 2023)" in n and "67%" in n and "Vale" not in n and " que " not in n
    assert "\n\n[lista] - " in n


def test_local_notes_is_default_and_saves_a_call(tmp_path):
    src = tmp_path / "a.md"; src.write_text("Foram 10 alunos (Silva, 2020) que vieram ontem.\n", encoding="utf-8")
    m = Fake(redator=lambda u: "Vieram 10 alunos ontem (Silva, 2020).")
    rep = d.run([src], tmp_path / "o", None, m, None, 0, 150, None, False, "pt", secret=False)
    assert m.calls == ["redator"] and rep["arquivos"][0]["itens_perdidos"] == []


def test_strip_marks_and_coverage():
    s, n, cov = d.strip_marks("Os ¦ alunos vieram ¦, e 10 ¦ ficaram (Silva, ¦ 2020).\n¦ Fim.")
    assert s == "Os alunos vieram, e 10 ficaram (Silva, 2020).\nFim." and n == 5
    assert 0.9 <= cov <= 1.0
    assert d.strip_marks("sem marcador nenhum aqui")[2] == 0.0


def test_markers_asked_from_last_writer_and_removed(tmp_path):
    systems = []
    def red(u): return "Vieram ¦ 10 alunos ¦ ontem (Silva, 2020)."
    class Spy(Fake):
        def __call__(self, system, user, max_tokens=0):
            systems.append(system); return super().__call__(system, user, max_tokens)
    s = Spy(redator=red)
    p = d.process_piece(d.Piece("Foram 10 alunos (Silva, 2020) que vieram ontem.", True), "pt", None, s, None, marks=5, secret=False)
    assert "¦" in systems[0] and "um a cada 2 palavras" in systems[0] and "Reorganize" in systems[0]
    assert "¦" not in p.out and p.marks == 2 and p.lost == []
    systems.clear()
    d.process_piece(d.Piece("Foram 10 alunos.", True), "pt", None, s, None, marks=0, logic=False, secret=False)
    assert "¦" not in systems[0] and "Reorganize" not in systems[0]


def test_vault_round_trip_and_api_never_sees_data():
    import random
    md = "O IBGE contou 103,5 milhões de pessoas em 2026, segundo o Banco Central (Silva, 2023). Veja [a série](https://x.org/a)."
    v = d.Vault(random.Random(1))
    enc = v.encode(md, d.known_names(md))
    for secret in ("IBGE", "103,5", "2026", "Banco Central", "Silva", "https://x.org/a"):
        assert secret not in enc
    assert d.Vault  # códigos têm forma {Xab}
    assert all(d._CODE.fullmatch(c) for c in v.back)
    back, lost = v.decode(enc)
    assert back == md and lost == []


def test_sigilo_end_to_end_writer_sees_only_codes():
    seen = []
    def red(u):
        seen.append(u)
        codes = re.findall(r"\{[NELC][mf]?[a-z]{2}\}", u)
        return "Texto novo com " + " e ".join(codes) + "."
    import re
    m = Fake(redator=red)
    p = d.process_piece(d.Piece("Em 2023, o IBGE contou 42 escolas (Silva, 2020).", True), "pt", None, m, None,
                        names={"IBGE"})
    assert "2023" not in seen[0] and "IBGE" not in seen[0] and "Silva" not in seen[0]
    assert "2023" in p.out and "IBGE" in p.out and "(Silva, 2020)" in p.out and p.lost == []


def test_sigilo_lost_code_reported_and_retry_uses_codes_only():
    calls = []
    def red(u):
        calls.append(u)
        return "Texto sem os códigos."
    m = Fake(redator=red)
    p = d.process_piece(d.Piece("O IBGE contou 42 escolas.", True), "pt", None, m, None, names={"IBGE"})
    assert len(calls) == 3 and all("42" not in c and "IBGE" not in c for c in calls[1:])
    assert "42" in p.lost and "IBGE" in p.lost


def test_sigilo_forte_hides_direction_and_round_trips():
    import random
    md = "A inflação subiu 2% e o desemprego caiu, mas o salário não acompanhou. Rates rose; jobs fell."
    v = d.Vault(random.Random(3), strong=True)
    enc = v.encode(md, set())
    for w in ("subiu", "caiu", "não", "rose", "fell"):
        assert w not in enc
    assert v.decode(enc) == (md, [])


def test_limiter_pause_recovers():
    import time as _t
    lim = d.Limiter()
    lim.pause(0.2)
    t0 = _t.monotonic(); lim.wait(); first = _t.monotonic() - t0
    t0 = _t.monotonic(); lim.wait(); second = _t.monotonic() - t0
    assert first >= 0.15 and second < 0.05


def test_gate_halves_on_concurrency_error_and_recovers():
    g = d.Gate(start=64)
    g.busy = 40
    g.too_many()
    assert g.limit == 20
    g.busy = 0
    for _ in range(40):
        g.success()
    assert g.limit == 30
