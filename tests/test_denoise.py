import re
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
        codes = re.findall(r"\{[NAELC][mf]?[a-z]{2}\}", u)
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


def test_year_and_quantity_codes_keep_their_slot_in_notes():
    import random
    v = d.Vault(random.Random(1))
    enc = v.encode("O Brasil chegou a 93,1 milhões em 1970.", {"Brasil"})
    assert "93,1" not in enc and "1970" not in enc
    assert re.search(r"\{N[a-z]{2}\} em \{A[a-z]{2}\}", enc)
    notes = d.local_notes(enc, "pt")
    assert re.search(r"em \{A[a-z]{2}\}", notes) and re.search(r"[Oo] \{Em[a-z]{2}\}", notes)


def test_decode_contracts_article_and_drops_fake_codes():
    import random
    v = d.Vault(random.Random(2))
    enc = v.encode("O Brasil cresceu.", {"Brasil"})
    c = re.search(r"\{Em[a-z]{2}\}", enc).group(0)
    out, lost = v.decode(f"A população de {c} começará sua {{P..}}.")
    assert out == "A população do Brasil começará sua." and lost == []


def test_echoed_headings_are_removed():
    assert d._no_heads("# Um século\nSeção: Um século\nUm século\n\nTexto real.", "Um século") == "Texto real."


def test_lead_keeps_italic_byline_and_sources_untouched():
    md = ("# Título\n\n*Em 2024 nasceram 2,39 milhões de brasileiros, o menor número em 29 anos de registro nacional de nascidos vivos.*\n\n"
          "wasemu · Demografia · 3 de novembro de 2026\n\n"
          "Fontes: [SIDRA 5932](https://sidra.ibge.gov.br/tabela/5932), [6579](https://sidra.ibge.gov.br/tabela/6579).\n\n"
          "Texto comum do artigo com várias palavras para reescrever aqui.")
    ps = d.make_pieces(md, 150)
    lead = [p for p in ps if p.wrap]
    assert len(lead) == 1 and lead[0].wrap == "*" and not lead[0].text.startswith("*")
    assert any(p.text.startswith("wasemu") and not p.rewrite for p in ps)
    assert any(p.text.startswith("Fontes") and not p.rewrite for p in ps)
    m = Fake(redator=lambda u: "Nova linha fina.")
    out = d.process_piece(lead[0], "pt", None, m, None, secret=False)
    assert out.out == "*Nova linha fina.*"


def test_number_with_unit_is_one_code_and_nomes_mode_keeps_numbers():
    import random
    v = d.Vault(random.Random(4))
    enc = v.encode("O menor número em 29 anos.", set())
    assert "anos" not in enc and len(v.back) == 1
    v2 = d.Vault(random.Random(4), numbers=False)
    enc2 = v2.encode("O menor número em 29 anos, diz o IBGE.", {"IBGE"})
    assert "29 anos" in enc2 and "IBGE" not in enc2


def test_model_meta_notes_are_removed():
    s = "Texto bom.\n\n---\n**Nota:** O item 1996 não tem lugar marcado."
    assert d._no_heads(s, "") == "Texto bom."


def test_decode_fixes_repeated_unit_missing_em_and_stray_prefix():
    import random
    v = d.Vault(random.Random(5))
    enc = v.encode("O menor número em 29 anos. Chegou a 93,1 milhões em 1970.", set())
    n29 = v.fwd["29 anos"]; n93 = v.fwd["93,1 milhões"]; a70 = v.fwd["1970"]
    out, lost = v.decode(f"O menor em {n29} anos. Atingiu {n93} {a70}. Em A1940 havia menos.")
    assert out == "O menor em 29 anos. Atingiu 93,1 milhões em 1970. Em 1940 havia menos." and lost == []


def test_codes_rule_lists_only_present_kinds():
    seen = []
    def red(u):
        return u
    class Cap(Fake):
        def __call__(self, system, user, max_tokens=0):
            seen.append(system)
            return super().__call__(system, user, max_tokens)
    m = Cap(redator=red)
    d.process_piece(d.Piece("O IBGE contou 42 escolas.", True), "pt", None, m, None, names={"IBGE"})
    assert " P palavra" not in seen[0] and "A ano" not in seen[0] and "N número" in seen[0]


def test_decode_inserts_de_between_scale_number_and_noun():
    import random
    v = d.Vault(random.Random(6))
    v.encode("Nasceram 2,39 milhões de brasileiros e 3 escolas.", set())
    c = v.fwd["2,39 milhões"]
    out, _ = v.decode(f"Foram {c} nascimentos, {c} mais que antes.")
    assert out.startswith("Foram 2,39 milhões de nascimentos, 2,39 milhões mais")


def test_short_italic_note_untouched():
    ps = d.make_pieces("Texto comum do artigo com várias palavras aqui.\n\n*Estimado a partir dos óbitos de 2019.*", 150)
    assert ps[-1].text == "*Estimado a partir dos óbitos de 2019.*" and not ps[-1].rewrite


def test_de_only_when_source_had_it_and_not_before_participle():
    import random
    v = d.Vault(random.Random(7))
    v.encode("Foram 2,32 milhões de bebês e 817 mil crianças a menos.", set())
    a, b = v.fwd["2,32 milhões"], v.fwd["817 mil"]
    out, _ = v.decode(f"Os {a} observados e os {a} bebês; {b} crianças.")
    assert out == "Os 2,32 milhões observados e os 2,32 milhões de bebês; 817 mil crianças."


def test_pair_check_catches_swapped_years_and_retry_fixes():
    import random
    v = d.Vault(random.Random(8))
    src = v.encode("Foi de 2,43 milhões em 2019 para 2,12 milhões em 2024.", set())
    pairs = d.code_pairs(src)
    n1, n2, y1, y2 = v.fwd["2,43 milhões"], v.fwd["2,12 milhões"], v.fwd["2019"], v.fwd["2024"]
    assert pairs == [(n1, y1), (n2, y2)]
    assert d.broken_pairs(pairs, f"Em {y2} eram {n1}, contra {n2} em {y1}.") != []
    assert d.broken_pairs(pairs, f"Em {y1} eram {n1}; em {y2}, {n2}.") == []
    outs = iter([f"Em {y2} eram {n1}, contra {n2} em {y1}.", f"Em {y1} eram {n1}; em {y2}, {n2}."])
    m = Fake(redator=lambda u: next(outs))
    p = d.process_piece(d.Piece("Foi de 2,43 milhões em 2019 para 2,12 milhões em 2024.", True), "pt", None, m, None,
                        rng=random.Random(8))
    assert "2019 eram 2,43 milhões" in p.out and not any("par trocado" in x for x in p.lost)


def test_prosa_mode_sends_coded_prose_not_notes():
    seen = []
    def red(u):
        seen.append(u)
        return "Reescrito."
    class Cap(Fake):
        def __call__(self, system, user, max_tokens=0):
            if system.startswith("Reescreva"):
                seen.append(user)
                return "Reescrito."
            return super().__call__(system, user, max_tokens)
    p = d.process_piece(d.Piece("Foram 817 mil crianças a menos em 2024.", True), "pt", "prosa", Cap(redator=red), None)
    assert "a menos" in seen[0] and "817" not in seen[0] and p.out == "Reescrito."


def test_stray_braces_removed_when_source_had_none():
    import random
    v = d.Vault(random.Random(9))
    enc = v.encode("O IBGE estimou.", {"IBGE"})
    c = re.search(r"\{E[mf]?[a-z]{2}\}", enc).group(0)
    assert v.decode(f"As estimativas do {c}}} saíram.")[0] == "As estimativas do IBGE saíram."


def test_structure_recipe_is_drawn_and_ai_patterns_counted():
    import random
    seen = []
    class Cap(Fake):
        def __call__(self, system, user, max_tokens=0):
            seen.append(system)
            return "Texto."
    d.process_piece(d.Piece("Texto comum com várias palavras para reescrever.", True), "pt", None, Cap(), None,
                    secret=False, rng=random.Random(1))
    assert "sorteada" in seen[0] and "'não é X, é Y'" in seen[0]
    assert d.ai_patterns("Não é só crise, é mudança. Juros, câmbio e inflação.", "pt") == 2
    assert d.ai_patterns("A inflação subiu em outubro.", "pt") == 0


def test_flipped_claims_aligned_by_codes():
    src = "A taxa {Nab} subiu em {Acd}. O {Emef} não aprovou a lei."
    assert d.flipped_claims(src, "Em {Acd}, a taxa {Nab} caiu. A lei não passou no {Emef}.", "pt") == [
        "A taxa {Nab} subiu em {Acd}."]
    assert d.flipped_claims(src, "Em {Acd} a taxa {Nab} cresceu. O {Emef} não aprovou a lei.", "pt") == []
    # pseudo-negação: "não só" não inverte
    assert d.flipped_claims("O {Eab} cresceu.", "O {Eab} não só cresceu, mas acelerou.", "pt") == []


def test_meaning_retry_and_judge():
    outs = iter(["A taxa {X} caiu em {Y}.", "A taxa {X} subiu em {Y}."])
    import random
    calls = []
    def red(u):
        calls.append(u)
        t = next(outs)
        codes = re.findall(r"\{[NA][a-z]{2}\}", u)
        return t.replace("{X}", codes[0]).replace("{Y}", codes[1])
    p = d.process_piece(d.Piece("A taxa 5% subiu em 2020.", True), "pt", None, Fake(redator=red), None,
                        rng=random.Random(3))
    assert p.out == "A taxa 5% subiu em 2020." and len(calls) == 2 and "mudaram de sentido" in calls[1]
    class J:
        def __call__(self, system, user, max_tokens=0):
            return "OK" if "subiu" in user.split("REESCRITA:")[1] else "Comparando as duas versões:\n- A taxa {Nab} subiu."
    assert d.judge(J(), "A taxa {Nab} subiu.", "A taxa {Nab} subiu.", "pt") == []
    assert d.judge(J(), "A taxa {Nab} subiu.", "A taxa {Nab} caiu.", "pt") == ["A taxa {Nab} subiu."]


def test_homoglyphs_folded_only_in_mixed_words():
    t, n = d.strip_invisible("O Bаnco subiu. Россия fica.")
    assert t == "O Banco subiu. Россия fica." and n == 1


def test_form_issues_and_dash_and_em_rules():
    assert d.form_issues("- a\n- b\n- c", "a, b e c.", "pt")
    assert d.form_issues("Texto.", "Será?  Sim.", "pt")
    assert d.form_issues("- a\n- b", "- x\n- y", "pt") == []
    import random
    v = d.Vault(random.Random(11))
    v.encode("Em 2024 nasceram 2,86 milhões em 2010.", set())
    a, n = v.fwd["2024"], v.fwd["2,86 milhões"]
    out, _ = v.decode(f"Durante {a}, foram {n} bebês.")
    assert out == "Durante 2024, foram 2,86 milhões de bebês."


def test_list_block_is_own_piece_and_units_get_de():
    ps = d.make_pieces("Parágrafo comum com várias palavras aqui.\n\n- item um longo\n- item dois longo\n\nOutro parágrafo comum aqui mesmo.", 150)
    assert [p.text.startswith("- ") for p in ps if p.rewrite] == [False, True, False]
    import random
    v = d.Vault(random.Random(12))
    v.encode("Colheu 14,3 milhões em 2020.", set())
    c = v.fwd["14,3 milhões"]
    assert v.decode(f"Foram {c} toneladas.")[0] == "Foram 14,3 milhões de toneladas."


def test_list_items_come_back_in_place_or_original_kept():
    class Cap(Fake):
        def __init__(self, reply):
            super().__init__(); self.reply = reply
        def __call__(self, system, user, max_tokens=0):
            return self.reply(user)
    lst = "- **Nascimentos.** Caem desde o pico.\n- **Óbitos.** Sobem com a idade."
    ok = Cap(lambda u: "[2] Os óbitos crescem com o envelhecimento.\n[1] Os nascimentos recuam desde o pico.")
    p = d.process_piece(d.Piece(lst, True), "pt", "prosa", ok, None, secret=False)
    assert p.out == "- Os nascimentos recuam desde o pico.\n- Os óbitos crescem com o envelhecimento."
    bad = Cap(lambda u: "Um parágrafo só, juntando tudo.")
    p = d.process_piece(d.Piece(lst, True), "pt", "prosa", bad, None, secret=False)
    assert p.out == lst and p.lost == ["(lista mantida como estava)"]


def test_fatos_mode_summarizes_then_reframes_and_restores_lost_code():
    seen = {}
    class Cap(Fake):
        def __call__(self, system, user, max_tokens=0):
            if system.startswith("Escreva"):
                seen["sys"] = system
            return super().__call__(system, user, max_tokens)
    def notas(u):
        seen["src"] = u
        codes = re.findall(r"\{N\w+\}", u)
        return "- foram " + codes[0] + " crianças a menos"  # o resumo perde o ano
    def red(u):
        seen["notes"] = u
        return " ".join(re.findall(r"\{[NA]\w+\}", u)[:2]) + " crianças a menos."
    import random
    p = d.process_piece(d.Piece("Foram 817 mil crianças a menos em 2024.", True), "pt", "fatos",
                        Cap(notas=notas, redator=red), None, rng=random.Random(1))
    assert "Ponto de vista deste trecho" in seen["sys"] and "Agrupe as ideias" in seen["sys"]
    assert re.search(r"\{A\w+\}", seen["notes"]), "frase com o ano perdido volta às notas"
    assert "817 mil" in p.out and "2024" in p.out and not p.error


def test_fatos_keeps_paragraphs_and_fixes_doubled_prep():
    seen = {}
    class Cap(Fake):
        def __call__(self, system, user, max_tokens=0):
            if system.startswith("Escreva"):
                seen["sys"] = system
            return super().__call__(system, user, max_tokens)
    src = "Em 2015 nasceram 3 milhões de crianças.\n\nEm 2016 foram 2 milhões de crianças."
    red = lambda u: "Os nascimentos caíram entre em 2015 e em 2016, de 3 milhões crianças a 2 milhões crianças."
    p = d.process_piece(d.Piece(src, True), "pt", "fatos", Cap(notas=lambda u: u, redator=red), None, secret=False)
    assert "2 parágrafos" in seen["sys"]
    assert "entre 2015 e 2016" in p.out and "3 milhões de crianças" in p.out and "milhões de de" not in p.out


def test_style_examples_go_to_writer_and_copy_is_redone(tmp_path):
    (tmp_path / "a.md").write_text("# Título\n\n" + " ".join(["palavra"] * 10) + "\n\n" +
                                   "O governo anunciou ontem, depois de semanas de conversa com os estados, que o plano "
                                   "de obras seria revisto porque faltava dinheiro e sobrava pressa, embora ninguém "
                                   "quisesse admitir isso em público naquela manhã de chuva fina no planalto central.",
                                   encoding="utf-8")
    style = d.load_style(tmp_path)
    assert len(style) == 1 and style[0].startswith("O governo")
    systems = []
    class Cap(Fake):
        def __call__(self, system, user, max_tokens=0):
            systems.append(system)
            return super().__call__(system, user, max_tokens)
    outs = iter(["Caiu a venda naquela manhã de chuva fina no planalto.", "A venda caiu 5% no mês."])
    p = d.process_piece(d.Piece("A venda caiu 5% no mês.", True), "pt", None,
                        Cap(redator=lambda u: next(outs)), None, secret=False, style=style)
    assert "FORMA" in systems[0] and "O governo anunciou" in systems[0]
    assert "O governo anunciou" not in systems[1]
    assert p.out == "A venda caiu 5% no mês." and any("exemplo" in x for x in p.lost)


def test_copied_ignores_ngrams_from_source():
    assert not d.copied("a casa de pedra do rio", ["uma casa de pedra do rio azul"], "a casa de pedra do rio")
    assert d.copied("vimos a casa de pedra do rio", ["uma casa de pedra do rio azul"], "vimos tudo")


def test_no_before_year_becomes_em():
    red = lambda u: "No 2015 nasceram 3 crianças, e no 2016 menos."
    p = d.process_piece(d.Piece("Em 2015 nasceram 3 crianças, e em 2016 menos.", True), "pt", None,
                        Fake(redator=red), None, secret=False)
    assert p.out.startswith("Em 2015") and "em 2016" in p.out
