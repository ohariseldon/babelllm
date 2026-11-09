"""Gera 20 artigos sintéticos de ~1000 palavras (pt e en) para o teste de tempo."""
import random, sys
from pathlib import Path
out = Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
PT = ["Vale ressaltar que a inteligência artificial desempenha um papel crucial na educação moderna, e {n}% das escolas já usam ferramentas desse tipo (Silva, 2023).",
      "Além disso, ela possibilita uma abordagem personalizada para diversos alunos, o que permite que professores otimizem o tempo em sala de aula.",
      "No entanto, existem desafios significativos, como a privacidade dos dados, que precisam ser enfrentados por gestores e famílias.",
      "Em {y}, um levantamento com {m} professores mostrou que a maioria ainda não recebeu formação específica para usar esses sistemas.",
      "Em suma, a tecnologia não é apenas uma tendência, mas também uma transformação profunda e abrangente do modo como se aprende."]
EN = ["It is worth noting that artificial intelligence plays a crucial role in modern healthcare, and {n}% of hospitals already use such tools (Smith, 2022).",
      "Moreover, it enables a personalized approach for many patients, which allows doctors to make better use of their time.",
      "However, there are significant challenges, such as data privacy, that administrators and regulators must address.",
      "In {y}, a survey of {m} clinicians showed that most had not received specific training to use these systems.",
      "In conclusion, the technology is not only a trend but also a deep and far-reaching transformation of care."]
r = random.Random(7)
for i in range(20):
    S = PT if i % 2 == 0 else EN
    parts = [f"# {'Artigo' if S is PT else 'Article'} {i+1}"]
    words = 0; sec = 0
    while words < 1000:
        if words // 330 >= sec: sec += 1; parts.append(f"## {'Seção' if S is PT else 'Section'} {sec}")
        para = " ".join(s.format(n=r.randint(10, 90), y=r.randint(2015, 2025), m=r.randint(100, 5000)) for s in r.sample(S, 4))
        parts.append(para); words += len(para.split())
    (out / f"artigo_{i+1:02d}.md").write_text("\n\n".join(parts) + "\n", encoding="utf-8")
print("ok")
