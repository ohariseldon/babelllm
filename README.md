# denoise.py (llmnoiser v0.5 leve)

Um script só (`denoise.py`, biblioteca padrão do Python) que tira a marca d'água
de LLM de um lote de artigos. Não é humanizador: o texto continua sendo seu,
feito com ajuda de IA; o que sai é o carimbo estatístico do modelo de origem.

## Uso

```bash
# API barata (o redator é o último a escrever: não use o modelo que gerou o artigo)
OPENAI_API_KEY=... python denoise.py artigos/ -o limpos/ -m openai:gpt-4o-mini

# modelo local (o mais seguro contra marca; roda 8 trechos por vez)
python denoise.py artigos/ -o limpos/ -m ollama:qwen2.5:7b-instruct

# Gemini (chave em GEMINI_API_KEY)
GEMINI_API_KEY=... python denoise.py artigos/ -o limpos/ -m gemini:gemini-2.5-flash

# Claude antigo, OpenRouter/Groq/LM Studio (qualquer endpoint compatível com OpenAI)
ANTHROPIC_API_KEY=... python denoise.py artigos/ -m anthropic:claude-3-haiku-20240307
LLM_API_KEY=... python denoise.py artigos/ -m compat:meta-llama/llama-3.1-8b-instruct --base-url https://openrouter.ai/api/v1

# mais fundo (custa mais): notas feitas por LLM e um passe extra de revisão
python denoise.py artigos/ -m ollama:qwen2.5:7b-instruct --notas openai:gpt-4o-mini --revisao ollama:llama3.2:3b

# só limpeza grátis (invisíveis + metadados de .docx/.pdf/.png/.jpg), sem LLM
python denoise.py artigo.docx figura.png -o limpos/ --so-limpar
```

Opções: `-j N` chamadas simultâneas (padrão: todos os trechos de uma vez, até 256;
baixe se a API devolver 429), `--palavras 150` tamanho do trecho, `--formato docx`,
`--idioma pt|en`, `--relatorio rel.json`. Entradas: .md .txt .docx .pdf.

## O que ele faz, por trecho de até ~150 palavras

1. **Limpeza grátis**: caracteres invisíveis (largura zero, tags, bidi, seletores) saem;
   .docx gravado sai sem autor, datas e partes C2PA.
2. **Notas locais, sem LLM**: cada frase vira só as palavras de conteúdo. Saem artigos,
   preposições e muletas ("vale ressaltar", "além disso"); ficam números, nomes, citações,
   negação e conectivos lógicos (mas, portanto, however).
3. **Redator**: um modelo escreve o trecho do zero a partir das notas, sem ver a prosa
   original. Se faltar número, citação ou URL, tenta de novo uma vez.

Uma chamada por trecho. Títulos, código e tabelas não são tocados. Todos os trechos
de todos os artigos saem ao mesmo tempo, em conexões HTTP persistentes.

## Marcadores removíveis (`--marcadores N`)

Com `--marcadores N`, o último modelo a escrever (o redator, ou o revisor se houver)
insere o símbolo `¦` cerca de N vezes a cada 10 palavras, e o script apaga todos
antes de gravar. Uma marca d'água do tipo SynthID escolhe cada token olhando os
poucos tokens anteriores. Os tokens escritos logo depois de um `¦` foram escolhidos
num contexto que não existe mais no texto final, então ali o detector vê cara ou coroa.

- **Quanto mais marcadores, melhor a remoção.** Com `2.5` (um a cada 4 palavras),
  quase toda palavra fica coberta se a janela da marca for de até 4 tokens; com `5` ou `10`,
  janelas menores também caem. O preço é fluidez: o modelo escreve "em volta" dos marcadores.
- O relatório mostra quantos marcadores o modelo de fato pôs e a `cobertura_marcadores`
  (fração das palavras com um marcador nas 4 posições anteriores).
- Serve para quando o redator é um modelo que também marca texto. Com redator local
  que não marca, não é preciso.
- Não afeta marcas que não dependem de contexto (lista verde fixa).
- Custa até N×10% a mais de tokens de saída.

## Reorganização da lógica

Por padrão o redator também muda a ordem das ideias, junta e divide frases e inverte
causa e efeito quando o sentido deixa. Isso mira marcas de nível de frase ou de sentido
que ainda não conhecemos (que olham a sequência de frases, não só as palavras).
`--manter-logica` desliga.

## Regra que importa

O último modelo a escrever deixa a marca dele. Então o redator (e o revisor, se usar) não podem
ser o modelo que gerou o artigo, nem outro que também marque texto. Modelo local
(Ollama) é o caminho mais seguro; uma API de outro fornecedor também serve.

## O relatório

Por arquivo: `rho_3gramas` (fração dos trigramas da saída que já estavam no original;
quanto menor, menos da sequência original sobreviveu), `itens_perdidos` (números e
citações que não voltaram, confira esses à mão), invisíveis removidos e erros.
No fim: chamadas e tokens por modelo. O script não diz que o texto está "seguro":
nenhuma ferramenta sem a chave da marca pode afirmar isso.

## Tempo e custo (medido com servidor simulado)

20 artigos de ~1.000 palavras, API simulada (0,6 s até o primeiro token).

| versão | chamadas | tokens (entrada + saída) | 90 tokens/s | 50 tokens/s |
|---|---|---|---|---|
| v0.3 (notas LLM + redator + revisão, trechos de 350) | 234 | 90 mil + 79 mil | 21,2 s | 35,2 s |
| **v0.4 padrão** (notas locais + redator, trechos de 150) | 241 | 35 mil + 33 mil | **3,4 s** | **5,5 s** |
| v0.4 com `--revisao` | 482 | 61 mil + 51 mil | 5,5 s | 8,8 s |
| v0.4 com `-j 64` | 241 | 35 mil + 33 mil | 9,5 s | 15,1 s |

O que é estimativa: a contagem de tokens vem do simulador (1,4 token por palavra;
português gasta um pouco mais) e a API real pode limitar pedidos simultâneos (429).
Com modelo local o tempo depende da máquina.

Para repetir: `python bench/make_articles.py bench/artigos`, `python bench/mock_server.py --tps 90 &`,
`python denoise.py bench/artigos -o /tmp/o -m compat:mock --base-url http://127.0.0.1:8765/v1`.

## Testes

`python -m pytest -q tests` (15 testes, sem rede).

## O que saiu da v0.2 e por quê

Modo de regras e perturbação (estragava o texto e relatava "seguro"), métricas de
cobertura e z estimado, tradução de ida e volta, rodízio de modelos, tipografia,
comparação de textos, embeddings e perplexidade. Na v0.4 as notas por LLM e a revisão viraram opcionais.

Troca da v0.4: as notas locais mantêm as palavras de conteúdo do original na ordem, então
sobra mais vocabulário do original do que com notas feitas por LLM. A costura das frases
é toda refeita. Se quiser reescrita mais funda, use `--notas provedor:modelo`. (A v0.2, com todos esses caminhos, não faz parte deste repositório.)
