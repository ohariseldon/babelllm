# babelllm

**Tira a marca d'água estatística de LLM de um artigo ou de um lote de artigos, removendo o carimbo do modelo de origem, com um modelo barato à sua escolha.**

*English summary below.*

Você escreve com ajuda de IA e assume isso. O que você não quer é o carimbo invisível que alguns
fornecedores embutem na escolha das palavras (marcas do tipo SynthID). O babelllm reescreve cada
trecho a partir de notas, com um modelo que você escolhe (local, antigo ou barato), e entrega o mesmo
conteúdo, com os mesmos números, sem a sequência de palavras que carrega a marca.

- **Um arquivo só**, biblioteca padrão do Python, sem dependências obrigatórias.
- **Rápido**: 20 artigos em ~38 s com Claude Haiku 4.5 (medido; ver [Resultados](#resultados)).
- **Barato**: uma chamada de modelo por trecho de ~150 palavras.
- **Sigilo**: números, nomes, siglas, links e citações vão para a API como códigos (`{Nkq}`, `{Exm}`)
  e são destrocados no seu computador. A API nunca vê os dados.
- **Qualquer modelo**: Ollama, OpenAI, Anthropic, Gemini, OpenRouter, Groq, LM Studio, vLLM.

## Instalação

```bash
pip install git+https://github.com/ohariseldon/babelllm            # comando `babelllm`
pip install "babelllm[docs] @ git+https://github.com/ohariseldon/babelllm"  # + .docx e .pdf
```

Ou só baixe o `denoise.py` e rode com `python denoise.py` (Python 3.10+).

## Uso rápido

```bash
# modelo local (o mais seguro: não marca o que escreve)
babelllm artigos/ -o limpos/ -m ollama:qwen2.5:7b-instruct

# API barata de outro fornecedor que não o que gerou o texto
ANTHROPIC_API_KEY=... babelllm artigos/ -o limpos/ -m anthropic:claude-haiku-4-5
OPENAI_API_KEY=...    babelllm artigos/ -o limpos/ -m openai:gpt-4o-mini
GEMINI_API_KEY=...    babelllm artigos/ -o limpos/ -m gemini:gemini-2.5-flash --rpm 15   # plano grátis
LLM_API_KEY=...       babelllm artigos/ -o limpos/ -m compat:meta-llama/llama-3.1-8b-instruct \
                               --base-url https://openrouter.ai/api/v1

# remoção mais forte contra marcas de contexto
babelllm artigos/ -o limpos/ -m anthropic:claude-haiku-4-5 --marcadores 5

# só a limpeza grátis: caracteres invisíveis e metadados (.docx .pdf .png .jpg), sem modelo
babelllm artigo.docx figura.png -o limpos/ --so-limpar
```

Entradas: `.md`, `.txt`, `.docx`, `.pdf` (arquivos ou pastas). Saída no mesmo formato, ou `--formato md|txt|docx`.

## A regra que importa

**O último modelo a escrever deixa a marca dele.** Não use como redator o modelo que gerou o artigo,
nem outro que também marque texto. Modelo local é o caminho mais seguro; se usar uma API que marca,
use `--marcadores` (abaixo).

## Como funciona

Para cada trecho de até ~150 palavras (títulos, código e tabelas não são tocados):

1. **Limpeza**: saem caracteres invisíveis (largura zero, tags Unicode, bidi, seletores de variação) e,
   nos `.docx` gravados, autor, datas e partes C2PA.
2. **Sigilo** (padrão): números, nomes próprios, siglas, links e citações viram códigos com chave
   aleatória. O modelo não sabe os valores, não faz conta e não "corrige" fatos.
3. **Notas locais**, sem modelo: cada frase vira só as palavras de conteúdo. Saem artigos, preposições e
   muletas ("vale ressaltar"); ficam negação e conectivos lógicos.
4. **Redator**: o modelo escreve o trecho do zero a partir das notas, sem ver a prosa original. Ele também
   reorganiza a ordem das ideias e troca por sinônimos pelo menos metade das palavras de conteúdo.
5. **Conferência**: todo código precisa voltar. Se faltar algum, o redator recebe só o que faltou (o código,
   nunca o dado) e insere no próprio texto. O que ainda faltar aparece no relatório.
6. **Destroca** local dos códigos e gravação.

Todos os trechos de todos os arquivos vão ao mesmo tempo, em conexões persistentes. Ao receber 429, o
script pausa pelo tempo que a API pede e retoma.

### Marcadores removíveis (`--marcadores N`)

O redator escreve com o símbolo `¦` cerca de N vezes a cada 10 palavras, e o script apaga todos antes de
gravar. Uma marca do tipo SynthID escolhe cada token olhando os poucos tokens anteriores; depois de um `¦`
apagado, o contexto que o modelo viu não existe mais no texto final e o detector vê cara ou coroa.
Quanto maior N, melhor a remoção e menor a fluidez. O relatório mostra a cobertura real alcançada.

### Sigilo forte (`--sigilo forte`)

Também esconde as palavras de direção e polaridade (subiu, caiu, mais, menos, maior, não, recorde; rose,
fell, more, less, not...). O modelo escreve "a taxa {Pab} {Ncd}" sem saber se a taxa subiu ou caiu.
A volta é exata. Custa fluidez.

## Opções

| opção | o que faz |
|---|---|
| `-m provedor:modelo` | redator (`ollama`, `openai`, `anthropic`, `gemini`, `compat` + `--base-url`) |
| `--marcadores N` | 0 a 10 marcadores removíveis por 10 palavras (padrão 0) |
| `--sigilo basico\|forte\|nao` | o que vai como código para a API (padrão `basico`) |
| `--manter-logica` | não reorganiza ideias nem troca sinônimos |
| `--notas provedor:modelo` | notas feitas por um modelo em vez de locais (reescrita mais funda, ~2× o custo) |
| `--revisao provedor:modelo` | um passe extra de correção depois do redator |
| `--rpm N` | teto de pedidos por minuto (ex.: 15 no plano grátis do Gemini) |
| `-j N` | chamadas simultâneas (padrão: todos os trechos, até 256; Ollama: 8) |
| `--palavras N` | tamanho máximo do trecho (padrão 150) |
| `--idioma auto\|pt\|en` | idioma (padrão: detecta) |
| `--relatorio arquivo.json` | relatório por arquivo |
| `--so-limpar` | só invisíveis e metadados, sem modelo |

O modelo não recebe `temperature` a menos que você passe `--temperatura` (vários modelos novos recusam).

## Relatório

Por arquivo: trechos, erros, `itens_perdidos` (números, nomes ou citações que não voltaram; confira à mão),
`rho_3gramas` (fração dos trigramas da saída que já estavam no original), marcadores e cobertura, invisíveis
removidos. No fim: chamadas e tokens por modelo. O script não diz que o texto está "seguro": sem a chave
de uma marca, ninguém pode afirmar isso.

## Resultados

Medidos em 20 artigos reais em português (~1.500 palavras cada), com Claude Haiku 4.5:

- **Tempo**: 20 artigos em 37,6 s; 4 artigos em 14 s.
- **Tokens**: ~123 mil de entrada e ~66 mil de saída nos 20 artigos.

**Teste de marca d'água com chave própria** (`bench/marca_teste.py`): marcamos o artigo escolhendo, para cada
frase, a paráfrase com mais palavras "verdes" (o princípio do SynthID), e medimos o z-score antes e depois.
z > 4 = marca detectada; |z| < 2 = indistinguível de texto sem marca.

| artigo de demografia | sem contexto | contexto de 1 palavra | contexto de 3 palavras |
|---|---|---|---|
| original | 0,2 | −0,1 | 0,1 |
| marcado | 2,1 | 2,6 | **4,2** |
| babelllm padrão | 1,2 | 1,5 | 2,1 |
| babelllm `--marcadores 5` | 1,7 | 1,9 | **0,1** |

Num segundo artigo, a marca de contexto 3 foi de 5,2 para 0,2 com `--marcadores 5`. A lista verde fixa
(sem contexto) é a que mais resiste, porque as palavras de conteúdo sobrevivem; a troca de sinônimos no
redator mira esse caso.

Para repetir: `python bench/marca_teste.py artigo.md -m anthropic:claude-haiku-4-5`.
Para medir só o tempo sem gastar tokens: `bench/mock_server.py` simula uma API compatível com OpenAI.

## Limites

- **Detectores de estilo não são o alvo.** Classificadores de "texto de IA" (do tipo Pangram/GPTZero)
  olham o estilo, e o texto de saída é escrito por um modelo. O babelllm remove marcas d'água, não a
  aparência de texto de IA.
- **Marcas que não dependem de contexto** (lista verde fixa) caem menos que as de contexto.
- **Marcas futuras**: não há como testar contra o que ainda não foi publicado. O desenho (notas sem a
  prosa, códigos, reorganização, sinônimos, marcadores) ataca os mecanismos conhecidos, não garante o resto.
- **Qualidade**: modelos pequenos às vezes perdem um código ou escrevem pior em volta deles. Confira os
  `itens_perdidos` e leia o texto antes de publicar.
- **C2PA em imagens**: os metadados saem; credenciais duráveis (marca nos pixels) não.

## Uso responsável

Feito para quem publica o próprio trabalho com ajuda de IA e assume isso. Não use para se passar por
autor de texto alheio, nem para burlar regras de uma escola, revista ou empregador que proíbam IA.

## Desenvolvimento

```bash
pip install -e '.[docs,dev]'
python -m pytest -q tests      # 20 testes, sem rede
```

Contribuições são bem-vindas: abra uma issue antes de mudanças grandes.

---

## English summary

**babelllm** removes statistical LLM text watermarks (SynthID-style) from your own AI-assisted articles,
in batch, using any cheap model you choose (local Ollama, OpenAI, Anthropic, Gemini, any OpenAI-compatible
endpoint). It is a single stdlib-only Python file.

For each ~150-word chunk it strips invisible Unicode and file metadata, replaces numbers, names, acronyms,
links and citations with keyed codes (the API never sees them), builds local content-word notes, and has the
writer model rewrite the chunk from the notes only, reordering ideas and swapping synonyms. Codes are verified
and restored locally. Optional removable markers (`--marcadores N`) break the context window of the writer's
own watermark. Measured: 20 real articles in 37.6 s with Claude Haiku 4.5; an own-key context watermark went
from z = 4.2–5.2 to 0.1–0.2. It is not a humanizer and does not target AI-style classifiers.

```bash
pip install git+https://github.com/ohariseldon/babelllm
babelllm articles/ -o clean/ -m ollama:qwen2.5:7b-instruct --idioma en
```

License: MIT.
