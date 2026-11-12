# AutoImpress

Aplicativo desktop em Python que **gera e imprime cartazes de preço automaticamente** a partir de uma planilha Excel.

Antes, cada cartaz era preenchido à mão no Word. Agora o programa lê a planilha semanal, descobre quais produtos mudaram de preço, escolhe o modelo certo para cada um, gera os PDFs e envia tudo para a impressora.

## Como funciona

```
dados.xlsx ──► filtra produtos ──► decide o tipo ──► preenche o modelo ──► PDF ──► impressora
                (mudou o preço)     (promo ou cheio)   (.docx com campos)
```

1. **Lê a planilha** `dados.xlsx` (primeira aba) e ignora produtos sem mudança de preço e produtos 220V.
2. **Decide o tipo de cartaz** comparando o preço:
   - `Por` menor que `De` (ex.: de 99 por 79) → **promo** (`cartaz_promo.docx`)
   - `Por` maior que `De` (ex.: de 99 por 129) → **cheio** (`cartaz_cheio.docx`)
   - Sem `De` na planilha → usa a coluna `CI Anterior` (abaixo de 1 = baixou, acima de 1 = subiu)
3. **Mostra tudo numa tabela** para você conferir. O modelo de cada produto pode ser trocado, ou o produto pode ser marcado como "Pular".
4. **Preenche o modelo Word** (biblioteca `docxtpl`), **converte para PDF** (`docx2pdf`) e **manda imprimir**.

## Requisitos

- Windows (ou macOS) com **Microsoft Word instalado**: o `docx2pdf` usa o Word para converter.
- Python 3.10 ou superior (o `tkinter` já vem com o instalador oficial).

```bash
pip install -r requirements.txt
```

## Testando com os arquivos de exemplo

A pasta `exemplo/` tem modelos genéricos e uma planilha com dados fictícios. Copie os três arquivos para a mesma pasta do `main.py` e rode:

```bash
# Windows (PowerShell), dentro da pasta do projeto
copy exemplo\* .
py main.py
```

Com os dados de exemplo, o programa lista 10 produtos (7 promo e 3 cheio). Antes de clicar em **Gerar e Imprimir**, marque como "Pular" os produtos que não quer imprimir, para não mandar tudo para a impressora durante o teste. Os PDFs ficam na pasta `impressoes/`.

## Usando com seus próprios arquivos

Coloque na mesma pasta do `main.py`:

| Arquivo | O que é |
|---|---|
| `dados.xlsx` | Planilha com os preços (formato abaixo) |
| `cartaz_promo.docx` | Modelo para preço que **baixou** |
| `cartaz_cheio.docx` | Modelo para preço que **subiu** |

### Formato da planilha

O programa lê a **primeira aba**, com o cabeçalho na **linha 5** e as colunas **G a K**:

| Coluna | Cabeçalho | Conteúdo |
|---|---|---|
| G | `Artigo` | Código do produto |
| H | `Desc. Artigo` | Descrição |
| I | `De:` | Preço anterior (pode ficar vazio) |
| J | `Por:` | Preço novo |
| K | `CI Anterior` | Índice de variação (`1` = mesmo preço, o produto é ignorado) |

### Campos dos modelos Word

Os modelos são arquivos `.docx` comuns com campos entre chaves duplas, que o programa troca pelos valores:

| Campo | Vira | Obrigatório |
|---|---|---|
| `{{Desc_Artigo}}` | Descrição do produto | sim |
| `{{Artigo}}` | Código do produto | sim |
| `{{Por}}` | Preço novo (ex.: `R$ 149,99`) | sim |
| `{{De}}` | Preço anterior (ex.: `R$ 179,99`) | não (só o promo usa) |

Digite cada campo de uma vez só, sem mudar a formatação no meio, para o Word não quebrar o campo em pedaços. Se um modelo não tiver os campos obrigatórios, o programa **não gera o cartaz** e avisa qual campo está faltando.

Os modelos precisam ser `.docx`. Arquivos `.doc` antigos não funcionam: abra no Word e use *Salvar como → .docx*.

### Trocar os nomes dos modelos

No topo do `main.py`:

```python
MODELO_PROMO = 'cartaz_promo.docx'
MODELO_CHEIO = 'cartaz_cheio.docx'
```

## Se o programa não abrir

Qualquer erro ao iniciar aparece num pop-up e é gravado no arquivo `erro.log`, na pasta do programa. O motivo mais comum é biblioteca faltando: o pop-up mostra o comando exato de instalação para o Python que está sendo usado.

## Estrutura do projeto

```
AutoImpress/
├── main.py              # o aplicativo inteiro (interface + lógica)
├── requirements.txt     # dependências
└── impressoes/          # PDFs gerados (criada automaticamente, não vai para o Git)
```

## Tecnologias

Python · tkinter · pandas · openpyxl · docxtpl · docx2pdf

## Observação

Os modelos e a planilha deste repositório são **exemplos genéricos**, com dados fictícios. Os arquivos reais de uso (marca, preços, ofertas) ficam fora do repositório.
