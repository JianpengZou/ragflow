# Building RagTokenizer Resources from MeCab IPA Dictionary

The repository ships the original IPA dictionary files under `mecab-ipadic`.  
Run the helper script to collapse those CSV sources into a RagTokenizer-ready
text dictionary and (optionally) a serialized trie:

```bash
python rag/nlp/build_ipadic_trie.py
```

The command writes `rag/res/ipadic.txt` and `rag/res/ipadic.txt.trie`.  
If you only need the text dictionary (for example when `datrie` is not
available), append `--skip-trie`.

To merge extra term lists (e.g., the 教科書 IT vocabulary Excel), pass the
workbook via `--extra-xlsx` and target the relevant sheet with `--extra-sheet`:

```bash
python rag/nlp/build_ipadic_trie.py \
  --extra-xlsx 情報科全教科書用語リスト改240509.xlsx \
  --extra-sheet "情報科全教科書用語説明付き" \
  --extra-freq 10000000
```

All rows in the specified sheet are deduplicated and appended by the value in
the `用語` column (split first on whitespace, then on `|`/`／`). Pure ASCII
tokens are ignored, and mixed ASCII/日本語 expressions have their ASCII
portion stripped before being inserted. The optional `--extra-freq` lets you
control the weight assigned to those injected terms.

To use the generated resources with `RagTokenizer`, load the dictionary at
runtime:

```python
from rag.nlp.rag_tokenizer import RagTokenizer

tokenizer = RagTokenizer()
tokenizer.loadUserDict("rag/res/ipadic.txt")
```

The loader will pick up the cached trie when it is present, or build it from
`ipadic.txt` on demand.
