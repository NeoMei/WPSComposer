# Mac Word list argument native acceptance

Source-only checks:

```bash
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python build/word-list-argument-20260910/source/test_macos_word_list_arguments_native.py
```

Native run (parent-owned lease only; output directory must not exist):

```bash
PYTHONDONTWRITEBYTECODE=1 /Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python \
  build/word-list-argument-20260910/source/macos_word_list_arguments_native.py \
  --execute \
  --source-freeze build/word-list-argument-20260910/source/SOURCE-FREEZE.json \
  --output build/word-list-argument-20260910/run-01 \
  --timeout 240
```

A PASS requires exact literal prefixes for default/custom/empty/ordered cases, requested left and hanging indents, List Paragraph font/reset OOXML, Body Text 24pt first-line indent after lists, A4 page setup, readable PDF text, exact read-only reopen rows, unchanged DOCX bytes, unchanged document inventory, exact owned close, and a still-matching source freeze. PNG/manual visual review is outside this bounded fixture.
