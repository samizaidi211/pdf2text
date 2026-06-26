# pdf_to_text

Drop PDFs into `input/`, run the script, get `.docx` files in `output/`.

- Embedded text is extracted directly (fast, lossless) where available.
- Pages with little or no embedded text fall back to Tesseract OCR.
- Conversions are cached by content hash, so re-running is cheap.
- The original PDFs stay in `input/` — you can verify what was processed any time.

## Layout

```
pdf_to_text/
├── input/             # drop PDFs here (never deleted)
├── output/            # .docx files land here
│   └── .manifest.json # cache: filename -> SHA256 of input + output path
├── pdf_to_text.py
├── requirements.txt
└── README.md
```

## Install

Python deps:

```bash
pip install -r requirements.txt
```

System deps (only needed if any PDF must be OCR'd):

- **Tesseract OCR** — Windows: <https://github.com/UB-Mannheim/tesseract/wiki>; macOS: `brew install tesseract`; Debian/Ubuntu: `sudo apt install tesseract-ocr`
- **Poppler** (used to rasterize pages) — Windows: <https://github.com/oschwartz10612/poppler-windows/releases> (pass the `bin` folder via `--poppler-path`); macOS: `brew install poppler`; Debian/Ubuntu: `sudo apt install poppler-utils`

## Use

Drop PDFs into `input/`, then:

```bash
python pdf_to_text.py
```

That's it. Subdirectories under `input/` are mirrored under `output/`.

### Cache behavior

The cache key is the **SHA-256 of the PDF's bytes** plus the output format. So:

- Re-running with no new files → everything is skipped.
- Adding a new PDF → only that one is processed.
- Replacing a PDF (same name, different content) → it's re-converted automatically.
- Want to re-run everything? `python pdf_to_text.py --force`.

### Other modes

Pure embedded-text (no OCR), useful for testing without Tesseract installed:

```bash
python pdf_to_text.py --no-ocr
```

Force OCR on every page (when embedded text is garbled or you want consistent output):

```bash
python pdf_to_text.py --ocr-only --dpi 400 --lang eng
```

Convert to `.txt` instead of `.docx`:

```bash
python pdf_to_text.py --format txt
```

Speed up batches:

```bash
python pdf_to_text.py --workers 4
```

Point at a different folder explicitly:

```bash
python pdf_to_text.py path/to/pdfs path/to/outdir
```

Windows with bundled binaries on a custom path:

```bash
python pdf_to_text.py ^
    --tesseract "C:\Program Files\Tesseract-OCR\tesseract.exe" ^
    --poppler-path "C:\poppler\Library\bin"
```

## Flags

| Flag | Default | Notes |
|---|---|---|
| `--format {docx,txt}` | `docx` | Output format |
| `--force` | off | Re-convert even if cache says input is unchanged |
| `--ocr-only` | off | Skip embedded text; OCR every page |
| `--no-ocr` | off | Embedded text only; no fallback |
| `--min-chars N` | 50 | Per-page char threshold below which a page is OCR'd |
| `--dpi N` | 300 | Rasterization DPI for OCR |
| `--lang STR` | `eng` | Tesseract language(s), e.g. `eng+fra` |
| `--poppler-path PATH` | auto | Poppler `bin` directory |
| `--tesseract PATH` | auto | tesseract executable |
| `--workers N` | 1 | Parallel workers in batch mode |
| `-q / --quiet` | off | Suppress per-file progress |
