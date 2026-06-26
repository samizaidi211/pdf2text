"""Convert PDFs to .docx (or .txt) using embedded text + Tesseract OCR fallback.

DEFAULT WORKFLOW (no args):
    Drop PDFs into ./input/, run `python pdf_to_text.py`, get .docx files in
    ./output/. Already-converted PDFs are skipped via a content-hash cache
    (./output/.manifest.json) — so re-running is cheap, and you can keep the
    originals in input/ to verify what was processed later.

EXPLICIT MODE:
    python pdf_to_text.py INPUT [OUTPUT] [options]
    INPUT  is a .pdf file or a directory of PDFs.
    OUTPUT is a file or directory (defaults next to input).

Options:
    --format {docx,txt}  Output format (default: docx).
    --force              Re-convert even if cache says input is unchanged.
    --ocr-only           Skip embedded text; OCR every page.
    --no-ocr             Embedded text only; no OCR fallback.
    --min-chars N        Per-page char threshold for OCR fallback (default: 50).
    --dpi N              OCR rasterization DPI (default: 300).
    --lang STR           Tesseract language code(s) (default: eng).
    --poppler-path PATH  Poppler 'bin' directory (Windows).
    --tesseract PATH     Path to the tesseract executable.
    --workers N          Parallel workers for batch mode (default: 1).
    -q / --quiet         Suppress per-file progress lines.

Requires: pypdf, pdf2image, pytesseract, Pillow, python-docx.
System:   Tesseract OCR + Poppler must be installed (or pointed to with flags)
          if any PDF requires OCR. Pure embedded-text PDFs don't need them.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

MANIFEST_NAME = ".manifest.json"


@dataclass
class Options:
    ocr_only: bool = False
    no_ocr: bool = False
    min_chars: int = 50
    dpi: int = 300
    lang: str = "eng"
    poppler_path: str | None = None
    tesseract_path: str | None = None
    quiet: bool = False
    fmt: str = "docx"
    force: bool = False


def _log(opts: Options, msg: str) -> None:
    if not opts.quiet:
        print(msg, flush=True)


def _sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def _extract_embedded_pages(pdf_path: Path) -> list[str]:
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    pages: list[str] = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text() or "")
        except Exception:
            pages.append("")
    return pages


def _ocr_page(pdf_path: Path, page_index: int, opts: Options) -> str:
    from pdf2image import convert_from_path
    import pytesseract

    if opts.tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = opts.tesseract_path

    images = convert_from_path(
        str(pdf_path),
        dpi=opts.dpi,
        first_page=page_index + 1,
        last_page=page_index + 1,
        poppler_path=opts.poppler_path,
    )
    if not images:
        return ""
    return pytesseract.image_to_string(images[0], lang=opts.lang)


def pdf_to_pages(pdf_path: Path, opts: Options) -> list[str]:
    """Per-page text via the hybrid strategy."""
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)

    if opts.ocr_only:
        from pypdf import PdfReader
        n_pages = len(PdfReader(str(pdf_path)).pages)
        embedded: list[str] = []
    else:
        embedded = _extract_embedded_pages(pdf_path)
        n_pages = len(embedded)

    out_pages: list[str] = []
    ocr_count = 0
    for i in range(n_pages):
        text = "" if opts.ocr_only else embedded[i]
        needs_ocr = opts.ocr_only or (
            not opts.no_ocr and len(text.strip()) < opts.min_chars
        )
        if needs_ocr:
            try:
                text = _ocr_page(pdf_path, i, opts)
                ocr_count += 1
            except Exception as exc:
                if opts.no_ocr:
                    raise
                text = text or f"[OCR failed on page {i + 1}: {exc}]"
        out_pages.append(text.rstrip())

    _log(
        opts,
        f"  {pdf_path.name}: {n_pages} pages "
        f"({ocr_count} OCR'd, {n_pages - ocr_count} embedded)",
    )
    return out_pages


def pages_to_txt(pages: list[str]) -> str:
    return ("\n\n".join(p.rstrip() for p in pages)).rstrip() + "\n"


# XML 1.0 allows \t \n \r and \x20-\xD7FF / \xE000-\xFFFD / \x10000-\x10FFFF only.
# Strip everything else so python-docx can serialize PDF-extracted text.
_XML_INVALID_RE = re.compile("[^\t\n\r -퟿-�\U00010000-\U0010ffff]")


def _xml_clean(s: str) -> str:
    return _XML_INVALID_RE.sub("", s)


def pages_to_docx(pages: list[str], out_path: Path) -> None:
    from docx import Document

    doc = Document()
    for i, page in enumerate(pages):
        if i > 0:
            doc.add_page_break()
        lines = page.splitlines() or [""]
        for line in lines:
            doc.add_paragraph(_xml_clean(line))
    doc.save(str(out_path))


def _load_manifest(output_dir: Path) -> dict:
    mf = output_dir / MANIFEST_NAME
    if not mf.exists():
        return {}
    try:
        return json.loads(mf.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_manifest(output_dir: Path, manifest: dict) -> None:
    mf = output_dir / MANIFEST_NAME
    mf.write_text(json.dumps(manifest, indent=2, sort_keys=True), encoding="utf-8")


def _iter_pdfs(directory: Path) -> list[Path]:
    seen: set[Path] = set()
    out: list[Path] = []
    for pat in ("*.pdf", "*.PDF"):
        for p in sorted(directory.rglob(pat)):
            if p.is_file() and p.resolve() not in seen:
                seen.add(p.resolve())
                out.append(p)
    return out


def _convert_one(pdf_path: Path, out_path: Path, opts: Options) -> tuple[Path, int]:
    pages = pdf_to_pages(pdf_path, opts)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if opts.fmt == "docx":
        pages_to_docx(pages, out_path)
    else:
        out_path.write_text(pages_to_txt(pages), encoding="utf-8")
    return out_path, len(pages)


def _process_job(pdf_path: Path, out_path: Path, opts: Options) -> tuple[Path, bool, str, int]:
    try:
        _, n_pages = _convert_one(pdf_path, out_path, opts)
        return pdf_path, True, "", n_pages
    except Exception as exc:
        return pdf_path, False, f"{type(exc).__name__}: {exc}", 0


def run_batch(input_dir: Path, output_dir: Path, opts: Options, workers: int) -> int:
    pdfs = _iter_pdfs(input_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest = _load_manifest(output_dir)

    if not pdfs:
        _log(opts, f"No PDFs found in {input_dir}. Drop files there and re-run.")
        return 0

    ext = "." + opts.fmt
    jobs: list[tuple[Path, Path, str]] = []
    skipped = 0
    for pdf in pdfs:
        rel = pdf.relative_to(input_dir).as_posix()
        out_path = output_dir / Path(rel).with_suffix(ext)
        sha = _sha256(pdf)
        entry = manifest.get(rel)
        if (
            not opts.force
            and entry
            and entry.get("sha256") == sha
            and entry.get("format") == opts.fmt
            and out_path.exists()
        ):
            skipped += 1
            continue
        jobs.append((pdf, out_path, sha))

    _log(
        opts,
        f"Found {len(pdfs)} PDF(s) in {input_dir}: "
        f"{len(jobs)} to convert, {skipped} cached.",
    )
    if not jobs:
        return 0

    started = time.time()
    failures: list[tuple[Path, str]] = []
    successes = 0

    def _record(pdf: Path, out_path: Path, sha: str, n_pages: int) -> None:
        rel = pdf.relative_to(input_dir).as_posix()
        manifest[rel] = {
            "sha256": sha,
            "output": out_path.relative_to(output_dir).as_posix(),
            "format": opts.fmt,
            "pages": n_pages,
            "converted_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        }
        _save_manifest(output_dir, manifest)

    if workers <= 1:
        for pdf, out_path, sha in jobs:
            _, ok, err, n_pages = _process_job(pdf, out_path, opts)
            if ok:
                successes += 1
                _record(pdf, out_path, sha, n_pages)
            else:
                failures.append((pdf, err))
                print(f"  FAIL {pdf.name}: {err}", file=sys.stderr)
    else:
        with ProcessPoolExecutor(max_workers=workers) as ex:
            fut_to_meta = {
                ex.submit(_process_job, pdf, out_path, opts): (pdf, out_path, sha)
                for pdf, out_path, sha in jobs
            }
            for fut in as_completed(fut_to_meta):
                pdf_in, out_path, sha = fut_to_meta[fut]
                _, ok, err, n_pages = fut.result()
                if ok:
                    successes += 1
                    _record(pdf_in, out_path, sha, n_pages)
                else:
                    failures.append((pdf_in, err))
                    print(f"  FAIL {pdf_in.name}: {err}", file=sys.stderr)

    elapsed = time.time() - started
    _log(
        opts,
        f"Done in {elapsed:.1f}s — {successes}/{len(jobs)} succeeded, "
        f"{skipped} already cached.",
    )
    return 0 if not failures else 2


def _resolve_single_output(input_path: Path, output: Path | None, fmt: str) -> Path:
    ext = "." + fmt
    if output is None:
        return input_path.with_suffix(ext)
    if output.exists() and output.is_dir():
        return output / input_path.with_suffix(ext).name
    if output.suffix == "":
        output.mkdir(parents=True, exist_ok=True)
        return output / input_path.with_suffix(ext).name
    return output


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Convert PDFs to .docx (or .txt) with OCR fallback.",
    )
    ap.add_argument("input", type=Path, nargs="?", default=None,
                    help="PDF file or directory. Defaults to ./input/")
    ap.add_argument("output", type=Path, nargs="?", default=None,
                    help="Output file or directory. Defaults to ./output/")
    ap.add_argument("--format", choices=["docx", "txt"], default="docx",
                    help="Output format (default: docx).")
    ap.add_argument("--force", action="store_true",
                    help="Re-convert even if cache says input is unchanged.")
    ap.add_argument("--ocr-only", action="store_true",
                    help="Skip embedded text; OCR every page.")
    ap.add_argument("--no-ocr", action="store_true",
                    help="Disable OCR fallback (embedded text only).")
    ap.add_argument("--min-chars", type=int, default=50,
                    help="Per-page char threshold for OCR fallback (default: 50).")
    ap.add_argument("--dpi", type=int, default=300,
                    help="Rasterization DPI for OCR (default: 300).")
    ap.add_argument("--lang", default="eng",
                    help="Tesseract language code(s) (default: eng).")
    ap.add_argument("--poppler-path", default=None,
                    help="Path to Poppler 'bin' directory (Windows).")
    ap.add_argument("--tesseract", default=None,
                    help="Path to tesseract executable.")
    ap.add_argument("--workers", type=int, default=1,
                    help="Parallel workers for batch mode (default: 1).")
    ap.add_argument("-q", "--quiet", action="store_true",
                    help="Suppress per-file progress lines.")
    args = ap.parse_args(argv)

    if args.ocr_only and args.no_ocr:
        ap.error("--ocr-only and --no-ocr are mutually exclusive")

    opts = Options(
        ocr_only=args.ocr_only,
        no_ocr=args.no_ocr,
        min_chars=args.min_chars,
        dpi=args.dpi,
        lang=args.lang,
        poppler_path=args.poppler_path,
        tesseract_path=args.tesseract,
        quiet=args.quiet,
        fmt=args.format,
        force=args.force,
    )

    # No-arg default workflow: ./input/ -> ./output/, anchored to script dir.
    if args.input is None:
        script_dir = Path(__file__).resolve().parent
        input_dir = script_dir / "input"
        output_dir = args.output.resolve() if args.output else script_dir / "output"
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)
        return run_batch(input_dir, output_dir, opts, max(1, args.workers))

    input_path = args.input.resolve()
    if input_path.is_dir():
        output_dir = (
            args.output.resolve()
            if args.output
            else input_path.parent / f"{input_path.name}_out"
        )
        return run_batch(input_path, output_dir, opts, max(1, args.workers))

    if not input_path.is_file():
        print(f"Input not found: {input_path}", file=sys.stderr)
        return 1

    out_path = _resolve_single_output(input_path, args.output, opts.fmt)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    _, ok, err, _ = _process_job(input_path, out_path, opts)
    if not ok:
        print(f"FAIL {input_path.name}: {err}", file=sys.stderr)
        return 2
    _log(opts, f"Wrote {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
