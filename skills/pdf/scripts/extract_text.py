from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

try:
    import pdfplumber
except Exception:  # pragma: no cover - optional dependency in some environments
    pdfplumber = None

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - optional dependency in some environments
    PdfReader = None


def _normalize_path(file_path: str) -> Path:
    path = Path(str(file_path or '').strip()).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f'PDF 文件不存在：{path}')
    return path


def _format_table(table: list[list[object]] | None, table_index: int) -> str:
    rows: list[str] = []
    for row in table or []:
        cells = [str(cell).strip() if cell is not None else '' for cell in row]
        if any(cells):
            rows.append(' | '.join(cells))
    if not rows:
        return ''
    return f'[表格 {table_index}]\n' + '\n'.join(rows)


def _extract_with_pdfplumber(path: Path, max_pages: int) -> str:
    if pdfplumber is None:
        return ''

    sections: list[str] = []
    with pdfplumber.open(str(path)) as pdf:
        for page_index, page in enumerate(pdf.pages[:max_pages], start=1):
            text = str(page.extract_text() or '').strip()
            tables = [
                _format_table(table, table_index)
                for table_index, table in enumerate(page.extract_tables() or [], start=1)
            ]
            page_parts = [part for part in [text, *tables] if part]
            if page_parts:
                sections.append(f'第 {page_index} 页\n' + '\n\n'.join(page_parts))
    return '\n\n'.join(sections).strip()


def _extract_with_pypdf(path: Path, max_pages: int) -> str:
    if PdfReader is None:
        return ''

    reader = PdfReader(str(path))
    sections: list[str] = []
    for page_index, page in enumerate(reader.pages[:max_pages], start=1):
        text = str(page.extract_text() or '').strip()
        if text:
            sections.append(f'第 {page_index} 页\n{text}')
    return '\n\n'.join(sections).strip()


def _extract_with_pdftotext(path: Path, max_pages: int) -> str:
    runner = shutil.which('pdftotext')
    if not runner:
        return ''

    try:
        completed = subprocess.run(
            [
                runner,
                '-enc',
                'UTF-8',
                '-layout',
                '-f',
                '1',
                '-l',
                str(max_pages),
                str(path),
                '-',
            ],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            timeout=60,
            check=False,
        )
    except Exception:
        return ''

    if completed.returncode != 0:
        return ''
    return str(completed.stdout or '').strip()


def extract_text(file_path: str, max_pages: int = 24) -> str:
    path = _normalize_path(file_path)
    page_limit = max(1, min(int(max_pages or 24), 200))

    text = _extract_with_pdfplumber(path, page_limit)
    if text:
        return text

    text = _extract_with_pypdf(path, page_limit)
    if text:
        return text

    text = _extract_with_pdftotext(path, page_limit)
    if text:
        return text

    raise RuntimeError('当前未提取到可读 PDF 正文；该文件可能是扫描件或导图型 PDF，需要 OCR。')
