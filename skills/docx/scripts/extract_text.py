from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree


def extract_text(file_path: str) -> str:
    path = Path(str(file_path or '').strip()).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f'DOCX 文件不存在：{path}')

    with zipfile.ZipFile(path) as archive:
        xml_bytes = archive.read('word/document.xml')

    root = ElementTree.fromstring(xml_bytes)
    paragraphs: list[str] = []
    for paragraph in root.iterfind('.//{*}p'):
        texts = [
            str(node.text).strip()
            for node in paragraph.iterfind('.//{*}t')
            if node.text and str(node.text).strip()
        ]
        if texts:
            paragraphs.append(''.join(texts))
    return '\n'.join(paragraphs).strip()
