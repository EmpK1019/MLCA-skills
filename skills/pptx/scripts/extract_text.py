from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree


def extract_text(file_path: str) -> str:
    path = Path(str(file_path or '').strip()).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f'PPTX 文件不存在：{path}')

    slide_sections: list[str] = []
    with zipfile.ZipFile(path) as archive:
        slide_names = sorted(
            name
            for name in archive.namelist()
            if name.startswith('ppt/slides/slide') and name.endswith('.xml')
        )
        for slide_index, slide_name in enumerate(slide_names, start=1):
            root = ElementTree.fromstring(archive.read(slide_name))
            paragraphs: list[str] = []
            for paragraph in root.iterfind('.//{*}p'):
                texts = [
                    str(node.text).strip()
                    for node in paragraph.iterfind('.//{*}t')
                    if node.text and str(node.text).strip()
                ]
                if texts:
                    paragraphs.append(''.join(texts))
            if paragraphs:
                slide_sections.append(f'幻灯片 {slide_index}\n' + '\n'.join(paragraphs))
    return '\n\n'.join(slide_sections).strip()
