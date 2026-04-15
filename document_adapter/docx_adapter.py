"""DOCX 어댑터: python-docx (편집) + docxtpl (템플릿 렌더)."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx import Document
from docxtpl import DocxTemplate

from .base import DocumentAdapter, TableSchema

TAG_PATTERN = re.compile(r"\{\{\s*(\w+)\s*\}\}")


class DocxAdapter(DocumentAdapter):
    format = "docx"

    def _open(self) -> None:
        self._doc = Document(self.path)

    def save(self, path: Path | str | None = None) -> Path:
        target = Path(path) if path else self.path
        self._doc.save(target)
        self.path = target
        return target

    # ---- inspection ----
    def get_placeholders(self) -> list[str]:
        keys: set[str] = set()
        for p in self._doc.paragraphs:
            keys.update(TAG_PATTERN.findall(p.text))
        for table in self._doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    keys.update(TAG_PATTERN.findall(cell.text))
        return sorted(keys)

    def get_tables(self, min_rows: int = 1, min_cols: int = 1,
                   preview_rows: int = 4, max_cell_len: int = 40) -> list[TableSchema]:
        schemas: list[TableSchema] = []
        for i, t in enumerate(self._doc.tables):
            rows, cols = len(t.rows), len(t.columns)
            if rows < min_rows or cols < min_cols:
                continue
            preview: list[list[str]] = []
            for row in list(t.rows)[:preview_rows]:
                preview.append([c.text.strip()[:max_cell_len] for c in row.cells])
            schemas.append(TableSchema(index=i, rows=rows, cols=cols, preview=preview))
        return schemas

    # ---- editing ----
    def render_template(self, context: dict[str, Any]) -> None:
        """docxtpl 기반 Jinja2 렌더. 참고:
        - `{%tr for row in rows %}` / `{%tr endfor %}`는 **각각 별도 행**에 두어야 함
        - 같은 행에 두면 `<w:tr>` 전체가 `{% for %}`로 교체되어 endfor 손실
        """
        tpl = DocxTemplate(self.path)
        tpl.render(context)
        tpl.save(self.path)
        # 렌더 후 _doc 재로드
        self._doc = Document(self.path)

    def set_cell(self, table_index: int, row: int, col: int, value: str) -> str:
        cell = self._doc.tables[table_index].rows[row].cells[col]
        old = cell.text
        _set_cell_preserving_format(cell, value)
        return old

    def append_row(self, table_index: int, values: list[str]) -> None:
        table = self._doc.tables[table_index]
        new_row = table.add_row()
        for i, v in enumerate(values):
            if i < len(new_row.cells):
                _set_cell_preserving_format(new_row.cells[i], v)


def _set_cell_preserving_format(cell, value: str) -> None:
    """Write ``value`` into ``cell`` without dropping existing run formatting.

    ``python-docx`` exposes ``cell.text = value`` but that setter deletes all
    existing runs and creates a fresh one with default font/size, so any
    template formatting (font name, size, bold, color) is lost. Instead, reuse
    the first existing run so its formatting survives and blank the rest.

    If the cell has no runs at all we fall back to the default setter, which is
    the only code path that is forced to accept the default font.

    Paragraph identity (``para is first_para``) is unreliable across repeated
    ``cell.paragraphs`` calls on some python-openxml versions, so we compare
    by index instead.
    """
    paragraphs = list(cell.paragraphs)
    first_idx = next((i for i, para in enumerate(paragraphs) if para.runs), None)

    if first_idx is None:
        cell.text = value
        return

    first_para = paragraphs[first_idx]
    first_para.runs[0].text = value
    for run in first_para.runs[1:]:
        run.text = ""

    for i, para in enumerate(paragraphs):
        if i == first_idx:
            continue
        for run in para.runs:
            run.text = ""
