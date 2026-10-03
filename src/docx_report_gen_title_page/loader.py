"""Load a .docx file and append its content to another document."""
import copy
from pathlib import Path
from typing import Any, Union

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn


PathLike = Union[str, Path]

_W_PSTYLE = qn('w:pStyle')
_W_RSTYLE = qn('w:rStyle')
_W_TBLSTYLE = qn('w:tblStyle')
_W_SECTPR = qn('w:sectPr')
_W_VAL = qn('w:val')


def load_docx_content(source: PathLike, target: DocxDocument) -> int:
    """Append the body content of `source` to `target`.

    Body elements are inserted at the end of `target`'s body, before
    its final section properties. Style definitions referenced by the
    copied content are added to `target` only if a style with the same
    ID is not already present.

    Args:
        source: path to the .docx file to read.
        target: the python-docx Document to append to. Modified in place.

    Returns:
        The number of body elements copied (paragraphs + tables).

    Raises:
        FileNotFoundError: the source file does not exist.
    """
    source_path = Path(source)
    if not source_path.exists():
        raise FileNotFoundError(
            f'load_docx_content(): source not found: {source_path}'
        )

    src = Document(str(source_path))
    src_body = src.element.body
    tgt_body = target.element.body

    elements: list[Any] = [
        child for child in list(src_body)
        if child.tag != _W_SECTPR
    ]

    if not elements:
        return 0

    _merge_styles(src, target, elements)

    tgt_sect_pr = tgt_body.find(_W_SECTPR)

    for el in elements:
        new_el = copy.deepcopy(el)
        if tgt_sect_pr is not None:
            tgt_sect_pr.addprevious(new_el)
        else:
            tgt_body.append(new_el)

    return len(elements)


def _merge_styles(
    src: DocxDocument,
    tgt: DocxDocument,
    elements: list[Any],
) -> None:
    """Copy style definitions referenced by `elements` from src to tgt."""
    used_ids = _collect_used_style_ids(elements)
    if not used_ids:
        return

    existing_ids: set[str] = {
        style.style_id for style in tgt.styles if style.style_id
    }

    styles_element = tgt.styles.element
    for style in src.styles:
        sid = style.style_id
        if sid in used_ids and sid not in existing_ids:
            styles_element.append(copy.deepcopy(style.element))


def _collect_used_style_ids(elements: list[Any]) -> set[str]:
    """Return the set of style IDs referenced anywhere in `elements`."""
    used: set[str] = set()
    for el in elements:
        for tag in (_W_PSTYLE, _W_RSTYLE, _W_TBLSTYLE):
            for ref in el.iter(tag):
                val = ref.get(_W_VAL)
                if val:
                    used.add(val)
    return used