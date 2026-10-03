"""Configuration dataclass for the title page plugin.

All content fields are optional. Missing fields are skipped entirely —
no empty paragraph is inserted in their place. Layout fields have
sensible defaults that fit a standard A4 page with ~2 cm margins.
"""
from dataclasses import dataclass
from typing import Optional


@dataclass
class TitlePageConfig:
    """Content and layout for a Russian academic title page.

    Content sections (each line is optional):
        Header block:   institution, institute, school
        Main block:     work_type, work_number, work_title,
                        discipline_prefix, discipline
        Signature block: student_label, student,
                        supervisor_label, supervisor
        Bottom block:   city, year

    Layout fields:
        font            Font family for every paragraph.
        base_size       Base point size (headers, discipline, signature).
        title_size      Point size for work type and title.
        institution_size Point size for the header block.
        line_spacing    Multiplier; 1.5 is the usual academic standard.
        bottom_gap_lines Number of blank lines between the signature
                        block and city/year. Controls where the bottom
                        block lands. The default targets an A4 page
                        with 2 cm margins so the whole page fits on one
                        sheet.
    """

    # ---------- header ----------
    institution: Optional[str] = None
    institute: Optional[str] = None
    school: Optional[str] = None

    # ---------- main ----------
    work_type: Optional[str] = None
    work_number: Optional[str] = None
    work_title: Optional[str] = None
    discipline_prefix: Optional[str] = 'по дисциплине'
    discipline: Optional[str] = None

    # ---------- signature ----------
    student_label: Optional[str] = 'Выполнил'
    student: Optional[str] = None
    supervisor_label: Optional[str] = 'Руководитель'
    supervisor: Optional[str] = None

    # ---------- bottom ----------
    city: Optional[str] = None
    year: Optional[int] = None

    # ---------- layout ----------
    font: str = 'Times New Roman'
    base_size: int = 14
    title_size: int = 16
    institution_size: int = 14
    line_spacing: float = 1.5
    bottom_gap_lines: int = 6