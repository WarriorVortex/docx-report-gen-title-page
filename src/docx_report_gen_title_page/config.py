"""Configuration dataclass for the title page plugin.

All content fields are optional. Missing fields are skipped entirely —
no empty paragraph is inserted in their place.

Layout is computed from relative gap weights. The renderer measures
the page geometry, estimates how many lines each text field will take,
and distributes the remaining vertical space among the gaps in
proportion to their weights. The result is a title page that fits
exactly one A4 sheet, regardless of margins or font sizes.
"""
from dataclasses import dataclass
from typing import Optional


@dataclass
class TitlePageConfig:
    """Content and layout for a Russian academic title page.

    Content sections (each field is optional):
        Header block:   institution, institute, school
        Main block:     work_type, work_number, work_title,
                        discipline_prefix, discipline
        Signature block: student_label, student,
                        supervisor_label, supervisor
        Bottom block:   city, year

    Layout:
        font             Font family for every paragraph.
        base_size        Base point size (discipline, signature, bottom).
        title_size       Point size for work type, number and title.
        institution_size Point size for the header block.
        line_height_pt   Exact line height in points. If None, derived
                         from the largest font size (×1.4). Must be at
                         least large enough to avoid clipping.
        gap_*            Relative weights of the vertical gaps between
                         blocks. Only the ratios matter; absolute values
                         are scaled so the page fits exactly. Set a
                         weight to 0 to eliminate a gap, or raise it to
                         push blocks further apart.
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

    # ---------- fonts ----------
    font: str = 'Times New Roman'
    base_size: int = 14
    title_size: int = 16
    institution_size: int = 14
    line_height_pt: Optional[int] = None

    # ---------- gap weights ----------
    # Higher weight = larger gap. Only the ratios matter.
    gap_after_header: float = 2.0
    gap_after_work: float = 0.5
    gap_after_title: float = 0.5
    gap_after_discipline_prefix: float = 0.3
    gap_after_discipline: float = 2.0
    gap_after_student_label: float = 0.3
    gap_after_student: float = 0.5
    gap_after_supervisor: float = 3.0