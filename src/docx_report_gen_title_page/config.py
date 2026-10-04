"""Configuration dataclasses for the title page plugin.

All content fields are optional. Missing fields are skipped without
leaving a placeholder.

Vertical rhythm:

    gap_* weights     relative whitespace between blocks; only ratios
                      matter, absolute points are computed so the page
                      fits exactly one sheet

Line-to-line spacing is not configured here — the plugin uses Word's
single spacing (line_spacing = 1.0). Actual line height depends on
the font size of each line. Set `line_height_pt` to override with an
exact value in points if needed.
"""
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Student:
    """One student who contributed to the report."""
    name: str
    group: Optional[str] = None
    prefix: str = 'студент'

    def left(self) -> str:
        parts: list[str] = []
        if self.prefix:
            parts.append(self.prefix)
        if self.group:
            parts.append(f'гр. {self.group}')
        return ' '.join(parts)

    def render(self) -> str:
        left = self.left()
        if left and self.name:
            return f'{left}\t{self.name}'
        return left or self.name or ''


@dataclass
class Supervisor:
    """One supervisor of the report."""
    name: str
    prefix: str = ''

    def left(self) -> str:
        return self.prefix

    def render(self) -> str:
        if self.prefix and self.name:
            return f'{self.prefix}\t{self.name}'
        return self.prefix or self.name or ''


@dataclass
class TitlePageConfig:
    """Content and layout for a Russian academic title page.

    Content sections (each field is optional):

        Header:    institution, institute, school
        Main:      work_type, work_number, work_title,
                   discipline_prefix, discipline
        Students:  students_label + list[Student]
        Supervisors: supervisors_label + list[Supervisor]
        Bottom:    city, year

    Replaceable labels: students_label and supervisors_label are plain
    strings. Set them to any text ("Выполнил", "Работу выполнил:",
    "Научный руководитель:") or to None / '' to omit them entirely.
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

    # ---------- signature: students ----------
    students_label: Optional[str] = 'Выполнил'
    students: list[Student] = field(default_factory=list)

    # ---------- signature: supervisors ----------
    supervisors_label: Optional[str] = 'Руководитель'
    supervisors: list[Supervisor] = field(default_factory=list)

    # ---------- bottom ----------
    city: Optional[str] = None
    year: Optional[int] = None

    # ---------- fonts ----------
    font: str = 'Times New Roman'
    base_size: int = 14
    title_size: int = 16
    institution_size: int = 14

    line_height_pt: Optional[int] = None
    """Override for line height in points. None (default) means single
    spacing — Word computes the line height from the font's metrics."""

    # ---------- gap weights ----------
    gap_after_header: float = 2.0
    gap_after_work: float = 0.5
    gap_after_title: float = 0.5
    gap_after_discipline_prefix: float = 0.3
    gap_after_discipline: float = 2.0
    gap_after_students_label: float = 0.3
    gap_after_students: float = 0.7
    gap_after_supervisors_label: float = 0.3
    gap_after_supervisors: float = 6.0

    # ---------- page margins (cm) ----------
    # Applied to the title page's section only. A new section is
    # started afterwards, restoring the previous margins for the rest
    # of the document. Set to None to inherit whatever the document
    # already has (python-docx default: 2.54 cm on all sides).
    margin_top: Optional[float] = 2.0
    margin_bottom: Optional[float] = 2.0
    margin_left: Optional[float] = 3.0
    margin_right: Optional[float] = 1.0

    # ---------- isolation ----------
    isolate: bool = False
    """False (default): the title page uses its own section only if
    custom margins are set, or a plain page break otherwise.
    True: forces a section break even without custom margins, so the
    title page's headers and footers can be cleared independently."""