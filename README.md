# docx-report-gen-title-page

Title page plugin for [`docx-report-gen`](https://github.com/WarriorVortex/docx-report-gen).

Adds a `title_page()` method to a Report that renders a formatted
title page from fields. Supports any number of students and
supervisors, replaceable labels, configurable margins and an optional
isolated section for headers/footers.

- **Repository:** https://github.com/WarriorVortex/docx-report-gen-title-page
- **Issues:** https://github.com/WarriorVortex/docx-report-gen-title-page/issues
- **Depends on:** [`docx-report-gen`](https://github.com/WarriorVortex/docx-report-gen) 0.3.1+
- **Python:** 3.9+
- **License:** MIT

---

## Status

Alpha, version `0.1.1`. The plugin API is stable within the `0.x`
series, but may change in minor releases.

---

## Installation

```bash
pip install docx-report-gen-title-page
```

From source:

```bash
pip install git+https://github.com/WarriorVortex/docx-report-gen-title-page.git
```

For development:

```bash
git clone https://github.com/WarriorVortex/docx-report-gen-title-page.git
cd docx-report-gen-title-page
pip install -e ".[dev]"
pytest
```

---

## Quick start

```python
from docx_report_gen import Report
from docx_report_gen_title_page import (
    TitlePagePlugin, Student, Supervisor,
)

r = Report(plugins=[TitlePagePlugin()])
r.title_page(
    institution='Санкт-Петербургский политехнический '
                'университет Петра Великого',
    institute='Институт компьютерных наук и технологий',
    school='Высшая школа программной инженерии',
    work_type='ЛАБОРАТОРНАЯ РАБОТА',
    work_number='№2',
    work_title='«Разложение сигналов в ряд Фурье»',
    discipline='«Применение методов ИИ для ЦОС»',
    students=[Student('Ерохин В.С.', group='5130904/30103')],
    supervisors=[Supervisor('Тутыгин В.С.', prefix='доцент')],
    city='Санкт-Петербург',
    year=2026,
)
r.h1('Введение')
r.save('report.docx')
```

The title page is rendered on the first sheet. Content added after
`title_page()` starts on the second page.

---

## Contents

- [Overview](#overview)
- [Usage](#usage)
- [Participants](#participants)
- [Replaceable labels](#replaceable-labels)
- [Layout](#layout)
- [Page margins](#page-margins)
- [Combining with a source document](#combining-with-a-source-document)
- [Auto-render](#auto-render)
- [API reference](#api-reference)
- [Examples](#examples)
- [Development](#development)
- [License](#license)

---

## Overview

`docx-report-gen` renders reports from Python. This plugin
contributes a single method — `Report.title_page(...)` — which draws
a formatted title page from a set of fields and optionally adds a
page break after it.

The plugin does not touch headers or footers unless `isolate=True`.
It does not modify the document's metadata. It renders into whichever
document the Report currently holds — whether that document was
created fresh or loaded via `Report(source=...)` from the base
package.

---

## Usage

### Attach the plugin and call the method

```python
from docx_report_gen import Report
from docx_report_gen_title_page import TitlePagePlugin

r = Report(plugins=[TitlePagePlugin()])
r.title_page(
    institution='СПбПУ',
    work_type='ЛАБОРАТОРНАЯ РАБОТА',
    work_title='«...»',
    students=[Student('Иванов И.И.', group='101')],
    year=2026,
)
r.h1('Введение')
r.save('report.docx')
```

### Or via a TitlePageConfig object

```python
from docx_report_gen import Report
from docx_report_gen_title_page import (
    TitlePagePlugin, TitlePageConfig, Student,
)

cfg = TitlePageConfig(
    institution='СПбПУ',
    work_title='«...»',
    students=[Student('Иванов И.И.', group='101')],
    year=2026,
)

r = Report(plugins=[TitlePagePlugin()])
r.title_page(cfg)
r.h1('Введение')
```

### Or globally, for every Report

```python
from docx_report_gen.writer import register_plugin
from docx_report_gen_title_page import TitlePagePlugin

register_plugin(TitlePagePlugin())    # registers the method
```

After this, every new `Report` gets a `title_page()` method
automatically.

### With the writer API

```python
from docx_report_gen.writer import new, h1, save, use_plugin

new()
r = use_plugin(TitlePagePlugin := __import__(
    'docx_report_gen_title_page'
).TitlePagePlugin)
# Simpler: use the object API to register, then writer continues
```

For writer-style usage, register the plugin globally before starting
the session:

```python
from docx_report_gen.writer import register_plugin, new, h1, save
from docx_report_gen_title_page import TitlePagePlugin

register_plugin(TitlePagePlugin())

new()
# title_page() is now available on the current Report
from docx_report_gen.writer import current
current().title_page(
    institution='СПбПУ',
    year=2026,
)
h1('Введение')
save('report.docx')
```

---

## Participants

Students and supervisors are given as lists of `Student` and
`Supervisor` dataclasses. Any number of each is supported; they are
rendered in the order they appear in the list.

### Student

```python
Student(
    name='Ерохин В.С.',      # required
    group='5130904/30103',    # optional, rendered as "гр. 5130904/30103"
    prefix='студент',         # optional, default is 'студент'
)
```

Rendered as two tab-separated parts on one line:

```
студент гр. 5130904/30103                       Ерохин В.С.
```

The prefix is fully configurable:

| prefix | Rendered |
|---|---|
| `'студент'` (default) | `студент гр. 101 Иванов И.И.` |
| `'студентка'` | `студентка гр. 101 Петрова А.А.` |
| `'магистр'` | `магистр Сидоров П.П.` |
| `''` (empty) | `гр. 101 Иванов И.И.` or just the name |

### Supervisor

```python
Supervisor(
    name='Тутыгин В.С.',     # required
    prefix='доцент',          # optional, default is ''
)
```

Rendered as two tab-separated parts on one line:

```
доцент                                             Тутыгин В.С.
```

Common prefixes: `'преподаватель'`, `'доцент'`, `'профессор'`,
`'ст. преподаватель'`, `'к.т.н.'`.

### Multiple participants

```python
r.title_page(
    students=[
        Student('Ерохин В.С.', group='5130904/30103'),
        Student('Петрова А.А.', group='5130904/30103',
                prefix='студентка'),
        Student('Сидоров П.П.', prefix='магистр'),
    ],
    supervisors=[
        Supervisor('Тутыгин В.С.', prefix='доцент'),
        Supervisor('Иванов И.И.', prefix='профессор'),
    ],
)
```

Each list is rendered as a block: label on top, participants below.
If a list is empty, both the label and the block are omitted.

---

## Replaceable labels

The words above each block are plain string fields. Set them to any
text you need, or to `None` / `''` to omit them.

```python
r.title_page(
    students=[Student('Иванов И.И.', group='101')],
    students_label='Работу выполнил:',

    supervisors=[Supervisor('Тутыгин В.С.', prefix='доцент')],
    supervisors_label='Научный руководитель:',
)
```

The default labels are `'Выполнил'` and `'Руководитель'`.
`discipline_prefix` (default `'по дисциплине'`) works the same way.

---

## Layout

### Fonts

```python
r.title_page(
    ...,
    font='Times New Roman',      # applied to every line
    base_size=14,                # discipline, signature, bottom
    title_size=16,               # work type and title
    institution_size=14,         # header block
)
```

### Line spacing

By default the plugin uses Word's single spacing
(`line_spacing = 1.0`). Actual line height depends on each line's
font size — roughly 16 pt for 14 pt text, 18 pt for 16 pt text.

To override with an exact height in points:

```python
r.title_page(..., line_height_pt=20)
```

### Gap weights

The vertical space between blocks is distributed proportionally to
gap weights. Only the ratios matter — absolute points are computed
so the title page fits exactly one A4 sheet.

```python
r.title_page(
    ...,
    gap_after_header=2.0,            # after institution/institute/school
    gap_after_work=0.5,              # after "ЛАБОРАТОРНАЯ РАБОТА №2"
    gap_after_title=0.5,             # after the title
    gap_after_discipline_prefix=0.3, # between "по дисциплине" and discipline
    gap_after_discipline=2.0,        # after the discipline block
    gap_after_students_label=0.3,    # between "Выполнил" and the student line
    gap_after_students=0.7,          # after the student block
    gap_after_supervisors_label=0.3, # between "Руководитель" and supervisor
    gap_after_supervisors=6.0,       # before city/year — usually large
)
```

The default `gap_after_supervisors = 6.0` pushes city and year to the
bottom of the page. The last line's baseline lands on the bottom
margin, so "2026" appears on the last line of the sheet.

If content overflows the page, gaps collapse to zero and the plugin
stops there. Reduce font sizes in that case.

---

## Page margins

The title page can have its own margins, different from the rest of
the document. Set `margin_top/bottom/left/right` in centimeters. When
any of them is not None, the plugin creates a section break after the
title page and restores the original margins for the rest of the
document.

```python
r.title_page(
    ...,
    margin_top=2.0,
    margin_bottom=2.0,
    margin_left=3.0,             # standard for Russian academic work
    margin_right=1.0,
)
```

Defaults target a typical Russian title page: 2 / 2 / 3 / 1 cm. Set
any to `None` to inherit the document's existing margin for that
side.

To skip custom margins entirely and use the document's defaults:

```python
r.title_page(
    ...,
    margin_top=None,
    margin_bottom=None,
    margin_left=None,
    margin_right=None,
)
```

When all margins are `None` and `isolate=False`, the plugin renders
into the existing section and uses an inline page break. No new
section is created.

---

## Combining with a source document

Loading an existing `.docx` as the base document is a feature of the
base package — use `Report(source=...)`:

```python
from docx_report_gen import Report
from docx_report_gen_title_page import TitlePagePlugin

r = Report(source='Титульный лист.docx',
           plugins=[TitlePagePlugin()])
r.h1('Введение')
r.save('report.docx')
```

In this mode the plugin only registers its method; it does not render
anything unless you call `title_page()` explicitly. The source file
is used as-is, with all its styles, docDefaults, headers and
formatting preserved.

If both a pre-made title page and additional generated content are
needed, combine them — the source is loaded first, then the plugin's
title page is appended:

```python
r = Report(source='header.docx', plugins=[TitlePagePlugin()])
r.title_page(
    students=[Student('Иванов И.И.', group='101')],
    year=2026,
)
r.h1('Введение')
```

---

## Auto-render

Passing a config or content kwargs to `TitlePagePlugin(...)` renders
the title page automatically in `setup()`, before any user content:

```python
r = Report(plugins=[
    TitlePagePlugin(
        institution='СПбПУ',
        work_title='«...»',
        students=[Student('Иванов И.И.', group='101')],
        year=2026,
    ),
])
# Title page already on page 1.
r.h1('Введение')
```

Constructor arguments:

| Argument | Meaning |
|---|---|
| `config=TitlePageConfig(...)` | Config for auto-render |
| `page_break=True` | Default for the page break after the title page |
| `isolate=False` | Default for section-break isolation |
| `**kwargs` | Any `TitlePageConfig` field, used when `config` is absent |

`config` and `**kwargs` are mutually exclusive.

---

## API reference

### Public names

```python
from docx_report_gen_title_page import (
    TitlePagePlugin,       # the plugin class
    TitlePageConfig,       # config dataclass
    Student,               # participant dataclass
    Supervisor,            # participant dataclass
)
```

### `TitlePagePlugin`

| Method | Purpose |
|---|---|
| `__init__(config=None, page_break=True, isolate=False, **kwargs)` | Create the plugin |
| `setup(report)` | Plugin lifecycle hook; registers `title_page()` and optionally auto-renders |

### `Report.title_page` (added by the plugin)

| Parameter | Meaning |
|---|---|
| `config=TitlePageConfig` | Use a fully constructed config |
| `page_break=True` | Add a page break after the title page |
| `isolate=False` | Clear headers/footers of the title page's section |
| `**kwargs` | Any `TitlePageConfig` field |

Returns the `Report`, so it chains with other methods.

### `TitlePageConfig`

| Field | Default | Meaning |
|---|---|---|
| `institution` | `None` | Top line, university name |
| `institute` | `None` | Second line, faculty/institute |
| `school` | `None` | Third line, department (bold) |
| `work_type` | `None` | E.g. `'ЛАБОРАТОРНАЯ РАБОТА'` |
| `work_number` | `None` | E.g. `'№2'` — joined with work_type |
| `work_title` | `None` | Title in quotes, bold |
| `discipline_prefix` | `'по дисциплине'` | Label above the discipline |
| `discipline` | `None` | Discipline name |
| `students_label` | `'Выполнил'` | Label above the student block |
| `students` | `[]` | List of `Student` |
| `supervisors_label` | `'Руководитель'` | Label above the supervisor block |
| `supervisors` | `[]` | List of `Supervisor` |
| `city` | `None` | Bottom line, city |
| `year` | `None` | Bottom line, year |
| `font` | `'Times New Roman'` | Font family |
| `base_size` | `14` | Base point size |
| `title_size` | `16` | Point size for work type and title |
| `institution_size` | `14` | Point size for the header block |
| `line_height_pt` | `None` | Override line height in points; None = single spacing |
| `gap_after_*` | see [Gap weights](#gap-weights) | Relative gaps between blocks |
| `margin_top` | `2.0` | Top margin in cm (title page only) |
| `margin_bottom` | `2.0` | Bottom margin in cm |
| `margin_left` | `3.0` | Left margin in cm |
| `margin_right` | `1.0` | Right margin in cm |
| `isolate` | `False` | Section break for header/footer isolation |

### `Student`

| Field | Default | Meaning |
|---|---|---|
| `name` | required | Full name |
| `group` | `None` | Group number |
| `prefix` | `'студент'` | Role word |

### `Supervisor`

| Field | Default | Meaning |
|---|---|---|
| `name` | required | Full name |
| `prefix` | `''` | Role word (optional) |

---

## Examples

See [`examples/`](examples/):

| File | Shows |
|---|---|
| `01_from_fields.py` | Title page built from fields, custom layout |
| `02_from_file.py` | Title page loaded from an existing .docx |

Run any example from the repository root:

```bash
python examples/01_from_fields.py
```

---

## Development

```bash
pip install -e ".[dev]"
pytest -v
mypy src/docx_report_gen_title_page --strict
```

Target coverage: 80%+.

---

## License

MIT. See [LICENSE](LICENSE).

---

## Repository

https://github.com/WarriorVortex/docx-report-gen-title-page