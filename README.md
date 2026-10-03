# docx-report-gen-title-page

Title page plugin for [`docx-report-gen`](https://github.com/WarriorVortex/docx-report-gen).

Generates a standardized title page as the first element of a report.
The title page is inserted via a plugin hook so the caller does not
need to remember the correct sequence of headings, spacing and page
breaks.

- **Repository:** https://github.com/WarriorVortex/docx-report-gen-title-page
- **Issues:** https://github.com/WarriorVortex/docx-report-gen-title-page/issues
- **Depends on:** [`docx-report-gen`](https://github.com/WarriorVortex/docx-report-gen) 0.3.0+
- **Python:** 3.9+
- **License:** MIT

---

## Status

Alpha, version `0.1.0`. The plugin API is stable within the `0.x`
series but may change in minor releases.

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
from docx_report_gen_title_page import TitlePagePlugin

r = Report(plugins=[TitlePagePlugin(...)])
r.h1('Introduction')
r.save('report.docx')
```

The plugin inserts the title page in its `setup(report)` hook, so
it appears before any content added by the caller.

---

## Contents

- [Overview](#overview)
- [Usage](#usage)
- [Title page standard](#title-page-standard)
- [Configuration](#configuration)
- [API reference](#api-reference)
- [Development](#development)
- [License](#license)

---

## Overview

`docx-report-gen` exposes a plugin API that lets an extension register
block methods, inline nodes and lifecycle hooks. This plugin uses the
`setup` hook to write a title page as the first element of the
document.

The plugin does not modify any global state. It only interacts with
the `Report` it is attached to.

---

## Usage

### Local to a single report

```python
from docx_report_gen import Report
from docx_report_gen_title_page import TitlePagePlugin

r = Report(plugins=[
    TitlePagePlugin(
        institution='Университет ИТМО',
        department='Факультет программной инженерии',
        discipline='Физика',
        work_type='Лабораторная работа',
        work_number='1',
        work_title='Измерение ускорения свободного падения',
        student='Иванов И. И.',
        group='ИУ7-31',
        teacher='Петров П. П.',
        city='Санкт-Петербург',
        year=2026,
    ),
])
r.h1('Introduction')
r.save('report.docx')
```

### Global for every report

```python
from docx_report_gen.writer import register_plugin
from docx_report_gen_title_page import TitlePagePlugin

register_plugin(TitlePagePlugin(...))
```

After this call, every subsequent `Report` receives the title page.

---

## Title page standard

<!-- TODO: fill in once the layout is finalized. -->

The current implementation targets a Russian academic standard:
institution header, department, discipline, work type and number,
work title, author block, supervisor block, city and year. Specific
alignment, font sizes and vertical spacing will be defined in the
implementation phase.

---

## Configuration

<!-- TODO: fill in once the configuration dataclass is finalized. -->

Configuration is passed to `TitlePagePlugin(...)` as keyword
arguments. All fields are optional; missing fields are omitted from
the rendered page rather than rendered as empty lines.

---

## API reference

<!-- TODO: fill in once the public API is finalized. -->

Planned exports:

```python
from docx_report_gen_title_page import (
    TitlePagePlugin,          # Plugin subclass
    TitlePageConfig,          # configuration dataclass
)
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