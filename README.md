# Shift Automator Pro

Shift Automator Pro is a Windows desktop app that automates weekly schedule-template printing through Microsoft Word COM.

![Platform](https://img.shields.io/badge/platform-Windows-0a7ea4) ![Language](https://img.shields.io/badge/language-Python%203.12-2ea043) ![UI](https://img.shields.io/badge/ui-Tkinter%2Fttk-4b5563) ![Automation](https://img.shields.io/badge/automation-Word%20COM-1f6feb)

## Snapshot

- Production-focused Word automation: open, replace, print, and close with cleanup safeguards
- Domain logic includes monthly clinical rotation handling ("Third Thursday")
- Defensive workflow with preflight checks, retry paths, and CSV failure reporting
- Modular architecture with strong unit-test coverage and static quality gates
- Single-file executable packaging for non-technical end users

## Preview

![Shift Automator Pro icon](icon.png)

## Core Features

- Batch print processing for date ranges across day/night template folders
- Date replacement automation with optional header/footer-only mode
- Template, path, printer, and date-range preflight validation
- Per-document retry handling for transient COM errors
- Structured logs and timestamped CSV failure reports
- Cancelable background processing with responsive UI progress updates

## Architecture

- `src/main.py`: orchestration and workflow control
- `src/ui.py`: Tkinter interface layer
- `src/word_processor.py`: Word COM automation lifecycle
- `src/scheduler.py`: date and template resolution logic
- `src/config.py`: config management and migration
- `src/path_validation.py`: path traversal and filename safety checks

## Tech Stack

| Layer | Technology |
| --- | --- |
| Language | Python 3.12 |
| UI | Tkinter/ttk |
| Office integration | pywin32 (Word COM) |
| Date picker | tkcalendar |
| Testing | pytest + pytest-cov |
| Quality | black + mypy + pylint |
| Packaging | PyInstaller |

## Quick Start

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

Windows helper scripts:

- `setup.bat` installs dependencies
- `start_app.bat` launches the app

## Quality and Testing

```bash
pip install -r requirements-dev.txt
pytest
black --check src tests
mypy src
pylint src --fail-under=8.0
```

The test suite mocks Windows-only modules so CI can run on non-Windows hosts.

## Security and Reliability

- Word documents open read-only during processing
- Word macros are force-disabled during automation
- Path validation blocks traversal outside configured template roots
- Date range limits prevent runaway batch operations
- Config writes are atomic and logs capture operational detail

## Project Layout

- `main.py`: top-level entry point
- `src/`: application modules (controller, UI, scheduler, COM processor, config, validation)
- `tests/`: unit tests and module fixtures
- `.github/workflows/build.yml`: Windows build/release workflow

## License

MIT (see `LICENSE`).
