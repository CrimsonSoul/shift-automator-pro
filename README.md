# Shift Automator Pro

![App Icon](icon.png)

A high-performance Windows desktop application for automating the management and printing of weekly shift schedules via Microsoft Word COM automation.

## Features

- **Batch Processing** — Print any date range with automated header/footer date replacement across day and night shift templates
- **Preflight Validation** — Validates template availability, folder paths, printer selection, and date ranges before printing begins
- **Third Thursday Detection** — Intelligent scheduling logic for monthly clinical rotation templates
- **Dark UI** — Fluent Design inspired dark mode built with Tkinter/ttk
- **Error Recovery** — Graceful handling of COM failures with per-document retry logic and detailed failure summaries
- **Failure Reports** — Automatically writes a CSV report when any documents fail to print
- **Portable** — Ships as a single standalone Windows executable via PyInstaller
- **Comprehensive Logging** — Structured logging to a per-user data directory for debugging and audit trails

## Installation

### Prerequisites

- Python 3.12 or higher
- Microsoft Word (required for COM document processing)
- Windows operating system
- Runtime dependencies: `pywin32`, `tkcalendar` (installed via `requirements.txt`)

### Option 1: Portable EXE (Recommended)

Download the latest `Shift Automator Pro.exe` from the [Releases](https://github.com/CrimsonSoul/shift-automator-pro/releases) page. No installation required.

### Option 2: Run from Source

1. Clone the repository:
   ```bash
   git clone https://github.com/CrimsonSoul/shift-automator-pro.git
   cd shift-automator-pro
   ```

2. Create a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # On Windows
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the application:
   ```bash
   python main.py
   ```

Or use the provided batch files:
- `setup.bat` — Install dependencies
- `start_app.bat` — Launch the application

## Configuration

On first launch, you will be prompted to select:

1. **Day Shift Folder** — Directory containing your daytime clinical shift templates (`.docx`)
2. **Night Shift Folder** — Directory containing your nighttime clinical shift templates (`.docx`)
3. **Printer** — Your target local or network printer

Settings are saved automatically to `%APPDATA%\Shift Automator Pro\config.json` and restored on the next launch.

### Advanced Options

- **Replace dates in headers/footers only** — When enabled, date patterns inside the document body are left unchanged. This can reduce unintended replacements in complex templates.

### Keyboard Shortcuts

| Key | Action |
|-----|--------|
| `Enter` | Start batch processing |
| `Escape` | Cancel the current batch |

### Tips

- Use the **Open Logs** button in the app footer to open the configuration, log, and report folder.
- Hover over the checkbox and buttons for tooltip explanations.

## Project Structure

```
shift-automator-pro/
├── src/
│   ├── __init__.py           # Package initialization and lazy imports
│   ├── app_paths.py          # Per-user data directory helpers
│   ├── config.py             # Configuration management (load/save/migrate)
│   ├── constants.py          # Application constants, colors, fonts, and styling
│   ├── logger.py             # Logging setup with file and console handlers
│   ├── main.py               # Main application controller and batch processing
│   ├── path_validation.py    # Path validation and security (traversal checks)
│   ├── scheduler.py          # Date range logic and template name resolution
│   ├── ui.py                 # Tkinter/ttk UI components and styling
│   └── word_processor.py     # Word COM automation (open, replace, print, close)
├── tests/
│   ├── __init__.py           # Test package marker
│   ├── conftest.py           # Mock Windows modules for cross-platform testing
│   ├── test_app_paths.py     # App paths tests
│   ├── test_config.py        # Configuration tests (load, save, migrate, edge cases)
│   ├── test_logger.py        # Logger setup tests
│   ├── test_main.py          # Main controller tests (validation, batch, failures)
│   ├── test_path_validation.py # Path validation tests
│   ├── test_scheduler.py     # Scheduler and date logic tests
│   ├── test_ui.py            # UI component tests
│   └── test_word_processor.py # Word processor tests (COM, templates, printing)
├── .github/workflows/
│   ├── build.yml             # On-demand Windows build and GitHub Release
│   └── ci.yml                # CI pipeline (black, mypy, pylint, pytest)
├── icon.ico                  # Application icon (Windows taskbar/exe)
├── icon.png                  # Application icon (fallback for non-Windows)
├── main.py                   # Application entry point
├── start_app.bat             # Windows launcher
├── setup.bat                 # Windows dependency installer
├── requirements.txt          # Runtime dependencies (pywin32, tkcalendar)
├── requirements-dev.txt      # Development dependencies (pytest, mypy, black, pylint)
├── pytest.ini                # Pytest configuration with coverage
├── LICENSE                   # MIT license
├── .gitignore
└── README.md
```

## Testing

The test suite runs cross-platform by mocking Windows-specific modules (`pywin32`, `win32com`) in `conftest.py`.

```bash
# Install development dependencies
pip install -r requirements-dev.txt

# Run tests (coverage is enabled by default via pytest.ini)
pytest

# Run tests with explicit coverage flags
pytest --cov=src --cov-report=term-missing
```

### Current Stats

- **156 tests**, all passing
- **76% code coverage**
- **0 mypy errors**, **0 black reformats**

## Development

### Code Quality Tools

| Tool | Purpose | Command |
|------|---------|---------|
| **black** | Code formatting | `black --check src tests` |
| **mypy** | Static type checking | `mypy src` |
| **pylint** | Linting | `pylint src --fail-under=8.0` |
| **pytest** | Testing + coverage | `pytest` |

### Standards

- **Type annotations** on all functions (parameters and return types)
- **Docstrings** with `Args`/`Returns`/`Raises` on all public APIs
- **No silently swallowed exceptions** — every `except` block logs at minimum `debug` level
- **No magic numbers** — all constants extracted to `src/constants.py`
- **Thread safety** — all UI updates from the worker thread go through `_safe_after()`

### Building the Executable

```bash
pip install pyinstaller

pyinstaller --onefile --windowed --icon=icon.ico \
  --add-data "icon.ico;." --add-data "icon.png;." \
  --name="Shift Automator Pro" main.py
```

The executable will be created in the `dist/` directory.

## Privacy

This application processes all documents locally and does not upload data to any external servers.

## Troubleshooting

### Common Issues

**Word not found** — Ensure Microsoft Word is installed and accessible via COM automation.

**Printer not listed** — Check that the printer is properly installed. Click the **Refresh** button next to the printer dropdown to re-scan.

**Template not found** — Verify that template files exist in the specified folders with the correct naming convention (e.g., `Monday.docx`, `Tuesday Night.docx`, `THIRD Thursday.docx`).

### Logs

The application creates a log file at:
```
%APPDATA%\Shift Automator Pro\shift_automator.log
```

### Failure Reports

If any documents fail to print, the app writes a CSV report to:
```
%APPDATA%\Shift Automator Pro\failure_report_YYYYMMDD_HHMMSS.csv
```

## License

MIT — see [LICENSE](LICENSE).

## Contributing

Contributions are welcome. Please ensure:

1. All tests pass (`pytest`)
2. Code is formatted (`black src tests`)
3. Types check cleanly (`mypy src`)
4. New features include tests
5. Documentation is updated

## Support

For issues and questions, please [open an issue](https://github.com/CrimsonSoul/shift-automator-pro/issues) on GitHub.
