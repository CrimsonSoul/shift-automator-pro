# Shift Automator Pro

![App Icon](icon.png)

A high-performance, modern desktop application for automating the management and printing of weekly shift schedules.

## 🚀 Features

- **Turbo-Fast Wildcard Engine**: Processes multiple weeks of clinical schedules in seconds
- **Third Thursday Detection**: Intelligent logic for monthly clinical rotation templates
- **Midnight Aesthetic**: Fluent Design inspired dark mode for professional environments
- **Batch Processing**: Print any date range with automated header/footer date updates
- **Portable**: Built for Windows as a single standalone executable
- **Comprehensive Logging**: Detailed logging for debugging and troubleshooting
- **Input Validation**: Robust validation of all user inputs
- **Error Recovery**: Graceful handling of errors with detailed failure summaries

## 🛠️ Installation

### Prerequisites

- Python 3.10 or higher
- Microsoft Word (required for document processing)
- Windows operating system

### Option 1: Portable EXE (Recommended)

Download the latest `Shift Automator Pro.exe` from the [Releases](https://github.com/CrimsonSoul/shift-automator-pro/releases) page. No installation required.

### Option 2: Run from Source

1. Clone the repository:
   ```bash
   git clone https://github.com/CrimsonSoul/shift-automator-pro.git
   cd shift-automator
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
- `setup.bat` - Install dependencies
- `start_app.bat` - Launch the application

## ⚙️ Configuration

On first launch, you will be prompted to select:

1. **Day Shift Folder**: Where your daytime clinical templates are stored
2. **Night Shift Folder**: Where your nighttime clinical templates are stored
3. **Printer**: Your target local or network printer

Settings are saved automatically to `config.json`.

## 📁 Project Structure

```
schedule_app/
├── src/
│   ├── __init__.py          # Package initialization
│   ├── constants.py          # Application constants and styling
│   ├── config.py             # Configuration management
│   ├── logger.py             # Logging setup
│   ├── path_validation.py    # Path validation and security
│   ├── scheduler.py          # Date and scheduling logic
│   ├── ui.py                 # UI components
│   ├── word_processor.py     # Word document processing
│   └── main.py              # Main application controller
├── tests/
│   ├── conftest.py             # Mock Windows modules for cross-platform testing
│   ├── test_config.py          # Configuration tests
│   ├── test_path_validation.py # Path validation tests
│   ├── test_scheduler.py       # Scheduler tests
│   ├── test_ui.py              # UI component tests
│   └── test_word_processor.py  # Word processor tests
├── main.py                  # Application entry point
├── requirements.txt          # Runtime dependencies
├── requirements-dev.txt      # Development dependencies
├── pytest.ini              # Pytest configuration
└── README.md
```

## 🧪 Testing

Run the test suite:

```bash
# Install development dependencies
pip install -r requirements-dev.txt

# Run tests
pytest

# Run tests with coverage
pytest --cov=src --cov-report=html
```

## 🔒 Privacy

This application processes all documents locally and does not upload data to any external servers.

## 📝 Development

### Code Quality

The project uses several tools to maintain code quality:

- **Type Hints**: All functions include type annotations
- **Docstrings**: Comprehensive documentation for all modules and functions
- **Logging**: Structured logging throughout the application
- **Error Handling**: Proper exception handling with specific error messages

### Building the Executable

```bash
# Install PyInstaller
pip install pyinstaller

# Build the executable
pyinstaller --onefile --windowed --icon=icon.ico --name="Shift Automator Pro" main.py
```

The executable will be created in the `dist/` directory.

## 🐛 Troubleshooting

### Common Issues

**Word not found**: Ensure Microsoft Word is installed and accessible.

**Printer not listed**: Check that the printer is properly installed and accessible.

**Template not found**: Verify that template files exist in the specified folders and have the correct naming convention.

### Logs

The application creates a log file (`shift_automator.log`) in the application directory. Check this file for detailed error information.

## 📄 License

MIT

## 🤝 Contributing

Contributions are welcome! Please ensure:

1. All tests pass
2. Code follows the existing style
3. New features include tests
4. Documentation is updated

## 📞 Support

For issues and questions, please open an issue on GitHub.
