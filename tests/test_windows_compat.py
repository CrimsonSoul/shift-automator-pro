import sys
import pytest
from unittest.mock import MagicMock, patch

@pytest.mark.skipif(sys.platform != "win32", reason="Windows specific tests")
def test_word_processor_initialization_real():
    """Test that WordProcessor can initialize on Windows (requires MS Word)."""
    from src.word_processor import WordProcessor
    wp = WordProcessor()
    try:
        wp.initialize()
        assert wp._initialized is True
        assert wp.word_app is not None
    except Exception as e:
        pytest.fail(f"Failed to initialize Word on Windows: {e}")
    finally:
        wp.shutdown()

@pytest.mark.skipif(sys.platform != "win32", reason="Windows specific tests")
def test_printer_enumeration_real():
    """Test that we can actually enumerate printers on Windows."""
    import win32print
    try:
        from src.constants import PRINTER_ENUM_LOCAL
        printers = win32print.EnumPrinters(PRINTER_ENUM_LOCAL)
        assert isinstance(printers, tuple)
    except Exception as e:
        pytest.fail(f"Failed to enumerate printers: {e}")

def test_platform_check_logic():
    """Verify that the app correctly identifies non-Windows platforms."""
    # Temporarily force HAS_PYWIN32 to False to simulate non-windows or missing deps
    with patch("src.word_processor.HAS_PYWIN32", False):
        from src.word_processor import WordProcessor
        wp = WordProcessor()
        with pytest.raises(RuntimeError) as excinfo:
            wp.initialize()
        assert "requires Windows" in str(excinfo.value)

def test_ui_no_win32print_behavior():
    """Verify UI behavior when win32print is missing (cross-platform safety)."""
    with patch("src.ui.HAS_WIN32PRINT", False):
        from src.ui import ScheduleAppUI
        root = MagicMock()
        ui = ScheduleAppUI(root)
        # Should show an error message in the printer row
        # This is a bit hard to verify without deep inspection, but we check it doesn't crash
        assert ui.printer_dropdown is not None
