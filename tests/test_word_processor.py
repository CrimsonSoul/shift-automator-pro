"""
Unit tests for Word document processing module.

These tests verify the Word document operations including template finding,
date replacement, and COM call handling.
"""

import tempfile
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch, call
import pytest

from src.word_processor import WordProcessor, HAS_PYWIN32
from src.constants import DOCX_EXTENSION


@pytest.mark.skipif(not HAS_PYWIN32, reason="pywin32 not available")
class TestWordProcessor:
    """Unit tests for WordProcessor class."""

    def test_initialization(self):
        """Test WordProcessor initialization."""
        processor = WordProcessor()
        assert processor.word_app is None
        assert processor._initialized is False

    @patch('src.word_processor.pythoncom')
    @patch('src.word_processor.win32com.client')
    def test_initialize_success(self, mock_win32, mock_pythoncom):
        """Test successful Word application initialization."""
        mock_word_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_word_app

        processor = WordProcessor()
        processor.initialize()

        assert processor._initialized is True
        assert processor.word_app == mock_word_app
        mock_word_app.Visible = False
        mock_word_app.DisplayAlerts = 0

    def test_initialize_without_pywin32(self):
        """Test initialization fails gracefully without pywin32."""
        with patch('src.word_processor.HAS_PYWIN32', False):
            processor = WordProcessor()
            with pytest.raises(RuntimeError) as exc_info:
                processor.initialize()
            assert "requires Windows with pywin32" in str(exc_info.value)

    @patch('src.word_processor.pythoncom')
    @patch('src.word_processor.win32com.client')
    def test_initialize_already_initialized(self, mock_win32, mock_pythoncom):
        """Test that initialize is idempotent."""
        mock_word_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_word_app

        processor = WordProcessor()
        processor.initialize()
        processor.initialize()  # Should not re-initialize

        # Should only call Dispatch once
        assert mock_win32.Dispatch.call_count == 1

    @patch('src.word_processor.pythoncom')
    @patch('src.word_processor.win32com.client')
    def test_shutdown(self, mock_win32, mock_pythoncom):
        """Test Word application shutdown."""
        mock_word_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_word_app

        processor = WordProcessor()
        processor.initialize()
        processor.shutdown()

        mock_word_app.Quit.assert_called_once()
        assert processor.word_app is None
        assert processor._initialized is False

    @patch('src.word_processor.pythoncom')
    @patch('src.word_processor.win32com.client')
    def test_shutdown_error_handling(self, mock_win32, mock_pythoncom):
        """Test shutdown handles errors gracefully."""
        mock_word_app = MagicMock()
        mock_word_app.Quit.side_effect = Exception("Quit error")
        mock_win32.Dispatch.return_value = mock_word_app

        processor = WordProcessor()
        processor.initialize()
        processor.shutdown()  # Should not raise

        assert processor.word_app is None
        assert processor._initialized is False

    def test_find_template_file_success(self, tmp_path):
        """Test finding a template file successfully."""
        # Create test template
        template_folder = tmp_path / "templates"
        template_folder.mkdir()
        (template_folder / f"Monday{DOCX_EXTENSION}").write_text("test")

        processor = WordProcessor()
        result = processor.find_template_file(str(template_folder), "Monday")

        assert result is not None
        assert result.endswith(f"Monday{DOCX_EXTENSION}")

    def test_find_template_file_not_found(self, tmp_path):
        """Test finding a non-existent template file."""
        template_folder = tmp_path / "templates"
        template_folder.mkdir()

        processor = WordProcessor()
        result = processor.find_template_file(str(template_folder), "MissingDay")

        assert result is None

    def test_find_template_file_case_insensitive(self, tmp_path):
        """Test template finding is case-insensitive."""
        template_folder = tmp_path / "templates"
        template_folder.mkdir()
        (template_folder / "monday.docx").write_text("test")  # lowercase

        processor = WordProcessor()
        result = processor.find_template_file(str(template_folder), "Monday")  # uppercase

        assert result is not None
        assert "monday.docx" in result.lower()

    def test_find_template_file_invalid_folder(self):
        """Test finding template in invalid folder."""
        processor = WordProcessor()
        result = processor.find_template_file("/nonexistent/folder", "Monday")

        assert result is None

    def test_find_template_file_filters_extensions(self, tmp_path):
        """Test that only .docx files are considered."""
        template_folder = tmp_path / "templates"
        template_folder.mkdir()
        (template_folder / "Monday.docx").write_text("test")
        (template_folder / "Monday.txt").write_text("test")
        (template_folder / "Tuesday.docx").write_text("test")

        processor = WordProcessor()
        result = processor.find_template_file(str(template_folder), "Monday")

        assert result is not None
        assert result.endswith(".docx")

    def test_safe_com_call_success(self):
        """Test safe COM call with successful execution."""
        processor = WordProcessor()
        mock_func = MagicMock(return_value="success")

        result = processor.safe_com_call(mock_func, "arg1", "arg2")

        assert result == "success"
        mock_func.assert_called_once_with("arg1", "arg2")


    def test_safe_com_call_retry_on_rejection(self):
        """Test safe COM call retries on call rejection."""
        from unittest.mock import MagicMock, patch

        processor = WordProcessor()

        mock_func = MagicMock()
        mock_func.side_effect = [
            Exception("Call was rejected by callee"),
            "success"
        ]

        # Mock time.sleep to avoid actual delays in tests
        with patch('time.sleep'):
            result = processor.safe_com_call(mock_func, retries=2, delay=0.01)

        assert result == "success"
        assert mock_func.call_count == 2

    def test_safe_com_call_raises_after_retries_exhausted(self):
        """Test safe COM call raises after all retries fail."""
        processor = WordProcessor()
        mock_func = MagicMock()
        mock_func.side_effect = Exception("Call was rejected by callee")

        with pytest.raises(Exception) as exc_info:
            processor.safe_com_call(mock_func, retries=2, delay=0.01)

        assert "rejected" in str(exc_info.value).lower()
        assert mock_func.call_count == 2

    def test_safe_com_call_non_retriable_error(self):
        """Test safe COM call doesn't retry non-retriable errors."""
        processor = WordProcessor()
        mock_func = MagicMock()
        mock_func.side_effect = Exception("Some other error")

        with pytest.raises(Exception) as exc_info:
            processor.safe_com_call(mock_func, retries=3)

        assert "Some other error" in str(exc_info.value)
        assert mock_func.call_count == 1  # No retries

    def test_context_manager(self):
        """Test WordProcessor as context manager."""
        with patch('src.word_processor.pythoncom'), \
             patch('src.word_processor.win32com.client') as mock_win32:

            mock_word_app = MagicMock()
            mock_win32.Dispatch.return_value = mock_word_app

            with WordProcessor() as processor:
                assert processor._initialized is True
                assert processor.word_app == mock_word_app

            # After context, should be shut down
            mock_word_app.Quit.assert_called_once()

    def test_replace_dates_patterns(self):
        """Test date replacement pattern generation."""
        processor = WordProcessor()

        # Mock document
        mock_doc = MagicMock()
        mock_story_range = MagicMock()
        mock_doc.StoryRanges = [mock_story_range]
        mock_story_range.NextStoryRange = None

        # Mock find object
        mock_find = MagicMock()
        mock_story_range.Find = mock_find

        test_date = date(2026, 1, 15)  # January 15, 2026

        processor.replace_dates(mock_doc, test_date)

        # Should call Execute 3 times: 1 for placeholder, 2 for patterns
        assert mock_find.Execute.call_count == 3
        
        # Verify first call is for the DATE_PLACEHOLDER
        from src.constants import DATE_PLACEHOLDER
        first_call_args = mock_find.Execute.call_args_list[0]
        assert first_call_args[0][0] == DATE_PLACEHOLDER

    @patch('src.word_processor.pythoncom')
    @patch('src.word_processor.win32com.client')
    @patch('src.word_processor.win32print')
    def test_print_document_success(self, mock_win32print, mock_win32, mock_pythoncom, tmp_path):
        """Test successful document printing."""
        # Mock printer status to be ready
        mock_win32print.OpenPrinter.return_value = 1
        mock_win32print.GetPrinter.return_value = {'Status': 0}
        mock_word_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_word_app

        # Create test template
        template_folder = tmp_path / "templates"
        template_folder.mkdir()
        (template_folder / "Monday.docx").write_text("test")

        mock_doc = MagicMock()
        mock_doc.ProtectionType = -1  # PROTECTION_NONE
        mock_doc.PrintOut = MagicMock()  # Make PrintOut accept any args
        mock_word_app.Documents.Open.return_value = mock_doc

        processor = WordProcessor()
        processor.initialize()

        success, error = processor.print_document(
            str(template_folder),
            "Monday",
            date(2026, 1, 15),
            "Test Printer"
        )

        assert success is True
        assert error is None
        mock_word_app.Documents.Open.assert_called_once()
        mock_doc.PrintOut.assert_called_once()
        mock_doc.Close.assert_called_once()

    @patch('src.word_processor.pythoncom')
    @patch('src.word_processor.win32com.client')
    @patch('src.word_processor.win32print')
    def test_print_document_template_not_found(self, mock_win32print, mock_win32, mock_pythoncom):
        """Test printing with missing template."""
        # Mock printer status to be ready
        mock_win32print.OpenPrinter.return_value = 1
        mock_win32print.GetPrinter.return_value = {'Status': 0}
        mock_word_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_word_app

        processor = WordProcessor()
        processor.initialize()

        success, error = processor.print_document(
            "/nonexistent/folder",
            "MissingDay",
            date(2026, 1, 15),
            "Test Printer"
        )

        assert success is False
        assert "not found" in error.lower()

    @patch('src.word_processor.pythoncom')
    @patch('src.word_processor.win32com.client')
    @patch('src.word_processor.win32print')
    def test_print_document_unprotect(self, mock_win32print, mock_win32, mock_pythoncom):
        """Test document is unprotected before modification."""
        # Mock printer status to be ready
        mock_win32print.OpenPrinter.return_value = 1
        mock_win32print.GetPrinter.return_value = {'Status': 0}
        mock_word_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_word_app

        # Create test template
        with tempfile.TemporaryDirectory() as tmp_dir:
            template_folder = Path(tmp_dir)
            (template_folder / "Monday.docx").write_text("test")

            mock_doc = MagicMock()
            mock_doc.ProtectionType = 1  # Protected
            mock_doc.PrintOut = MagicMock()  # Accept any args
            mock_word_app.Documents.Open.return_value = mock_doc

            processor = WordProcessor()
            processor.initialize()

            processor.print_document(
                str(template_folder),
                "Monday",
                date(2026, 1, 15),
                "Test Printer"
            )

            mock_doc.Unprotect.assert_called_once()


@pytest.mark.skipif(not HAS_PYWIN32, reason="pywin32 not available")
class TestWordProcessorIntegration:
    """Integration tests for WordProcessor with mocked COM."""

    @patch('src.word_processor.pythoncom')
    @patch('src.word_processor.win32com.client')
    @patch('src.word_processor.win32print')
    def test_print_multiple_documents(self, mock_win32print, mock_win32, mock_pythoncom, tmp_path):
        """Test printing multiple documents in sequence."""
        # Mock printer status to be ready
        mock_win32print.OpenPrinter.return_value = 1
        mock_win32print.GetPrinter.return_value = {'Status': 0}
        mock_word_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_word_app

        # Create test templates
        template_folder = tmp_path / "templates"
        template_folder.mkdir()
        for day in ["Monday", "Tuesday", "Wednesday"]:
            (template_folder / f"{day}.docx").write_text(f"{day} template")

        mock_doc = MagicMock()
        mock_doc.ProtectionType = -1
        mock_doc.PrintOut = MagicMock()  # Accept any args
        mock_word_app.Documents.Open.return_value = mock_doc

        processor = WordProcessor()
        processor.initialize()

        results = []
        for day in ["Monday", "Tuesday", "Wednesday"]:
            success, error = processor.print_document(
                str(template_folder),
                day,
                date(2026, 1, 15),
                "Test Printer"
            )
            results.append((success, error))

        # All should succeed
        for success, error in results:
            assert success is True
            assert error is None

        # Should print 3 times
        assert mock_doc.PrintOut.call_count == 3
