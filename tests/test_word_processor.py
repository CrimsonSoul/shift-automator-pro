"""
Unit tests for word_processor module.
"""

import os
import pytest
from unittest.mock import MagicMock, patch
from pathlib import Path
from datetime import date

from src.word_processor import WordProcessor

class TestWordProcessor:
    """Tests for WordProcessor class."""

    @pytest.fixture
    def wp(self):
        """Create a WordProcessor instance."""
        # Patch pythoncom and win32com to avoid errors during initialization
        with patch("pythoncom.CoInitialize"), \
             patch("win32com.client.Dispatch"):
            wp = WordProcessor()
            yield wp

    def test_init(self, wp):
        """WordProcessor should initialize with default values."""
        assert wp.word_app is None
        assert wp._initialized is False
        assert wp._template_cache == {}

    def test_find_template_file_exact_match(self, wp, tmp_path):
        """Should find template with exact match."""
        # Create dummy templates
        (tmp_path / "Monday.docx").write_text("dummy")
        (tmp_path / "Tuesday.docx").write_text("dummy")
        
        result = wp.find_template_file(str(tmp_path), "Monday")
        assert result is not None
        assert result.endswith("Monday.docx")

    def test_find_template_file_cache_usage(self, wp, tmp_path):
        """Should use cache for subsequent lookups."""
        (tmp_path / "Monday.docx").write_text("dummy")
        
        # First call builds cache
        wp.find_template_file(str(tmp_path), "Monday")
        assert str(tmp_path.resolve()) in wp._template_cache
        
        # Modify folder (add file) but cache should still be used
        (tmp_path / "Tuesday.docx").write_text("dummy")
        result = wp.find_template_file(str(tmp_path), "Tuesday")
        # Tuesday wasn't in cache, so it should NOT be found if cache is strict
        # Wait, the current implementation checks if folder is in cache.
        # If folder is in cache, it only looks at that cache.
        assert result is None 

    def test_robust_template_matching(self, wp, tmp_path):
        """Should match 'Thursday Night' when searching for 'Thursday' if unique."""
        (tmp_path / "Thursday Night.docx").write_text("dummy")
        (tmp_path / "Friday.docx").write_text("dummy")
        
        result = wp.find_template_file(str(tmp_path), "Thursday")
        assert result is not None
        assert result.endswith("Thursday Night.docx")

    def test_robust_template_matching_boundary(self, wp, tmp_path):
        """Should NOT match 'THIRD Thursday' when searching for 'Thursday'."""
        (tmp_path / "THIRD Thursday.docx").write_text("dummy")
        
        result = wp.find_template_file(str(tmp_path), "Thursday")
        assert result is None

    def test_ambiguous_template_matching(self, wp, tmp_path):
        """Should handle ambiguous matches gracefully."""
        (tmp_path / "Thursday.docx").write_text("dummy")
        (tmp_path / "Thursday Night.docx").write_text("dummy")
        
        # "Thursday" matches both. It should prefer the one starting with "Thursday"
        # or the exact match.
        result = wp.find_template_file(str(tmp_path), "Thursday")
        assert result is not None
        assert result.endswith("Thursday.docx")

    def test_replace_dates_logic(self, wp):
        """Should call find/replace with correct patterns."""
        mock_doc = MagicMock()
        current_date = date(2026, 1, 15) # Thursday
        
        with patch.object(wp, "_execute_replace") as mock_exec:
            wp.replace_dates(mock_doc, current_date)
            
            # Should be called 3 times (for 3 patterns)
            assert mock_exec.call_count == 3
            
            # Verify one of the calls
            # Pattern: "[A-Za-z]@, [A-Za-z]@ [0-9]{1,2}, [0-9]{4}"
            # Replacement: "Thursday, January 15, 2026"
            calls = [c[0][2] for c in mock_exec.call_args_list]
            assert "Thursday, January 15, 2026" in calls

    @patch("src.word_processor.pythoncom.CoInitialize")
    @patch("src.word_processor.win32com.client.Dispatch")
    def test_initialize_success(self, mock_dispatch, mock_coinit):
        """Initialize should set up COM correctly."""
        # Create a fresh WordProcessor without the fixture's patches
        wp = WordProcessor()
        wp.initialize()
        assert wp._initialized is True
        assert wp.word_app is not None
        mock_coinit.assert_called_once()
        mock_dispatch.assert_called_with("Word.Application")

    def test_safe_com_call_retry(self, wp):
        """Safe COM call should retry on rejection."""
        mock_func = MagicMock()
        # Fail twice with "rejected", then succeed
        mock_func.side_effect = [
            Exception("Call was rejected by callee"),
            Exception("Rejected"),
            "Success"
        ]
        
        with patch("time.sleep"): # Don't actually wait
            result = wp.safe_com_call(mock_func, "arg1", retries=3)
            
        assert result == "Success"
        assert mock_func.call_count == 3

    def test_safe_com_call_fail(self, wp):
        """Safe COM call should eventually fail."""
        mock_func = MagicMock(side_effect=Exception("Permanent Failure"))
        
        with pytest.raises(Exception, match="Permanent Failure"):
            wp.safe_com_call(mock_func, retries=2)

    def test_clear_template_cache(self, wp, tmp_path):
        """Should clear the template cache."""
        (tmp_path / "Monday.docx").write_text("dummy")
        wp.find_template_file(str(tmp_path), "Monday")
        assert str(tmp_path.resolve()) in wp._template_cache
        
        wp.clear_template_cache()
        assert wp._template_cache == {}
