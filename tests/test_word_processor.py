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

        with patch.object(wp, "_normalize_spaces_in_doc"), \
             patch.object(wp, "_execute_replace", return_value=True) as mock_exec:
            wp.replace_dates(mock_doc, current_date)

            # Should be called 4 times (for 4 patterns)
            assert mock_exec.call_count == 4

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

    def test_find_template_third_thursday_extra_spaces(self, wp, tmp_path):
        """Should find 'THIRD Thursday' even if filename has extra spaces."""
        (tmp_path / "THIRD  Thursday.docx").write_text("dummy")

        result = wp.find_template_file(str(tmp_path), "THIRD Thursday")
        assert result is not None
        assert "Thursday" in result

    def test_third_thursday_integration(self, wp, tmp_path):
        """Integration: scheduler template name should find the right file."""
        from src.scheduler import get_shift_template_name

        (tmp_path / "Thursday.docx").write_text("dummy")
        (tmp_path / "THIRD Thursday.docx").write_text("dummy")

        # January 15, 2026 is the third Thursday
        template_name = get_shift_template_name(date(2026, 1, 15), "day")
        assert template_name == "THIRD Thursday"

        result = wp.find_template_file(str(tmp_path), template_name)
        assert result is not None
        assert "third" in Path(result).name.lower()

    def test_replace_dates_no_match_warning(self, wp):
        """Should log warning when no date patterns match."""
        mock_doc = MagicMock()
        current_date = date(2026, 1, 14)  # Wednesday

        with patch.object(wp, "_normalize_spaces_in_doc"), \
             patch.object(wp, "_execute_replace", return_value=False), \
             patch("src.word_processor.logger") as mock_logger:
            wp.replace_dates(mock_doc, current_date)
            mock_logger.warning.assert_called()

    def test_normalize_spaces_called_before_patterns(self, wp):
        """Should normalize non-breaking spaces before running date patterns."""
        mock_doc = MagicMock()
        current_date = date(2026, 1, 14)

        call_order = []

        def track_normalize(doc):
            call_order.append("normalize")

        def track_execute(doc, find_text, replace_text):
            call_order.append("execute")
            return False

        with patch.object(wp, "_normalize_spaces_in_doc", side_effect=track_normalize), \
             patch.object(wp, "_execute_replace", side_effect=track_execute):
            wp.replace_dates(mock_doc, current_date)

        assert call_order[0] == "normalize"
        assert "execute" in call_order

    def test_run_find_replace_returns_bool(self, wp):
        """_run_find_replace should return True when pattern matches."""
        mock_range = MagicMock()
        mock_range.Find.Execute.return_value = True

        result = wp._run_find_replace(mock_range, "pattern", "replacement")
        assert result is True

    def test_run_find_replace_returns_false_on_no_match(self, wp):
        """_run_find_replace should return False when pattern doesn't match."""
        mock_range = MagicMock()
        mock_range.Find.Execute.return_value = False

        result = wp._run_find_replace(mock_range, "pattern", "replacement")
        assert result is False
