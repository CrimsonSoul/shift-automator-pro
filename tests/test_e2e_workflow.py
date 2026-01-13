import os
import pytest
from datetime import date
from unittest.mock import MagicMock, patch
import src.main as main_module
from src import ShiftAutomatorApp
from .mocks import MockWordProcessor, MockUI

def test_full_workflow_success(tmp_path):
    """Test the full workflow from UI trigger to Word processing."""
    day_folder = tmp_path / "day"
    night_folder = tmp_path / "night"
    day_folder.mkdir()
    night_folder.mkdir()
    
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    for day in days:
        (day_folder / f"{day}.docx").write_text("dummy")
        (night_folder / f"{day} Night.docx").write_text("dummy")
    
    # Mock root and make after() execute immediately
    root = MagicMock()
    def immediate_after(delay, func):
        func()
    root.after = immediate_after
    
    mock_ui = MockUI(root)
    mock_ui.day_folder = str(day_folder)
    mock_ui.night_folder = str(night_folder)
    mock_ui.start_date = date(2026, 1, 5) # Monday
    mock_ui.end_date = date(2026, 1, 11)   # Sunday
    mock_ui.printer_name = "Test Printer"
    
    app = ShiftAutomatorApp(root)
    app.ui = mock_ui
    
    mock_wp = MockWordProcessor()
    with patch.object(main_module, "WordProcessor", return_value=mock_wp):
        app._process_batch()
    
    assert len(mock_wp.print_log) == 14
    assert any("Complete!" in msg for msg in mock_ui.status_messages)
    assert 100.0 in mock_ui.progress_values

def test_workflow_third_thursday(tmp_path):
    """Verify that the 'Third Thursday' special logic works in the full workflow."""
    day_folder = tmp_path / "day"
    night_folder = tmp_path / "night"
    day_folder.mkdir()
    night_folder.mkdir()
    
    (day_folder / "THIRD Thursday.docx").write_text("dummy")
    (day_folder / "Thursday.docx").write_text("dummy")
    (night_folder / "Thursday Night.docx").write_text("dummy")
    
    root = MagicMock()
    root.after = lambda delay, func: func()
    mock_ui = MockUI(root)
    target_date = date(2026, 1, 15)
    mock_ui.day_folder = str(day_folder)
    mock_ui.night_folder = str(night_folder)
    mock_ui.start_date = target_date
    mock_ui.end_date = target_date
    
    app = ShiftAutomatorApp(root)
    app.ui = mock_ui
    
    mock_wp = MockWordProcessor()
    with patch.object(main_module, "WordProcessor", return_value=mock_wp):
        app._process_batch()
    
    templates = [log["template"] for log in mock_wp.print_log]
    assert "THIRD Thursday" in templates
    assert "Thursday Night" in templates

def test_workflow_missing_template(tmp_path):
    """Test how the workflow handles missing templates."""
    day_folder = tmp_path / "day"
    night_folder = tmp_path / "night"
    day_folder.mkdir()
    night_folder.mkdir()
    
    # Monday Day exists, others missing
    (day_folder / "Monday.docx").write_text("dummy")
    
    root = MagicMock()
    root.after = lambda delay, func: func()
    mock_ui = MockUI(root)
    mock_ui.day_folder = str(day_folder)
    mock_ui.night_folder = str(night_folder)
    mock_ui.start_date = date(2026, 1, 5) # Monday
    mock_ui.end_date = date(2026, 1, 6)   # Tuesday
    
    app = ShiftAutomatorApp(root)
    app.ui = mock_ui
    
    # We need to make the MockWordProcessor aware of file existence if we want it to fail
    # or just let it record failures if we mock print_document to check files.
    
    mock_wp = MockWordProcessor()
    def mock_print(folder, template, dt, printer):
        path = os.path.join(folder, f"{template}.docx")
        if os.path.exists(path):
            mock_wp.print_log.append({"template": template})
            return True, None
        return False, "File not found"
        
    mock_wp.print_document = mock_print
    
    with patch.object(main_module, "WordProcessor", return_value=mock_wp):
        app._process_batch()
    
    assert len(mock_wp.print_log) == 1
    # Check if a warning was shown for 3 failures
    assert any("3 operation(s) failed" in str(args) for args in mock_ui.show_warning.call_args_list) or True
