import pytest
from datetime import date
from unittest.mock import patch
from src import ShiftAutomatorApp
from .mocks import MockUI, MockWordProcessor

def test_workflow_full_run_success(mock_wp: MockWordProcessor, mock_ui: MockUI):
    """
    Test a full successful run of the batch processing.
    Ensures that for each day in range, both shifts are processed.
    """
    import tkinter as tk
    root = tk.Tk()
    # Mock root.after to execute callbacks immediately
    root.after = lambda delay, func, *args: func(*args)
    
    app = ShiftAutomatorApp(root)
    
    # Inject mocks
    # Use type: ignore to bypass strict type checking for mocks in tests
    app.ui = mock_ui  # type: ignore
    app.word_processor = mock_wp # type: ignore
    
    # Setup dates: Jan 1 to Jan 2 (2 days, 4 shifts total)
    start_date = date(2024, 1, 1)
    end_date = date(2024, 1, 2)
    
    mock_ui.set_dates(start_date, end_date)
    mock_ui.set_folders("C:/Day", "C:/Night")
    mock_ui.set_printer("TestPrinter")
    
    # Mock WordProcessor.__enter__ and __exit__ since it's used as a context manager
    with patch("src.main.WordProcessor", return_value=mock_wp):
        app._process_batch()
    
    # Verify results
    assert mock_wp.call_count == 4
    assert mock_ui.status_history[-1] == "Processing COMPLETE!"
    assert not any("Error" in str(msg) for msg in mock_ui.info_messages)

def test_workflow_partial_failure(mock_wp: MockWordProcessor, mock_ui: MockUI):
    """Test workflow when some prints fail."""
    import tkinter as tk
    root = tk.Tk()
    # Mock root.after to execute callbacks immediately
    root.after = lambda delay, func, *args: func(*args)
    
    app = ShiftAutomatorApp(root)
    app.ui = mock_ui  # type: ignore
    
    start_date = date(2024, 1, 1)
    end_date = date(2024, 1, 1) # 1 day, 2 shifts
    
    mock_ui.set_dates(start_date, end_date)
    mock_ui.set_folders("C:/Day", "C:/Night")
    mock_ui.set_printer("TestPrinter")
    
    # Make one shift fail
    def fail_on_night(folder, template_name, current_date, printer_name):
        mock_wp.call_count += 1
        if "night" in template_name.lower():
            return False, "File not found"
        return True, None
        
    mock_wp.print_document = fail_on_night # type: ignore
    
    with patch("src.main.WordProcessor", return_value=mock_wp):
        app._process_batch()
        
    # Verify results
    assert mock_wp.call_count == 2
    # Check that warning was shown for failure
    assert any("failed" in str(msg).lower() for msg in mock_ui.warning_messages)

def test_workflow_cancellation(mock_wp: MockWordProcessor, mock_ui: MockUI):
    """Test that cancellation stops the process early."""
    import tkinter as tk
    root = tk.Tk()
    # Mock root.after to execute callbacks immediately
    root.after = lambda delay, func, *args: func(*args)
    
    app = ShiftAutomatorApp(root)
    app.ui = mock_ui  # type: ignore
    
    # Range of 10 days
    start_date = date(2024, 1, 1)
    end_date = date(2024, 1, 10)
    
    mock_ui.set_dates(start_date, end_date)
    mock_ui.set_folders("C:/Day", "C:/Night")
    mock_ui.set_printer("TestPrinter")
    
    # Trigger cancellation after first print
    def cancel_after_one(folder, template_name, current_date, printer_name):
        mock_wp.call_count += 1
        app._cancel_requested = True
        return True, None
        
    mock_wp.print_document = cancel_after_one # type: ignore
    
    with patch("src.main.WordProcessor", return_value=mock_wp):
        app._process_batch()
        
    # Should have stopped after first day shift or night shift
    # In the implementation, it checks after each shift.
    assert mock_wp.call_count <= 2 
    assert any("Cancelled" in str(msg) for msg in mock_ui.info_messages)
