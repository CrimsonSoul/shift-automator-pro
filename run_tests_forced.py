import sys
import os
from unittest.mock import patch

# Add src to sys.path
sys.path.insert(0, os.path.join(os.getcwd(), 'src'))

import pytest

# Force HAS_PYWIN32 to True to run tests that would otherwise be skipped
with patch('src.word_processor.HAS_PYWIN32', True):
    # We also need to mock the modules that would be imported
    with patch('src.word_processor.pythoncom'), \
         patch('src.word_processor.win32com.client'), \
         patch('src.word_processor.win32com.client.dynamic'), \
         patch('src.word_processor.win32print'):
        exit_code = pytest.main(['tests/test_word_processor.py', '-v'])
        sys.exit(exit_code)
