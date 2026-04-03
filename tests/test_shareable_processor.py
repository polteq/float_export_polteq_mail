import pytest
import pandas as pd
from datetime import datetime
import sys
import os
from pathlib import Path

# Add the parent directory to the path so we can import shareable_processor
sys.path.insert(0, str(Path(__file__).parent.parent))

from shareable_processor import TimesheetProcessor

def test_escape_applescript_string():
    assert TimesheetProcessor.escape_applescript_string("hello") == "hello"
    assert TimesheetProcessor.escape_applescript_string("hello \"world\"") == 'hello \\"world\\"'
    assert TimesheetProcessor.escape_applescript_string("path\\to\\file") == "path\\\\to\\\\file"
    assert TimesheetProcessor.escape_applescript_string("mix\\\"match\"") == 'mix\\\\\\"match\\"'
    assert TimesheetProcessor.escape_applescript_string(None) == ""
    assert TimesheetProcessor.escape_applescript_string(123) == "123"

def test_extract_end_date_from_filename():
    # Valid filename should extract date correctly (2026-01-31)
    filename = "Jorre van Munster-LoggedTime-20260101-20260131.csv"
    end_date = TimesheetProcessor.extract_end_date(filename)
    assert end_date == datetime(2026, 1, 31)

def test_extract_end_date_from_dataframe_fallback():
    # Invalid filename, but DataFrame has a Date column
    filename = "my_custom_timesheet.csv"
    
    df = pd.DataFrame({
        'Date': ['2026-02-01', '2026-02-15', '2026-02-28'],
        'Hours': [8, 8, 8]
    })
    
    end_date = TimesheetProcessor.extract_end_date(filename, df)
    assert end_date == datetime(2026, 2, 28)

def test_extract_end_date_failure():
    # Invalid filename AND no DataFrame
    filename = "invalid_name.csv"
    
    with pytest.raises(ValueError, match="Could not determine end date"):
        TimesheetProcessor.extract_end_date(filename)
        
    # Invalid filename AND DataFrame without Date column
    df = pd.DataFrame({
        'Day': ['Monday'],
        'Hours': [8]
    })
    with pytest.raises(ValueError, match="Could not determine end date"):
        TimesheetProcessor.extract_end_date(filename, df)
