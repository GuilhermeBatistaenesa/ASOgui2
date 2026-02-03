import sys
import os
import pytest
from unittest.mock import MagicMock, patch

# Add project root to sys.path to import from main.py
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from main import _extract_gdrive_file_ids, _gdrive_name_matches, download_gdrive_file, GDRIVE_NAME_FILTER

def test_extract_gdrive_file_ids():
    # Test case 1: Normal link
    text1 = "Here is the link: https://drive.google.com/file/d/1bvdjgOE-LeX4eeEdKPhjUR1WBclEyDc1/view?usp=sharing"
    ids1 = _extract_gdrive_file_ids(text1)
    assert "1bvdjgOE-LeX4eeEdKPhjUR1WBclEyDc1" in ids1
    assert len(ids1) == 1

    # Test case 2: Multiple links
    text2 = """
    Link 1: https://drive.google.com/file/d/ID_ONE_123/view
    Link 2: https://drive.google.com/open?id=ID_TWO_456
    """
    ids2 = _extract_gdrive_file_ids(text2)
    assert "ID_ONE_123" in ids2
    assert "ID_TWO_456" in ids2
    assert len(ids2) == 2

    # Test case 3: HTML encoded
    text3 = "Link: https://drive.google.com/file/d/ID_ENCODED_789/view?usp=sharing"
    # Simulating HTML entity encoding if necessary, but the function calls unescape.
    text3_html = "Link: https://drive.google.com/file/d/ID_ENCODED_789/view?usp=sharing"
    ids3 = _extract_gdrive_file_ids(text3_html)
    assert "ID_ENCODED_789" in ids3

    # Test case 4: No links
    ids4 = _extract_gdrive_file_ids("Just some text without links.")
    assert len(ids4) == 0

def test_gdrive_name_matches():
    # Assuming default filter is "asos enesa"
    # We might need to mock GDRIVE_NAME_FILTER if we want to be independent of env var
    
    # Matches
    assert _gdrive_name_matches("ASOS ENESA.pdf")
    assert _gdrive_name_matches("aso admissional asos enesa 2024.pdf")
    
    # Does not match
    # Note: depends on actual env var. If env var is set, this might fail if we don't patch/mock it. 
    # But main.py reads env at import time. We can try to patch the constant in main if needed, 
    # or just assume the default if not set. 
    # Let's rely on the logic: if filter is set, it checks containment.
    
    if GDRIVE_NAME_FILTER:
        assert not _gdrive_name_matches("random_file.pdf")

@patch('main.urllib.request.build_opener')
@patch('main.GDRIVE_NAME_FILTER', "")  # Disable filter for this test
def test_download_gdrive_file_direct(mock_build_opener, tmp_path):
    # Setup mock response for direct download
    mock_response = MagicMock()
    mock_response.headers = {
        "Content-Disposition": 'attachment; filename="ASOS ENESA.pdf"',
        "Content-Type": "application/pdf"
    }
    # Return some bytes
    mock_response.read.side_effect = [b"%PDF-1.4 content", b""]
    mock_response.close = MagicMock()
    
    mock_opener = MagicMock()
    mock_opener.open.return_value = mock_response
    mock_build_opener.return_value = mock_opener

    dest_dir = str(tmp_path)
    file_id = "TEST_FILE_ID"
    
    # Execute
    result_path = download_gdrive_file(file_id, dest_dir)
    
    # Assert
    assert result_path is not None
    # _safe_filename does NOT replace spaces, so it should be preserved
    assert os.path.basename(result_path) == "ASOS ENESA.pdf" 
    assert os.path.exists(result_path)
    with open(result_path, "rb") as f:
        content = f.read()
        assert content == b"%PDF-1.4 content"

@patch('main.urllib.request.build_opener')
@patch('main.GDRIVE_NAME_FILTER', "")  # Disable filter
def test_download_gdrive_file_confirm(mock_build_opener, tmp_path):
    # Setup mock response 1: Confirmation page
    mock_resp1 = MagicMock() # First response (HTML with confirm token)
    mock_resp1.headers = {"Content-Type": "text/html"}
    html_content = '<html>...<a href="/uc?export=download&amp;confirm=TOKEN123&amp;id=TEST_ID">Download</a>...</html>'
    mock_resp1.read.side_effect = [html_content.encode('utf-8'), b""]
    mock_resp1.close = MagicMock()

    # Setup mock response 2: Actual file
    mock_resp2 = MagicMock()
    mock_resp2.headers = {
        "Content-Disposition": 'attachment; filename="BIG_FILE.pdf"',
        "Content-Type": "application/pdf"
    }
    mock_resp2.read.side_effect = [b"%PDF-1.4 big content", b""]
    mock_resp2.close = MagicMock()

    mock_opener = MagicMock()
    # First call returns resp1, second call returns resp2
    mock_opener.open.side_effect = [mock_resp1, mock_resp2] 
    mock_build_opener.return_value = mock_opener

    dest_dir = str(tmp_path)
    file_id = "TEST_ID"

    # Execute
    result_path = download_gdrive_file(file_id, dest_dir)

    # Assert
    assert result_path is not None
    assert os.path.basename(result_path) == "BIG_FILE.pdf"
    with open(result_path, "rb") as f:
        assert f.read() == b"%PDF-1.4 big content"
