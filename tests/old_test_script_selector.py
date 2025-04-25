import unittest
from unittest.mock import patch, MagicMock
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from main import run_script, open_excel, clear_excel

class TestScriptSelector(unittest.TestCase):
    @patch("script_selector.messagebox")
    @patch("script_selector.load_config")
    def test_run_script(self, mock_load_config, mock_messagebox):
        # Mock the config and file existence
        mock_load_config.return_value = {"CIFS": {}}
        with patch("os.path.exists", return_value=True):
            with patch("script_selector.generate_commands", return_value="dummy_path.txt"):
                run_script("CIFS", "Create")
                mock_messagebox.showinfo.assert_called_once()

    @patch("script_selector.messagebox")
    @patch("script_selector.load_config")
    def test_open_excel(self, mock_load_config, mock_messagebox):
        # Mock the config and file existence
        mock_load_config.return_value = {"CIFS": {}}
        with patch("os.path.exists", return_value=True):
            open_excel("CIFS", "Create")
            mock_messagebox.showerror.assert_not_called()

if __name__ == "__main__":
    unittest.main()