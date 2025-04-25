import unittest
from unittest.mock import patch, MagicMock
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from gui_helpers import apply_style, create_buttons_and_dropdown

class TestGUIHelpers(unittest.TestCase):
    @patch("tkinter.ttk.Combobox")
    def test_create_buttons_and_dropdown(self, mock_combobox):
        # Mock the root window and config
        mock_root = MagicMock()
        mock_config = {"CIFS": {"Create": {}}}
        script_var = MagicMock()
        command_var = MagicMock()

        # Call the function
        create_buttons_and_dropdown(mock_root, MagicMock(), MagicMock(), script_var, command_var, MagicMock(), mock_config)

        # Assert that the Combobox was created
        mock_combobox.assert_called_once()

if __name__ == "__main__":
    unittest.main()