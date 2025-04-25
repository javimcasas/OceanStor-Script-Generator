import unittest
import tkinter as tk
from unittest.mock import patch, MagicMock
import os
import sys

# Add the parent directory to the path so we can import the modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import run_script, main

class TestMainFunctions(unittest.TestCase):
    @patch('main.messagebox')
    @patch('main.toggle_loading')
    @patch('main.generate_commands')
    @patch('main.load_config')
    def test_run_script_success(self, mock_load_config, mock_generate_commands, mock_toggle_loading, mock_messagebox):
        # Setup mocks
        mock_load_config.return_value = {
            'test_script': {'operations': {'test_command': {}}}
        }
        mock_generate_commands.return_value = '/path/to/output.txt'
        
        # Create a mock root window
        root = tk.Tk()
        root.withdraw()  # Hide the window
        
        # Test successful execution
        run_script('test_script', 'test_command', 'test_device')
        
        # Verify calls
        mock_toggle_loading.assert_called()
        mock_generate_commands.assert_called_once_with('test_script', 'test_device')
        mock_messagebox.showinfo.assert_called()
        
    @patch('main.messagebox')
    @patch('main.toggle_loading')
    @patch('main.load_config')
    def test_run_script_invalid_type(self, mock_load_config, mock_toggle_loading, mock_messagebox):
        # Setup mocks
        mock_load_config.return_value = {
            'valid_script': {}
        }
        
        # Create a mock root window
        root = tk.Tk()
        root.withdraw()  # Hide the window
        
        # Test invalid script type
        run_script('invalid_script', 'test_command', 'test_device')
        
        # Verify calls
        mock_messagebox.showerror.assert_called()
        mock_toggle_loading.assert_called()
        
    @patch('main.tk.Tk')
    @patch('main.apply_style')
    @patch('main.create_buttons_and_dropdown')
    @patch('main.load_config')
    def test_main_function(self, mock_load_config, mock_create_buttons, mock_apply_style, mock_tk):
        # Setup mocks
        mock_load_config.return_value = {
            'OceanStor Dorado': {'script1': {}},
            'OceanStor Pacific': {'script2': {}}
        }
        
        # Mock Tkinter elements
        mock_root = MagicMock()
        mock_tk.return_value = mock_root
        
        # Run main function
        main()
        
        # Verify calls
        mock_tk.assert_called_once()
        mock_apply_style.assert_called_once_with(mock_root)
        mock_create_buttons.assert_called_once()

if __name__ == '__main__':
    unittest.main()