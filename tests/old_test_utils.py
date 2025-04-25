import unittest
from unittest.mock import patch, mock_open
import json
import pandas as pd
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from utils import load_config, read_file

class TestUtils(unittest.TestCase):
    def test_load_config(self):
        # Mock the config file
        mock_config = '{"key": "value"}'
        with patch("builtins.open", mock_open(read_data=mock_config)):
            config = load_config()
            self.assertEqual(config, {"key": "value"})

    def test_read_file(self):
        # Mock an Excel file
        mock_data = pd.DataFrame({"col1": [1, 2], "col2": [3, 4]})
        with patch("pandas.read_excel", return_value=mock_data):
            result = read_file("dummy_path.xlsx", "Sheet1")
            self.assertTrue(result.equals(mock_data))

if __name__ == "__main__":
    unittest.main()