import sys
import os
import unittest
from unittest.mock import patch, MagicMock

# Add the parent directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from file_operations import create_excel_for_resource

class TestFileOperations(unittest.TestCase):
    @patch("openpyxl.Workbook")
    @patch("os.path.exists", return_value=False)  # Mock os.path.exists to return False
    def test_create_excel_for_resource_new_file(self, mock_exists, mock_workbook):
        # Mock the workbook and its methods
        mock_workbook_instance = MagicMock()
        mock_workbook.return_value = mock_workbook_instance

        # Mock the sheet creation
        mock_sheet = MagicMock()
        mock_workbook_instance.create_sheet.return_value = mock_sheet

        # Mock the config loading
        with patch("file_operations.load_config") as mock_load_config:
            mock_load_config.return_value = {
                "CIFS": {
                    "Create": {
                        "mandatory": [{"name": "field1", "field_type": "text"}],
                        "optional": [{"name": "field2", "field_type": "text"}]
                    }
                }
            }

            # Call the function
            create_excel_for_resource("CIFS", "dummy_path.xlsx")

            # Assert that the workbook was saved
            mock_workbook_instance.save.assert_called_once_with("dummy_path.xlsx")

    @patch("openpyxl.Workbook")
    @patch("os.path.exists", return_value=True)  # Mock os.path.exists to return True
    def test_create_excel_for_resource_existing_file(self, mock_exists, mock_workbook):
        # Mock the workbook and its methods
        mock_workbook_instance = MagicMock()
        mock_workbook.return_value = mock_workbook_instance

        # Call the function
        create_excel_for_resource("CIFS", "dummy_path.xlsx")

        # Assert that the workbook was NOT saved
        mock_workbook_instance.save.assert_not_called()

if __name__ == "__main__":
    unittest.main()