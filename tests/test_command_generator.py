import unittest
import pandas as pd
import os
import sys
from unittest.mock import patch, MagicMock
from io import StringIO

# Add the parent directory to the path so we can import the modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from command_generator import CommandGenerator

class TestCommandGenerator(unittest.TestCase):
    def setUp(self):
        self.generator = CommandGenerator()
        
    def test_process_text_field(self):
        # Test normal text field
        field_config = {'name': 'test_field', 'field_type': 'text'}
        self.assertEqual(self.generator.process_text_field(" test ", field_config), "test")
        self.assertEqual(self.generator.process_text_field("", field_config), None)
        self.assertEqual(self.generator.process_text_field(None, field_config), None)
        
    def test_process_select_field(self):
        # Test select field with allowed values
        field_config = {
            'name': 'test_field',
            'field_type': 'select',
            'allowed_values': ['Yes', 'No'],
            'default': 'No'
        }
        self.assertEqual(self.generator.process_select_field("Yes", field_config), "yes")
        self.assertEqual(self.generator.process_select_field("NO", field_config), "no")
        self.assertEqual(self.generator.process_select_field("", field_config), "No")
        self.assertEqual(self.generator.process_select_field(None, field_config), "No")
        
        with self.assertRaises(ValueError):
            self.generator.process_select_field("Maybe", field_config)
            
    def test_process_list_field(self):
        # Test list field processing
        field_config = {
            'name': 'test_field',
            'field_type': 'list',
            'separator': ','
        }
        self.assertEqual(self.generator.process_list_field("a, b, c", field_config), "a,b,c")
        self.assertEqual(self.generator.process_list_field("a b c", field_config), "a,b,c")
        self.assertEqual(self.generator.process_list_field("", field_config), None)
        
    def test_transform_prefix_slash(self):
        # Test path prefix transformation
        self.assertEqual(self.generator.transform_prefix_slash("path"), "/path")
        self.assertEqual(self.generator.transform_prefix_slash("/path"), "/path")
        self.assertEqual(self.generator.transform_prefix_slash("//path"), "//path")
        
    def test_process_field_value(self):
        # Test field processing with transform
        field_config = {
            'name': 'path_field',
            'field_type': 'text',
            'transform': 'prefix_slash'
        }
        self.assertEqual(self.generator.process_field_value(field_config, "test"), "/test")
        
    def test_generate_command(self):
        # Test command generation
        row = {'name': 'test', 'size': '100', 'enabled': 'Yes'}
        operation_config = {
            'cli_prefix': 'create',
            'mandatory': [
                {'name': 'name', 'field_type': 'text'},
                {'name': 'size', 'field_type': 'text'}
            ],
            'optional': [
                {'name': 'enabled', 'field_type': 'select', 'allowed_values': ['Yes', 'No']}
            ]
        }
        expected = "create name=test size=100 enabled=yes"
        self.assertEqual(self.generator.generate_command(row, operation_config), expected)
        
    @patch('pandas.read_excel')
    @patch('os.path.exists')
    @patch('os.makedirs')
    def test_main(self, mock_makedirs, mock_exists, mock_read_excel):
        # Test the main function with mocked dependencies
        mock_exists.return_value = True
        
        # Mock DataFrame
        mock_df = pd.DataFrame({
            'name': ['test1', 'test2'],
            'size': ['100', '200']
        })
        mock_read_excel.return_value = mock_df
        
        # Mock config
        config = {
            'test_resource': {
                'operations': {
                    'create': {
                        'cli_prefix': 'create',
                        'mandatory': [
                            {'name': 'name', 'field_type': 'text'},
                            {'name': 'size', 'field_type': 'text'}
                        ]
                    }
                }
            }
        }
        
        with patch('command_generator.load_config', return_value=config):
            generator = CommandGenerator()
            output_path = generator.main('test_resource', 'test_device')
            
            self.assertTrue(output_path.endswith('test_device_test_resource_commands.txt'))
            mock_read_excel.assert_called_once()
            mock_makedirs.assert_not_called()  # Because we mocked exists to return True

if __name__ == '__main__':
    unittest.main()