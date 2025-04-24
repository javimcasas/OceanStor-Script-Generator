import os
import sys
import pandas as pd
import re
from utils import load_config, read_file, get_data_file_path

class CommandGenerator:
    def __init__(self):
        self.field_processors = {
            'text': self.process_text_field,
            'select': self.process_select_field,
            'list': self.process_list_field
        }
        self.field_transforms = {
            'prefix_slash': self.transform_prefix_slash
        }

    def process_text_field(self, value, field_config):
        """Process text field type"""
        if pd.isna(value) or value == '':
            return None
        return str(value).strip()

    def process_select_field(self, value, field_config):
        """Process select field type with allowed values"""
        if pd.isna(value) or value == '':
            return field_config.get('default')
        
        value = str(value).strip().lower()
        allowed = [v.lower() for v in field_config['allowed_values']]
        
        if value not in allowed:
            raise ValueError(f"Invalid value '{value}' for field {field_config['name']}. Allowed: {allowed}")
        return value

    def process_list_field(self, value, field_config):
        """Process list field type with separator"""
        if pd.isna(value) or value == '':
            return None
            
        separator = field_config.get('separator', ',')
        items = [item.strip() for item in re.split(r"[\s,]+", str(value).strip()) if item.strip()]
        return separator.join(items)

    def transform_prefix_slash(self, value):
        """Transform to ensure path starts with slash"""
        value = str(value).strip()
        if not value.startswith('/'):
            return f"/{value.lstrip('/')}"
        return value

    def process_field_value(self, field_config, value):
        """Process a field value based on its configuration"""
        if pd.isna(value) or value == '':
            return None
            
        field_type = field_config.get('field_type', 'text')
        processor = self.field_processors.get(field_type, self.process_text_field)
        
        # First apply any transforms
        if 'transform' in field_config:
            transform = self.field_transforms.get(field_config['transform'])
            if transform:
                value = transform(value)
        
        # Then process the field
        return processor(value, field_config)

    def generate_command(self, row, operation_config):
        """Generate a single command based on row data and operation config"""
        command_parts = [operation_config['cli_prefix']]
        
        # Process mandatory fields
        for field in operation_config.get('mandatory', []):
            field_name = field['name']
            value = row.get(field_name)
            processed_value = self.process_field_value(field, value)
            
            if processed_value is None:
                raise ValueError(f"Missing mandatory field: {field_name}")
            
            command_parts.append(f"{field_name}={processed_value}")
        
        # Process optional fields
        for field in operation_config.get('optional', []):
            field_name = field['name']
            if field_name in row:
                processed_value = self.process_field_value(field, row[field_name])
                if processed_value is not None:
                    command_parts.append(f"{field_name}={processed_value}")
        
        return ' '.join(command_parts)

    def generate_commands(self, data_frame, resource_type, command_type, config):
        """Generate all commands for a given operation"""
        try:
            operation_config = config[resource_type]['operations'][command_type]
        except KeyError:
            raise ValueError(f"Unsupported operation: {resource_type}/{command_type}")
        
        commands = []
        for index, row in data_frame.iterrows():
            try:
                commands.append(self.generate_command(row, operation_config))
            except ValueError as e:
                print(f"Warning: Skipping row {index + 1} - {str(e)}")
                continue
        
        return commands

    def main(self, resource_type, device_type):
        base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.abspath(".")
        excel_file_path = os.path.join(base_path, 'Documents', f'{device_type}_{resource_type}_commands.xlsx')
        results_dir = os.path.join(base_path, 'Results')

        if not os.path.exists(results_dir):
            os.makedirs(results_dir)

        config = load_config(device_type)
        resource_config = config.get(resource_type, {})

        if not os.path.exists(excel_file_path):
            print(f"Error: Excel file '{excel_file_path}' does not exist.")
            sys.exit(1)

        all_commands = []
        for command_type in resource_config.get('operations', {}).keys():
            data_frame = read_file(excel_file_path, sheet_name=command_type)
            if data_frame is None:
                continue

            commands = self.generate_commands(data_frame, resource_type, command_type, config)
            all_commands.extend(commands)
            all_commands.append('')  # Add empty line between command groups

        output_file_path = os.path.join(results_dir, f'{device_type.lower()}_{resource_type.lower()}_commands.txt')
        with open(output_file_path, 'w') as f:
            f.write('\n'.join(all_commands))

        print(f"Commands written to: {output_file_path}")
        return output_file_path

def main(resource_type, device_type):
    generator = CommandGenerator()
    return generator.main(resource_type, device_type)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python command_generator.py <resource_type>")
        sys.exit(1)

    resource_type = sys.argv[1]
    main(resource_type)