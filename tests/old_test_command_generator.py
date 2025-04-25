import sys
import os
import unittest
import pandas as pd
from unittest.mock import patch, MagicMock

# Add the parent directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from command_generator import generate_commands

class TestCommandGenerator(unittest.TestCase):
    def setUp(self):
        # Sample data for testing
        self.cifs_create_data = pd.DataFrame({
            "name": ["share1"],
            "local_path": ["/path/to/share"],
            "oplock_enabled": ["yes"],
            "notify_enabled": ["no"]
        })
        self.nfs_create_data = pd.DataFrame({
            "local_path": ["/path/to/nfs"],
            "charset": ["UTF-8"],
            "lock_type": ["Mandatory"]
        })
        self.filesystem_create_data = pd.DataFrame({
            "name": ["fs1"],
            "workload_type_id": ["1"],
            "pool_id": ["pool1"],
            "initial_distribute_policy": ["automatic"]
        })

    def test_generate_commands_cifs_create(self):
        # Test CIFS Create command generation
        resource_type = "CIFS"
        command_type = "Create"
        mandatory_fields = ["name", "local_path"]

        commands = generate_commands(self.cifs_create_data, resource_type, command_type, mandatory_fields)
        expected_commands = [
            "create share cifs name=share1 local_path=/path/to/share oplock_enabled=yes notify_enabled=no"
        ]
        self.assertEqual(commands, expected_commands)

    def test_generate_commands_nfs_create(self):
        # Test NFS Create command generation
        resource_type = "NFS"
        command_type = "Create"
        mandatory_fields = ["local_path"]

        commands = generate_commands(self.nfs_create_data, resource_type, command_type, mandatory_fields)
        expected_commands = [
            "create share nfs local_path=/path/to/nfs charset=UTF-8 lock_type=Mandatory"
        ]
        self.assertEqual(commands, expected_commands)

    def test_generate_commands_filesystem_create(self):
        # Test FileSystem Create command generation
        resource_type = "FileSystem"
        command_type = "Create"
        mandatory_fields = ["name"]

        commands = generate_commands(self.filesystem_create_data, resource_type, command_type, mandatory_fields)
        expected_commands = [
            "create file_system general name=fs1 workload_type_id=1 pool_id=pool1 initial_distribute_policy=automatic"
        ]
        self.assertEqual(commands, expected_commands)

    def test_generate_commands_missing_mandatory_fields(self):
        # Test handling of missing mandatory fields
        resource_type = "CIFS"
        command_type = "Create"
        mandatory_fields = ["name", "local_path"]

        # Data with missing mandatory field
        data = pd.DataFrame({
            "name": ["share1"]
        })

        commands = generate_commands(data, resource_type, command_type, mandatory_fields)
        self.assertEqual(commands, [])

    def test_generate_commands_empty_dataframe(self):
        # Test handling of empty DataFrame
        resource_type = "CIFS"
        command_type = "Create"
        mandatory_fields = ["name", "local_path"]

        # Empty DataFrame
        data = pd.DataFrame()

        commands = generate_commands(data, resource_type, command_type, mandatory_fields)
        self.assertEqual(commands, ["create share cifs"])

    def test_generate_commands_show_command(self):
        # Test Show command generation
        resource_type = "CIFS"
        command_type = "Show"
        mandatory_fields = []

        # Empty DataFrame for Show command
        data = pd.DataFrame()

        commands = generate_commands(data, resource_type, command_type, mandatory_fields)
        self.assertEqual(commands, ["show share cifs"])

if __name__ == "__main__":
    unittest.main()