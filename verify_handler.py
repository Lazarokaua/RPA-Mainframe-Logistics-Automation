import sys
import os
import unittest
from unittest.mock import patch

# Adjust path to import resource_handler
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))
from utils.resource_handler import get_resource_path

class TestResourceHandler(unittest.TestCase):
    def test_priority_external_file(self):
        # Normalize paths for Windows comparison
        mock_meipass = os.path.normpath('C:/tmp/meipass')
        mock_exe = os.path.normpath('C:/app/myapp.exe')

        # Setup: Mock sys._MEIPASS to simulate frozen app
        with patch.object(sys, '_MEIPASS', mock_meipass, create=True):
            with patch.object(sys, 'executable', mock_exe):
                with patch('os.path.exists') as mock_exists:
                    # Scenario 1: External file exists
                    exe_dir = os.path.dirname(mock_exe)
                    relative = os.path.normpath('output/base_filiais.xlsx')
                    expected_external_path = os.path.join(exe_dir, relative)

                    # Ensure mock_exists returns True ONLY for the expected external path
                    def side_effect(p):
                        return p == expected_external_path

                    mock_exists.side_effect = side_effect

                    # Call with the same relative path structure
                    path = get_resource_path('output/base_filiais.xlsx')

                    print(f"DEBUG: Expected: {expected_external_path}")
                    print(f"DEBUG: Actual:   {path}")

                    self.assertEqual(path, expected_external_path)

    def test_fallback_embedded_file(self):
        # Normalize paths for Windows comparison
        mock_meipass = os.path.normpath('C:/tmp/meipass')
        mock_exe = os.path.normpath('C:/app/myapp.exe')

        # Setup: Mock sys._MEIPASS to simulate frozen app
        with patch.object(sys, '_MEIPASS', mock_meipass, create=True):
            with patch.object(sys, 'executable', mock_exe):
                with patch('os.path.exists') as mock_exists:
                    # Scenario 2: External file DOES NOT exist
                    mock_exists.return_value = False

                    relative = os.path.normpath('output/base_filiais.xlsx')
                    # Expect fallback to MEIPASS
                    expected_embedded_path = os.path.join(mock_meipass, relative)

                    path = get_resource_path('output/base_filiais.xlsx')

                    print(f"DEBUG: Expected: {expected_embedded_path}")
                    print(f"DEBUG: Actual:   {path}")

                    self.assertEqual(path, expected_embedded_path)

if __name__ == '__main__':
    unittest.main()
