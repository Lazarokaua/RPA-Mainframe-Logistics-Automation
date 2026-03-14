import sys
import os

def get_resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS

        # Priority 1: Check if file exists next to the executable (User Editable)
        # sys.executable points to the .exe file
        exe_dir = os.path.dirname(sys.executable)
        external_path = os.path.normpath(os.path.join(exe_dir, relative_path))

        if os.path.exists(external_path):
            return external_path

    except AttributeError:
        # If not running as a PyInstaller bundle, find the project root
        # This file is in src/utils/resource_handler.py
        # We need to go up two levels: src/utils -> src -> dist_prep (project root)
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.dirname(current_dir)
        project_root = os.path.dirname(src_dir)

        base_path = project_root

    return os.path.normpath(os.path.join(base_path, relative_path))
