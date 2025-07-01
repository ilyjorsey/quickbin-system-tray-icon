import os
import sys


def resource_path(relative_path):
    """Returns a path considering execution in PyInstaller or normal mode"""

    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)

    return os.path.join(os.path.dirname(os.path.relpath(__file__)), relative_path)
