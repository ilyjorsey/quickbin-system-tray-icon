import ctypes
import os
import winreg
from typing import Optional

from PyQt6.QtCore import QTimer, QObject, pyqtSignal
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication, QSystemTrayIcon
from winotify import Notification

from settings import *


class RecycleBinManager:
    @staticmethod
    def open_bin():
        os.startfile(RECYCLEBINPATH)

    @staticmethod
    def clear_bin():
        she_empty_recycle_bin = ctypes.windll.shell32.SHEmptyRecycleBinW
        she_empty_recycle_bin.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
        she_empty_recycle_bin.restype = ctypes.c_int

        result = she_empty_recycle_bin(None, None, SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND)

        toast = Notification(
            app_id='QuickBin',
            title='Recycle Bin Cleared' if result == 0 else "Recycle Bin Error",
            msg="Successful emptying of the recycle bin." if result == 0 else "Failed to empty Recycle Bin.",
            duration="short",
        )
        toast.show()


class ThemeWatcher(QObject):
    theme_changed = pyqtSignal(bool)

    def __init__(self):
        super().__init__()
        self.last_theme = self.is_light_theme()

    @staticmethod
    def is_light_theme() -> bool:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, WINDOWS_THEME_PATH) as key:
            value = winreg.QueryValueEx(key, "AppsUseLightTheme")[0]
            return value == 1

    def check_theme(self) -> None:
        current_theme = self.is_light_theme()
        if current_theme != self.last_theme:
            self.last_theme = current_theme
            self.theme_changed.emit(current_theme)


class BinTrayIcon(QSystemTrayIcon):
    def __init__(self, parent: Optional[QObject] = None):
        super().__init__(parent)

        self.bin_manager = RecycleBinManager()
        self.single_click_timer = QTimer()
        self.single_click_timer.setSingleShot(True)
        self.double_click_detected = False
        self.theme_watcher = ThemeWatcher()

        self.single_click_timer.timeout.connect(self.handle_single_click)

        self.setup_theme_watcher()

        self.update_icon(self.theme_watcher.is_light_theme())

    def setup_theme_watcher(self) -> None:

        self.theme_watcher.theme_changed.connect(self.update_icon)

        theme_check_timer = QTimer()
        theme_check_timer.timeout.connect(self.theme_watcher.check_theme)
        theme_check_timer.start(1000)

    def handle_click(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.double_click_detected = False
            self.single_click_timer.start(200)
        elif reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.double_click_detected = True
            self.single_click_timer.stop()
            self.bin_manager.clear_bin()

    def handle_single_click(self) -> None:
        if not self.double_click_detected:
            self.bin_manager.open_bin()

    def update_icon(self, is_light_theme: bool) -> None:
        theme_folder = "icons/light_theme" if is_light_theme else "icons/dark_theme"
        icon_path = os.path.join(theme_folder, "empty.png")
        self.setIcon(QIcon(icon_path))


class TrayApplication:
    def __init__(self):
        self.app = QApplication([])
        self.app.setQuitOnLastWindowClosed(False)

    def run(self) -> None:
        tray_icon = BinTrayIcon()
        tray_icon.activated.connect(tray_icon.handle_click)
        tray_icon.show()
        self.app.exec()


if __name__ == "__main__":
    application = TrayApplication()
    application.run()
