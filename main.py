import ctypes
import os
import winreg
from typing import Optional

from PyQt6.QtCore import QTimer, QObject, pyqtSignal
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from winotify import Notification
import configparser

config = configparser.ConfigParser()
config.read('config.ini')


class RecycleBinManager:
    """Handles operations related to the Windows Recycle Bin"""

    @staticmethod
    def open_bin() -> None:
        os.startfile('shell:RecycleBinFolder')

    @staticmethod
    def clear_bin() -> None:
        she_empty_recycle_bin = ctypes.windll.shell32.SHEmptyRecycleBinW
        she_empty_recycle_bin.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_int]
        she_empty_recycle_bin.restype = ctypes.c_int

        noc = int(config['Settings']['SHERB_NOCONFIRMATION'], 16)
        nosound = int(config['Settings']['SHERB_NOSOUND'], 16)
        noprogress = int(config['Settings']['SHERB_NOPROGRESSUI'], 16)

        result = she_empty_recycle_bin(
            None,
            None,
            noc | nosound | noprogress
        )

        toast = Notification(
            app_id='QuickBin',
            title='Recycle Bin Cleared' if result == 0 else "Recycle Bin Error",
            msg="Successful emptying of the recycle bin." if result == 0 else "Failed to empty Recycle Bin.",
            duration="short",
        )
        toast.show()


class ThemeWatcher(QObject):
    """Monitors Windows theme changes"""

    theme_changed = pyqtSignal(bool)

    def __init__(self):
        super().__init__()
        self.last_theme = self.get_current_theme()
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_theme)
        self.timer.start(1000)

    @staticmethod
    def get_current_theme() -> bool:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, config['Settings']['WINDOWS_THEME_PATH']) as key:
            value = winreg.QueryValueEx(key, "AppsUseLightTheme")[0]
            return value == 1


    def check_theme(self) -> None:
        current_theme = self.get_current_theme()
        if current_theme != self.last_theme:
            self.last_theme = current_theme
            self.theme_changed.emit(current_theme)


class BinTrayIcon(QSystemTrayIcon):
    """System tray icon with Recycle Bin functionality"""

    def __init__(self, parent: Optional[QObject] = None):
        super().__init__(parent)

        # Components
        self.bin_manager = RecycleBinManager()
        self.theme_watcher = ThemeWatcher()

        # Click handling
        self.single_click_timer = QTimer()
        self.single_click_timer.setSingleShot(True)
        self.double_click_detected = False
        self.single_click_timer.timeout.connect(self._handle_single_click)

        # Setup
        self.theme_watcher.theme_changed.connect(self._update_icon)
        self._update_icon(self.theme_watcher.get_current_theme())
        self._create_tray_menu()

    def handle_click(self, reason: QSystemTrayIcon.ActivationReason) -> None:
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self._handle_single_click_init()
        elif reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self._handle_double_click()

    def _handle_single_click_init(self) -> None:
        self.double_click_detected = False
        self.single_click_timer.start(200)

    def _handle_double_click(self) -> None:
        if config.getboolean('Settings', 'EMPTYDOUBLECLICK'):
            self.double_click_detected = True
            self.single_click_timer.stop()
            self.bin_manager.clear_bin()

    def _handle_single_click(self) -> None:
        if not self.double_click_detected:
            self.bin_manager.open_bin()

    def _update_icon(self, is_light_theme: bool) -> None:
        theme_folder = "icons/light_theme" if is_light_theme else "icons/dark_theme"
        icon_path = os.path.join(theme_folder, "empty.png")
        self.setIcon(QIcon(icon_path))

    def _create_tray_menu(self) -> None:
        menu = QMenu()

        double_click_action = QAction("Double click emptying", self)
        double_click_action.setCheckable(True)
        double_click_action.setChecked(config.getboolean('Settings', 'EMPTYDOUBLECLICK'))
        double_click_action.triggered.connect(self._toggle_double_click)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self._exit_app)

        menu.addAction(double_click_action)
        menu.addSeparator()
        menu.addAction(exit_action)

        self.setContextMenu(menu)

    @staticmethod
    def _toggle_double_click(checked: bool) -> None:
        config.set('Settings', 'EMPTYDOUBLECLICK', str(checked))
        with open('config.ini', 'w') as configfile: # type: ignore
            config.write(configfile)

    @staticmethod
    def _exit_app() -> None:
        QApplication.quit()


class TrayApplication:
    """Main application"""

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
