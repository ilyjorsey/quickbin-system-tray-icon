import configparser
import ctypes
import os
import sys
import winreg
from typing import Optional

import win32api
import win32event
import winerror
from PyQt6.QtCore import QTimer, QObject, pyqtSignal
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import QApplication, QSystemTrayIcon, QMenu
from winotify import Notification

from utils import resource_path

config = configparser.ConfigParser()
config_path = resource_path('config.ini')
config.read(config_path)


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
            self._set_loading_icon(self.theme_watcher.get_current_theme())
            self.bin_manager.clear_bin()
            self._update_icon(self.theme_watcher.get_current_theme())

    def _handle_single_click(self) -> None:
        if not self.double_click_detected:
            self.bin_manager.open_bin()

    def _update_icon(self, is_light_theme: bool) -> None:
        theme_folder = "icons/light_theme" if is_light_theme else "icons/dark_theme"
        icon_path = resource_path(os.path.join(theme_folder, "empty.png"))
        self.setIcon(QIcon(icon_path))

    def _set_loading_icon(self, is_light_theme: bool) -> None:
        theme_folder = "icons/light_theme" if is_light_theme else "icons/dark_theme"
        icon_path = resource_path(os.path.join(theme_folder, "loading.png"))
        self.setIcon(QIcon(icon_path))

    def _set_start_at_boot(self) -> None:
        run_path = config.get('Settings', 'START_AT_BOOT_PATH')
        app_name = config.get('Settings', 'APP_NAME')
        exe_path = os.path.abspath(sys.argv[0])

        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, run_path, 0, winreg.KEY_WRITE) as key:
                if config.getboolean('Settings', 'IS_START_AT_BOOT'):
                    winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{exe_path}"')
                else:
                    try:
                        winreg.DeleteValue(key, app_name)
                    except FileNotFoundError:
                        pass
        except PermissionError:
            pass

    def _create_tray_menu(self) -> None:
        menu = QMenu()

        double_click_action = QAction("Double click emptying", self)
        double_click_action.setCheckable(True)
        double_click_action.setChecked(config.getboolean('Settings', 'EMPTYDOUBLECLICK'))
        double_click_action.triggered.connect(self._toggle_double_click)

        start_at_boot = QAction("Start at boot", self)
        start_at_boot.setCheckable(True)
        start_at_boot.setChecked(config.getboolean('Settings', 'IS_START_AT_BOOT'))
        start_at_boot.triggered.connect(self._toggle_start_at_boot)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self._exit_app)

        menu.addAction(start_at_boot)
        menu.addAction(double_click_action)
        menu.addSeparator()
        menu.addAction(exit_action)

        self.setContextMenu(menu)

    def _toggle_double_click(self, checked: bool) -> None:
        config.set('Settings', 'EMPTYDOUBLECLICK', str(checked))
        with open(config_path, 'w') as configfile:
            config.write(configfile)

    def _toggle_start_at_boot(self, checked: bool) -> None:
        config.set('Settings', 'IS_START_AT_BOOT', str(checked))
        with open(config_path, 'w') as configfile:
            config.write(configfile)
        self._set_start_at_boot()

    @staticmethod
    def _exit_app() -> None:
        QApplication.quit()


class TrayApplication:
    """Main application"""

    def __init__(self):
        self.app = QApplication([])
        self.app.setQuitOnLastWindowClosed(False)
        self.mutex = None

    def run(self) -> None:
        mutex_name = 'QuickBinMutex'
        _mutex = win32event.CreateMutex(None, False, mutex_name)  # Mutex for prevented double app launch
        error = win32api.GetLastError()

        if error == winerror.ERROR_ALREADY_EXISTS:
            sys.exit(0)
        else:
            tray_icon = BinTrayIcon()
            tray_icon.activated.connect(tray_icon.handle_click)
            tray_icon.show()
            self.app.exec()


if __name__ == "__main__":
    application = TrayApplication()
    application.run()
