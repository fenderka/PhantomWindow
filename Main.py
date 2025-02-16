
import sys
import psutil
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QComboBox,
                             QPushButton, QLabel, QLineEdit, QFormLayout,
                             QMessageBox, QGridLayout, QHBoxLayout, QScrollArea,
                             QGroupBox, QSpinBox, QSystemTrayIcon, QMenu, QAction,
                             QProgressBar, QShortcut)
from PyQt5.QtCore import Qt, QSize, QLocale, QSettings, QTimer
from PyQt5.QtGui import QIcon, QPixmap, QPalette, QColor

import win32gui
import win32con
import win32process

class OverlayApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PhantomWindow")

        # Settings
        self.settings = QSettings("MyCompany", "OverlayApp")
        self.load_settings()

        # UI elements
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск процесса...")
        self.search_edit.textChanged.connect(self.update_process_list)

        self.process_list_combo = QComboBox()
        self.process_list_combo.setIconSize(QSize(32, 32))
        self.process_list_combo.setMaxVisibleItems(10)
        self.process_list_combo.currentIndexChanged.connect(self.save_selected_process)

        # Use QSpinBox for integer inputs
        self.x_spin = QSpinBox()
        self.y_spin = QSpinBox()
        self.width_spin = QSpinBox()
        self.height_spin = QSpinBox()

        # Set ranges and initial values from settings
        self.x_spin.setRange(-9999, 9999)
        self.y_spin.setRange(-9999, 9999)
        self.width_spin.setRange(1, 9999)
        self.height_spin.setRange(1, 9999)

        self.x_spin.setValue(self.settings.value("x", 0))
        self.y_spin.setValue(self.settings.value("y", 0))
        self.width_spin.setValue(self.settings.value("width", 400))
        self.height_spin.setValue(self.settings.value("height", 300))

        self.overlay_button = QPushButton("Наложить")
        self.overlay_button.clicked.connect(self.overlay_selected_window)

        self.remove_overlay_button = QPushButton("Убрать наложение")
        self.remove_overlay_button.clicked.connect(self.remove_overlay)
        self.remove_overlay_button.setEnabled(False)

        # Toggle Theme Button
        self.toggle_theme_button = QPushButton("Переключить тему", self)
        self.toggle_theme_button.clicked.connect(self.toggle_theme)

        # Progress Bar
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 0)  # Set to busy indicator mode
        self.progress_bar.setVisible(False)


        self.hwnd = None
        self.target_pid = None
        self.dark_theme = False

        # Layout
        main_layout = QVBoxLayout()

        # Process selection group
        self.process_group = QGroupBox("Выбор процесса")
        process_group_layout = QVBoxLayout()
        process_group_layout.addWidget(QLabel("Выберите процесс:"))
        search_layout = QHBoxLayout()
        search_layout.addWidget(self.search_edit)
        process_group_layout.addLayout(search_layout)
        process_group_layout.addWidget(self.process_list_combo)
        self.process_group.setLayout(process_group_layout)
        main_layout.addWidget(self.process_group)

        # Position and size group
        self.position_group = QGroupBox("Позиция и размер")
        position_size_group_layout = QGridLayout()
        position_size_group_layout.addWidget(QLabel("X:"), 0, 0)
        position_size_group_layout.addWidget(self.x_spin, 0, 1)
        position_size_group_layout.addWidget(QLabel("Y:"), 0, 2)
        position_size_group_layout.addWidget(self.y_spin, 0, 3)
        position_size_group_layout.addWidget(QLabel("Ширина:"), 1, 0)
        position_size_group_layout.addWidget(self.width_spin, 1, 1)
        position_size_group_layout.addWidget(QLabel("Высота:"), 1, 2)
        position_size_group_layout.addWidget(self.height_spin, 1, 3)
        self.position_group.setLayout(position_size_group_layout)
        main_layout.addWidget(self.position_group)

        # Actions group
        self.button_group = QGroupBox("Действия")
        button_layout = QVBoxLayout()
        button_layout.addWidget(self.overlay_button)
        button_layout.addWidget(self.remove_overlay_button)
        button_layout.addWidget(self.toggle_theme_button)  # Add the toggle theme button
        button_layout.addWidget(self.progress_bar)
        self.button_group.setLayout(button_layout)
        main_layout.addWidget(self.button_group)

        self.setLayout(main_layout)

        # Styling (moved to a separate method)
        self.apply_styles()

        self.update_process_list()

        # System Tray Icon
        self.tray_icon = QSystemTrayIcon(self)  #  QIcon('icon.png')
        self.tray_icon.setIcon(self.get_app_icon())  # Use a placeholder get_app_icon function or load a real icon.
        self.tray_icon.setToolTip("PhantomWindow")
        tray_menu = QMenu()
        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.close)
        tray_menu.addAction(exit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

        # Hotkeys
        self.shortcut_overlay = QShortcut(Qt.CTRL + Qt.Key_O, self)
        self.shortcut_overlay.activated.connect(self.overlay_selected_window)

        self.shortcut_remove_overlay = QShortcut(Qt.CTRL + Qt.Key_R, self)
        self.shortcut_remove_overlay.activated.connect(self.remove_overlay)

        # Timer to refresh process list
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_process_list)
        self.timer.start(5000)  # Refresh every 5 seconds

    def get_app_icon(self):
        """Placeholder function to get application icon"""
        # Replace this with your actual method to get the application icon.
        return QIcon()

    def apply_styles(self):
        """Apply consistent styling to widgets."""
        if self.dark_theme:
            app_style = """
                QWidget {
                    background-color: #2E2E2E;
                    color: #FFFFFF;
                    font-family: Segoe UI, Arial, sans-serif;
                    font-size: 10pt;
                }
                QLabel {
                    color: #BBBBBB;
                }
            """

            groupbox_style = """
                QGroupBox {
                    border: 1px solid #555555;
                    border-radius: 4px;
                    margin-top: 0.5em;
                    color: #FFFFFF;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 7px;
                    padding: 0 5px 0 5px;
                    color: #FFFFFF;
                }
            """

            lineedit_style = """
                QLineEdit, QSpinBox {
                    border: 1px solid #777777;
                    border-radius: 3px;
                    padding: 4px;
                    background-color: #333333;
                    color: #FFFFFF;
                }
                QLineEdit:focus, QSpinBox:focus {
                    border: 1px solid #64B5F6;
                }
            """

            combobox_style = """
                QComboBox {
                    border: 1px solid #777777;
                    border-radius: 3px;
                    padding: 4px;
                    background-color: #333333;
                    color: #FFFFFF;
                }
                QComboBox:hover {
                    background-color: #555555;
                }
                QComboBox::drop-down {
                    border: 0px;
                }
                QComboBox::down-arrow {
                    image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAh0lEQVQ4T2NkoBAwUqifgbsDqCcMDcQ4gFgDyE5gaCmEAQVhQA0YkE5hE2gGiF8g4gEQTwPxAdQ7kMwJ0AcgjEcQeQPyLCAgAAKMA0AKQjQdgAAAABJRU5ErkJggg==);
                    width: 16px;
                    height: 16px;
                }
            """

            button_style = """
                QPushButton {
                    background-color: #3C3C3C;
                    border: none;
                    color: white;
                    padding: 8px 16px;
                    text-align: center;
                    text-decoration: none;
                    display: inline-block;
                    font-size: 10pt;
                    margin: 4px 2px;
                    cursor: pointer;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #555555;
                }
                QPushButton:disabled {
                    background-color: #555555;
                    color: #777777;
                }
            """

            remove_button_style = """
                QPushButton {
                    background-color: #D32F2F;
                    border: none;
                    color: white;
                    padding: 8px 16px;
                    text-align: center;
                    text-decoration: none;
                    display: inline-block;
                    font-size: 10pt;
                    margin: 4px 2px;
                    cursor: pointer;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #C62828;
                }
            """

        else:
            app_style = """
                QWidget {
                    background-color: #F5F5F5;
                    color: #333;
                    font-family: Segoe UI, Arial, sans-serif;
                    font-size: 10pt;
                }
                QLabel {
                    color: #555;
                }
            """

            groupbox_style = """
                QGroupBox {
                    border: 1px solid #8F8F91;
                    border-radius: 4px;
                    margin-top: 0.5em;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 7px;
                    padding: 0 5px 0 5px;
                    color: #555;
                }
            """

            lineedit_style = """
                QLineEdit, QSpinBox {
                    border: 1px solid #A0A0A0;
                    border-radius: 3px;
                    padding: 4px;
                    background-color: #FFFFFF;
                    color: #333;
                }
                QLineEdit:focus, QSpinBox:focus {
                    border: 1px solid #64B5F6;
                }
            """

            combobox_style = """
                QComboBox {
                    border: 1px solid #A0A0A0;
                    border-radius: 3px;
                    padding: 4px;
                    background-color: #FFFFFF;
                    color: #333;
                }
                QComboBox:hover {
                    background-color: #E0E0E0;
                }
                QComboBox::drop-down {
                    border: 0px;
                }
                QComboBox::down-arrow {
                    image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAh0lEQVQ4T2NkoBAwUqifgbsDqCcMDcQ4gFgDyE5gaCmEAQVhQA0YkE5hE2gGiF8g4gEQTwPxAdQ7kMwJ0AcgjEcQeQPyLCAgAAKMA0AKQjQdgAAAABJRU5ErkJggg==);
                    width: 16px;
                    height: 16px;
                }
            """

            button_style = """
                QPushButton {
                    background-color: #4CAF50;
                    border: none;
                    color: white;
                    padding: 8px 16px;
                    text-align: center;
                    text-decoration: none;
                    display: inline-block;
                    font-size: 10pt;
                    margin: 4px 2px;
                    cursor: pointer;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #A0A0A0;
                    color: #606060;
                }
            """

            remove_button_style = """
                QPushButton {
                    background-color: #F44336;
                    border: none;
                    color: white;
                    padding: 8px 16px;
                    text-align: center;
                    text-decoration: none;
                    display: inline-block;
                    font-size: 10pt;
                    margin: 4px 2px;
                    cursor: pointer;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #D32F2F;
                }
            """

        self.setStyleSheet(app_style)
        self.search_edit.setStyleSheet(lineedit_style)
        self.process_list_combo.setStyleSheet(combobox_style)
        self.x_spin.setStyleSheet(lineedit_style)
        self.y_spin.setStyleSheet(lineedit_style)
        self.width_spin.setStyleSheet(lineedit_style)
        self.height_spin.setStyleSheet(lineedit_style)
        self.overlay_button.setStyleSheet(button_style)
        self.remove_overlay_button.setStyleSheet(remove_button_style)
        self.toggle_theme_button.setStyleSheet(button_style) # Use the same button style or create a different one.
        self.process_group.setStyleSheet(groupbox_style)
        self.position_group.setStyleSheet(groupbox_style)
        self.button_group.setStyleSheet(groupbox_style)

    def load_settings(self):
        self.last_selected_pid = self.settings.value("last_selected_pid", type=int) if self.settings.contains(
            "last_selected_pid") else None

    def save_selected_process(self, index):
        if index >= 0:
            selected_pid = self.process_list_combo.itemData(index)
            self.settings.setValue("last_selected_pid", selected_pid)
            self.settings.sync()

    def closeEvent(self, event):
        self.save_settings()
        event.accept()

    def save_settings(self):
        self.settings.setValue("x", self.x_spin.value())
        self.settings.setValue("y", self.y_spin.value())
        self.settings.setValue("width", self.width_spin.value())
        self.settings.setValue("height", self.height_spin.value())
        self.settings.sync()

    def is_system_process(self, process):
        try:
            exe_path = process.exe()
            if exe_path.startswith((
                    "C:\\Windows",
                    "C:\\ProgramData\\Microsoft\\Windows",
                    "\\SystemRoot"
            )):
                return True
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            pass
        return False

    def update_process_list(self):
        self.process_list_combo.clear()
        search_term = self.search_edit.text().lower()

        def is_window_visible(hwnd):
            """Определяет, видимо ли окно (т.е., появляется ли в Alt+Tab)."""
            if not win32gui.IsWindowVisible(hwnd):
                return False
            if win32gui.GetWindowText(hwnd) == "":  # Ignore windows with empty titles
                return False

            # Exclude tooltips and other special windows
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            if ex_style & win32con.WS_EX_TOOLWINDOW:
                return False

            # Exclude windows that don't have an owner
            if win32gui.GetWindowLong(hwnd, win32con.GWL_HWNDPARENT) != 0:
                return False

            return True


        processes_with_visible_windows = set()  # Use a set to avoid duplicates
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                process = psutil.Process(proc.info['pid'])
                if self.is_system_process(process):
                    continue

                process_name = proc.info['name']
                process_pid = proc.info['pid']


                def callback(hwnd, hwnds):
                    try:
                        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                        if found_pid == process_pid and is_window_visible(hwnd):
                            hwnds.append(hwnd)
                    except win32gui.error:
                        pass
                    return True

                hwnds = []
                try:
                    win32gui.EnumWindows(callback, hwnds)

                    if hwnds:  # Only add the process if it has visible windows
                        processes_with_visible_windows.add(process_pid)

                        if search_term and search_term not in process_name.lower():
                            continue

                        icon = self.get_process_icon(process)
                        self.process_list_combo.addItem(icon, f"{process_name} (PID: {process_pid})", process_pid)

                        if self.last_selected_pid is not None and process_pid == self.last_selected_pid:
                            self.process_list_combo.setCurrentIndex(self.process_list_combo.count() - 1)


                except Exception as e:
                    print(f"Ошибка при обработке процесса {process_pid}: {e}")

            except (psutil.NoSuchProcess, psutil.AccessDenied, KeyError) as e:
                print(f"Ошибка при обработке процесса: {e}")
                pass

    def get_process_icon(self, process):
        try:
            exe_path = process.exe()
            icon = QIcon(exe_path)
            if not icon.isNull():
                return icon
        except (psutil.AccessDenied, psutil.NoSuchProcess) as e:
            print(f"Отказано в доступе или процесс не найден: {e}")
            pass
        except Exception as e:
            print(f"Произошла неожиданная ошибка: {e}")
        return QIcon()

    def overlay_selected_window(self):
        self.start_progress() # Starts progress bar

        self.target_pid = self.process_list_combo.currentData()
        self.hwnd = self.get_main_window_handle(self.target_pid)

        if self.hwnd:
            print(f"Найдено окно HWND: {self.hwnd}")
            self.set_window_topmost(self.hwnd)
            self.resize_and_move_window(self.hwnd)

            self.overlay_button.setEnabled(False)
            self.remove_overlay_button.setEnabled(True)

        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось найти дескриптор окна.")

        self.stop_progress() # Stops progress bar


    def remove_overlay(self):
        self.start_progress() # Starts progress bar

        if self.hwnd:
            try:
                win32gui.SetWindowPos(self.hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                print("Наложение убрано.")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при удалении наложения: {e}")
            finally:
                self.overlay_button.setEnabled(True)
                self.remove_overlay_button.setEnabled(False)
                self.hwnd = None

        self.stop_progress() # Stops progress bar

    def get_main_window_handle(self, pid):
        """
        Находит главное окно процесса, учитывая окна без заголовка.
        Определяет главное окно по наличию заголовка и меню,
        либо возвращает первое найденное окно, если ни одно не подходит.
        """
        def callback(hwnd, hwnds):
            try:
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    hwnds.append(hwnd)
            except win32gui.error:
                pass
            return True

        hwnds = []
        try:
            win32gui.EnumWindows(callback, hwnds)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при перечислении окон: {e}")
            return None

        if not hwnds:
            return None

        # Попытка найти окно с заголовком и меню
        for hwnd in hwnds:
            if win32gui.GetWindowText(hwnd) and win32gui.GetMenu(hwnd):
                return hwnd

        # Если окно с заголовком и меню не найдено, вернуть первое окно
        return hwnds[0]

    def set_window_topmost(self, hwnd):
        try:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка",
                                 f"Ошибка при установке окна на передний план: {e}")

    def resize_and_move_window(self, hwnd):
        try:
            x = self.x_spin.value()
            y = self.y_spin.value()
            width = self.width_spin.value()
            height = self.height_spin.value()

            win32gui.MoveWindow(hwnd, x, y, width, height, True)
        except ValueError:
            QMessageBox.warning(self, "Предупреждение",
                                "Неверный ввод для позиции или размера. Пожалуйста, введите целые числа.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка",
                                 f"Ошибка при изменении размера/перемещении окна: {e}")

    def toggle_theme(self):
        """Toggle the dark theme"""
        self.dark_theme = not self.dark_theme
        self.apply_styles()

    def start_progress(self):
        """Show progress bar"""
        self.progress_bar.setVisible(True)

    def stop_progress(self):
        """Hide progress bar"""
        self.progress_bar.setVisible(False)

if __name__ == '__main__':
    QLocale.setDefault(QLocale("ru_RU"))

    app = QApplication(sys.argv)

    # Global stylesheet
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#F5F5F5"))  # Slightly off-white
    palette.setColor(QPalette.WindowText, QColor("#333333"))
    palette.setColor(QPalette.Base, QColor("#FFFFFF"))
    palette.setColor(QPalette.AlternateBase, QColor("#F0F0F0"))
    palette.setColor(QPalette.Text, QColor("#333333"))
    palette.setColor(QPalette.Button, QColor("#4CAF50"))
    palette.setColor(QPalette.ButtonText, QColor("white"))
    palette.setColor(QPalette.Highlight, QColor("#388E3C"))
    palette.setColor(QPalette.HighlightedText, QColor("white"))
    app.setPalette(palette)


    window = OverlayApp()
    window.show()
    sys.exit(app.exec_())
