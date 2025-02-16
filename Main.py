import sys
import psutil
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QComboBox,
                             QPushButton, QLabel, QLineEdit, QFormLayout,
                             QMessageBox, QGridLayout, QHBoxLayout, QScrollArea,
                             QGroupBox)
from PyQt5.QtCore import Qt, QSize, QLocale, QSettings
from PyQt5.QtGui import QIcon, QPixmap, QPalette, QColor

import win32gui
import win32con
import win32process


class OverlayApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PhantomWindow")  # Window title

        # Settings
        self.settings = QSettings("MyCompany", "OverlayApp")
        self.load_settings()

        # UI elements
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск процесса...")  # Search placeholder

        # Apply style to search edit
        self.search_edit.setStyleSheet("""
            QLineEdit {
                border: 1px solid #4CAF50; /* Green border */
                border-radius: 5px;
                padding: 5px;
                background-color: #F0F0F0;
                color: #333;
            }
            QLineEdit:focus {
                border: 1px solid #388E3C; /* Darker green on focus */
                background-color: white;
            }
        """)

        self.search_edit.textChanged.connect(self.update_process_list)  # Connect search

        self.process_list_combo = QComboBox()
        self.process_list_combo.setIconSize(QSize(32, 32))  # Set icon size
        self.process_list_combo.setMaxVisibleItems(10)  # Limit visible items

        # Apply style to combo box
        self.process_list_combo.setStyleSheet("""
            QComboBox {
                border: 1px solid #4CAF50;
                border-radius: 5px;
                padding: 5px;
                background-color: white;
                color: #333;
            }
            QComboBox:hover {
                background-color: #E0E0E0;
            }
            QComboBox::drop-down {
                border: 0px;
            }
            QComboBox::down-arrow {
                image: url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAh0lEQVQ4T2NkoBAwUqifgbsDqCcMDcQ4gFgDyE5gaCmEAQVhQA0YkE5hE2gGiF8g4gEQTwPxAdQ7kMwJ0AcgjEcQeQPyLCAgAAKMA0AKQjQdgAAAABJRU5ErkJggg==); /* Replace with your own arrow icon */
                width: 16px;
                height: 16px;
            }
        """)

        self.x_edit = QLineEdit(str(self.settings.value("x", 0)))  # Load from settings
        self.y_edit = QLineEdit(str(self.settings.value("y", 0)))  # Load from settings
        self.width_edit = QLineEdit(str(self.settings.value("width", 400)))  # Load from settings
        self.height_edit = QLineEdit(str(self.settings.value("height", 300)))  # Load from settings

        # Apply style to line edits
        lineEditStyle = """
            QLineEdit {
                border: 1px solid #4CAF50;
                border-radius: 5px;
                padding: 5px;
                background-color: white;
                color: #333;
            }
            QLineEdit:focus {
                border: 1px solid #388E3C;
                background-color: white;
            }
        """
        self.x_edit.setStyleSheet(lineEditStyle)
        self.y_edit.setStyleSheet(lineEditStyle)
        self.width_edit.setStyleSheet(lineEditStyle)
        self.height_edit.setStyleSheet(lineEditStyle)

        self.overlay_button = QPushButton("Наложить")  # Overlay button text

        # Apply style to overlay button
        self.overlay_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 10px 20px;
                text-align: center;
                text-decoration: none;
                display: inline-block;
                font-size: 16px;
                margin: 4px 2px;
                cursor: pointer;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #388E3C;
            }
            QPushButton:disabled {
                background-color: #A0A0A0;
                color: #606060;
            }
        """)

        self.overlay_button.clicked.connect(self.overlay_selected_window)

        self.remove_overlay_button = QPushButton("Убрать наложение")  # Remove overlay button text

        # Apply style to remove overlay button
        self.remove_overlay_button.setStyleSheet("""
            QPushButton {
                background-color: #F44336;
                border: none;
                color: white;
                padding: 10px 20px;
                text-align: center;
                text-decoration: none;
                display: inline-block;
                font-size: 16px;
                margin: 4px 2px;
                cursor: pointer;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #D32F2F;
            }
            QPushButton:disabled {
                background-color: #A0A0A0;
                color: #606060;
            }
        """)

        self.remove_overlay_button.clicked.connect(self.remove_overlay)
        self.remove_overlay_button.setEnabled(False)  # Disable initially

        # Store the HWND and PID of the target window
        self.hwnd = None
        self.target_pid = None

        # Layout
        main_layout = QVBoxLayout()

        # Group processes list
        process_group = QGroupBox("Выбор процесса")  # Group box for process selection
        process_group_layout = QVBoxLayout()
        process_group_layout.addWidget(QLabel("Выберите процесс:"))  # Label
        search_layout = QHBoxLayout()
        search_layout.addWidget(self.search_edit)
        process_group_layout.addLayout(search_layout)
        process_group_layout.addWidget(self.process_list_combo)
        process_group.setLayout(process_group_layout)
        main_layout.addWidget(process_group)

        # Group positions and sizes list
        position_group = QGroupBox("Позиция и размер")  # Group box for position
        position_size_group_layout = QGridLayout()
        position_size_group_layout.addWidget(QLabel("X:"), 0, 0)
        position_size_group_layout.addWidget(self.x_edit, 0, 1)
        position_size_group_layout.addWidget(QLabel("Y:"), 0, 2)
        position_size_group_layout.addWidget(self.y_edit, 0, 3)
        position_size_group_layout.addWidget(QLabel("Ширина:"), 1, 0)
        position_size_group_layout.addWidget(self.width_edit, 1, 1)
        position_size_group_layout.addWidget(QLabel("Высота:"), 1, 2)
        position_size_group_layout.addWidget(self.height_edit, 1, 3)
        position_group.setLayout(position_size_group_layout)  # Set layout for the position

        # Apply style to group boxes
        groupbox_style = """
            QGroupBox {
                border: 2px solid #4CAF50;
                border-radius: 5px;
                margin-top: 1ex; /* leave space at the top for the title */
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top center; /* position at the top center */
                padding: 0 3px;
                background-color: #F0F0F0;
                color: #333;
            }
        """

        process_group.setStyleSheet(groupbox_style)
        position_group.setStyleSheet(groupbox_style)

        # Group buttons to do things with program
        button_group = QGroupBox("Действия")  # Group box for actions
        button_layout = QVBoxLayout()
        button_layout.addWidget(self.overlay_button)
        button_layout.addWidget(self.remove_overlay_button)
        button_group.setLayout(button_layout)  # Set layout for the buttons
        button_group.setStyleSheet(groupbox_style)

        # Apply styles to the main layout
        self.setStyleSheet("""
            QWidget {
                background-color: #F0F0F0;
                color: #333;
                font-family: Arial;
                font-size: 12pt;
            }
            QLabel {
                color: #333;
            }
        """)

        # Add all group to main layout
        main_layout.addWidget(position_group)  # Set layout for the positions
        main_layout.addWidget(button_group)  # Set layout for the actions

        self.setLayout(main_layout)
        self.update_process_list()  # Initial list update

        self.process_list_combo.currentIndexChanged.connect(self.save_selected_process)

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
        # Save window position and size
        self.settings.setValue("x", int(self.x_edit.text()))
        self.settings.setValue("y", int(self.y_edit.text()))
        self.settings.setValue("width", int(self.width_edit.text()))
        self.settings.setValue("height", int(self.height_edit.text()))
        self.settings.sync()

    def is_system_process(self, process):
        """Определяет, является ли процесс системным, на основе его пути."""  # Docstring in Russian
        try:
            # Get the path to the executable
            exe_path = process.exe()
            # Check if the path starts with a typical system directory
            if exe_path.startswith((
                    "C:\\Windows",
                    "C:\\ProgramData\\Microsoft\\Windows",
                    "\\SystemRoot"  # For OS that is not installed on C
            )):
                return True
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            # Ignore processes we can't get information about.  Treat as not system.
            pass
        return False

    def update_process_list(self):
        self.process_list_combo.clear()
        search_term = self.search_edit.text().lower()

        for proc in psutil.process_iter(['pid', 'name']):
            try:
                # Directly create a Process object with the PID
                process = psutil.Process(proc.info['pid'])

                if self.is_system_process(process):
                    continue  # Skip system processes

                process_name = proc.info['name']
                process_pid = proc.info['pid']

                if search_term and search_term not in process_name.lower():
                    continue  # Skip if not matching search

                icon = self.get_process_icon(process)  # Get the icon

                self.process_list_combo.addItem(icon, f"{process_name} (PID: {process_pid})", process_pid)

                if self.last_selected_pid is not None and process_pid == self.last_selected_pid:
                    self.process_list_combo.setCurrentIndex(self.process_list_combo.count() - 1)  # Select last selected
            except (psutil.NoSuchProcess, psutil.AccessDenied, KeyError) as e:
                print(f"Ошибка при обработке процесса: {e}")  # Error message in Russian
                # Handle cases where process information is inaccessible
                pass

    def get_process_icon(self, process):
        """Извлекает иконку процесса."""  # Docstring in Russian
        try:
            exe_path = process.exe()
            icon = QIcon(exe_path)  # Create icon from exe path
            if not icon.isNull():
                return icon
        except (psutil.AccessDenied, psutil.NoSuchProcess) as e:
            print(f"Отказано в доступе или процесс не найден: {e}")  # Error message in Russian
            pass
        except Exception as e:
            print(f"Произошла неожиданная ошибка: {e}")  # Error message in Russian
        return QIcon()  # Return empty icon on failure

    def overlay_selected_window(self):
        self.target_pid = self.process_list_combo.currentData()
        self.hwnd = self.get_window_handle_from_pid(self.target_pid)

        if self.hwnd:
            print(f"Найдено окно HWND: {self.hwnd}")  # Message in Russian
            self.set_window_topmost(self.hwnd)
            self.resize_and_move_window(self.hwnd)

            self.overlay_button.setEnabled(False)
            self.remove_overlay_button.setEnabled(True)

        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось найти дескриптор окна.")  # Error message in Russian

    def remove_overlay(self):
        if self.hwnd:
            try:
                # Restore window to normal state
                win32gui.SetWindowPos(self.hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                      win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                print("Наложение убрано.")  # Message in Russian
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Ошибка при удалении наложения: {e}")  # Error message in Russian
            finally:
                self.overlay_button.setEnabled(True)
                self.remove_overlay_button.setEnabled(False)
                self.hwnd = None  # Reset the handle

    def get_window_handle_from_pid(self, pid):
        def callback(hwnd, hwnds):
            try:
                _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
                if found_pid == pid:
                    hwnds.append(hwnd)
            except win32gui.error:
                # Handle potential errors during GetWindowThreadProcessId (e.g., access denied)
                pass  # Or log the error
            return True

        hwnds = []
        try:
            win32gui.EnumWindows(callback, hwnds)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при перечислении окон: {e}")  # Error message in Russian
            return None

        return hwnds[0] if hwnds else None

    def set_window_topmost(self, hwnd):
        try:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка",
                                 f"Ошибка при установке окна на передний план: {e}")  # Error message in Russian

    def resize_and_move_window(self, hwnd):
        try:
            x = int(self.x_edit.text())
            y = int(self.y_edit.text())
            width = int(self.width_edit.text())
            height = int(self.height_edit.text())

            win32gui.MoveWindow(hwnd, x, y, width, height, True)
        except ValueError:
            QMessageBox.warning(self, "Предупреждение",
                                "Неверный ввод для позиции или размера. Пожалуйста, введите целые числа.")  # Warning in Russian
        except Exception as e:
            QMessageBox.critical(self, "Ошибка",
                                 f"Ошибка при изменении размера/перемещении окна: {e}")  # Error message in Russian


if __name__ == '__main__':
    # Set locale for translation
    QLocale.setDefault(QLocale("ru_RU"))

    app = QApplication(sys.argv)

    # Apply a modern style to the entire application
    app.setStyle("Fusion")  # Use the Fusion style for a modern look
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#F0F0F0"))  # Light gray background
    palette.setColor(QPalette.WindowText, QColor("#333333"))  # Dark text
    palette.setColor(QPalette.Base, QColor("white"))  # White background for text inputs
    palette.setColor(QPalette.AlternateBase, QColor("#E0E0E0"))  # Slightly darker background for alternates
    palette.setColor(QPalette.Text, QColor("#333333"))  # Dark text
    palette.setColor(QPalette.Button, QColor("#4CAF50"))  # Green button background
    palette.setColor(QPalette.ButtonText, QColor("white"))  # White button text
    palette.setColor(QPalette.Highlight, QColor("#388E3C"))  # Darker green highlight color
    palette.setColor(QPalette.HighlightedText, QColor("white"))  # White highlighted text
    app.setPalette(palette)

    window = OverlayApp()
    window.show()
    sys.exit(app.exec_())
