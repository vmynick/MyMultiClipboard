#!/usr/bin/env python
import json
import os
import sys
import ctypes  # Add this import for console window hiding
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import Qt, QRect
from PyQt5.QtGui import QPainter, QColor, QBrush, QFont, QIcon, QPixmap
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QSpacerItem, QSizePolicy, QSystemTrayIcon, QMenu, QAction
import webbrowser
import pyperclip
import keyboard
import threading
import csv
from pystray import Icon, Menu, MenuItem
from PIL import Image, ImageDraw
import winsound  # Add this import for a more noticeable beep sound
import pynput  # Add this import for global hotkey
import base64
from icon_base64 import encoded_icon  # Import the base64 string

# Constants
DEFAULT_CONFIG = {
    "hotkey": "ctrl+alt+p",
    "window_width": 850,
    "window_height": 600,    
    "colors": ["#D3D3D3", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF", "#D1BAFF", "#FFB3E6", "#FFB3FF", "#E6B3FF"]  # Change first color to default gray
}
BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))  # Use sys.argv[0] for portability
DATA_FILE = os.path.join(BASE_DIR, "data.json")


class FocusThread(QtCore.QThread):
    def __init__(self, window):
        super().__init__()
        self.window = window

    def run(self):
        QtCore.QTimer.singleShot(100, self.window.set_focus_on_listbox)

class PopupApp(QtWidgets.QWidget):
    def __init__(self, config):
        super().__init__()
        self.config = config
        self.data = []
        self.filtered_data = []
        self.selected_index = -1
        self.load_data()
        self.init_ui()
        self.tray_icon = None
        # self.setWindowTitle("MyMultiClipboard")
        # self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)  # Keep the window on top        

    def quit(self):
        #Quit the program."""
        QtWidgets.QApplication.quit()

    def load_data(self):
        try:
            if not os.path.exists(DATA_FILE):
                # Create file with default data if it doesn't exist
                with open(DATA_FILE, "w") as f:
                    json.dump({
                        "hotkey": DEFAULT_CONFIG["hotkey"],
                        "window_width": DEFAULT_CONFIG["window_width"],
                        "window_height": DEFAULT_CONFIG["window_height"],
                        "data": [{"name": "Example", "data": "http://example.com", "color": "#FFB3BA"}]
                    }, f)
            with open(DATA_FILE, "r") as f:
                file_data = json.load(f)
                self.config["hotkey"] = file_data.get("hotkey", DEFAULT_CONFIG["hotkey"])
                self.config["window_width"] = file_data.get("window_width", DEFAULT_CONFIG["window_width"])
                self.config["window_height"] = file_data.get("window_height", DEFAULT_CONFIG["window_height"])
                self.data = file_data.get("data", [])
                if not isinstance(self.data, list):
                    raise ValueError("Data must be a list.")
                for item in self.data:
                    if "color" not in item:
                        item["color"] = DEFAULT_CONFIG["colors"][0]  # Default color if not present
        except (json.JSONDecodeError, ValueError):
            # Handle invalid data file format and reset to default
            QtWidgets.QMessageBox.critical(self, "Error", f"Invalid data in {DATA_FILE}. Resetting.")
            self.config["hotkey"] = DEFAULT_CONFIG["hotkey"]
            self.config["window_width"] = DEFAULT_CONFIG["window_width"]
            self.config["window_height"] = DEFAULT_CONFIG["window_height"]
            self.data = [{"name": "Example", "data": "http://example.com"}]
            with open(DATA_FILE, "w") as f:
                json.dump({
                    "hotkey": self.config["hotkey"],
                    "window_width": self.config["window_width"],
                    "window_height": self.config["window_height"],
                    "data": self.data
                }, f)
        self.filtered_data = self.data[:]

    def init_ui(self):
        self.setWindowTitle("MyMultiClipboard")
        self.update()
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setFixedSize(self.config["window_width"], self.config["window_height"])

        layout = QtWidgets.QVBoxLayout()

        # Layout for buttons
        MXbutton_layout = QtWidgets.QHBoxLayout()
        MXbutton_layout.setSpacing(5)
        MXbutton_layout.setContentsMargins(0, 0, 5, 0)

        # Add a spacer to push the buttons to the top-right
        MXbutton_layout.addSpacerItem(QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

        # Hotkey Button
        hotkey_button = QtWidgets.QPushButton("H")
        hotkey_button.setFixedSize(20, 20)
        hotkey_button.setStyleSheet("background-color: #cccccc; color: white; border: none; border-radius: 3px; padding: 3px; font-size: 12px; font-weight: bold;")
        hotkey_button.setToolTip("Adjust Hotkey")
        hotkey_button.clicked.connect(self.adjust_hotkey)

        # Minimize Button
        minimize_button = QtWidgets.QPushButton("_")
        minimize_button.setFixedSize(20,20)
        minimize_button.setStyleSheet("background-color: #99ccff; color: white; border: none; border-radius: 3px; padding: 3px; font-size: 12px; font-weight: bold;")
        minimize_button.setToolTip("Minimize")
        minimize_button.clicked.connect(self.hide_window)

        # Close Button
        close_button = QtWidgets.QPushButton("X")
        close_button.setFixedSize(20,20)
        close_button.setStyleSheet("background-color: #ff9999; color: white; border: none; border-radius: 3px; padding: 3px; font-size: 12px; font-weight: bold;")
        close_button.setToolTip("Close")
        close_button.clicked.connect(self.close)

        # Add buttons to the layout
        MXbutton_layout.addWidget(hotkey_button, alignment=QtCore.Qt.AlignRight)
        MXbutton_layout.addWidget(minimize_button, alignment=QtCore.Qt.AlignRight)
        MXbutton_layout.addWidget(close_button, alignment=QtCore.Qt.AlignRight)

        layout.addLayout(MXbutton_layout)

        # Listbox for displaying data
        self.listbox = QtWidgets.QListWidget(self)
        self.listbox.setStyleSheet("background-color: #7d7d7d; color: white; border-radius: 5px; font-size: 14px;")  # Lighten the base background color
        self.refresh_listbox()
        self.listbox.itemDoubleClicked.connect(self.handle_enter)
        self.listbox.itemSelectionChanged.connect(self.update_selected_index)  # Detect selection changes
        layout.addWidget(self.listbox)

        # Buttons for Add, Edit, Delete, Export, and Import
        button_layout = QtWidgets.QHBoxLayout()
        add_button = QtWidgets.QPushButton("Add", self)
        add_button.clicked.connect(self.add_line)
        add_button.setStyleSheet("background-color: #b3e6b3; color: white; border-radius: 3px; height:20px; font-size: 12px; font-weight: bold;")
        button_layout.addWidget(add_button)

        edit_button = QtWidgets.QPushButton("Edit", self)
        edit_button.clicked.connect(self.edit_line)
        edit_button.setStyleSheet("background-color: #f7c6a3; color: white;border-radius: 3px; height:20px; font-size: 12px; font-weight: bold;")
        button_layout.addWidget(edit_button)

        delete_button = QtWidgets.QPushButton("Delete", self)
        delete_button.clicked.connect(self.delete_line)
        delete_button.setStyleSheet("background-color: #ffb3b3; color: white;border-radius: 3px; height:20px; font-size: 12px; font-weight: bold;")
        button_layout.addWidget(delete_button)

        export_button = QtWidgets.QPushButton("Export", self)
        export_button.clicked.connect(self.export_data)
        export_button.setStyleSheet("background-color: #99ccff; color: white;border-radius: 3px; height:20px; font-size: 12px; font-weight: bold;")
        button_layout.addWidget(export_button)

        import_button = QtWidgets.QPushButton("Import", self)
        import_button.clicked.connect(self.import_data)
        import_button.setStyleSheet("background-color: #d1b3ff; color: white;border-radius: 3px; height:20px; font-size: 12px; font-weight: bold;")
        button_layout.addWidget(import_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)

        # Center the window on the screen
        self.center_window()

        # Bind Ctrl+Return to edit_line
        self.shortcut_edit = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Return"), self)
        self.shortcut_edit.activated.connect(self.edit_line)

        # Bind Enter to handle_enter
        self.shortcut_enter = QtWidgets.QShortcut(QtGui.QKeySequence("Return"), self)
        self.shortcut_enter.activated.connect(self.handle_enter)

        # Bind Ctrl+Delete to delete_line
        self.shortcut_delete = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Delete"), self)
        self.shortcut_delete.activated.connect(self.delete_line)

        # Bind Ctrl+Insert to add_line
        self.shortcut_insert = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Insert"), self)
        self.shortcut_insert.activated.connect(self.add_line)

        # Bind ESC to hide the window
        self.shortcut_esc = QtWidgets.QShortcut(QtGui.QKeySequence("Esc"), self)
        self.shortcut_esc.activated.connect(self.hide_window)

        # Bind Shift+Return to open_url
        self.shortcut_open_url = QtWidgets.QShortcut(QtGui.QKeySequence("Shift+Return"), self)
        self.shortcut_open_url.activated.connect(self.open_url)

        # Bind Ctrl+Up to move item up
        self.shortcut_move_up = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Up"), self)
        self.shortcut_move_up.activated.connect(self.move_item_up)

        # Bind Ctrl+Down to move item down
        self.shortcut_move_down = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Down"), self)
        self.shortcut_move_down.activated.connect(self.move_item_down)

        # Bind Ctrl+H to open hotkey adjust window
        self.shortcut_hotkey_adjust = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+H"), self)
        self.shortcut_hotkey_adjust.activated.connect(self.adjust_hotkey)

        # Bind Ctrl+M to send application to systray
        self.shortcut_systray = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+M"), self)
        self.shortcut_systray.activated.connect(self.hide_window)

        # Bind Ctrl+X to close the application
        self.shortcut_close = QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+X"), self)
        self.shortcut_close.activated.connect(self.quit)

    def paintEvent(self, event):
        # Paint the window with rounded corners and a 2px border
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Set the color for the border
        painter.setPen(QColor(119, 221, 119))  # Black border
        painter.setBrush(QBrush(QColor(102, 153, 153)))  # White background inside the border
        
        # Draw a rounded rectangle with a 2px border
        radius = 16  # Radius for rounded corners
        painter.drawRoundedRect(2, 2, self.width() - 4, self.height() - 4, radius, radius)

        # Draw custom title text inside the window
        painter.setPen(QColor(0, 128, 96)) 
        font = QFont("Arial", 12)
        font.setBold(True)
        painter.setFont(font)
        painter.drawText(QRect(10, 10, self.width() - 20, 40), Qt.AlignLeft, "MyMultiClipboard")

    def open_add_edit_popup(self, title, name_label, data_label, current_name=None, current_data=None, current_color=None, index=None):
        popup = QtWidgets.QDialog(self)
        popup.setWindowTitle(title)
        popup.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Dialog)
        popup.setStyleSheet("background-color: #009999; border-radius: 5px;")
        popup.setMinimumWidth(600)  # Ensures minimum width of 600px
        popup.setMinimumHeight(250)  # Increase height to ensure buttons are visible
        popup.setModal(True)

        layout = QtWidgets.QVBoxLayout()

        name_label_widget = QtWidgets.QLabel(name_label, popup)
        name_label_widget.setStyleSheet("color: white; height:30px; font-size: 14px;")
        layout.addWidget(name_label_widget)

        name_entry = QtWidgets.QLineEdit(popup)
        name_entry.setText(current_name or "")
        name_entry.setStyleSheet(f"background-color: {current_color or '#2d2d2d'}; color: #00008B; border: solid 1px #ccc; border-radius: 3px; height:30px; font-size: 14px;")
        layout.addWidget(name_entry)

        data_label_widget = QtWidgets.QLabel(data_label, popup)
        data_label_widget.setStyleSheet("color: white; height:30px; font-size: 14px;")
        layout.addWidget(data_label_widget)

        data_entry = QtWidgets.QLineEdit(popup)
        data_entry.setText(current_data or "")
        data_entry.setStyleSheet(f"background-color: {current_color or '#2d2d2d'}; color: #00008B; border: solid 1px #ccc; border-radius: 3px; height:30px; font-size: 14px;")
        layout.addWidget(data_entry)

        color_label_widget = QtWidgets.QLabel("Choose color:", popup)
        color_label_widget.setStyleSheet("color: white; height:30px; font-size: 14px;")
        layout.addWidget(color_label_widget)

        color_layout = QtWidgets.QHBoxLayout()
        color_buttons = []
        for color in DEFAULT_CONFIG["colors"]:
            color_button = QtWidgets.QPushButton("", popup)
            color_button.setStyleSheet(f"background-color: {color}; border-radius: 3px; height:30px; width:30px;")
            color_button.setCheckable(True)
            color_buttons.append(color_button)
            color_layout.addWidget(color_button)
            if color == current_color:
                color_button.setChecked(True)
                color_button.setStyleSheet(f"background-color: {color}; border: 2px solid black; border-radius: 3px; height:30px; width:30px;")
            else:
                color_button.setStyleSheet(f"background-color: {color}; border-radius: 3px; height:30px; width:30px;")
            color_button.clicked.connect(lambda ch, btn=color_button: self.highlight_selected_color(btn, color_buttons, name_entry, data_entry))
        layout.addLayout(color_layout)

        button_layout = QtWidgets.QHBoxLayout()

        ok_button = QtWidgets.QPushButton("OK", popup)
        ok_button.setStyleSheet("background-color: #0099cc; color: white; height:30px; border: solid 1px #6600cc;")
        ok_button.clicked.connect(lambda: self.submit_popup(name_entry, data_entry, color_buttons, popup, index))
        button_layout.addWidget(ok_button)

        cancel_button = QtWidgets.QPushButton("Cancel", popup)
        cancel_button.setStyleSheet("background-color: #ff9999; color: white; height:30px; border: solid 1px #6600cc;")
        cancel_button.clicked.connect(popup.close)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        popup.setLayout(layout)
        popup.exec_()

    def highlight_selected_color(self, selected_button, color_buttons, name_entry, data_entry):
        for button in color_buttons:
            if button == selected_button:
                button.setChecked(True)
                button.setStyleSheet(f"background-color: {button.palette().button().color().name()}; border: 2px solid black; border-radius: 3px; height:30px; width:30px;")
                name_entry.setStyleSheet(f"background-color: {button.palette().button().color().name()}; color: #00008B; border: solid 1px #ccc; border-radius: 3px; height:30px; font-size: 14px;")
                data_entry.setStyleSheet(f"background-color: {button.palette().button().color().name()}; color: #00008B; border: solid 1px #ccc; border-radius: 3px; height:30px; font-size: 14px;")
            else:
                button.setChecked(False)
                button.setStyleSheet(f"background-color: {button.palette().button().color().name()}; border-radius: 3px; height:30px; width:30px;")

    def center_window(self):
        screen = QtWidgets.QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width() - size.width()) // 2, (screen.height() - size.height()) // 2)

    def update_selected_index(self):
        # Update the last selected index when the item selection changes.
        selected_items = self.listbox.selectedItems()
        if selected_items:
            self.selected_index = self.listbox.row(selected_items[0])
        else:
            self.selected_index = -1
        self.update_selected_item_border()

    def handle_enter(self):
        item = self.listbox.currentItem()
        if item:
            index = self.listbox.row(item)
            content = self.data[index]["data"]  # Ensure the correct data is copied
            threading.Thread(target=pyperclip.copy, args=(content,)).start()  # Copy to clipboard in a separate thread
            threading.Thread(target=winsound.Beep, args=(1000, 500)).start()  # Make a more noticeable beep sound in a separate thread
            self.hide()
            self.send_to_systray()
            self.update_tray_menu()

    def add_line(self):
        self.open_add_edit_popup("Add Line", "Enter name:", "Enter data:", None, None, DEFAULT_CONFIG["colors"][0])

    def edit_line(self):
        item = self.listbox.currentItem()
        if item:
            index = self.listbox.row(item)
            current_item = self.filtered_data[index]
            self.open_add_edit_popup("Edit Line", "Edit name:", "Edit data:", current_item["name"], current_item["data"], current_item["color"], index)

    def delete_line(self):
        item = self.listbox.currentItem()
        if item:
            reply = QtWidgets.QMessageBox.question(self, "Delete Confirmation", 
                                                   "Are you sure you want to delete this item?", 
                                                   QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if reply == QtWidgets.QMessageBox.Yes:
                index = self.listbox.row(item)
                del self.data[index]
                self.filtered_data = self.data[:]
                self.refresh_listbox()
                self.save_data()

    def submit_popup(self, name_entry, data_entry, color_buttons, popup, index):
        new_name = name_entry.text().strip()
        new_data = data_entry.text().strip()
        new_color = next((button.palette().button().color().name() for button in color_buttons if button.isChecked()), DEFAULT_CONFIG["colors"][0])
        if not new_name or not new_data:
            QtWidgets.QMessageBox.warning(self, "Input Error", "Name and data cannot be empty.")
            return
        new_item = {"name": new_name, "data": new_data, "color": new_color}
        if index is None:
            self.data.append(new_item)
        else:
            self.data[index] = new_item
        self.filtered_data = self.data[:]
        self.refresh_listbox()
        self.save_data()
        popup.close()

    def import_data(self):
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Import Data", "MMC_item_list", "JSON files (*.json);;All files (*.*)")
        if file_path:
            try:
                with open(file_path, "r") as f:
                    imported_data = json.load(f)
                    if not isinstance(imported_data, dict) or not isinstance(imported_data.get("data", []), list):
                        raise ValueError("Imported JSON must be a dictionary with a 'data' list.")

                action = QtWidgets.QMessageBox.question(self, "Import Data", "Do you want to add new lines to the existing data? [Yes]\n\nClick To delete existing data and add new lines. [No]", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)				
                if action == QtWidgets.QMessageBox.Yes:
                    self.data.extend(imported_data["data"])
                else:
                    self.data = imported_data["data"]
                self.filtered_data = self.data[:]
                self.refresh_listbox()
                self.save_data()
            except (json.JSONDecodeError, ValueError) as e:
                QtWidgets.QMessageBox.critical(self, "Import Error", f"Error importing data: {str(e)}")

    def export_data(self):
        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export Data", "MMC_item_list", "JSON files (*.json);;All files (*.*)")
        if file_path:
            try:
                with open(file_path, "w") as f:
                    json.dump({
                        "hotkey": self.config["hotkey"],
                        "window_width": self.config["window_width"],
                        "window_height": self.config["window_height"],
                        "data": self.data
                    }, f, indent=4)
                QtWidgets.QMessageBox.information(self, "Export Successful", "Data exported successfully.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Export Error", f"Error exporting data: {str(e)}")

    def refresh_listbox(self):
        self.listbox.clear()
        for item in self.filtered_data:
            list_item = QtWidgets.QListWidgetItem("  " + item["name"])  # Add 2px padding to the left
            list_item.setBackground(QtGui.QColor(item["color"]))
            list_item.setForeground(QtGui.QColor("#00008B"))  # Set font color to dark blue
            self.listbox.addItem(list_item)
        self.update_selected_item_border()

    def update_selected_item_border(self):
        for index in range(self.listbox.count()):
            item = self.listbox.item(index)
            widget = self.listbox.itemWidget(item)
            if widget is None:
                widget = QtWidgets.QLabel(item.text())
                widget.setStyleSheet(f"background-color: {item.background().color().name()}; color: #00008B; padding-left: 2px;")
                self.listbox.setItemWidget(item, widget)
            if index == self.selected_index:
                widget.setStyleSheet(f"background-color: {item.background().color().name()}; color: #00008B; border: 1px solid red; padding-left: 2px; font-weight: bold;")
            else:
                widget.setStyleSheet(f"background-color: {item.background().color().name()}; color: #00008B; border: none; padding-left: 2px; font-weight: normal;")

    def save_data(self):
        with open(DATA_FILE, "w") as f:
            json.dump({
                "hotkey": self.config["hotkey"],
                "window_width": self.config["window_width"],
                "window_height": self.config["window_height"],
                "data": self.data
            }, f, indent=4)

    def move_item_up(self):
        current_row = self.listbox.currentRow()
        if current_row > 0:
            self.data.insert(current_row - 1, self.data.pop(current_row))
            self.filtered_data = self.data[:]
            self.refresh_listbox()
            self.listbox.setCurrentRow(current_row - 1)
            self.save_data()

    def move_item_down(self):
        current_row = self.listbox.currentRow()
        if current_row < self.listbox.count() - 1:
            self.data.insert(current_row + 1, self.data.pop(current_row))
            self.filtered_data = self.data[:]
            self.refresh_listbox()
            self.listbox.setCurrentRow(current_row + 1)
            self.save_data()
																			  
    def show(self):
        # Override the show method to focus on the last selected item.
        super().show()
        self.setFocus()  # Set focus on the listbox
        if self.selected_index == -1 and self.listbox.count() > 0:  # Ensure there is a first selected index
            self.selected_index = 0
        if self.selected_index != -1:  # Ensure there is a last selected index
            item = self.listbox.item(self.selected_index)
            self.listbox.setCurrentItem(item)  # Set the current item to the last selected one
        self.listbox.setFocus()  # Set focus on the listbox

    def show2(self):
        self.show()
        self.activateWindow()
        self.listbox.setFocus()
        self.update_tray_menu()

    def show_and_focus(self):
        self.show()
        # self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)  # Ensure the window stays on top
        self.show()
        self.focus_thread = FocusThread(self)
        self.focus_thread.start()
        keyboard.release('ctrl')  # Release the Ctrl key to prevent it from getting stuck
        self.activateWindow()

    def set_focus_on_listbox(self):
        self.listbox.setFocus()
        if self.selected_index != -1:  # Ensure there is a last selected index
            item = self.listbox.item(self.selected_index)
            self.listbox.setCurrentItem(item)  # Set the current item to the last selected one
        self.listbox.setFocus()  # Set focus on the listbox
        keyboard.release('ctrl')  # Release the Ctrl key to prevent it from getting stuck

    def send_to_systray(self):
        if hasattr(self, 'tray_icon') and self.tray_icon:
            return  # Icon is already running, do nothing

        # Decode the base64-encoded string
        icon_data = base64.b64decode(encoded_icon)

        # Convert the decoded data into a QPixmap
        pixmap = QPixmap()
        pixmap.loadFromData(icon_data)

        # Create the QIcon from the QPixmap
        icon = QIcon(pixmap)
        self.tray_icon = QSystemTrayIcon(QIcon(icon), self)

        self.tray_icon.setToolTip("MyMultiClipboard")
        self.tray_menu = QMenu()

        # Add "Open" and "Quit" actions
        self.open_action = QAction("Show", self)
        self.open_action.triggered.connect(self.show2)
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.close)

        self.tray_menu.addAction(self.open_action)
        self.tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.show()

        # Connect double-click to show the window
        self.tray_icon.activated.connect(self.on_systray_activated)

    def on_systray_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show2()

    def update_tray_menu(self):
        if self.isHidden():
            self.open_action.setText("Show")
            self.open_action.triggered.disconnect()
            self.open_action.triggered.connect(self.show2)
        else:
            self.open_action.setText("Hide")
            self.open_action.triggered.disconnect()
            self.open_action.triggered.connect(self.hide_window)

    def hide_window(self):
        self.hide()
        self.send_to_systray()
        self.update_tray_menu()

    def close(self, event):
        #Handle window close event to remove the system tray icon."""
        if self.tray_icon:
            self.tray_icon.hide()  # Hide the system tray icon before closing            
            # self.tray_icon.stop()  # Remove the system tray icon
        # event.accept()  # Accept the close event to close the window
        self.quit()
        # exit()

    def minimizeEvent(self, event):
        #Override the minimize event to hide the window instead of minimizing it."""
        event.ignore()
        self.hide()  # Hide the window when minimized
        self.tray_icon.showMessage("Application Running", "The app is still running in the system tray.")

    def exit_app(self):
        self.quit()

    def closeEvent(self, event):
        #Handle window close event to quit the application."""
        if self.tray_icon:
            self.tray_icon.hide()  # Hide the system tray icon before closing
        event.accept()  # Accept the close event, allowing the window to close
							
    def show_window(self):
        #Show the main window and focus on the list."""
        self.show()
        self.raise_()
        self.activateWindow()
        self.listbox.setFocus()  # Set focus on the listbox

    def show_window_ontop(self):
        #Show the main window and ensure it stays on top."""
        # self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)  # Set always on top
        self.show()
        self.raise_()
        self.activateWindow()
        self.listbox.setFocus()

    def open_url(self):
        #Open the selected item in the default browser if it is a valid URL."""
        item = self.listbox.currentItem()
        if item:
            index = self.listbox.row(item)
            content = self.filtered_data[index]["data"]
            if content.startswith("http://") or content.startswith("https://"):
                self.hide()
                self.send_to_systray()
                self.update_tray_menu()                
                webbrowser.open(content)
            else:
                QtWidgets.QMessageBox.warning(self, "Invalid URL", "The selected item is not a valid URL.")

    def adjust_hotkey(self):
        self.hotkey_dialog = QtWidgets.QDialog(self)
        self.hotkey_dialog.setWindowTitle("Adjust Hotkey")
        self.hotkey_dialog.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Dialog)
        self.hotkey_dialog.setStyleSheet("background-color: #009999; border-radius: 5px;")
        self.hotkey_dialog.setMinimumWidth(300)
        self.hotkey_dialog.setMinimumHeight(150)  # Increase height to ensure buttons are visible
        self.hotkey_dialog.setModal(True)

        layout = QtWidgets.QVBoxLayout()

        label = QtWidgets.QLabel("Press the new hotkey (Ctrl+Alt+Key):", self.hotkey_dialog)
        label.setStyleSheet("color: white; height:30px; font-size: 14px;")
        layout.addWidget(label)

        self.hotkey_input = QtWidgets.QLineEdit(self.hotkey_dialog)
        self.hotkey_input.setStyleSheet("background-color: #2d2d2d; color: white; border: solid 1px #ccc; border-radius: 3px; height:30px; font-size: 14px;")
        layout.addWidget(self.hotkey_input)

        button_layout = QtWidgets.QHBoxLayout()

        ok_button = QtWidgets.QPushButton("OK", self.hotkey_dialog)
        ok_button.setStyleSheet("background-color: #0099cc; color: white; height:30px; border: solid 1px #6600cc;")
        ok_button.clicked.connect(self.save_new_hotkey)
        button_layout.addWidget(ok_button)

        cancel_button = QtWidgets.QPushButton("Cancel", self.hotkey_dialog)
        cancel_button.setStyleSheet("background-color: #ff9999; color: white; height:30px; border: solid 1px #6600cc;")
        cancel_button.clicked.connect(self.hotkey_dialog.close)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

        self.hotkey_dialog.setLayout(layout)
        self.hotkey_dialog.show()

        self.hotkey_input.setFocus()
        self.hotkey_input.keyPressEvent = self.capture_hotkey

    def capture_hotkey(self, event):
        key = event.key()
        modifiers = event.modifiers()
        if key in [Qt.Key_Return, Qt.Key_Enter]:
            return
        if key == Qt.Key_Escape:
            self.hotkey_dialog.close()
            return
        if not (modifiers & Qt.ControlModifier and modifiers & Qt.AltModifier):
            return
        key_sequence = QtGui.QKeySequence(modifiers | key).toString()
        self.hotkey_input.setText(key_sequence)

    def save_new_hotkey(self):
        new_hotkey = self.hotkey_input.text().strip()
        if new_hotkey and "Ctrl+Alt+" in new_hotkey:
            self.config["hotkey"] = new_hotkey
            self.save_data()
            QtWidgets.QMessageBox.information(self, "Hotkey Changed", f"Hotkey changed to: {new_hotkey}")
            self.update_hotkey_listener()
        else:
            QtWidgets.QMessageBox.warning(self, "Invalid Hotkey", "Please enter a valid hotkey with Ctrl and Alt modifiers.")
        self.hotkey_dialog.close()

    def update_hotkey_listener(self):
        keyboard.clear_all_hotkeys()
        hotkey = self.config.get("hotkey", "ctrl+alt+p")
        keyboard.add_hotkey(hotkey, self.show_and_focus)

if __name__ == "__main__":
    # Hide the console window
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    
    # Load config
    config = DEFAULT_CONFIG.copy()
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            file_data = json.load(f)
            config["hotkey"]        = file_data.get("hotkey",        DEFAULT_CONFIG["hotkey"])
            config["window_width"]  = file_data.get("window_width",  DEFAULT_CONFIG["window_width"])
            config["window_height"] = file_data.get("window_height", DEFAULT_CONFIG["window_height"])

    # Create application
    app = QtWidgets.QApplication(sys.argv)
						  
    window = PopupApp(config)
    window.show()						 
    window.send_to_systray()
    window.hide()
    
    # Add hotkey listener in background thread
    def listen_hotkeys():
        hotkey = config.get("hotkey", "ctrl+alt+p")
        keyboard.add_hotkey(hotkey, window.show_and_focus)
							   
		 
    hotkey_thread = threading.Thread(target=listen_hotkeys, daemon=True)
    hotkey_thread.start()

    sys.exit(app.exec_())


