#!/usr/bin/env python3
'''
          Script :: keyring_tools.py
         Version :: v2.0 (04-17-2025)
          Author :: jason.thomaschefsky@cdw.com
         Purpose :: A cross-platform password management tool using system keyring.
     Information :: See 'README.md'

MIT License

Copyright (c) 2025 Jason Thomaschefsky

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''

# --- Import the modules needed for this script ---
import argparse
import csv
import os
from pathlib import Path
import subprocess
import sys
import threading
import time
from typing import Optional, Dict, List, Union, NoReturn
from PySide6.QtGui import QColor, QPalette
import keyring
import pyperclip
from PySide6.QtWidgets import (
    QApplication, QDialog, QFrame, QScrollArea, QStyle, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QRadioButton, QLineEdit, QPushButton, QLabel, QFileDialog
)
from PySide6.QtCore import Qt

# --- Metadata and Constants  ---
APP_NAME = "Keyring Tools"
APP_VERSION = "v2.0"
VERSION_DATE = "(04-17-2025)"
MAX_SYSTEM_LENGTH = 254
MAX_USERNAME_LENGTH = 32
MAX_PASSWORD_LENGTH = 400
# Get script directory (works even when frozen/imported)
try:
    SCRIPT_DIR = Path(__file__).parent
except NameError:
    SCRIPT_DIR = Path(sys.argv[0]).parent
RESULTS_FILE = SCRIPT_DIR / "keyring_results.txt"

# --- Dark mode state ---
DARK_MODE_STATE = True  # Start with dark mode enabled
# DARK_MODE_STATE = False  # Start with dark mode disabled (auto-set to OS)

# --- Debug output state ---
DEBUG = False

# --- Utility Functions ---
def clear_clipboard_after(delay: int = 30) -> None:
    """
    Clear clipboard after a specified delay (threaded).

    Args:
        delay: Number of seconds to wait before clearing clipboard. Default is 30.
    """
    def clear() -> None:
        time.sleep(delay)
        pyperclip.copy("")
    threading.Thread(target=clear, daemon=True).start()

# --- Unified bulk CSV processor for both CLI and GUI ---
def process_bulk_csv(file_path: Union[str, Path], result_display: Optional[QLabel] = None) -> None:
    """
    Process a CSV file containing bulk keyring operations.

    Handles both CLI and GUI modes with appropriate output display.

    Args:
        file_path: Path to the CSV file to process
        result_display: Optional QLabel widget for GUI output display

    Raises:
        ValueError: If CSV format is invalid
        Exception: For other processing errors
    """
    def debug_print(*args) -> None:
        """Print debug messages if DEBUG is enabled."""
        if DEBUG:
            print("DEBUG:", *args)

    output_file = RESULTS_FILE
    results_filename = os.path.basename(output_file)  # Display only the filename in the completion message

    def update_status(msg: str, is_error: bool = False, display: Optional[QLabel] = None) -> None:
        """
        Update status display in either GUI or CLI mode.

        Args:
            msg: Message to display
            is_error: Whether the message is an error
            display: Optional QLabel for GUI display
        """
        if display:  # GUI mode
            display.setText(msg)
            if is_error:
                display.setStyleSheet("font-style: italic; font-weight: bold; color: red;")
            else:
                display.setStyleSheet("font-style: italic; font-weight: bold; color: green;")
            QApplication.processEvents()
        else:  # CLI mode
            print(msg, file=sys.stderr if is_error else sys.stdout)

    try:
        debug_print(f"Keyring backend: {keyring.get_keyring()}")
        with open(file_path, mode="r", encoding="utf-8") as csv_file, \
             open(output_file, mode="w", encoding="utf-8") as out_file:

            reader = csv.DictReader(csv_file)

            if reader.fieldnames != ["function", "system", "username", "password"]:
                error_msg = "CSV must have headers: function,system,username,password"
                raise ValueError(error_msg)

            for row_num, row in enumerate(reader, 1):
                try:
                    # Validate required fields
                    if not row.get("function") or not row.get("system") or not row.get("username"):
                        raise ValueError("Missing required field(s): function, system, or username")

                    function = row["function"].strip().lower()
                    system = row["system"].strip()
                    username = row["username"].strip()
                    password = row.get("password", "").strip()

                    # Validate lengths using constants
                    if len(system) > MAX_SYSTEM_LENGTH:
                        raise ValueError(f"System name must be ≤{MAX_SYSTEM_LENGTH} characters")
                    if len(username) > MAX_USERNAME_LENGTH:
                        raise ValueError(f"Username must be ≤{MAX_USERNAME_LENGTH} characters")
                    if function == "set" and len(password) > MAX_PASSWORD_LENGTH:
                        raise ValueError(f"Password must be ≤{MAX_PASSWORD_LENGTH} characters")

                    if function == "set":
                        debug_print(f"(set) System: {system}, User: {username}, Pass: {password}")
                        keyring.set_password(system, username, password)
                        result = f"Set password for {system}/{username}"
                    elif function == "get":
                        debug_print(f"(get) System: {system}, User: {username}, Pass: {password}")
                        result = keyring.get_password(system, username)
                        result = f"Get password for {system}/{username}: {result}"
                    elif function == "del":
                        debug_print(f"(del) System: {system}, User: {username}, Pass: {password}")
                        keyring.delete_password(system, username)
                        result = f"Deleted password for {system}/{username}"
                    else:
                        raise ValueError(f"Invalid function: {function}")

                    update_status(f"SUCCESS: {result}", False, result_display)
                    out_file.write(f"SUCCESS: {result}\n")

                except Exception as e:
                    error_msg = f"ERROR: {system}/{username} - {str(e)}"
                    out_file.write(f"{error_msg}\n")
                    update_status(error_msg, True, result_display)

            completion_msg = f'Bulk processing complete.\nResults saved to "./{results_filename}"'
            update_status(completion_msg, False, result_display)

            # If in GUI mode, auto-open results file in the default text editor
            if result_display:
                try:
                    if sys.platform == "win32":
                        os.startfile(output_file)
                    elif sys.platform == "darwin":
                        subprocess.run(["open", str(output_file)])
                    else:
                        subprocess.run(["xdg-open", str(output_file)])
                except Exception as e:
                    update_status(f"Failed to auto-open results: {str(e)}", True, result_display)

            # Clear any passwords that might be in clipboard
            pyperclip.copy("")
            clear_clipboard_after(1)  # Force-clear after 1 second

    except Exception as e:
        error_msg = f"Failed to process CSV: {str(e)}"
        update_status(error_msg, True, result_display)
        if not result_display:
            sys.exit(1)

# --- CLI Functions ---
def cli_handler() -> None:
    """
    Handle command-line interface arguments and execution.

    Parses arguments and routes to appropriate keyring functions.
    Exits with status 1 on error.
    """
    parser = argparse.ArgumentParser(prog="keyring_tools.py", description="Keyring Tools")
    parser.add_argument("-v", "--version", action="version", version=f"%(prog)s {APP_VERSION} {VERSION_DATE}")

    # Single operation mode
    subparsers = parser.add_subparsers(dest="command", required=False)

    # Set parser
    set_parser = subparsers.add_parser("set", help="Set password in keyring")
    set_parser.add_argument("system", help="System/service name")
    set_parser.add_argument("username", help="Username for the service")
    set_parser.add_argument("password", help="Password to store")

    # Get parser
    get_parser = subparsers.add_parser("get", help="Get password from keyring")
    get_parser.add_argument("system", help="System/service name")
    get_parser.add_argument("username", help="Username for the service")

    # Del parser
    del_parser = subparsers.add_parser("del", help="Delete password from keyring")
    del_parser.add_argument("system", help="System/service name")
    del_parser.add_argument("username", help="Username for the service")

    # Bulk mode
    parser.add_argument("--bulk", help="Process a CSV file for bulk operations", metavar="FILE")

    args = parser.parse_args()

    if args.bulk:
        process_bulk_csv(args.bulk)
    elif args.command:
        try:
            if args.command == "set":
                keyring.set_password(args.system, args.username, args.password)
                print(f'Password set for "{args.system}/{args.username}".')
            elif args.command == "get":
                result = keyring.get_password(args.system, args.username) or "No password found"
                print(f"Get password for {args.system}/{args.username}: {result}")
            elif args.command == "del":
                keyring.delete_password(args.system, args.username)
                print(f'Credential store "{args.system}/{args.username}" has been deleted.')
        except keyring.errors.PasswordDeleteError:
            print(f'ERROR: Failed to delete credential store.')
            print(f'Credential store "{args.system}" may not exist, or username "{args.username}" may be incorrect.')
        except Exception as e:
            print(f"Error: {str(e)}", file=sys.stderr)
            sys.exit(1)
    else:
        parser.print_help()
        sys.exit(1)

# --- GUI Classes ---
class HelpDialog(QDialog):
    """A custom dialog box displaying help information for the application."""

    def __init__(self, title: str, version: str, version_date: str, parent: Optional[QWidget] = None) -> None:
        """
        Initialize the help dialog.

        Args:
            title: Application title
            version: Version string
            version_date: Version date string
            parent: Optional parent widget
        """
        super().__init__(parent)
        self.setWindowTitle("Help")

        # Inherit dark mode setting from parent if available
        self.dark_mode = parent.dark_mode if parent and hasattr(parent, 'dark_mode') else DARK_MODE_STATE

        # Set Help window size
        self.setFixedSize(410, 510)  # Size (width, height)

        # Create main layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)

        # Add title and version
        title_label = QLabel(f"<h1>{APP_NAME}</h1><h3>{APP_VERSION} {VERSION_DATE}</h3>")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        # Create scroll area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)

        # Create container widget for scroll area
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)

        # Create framed area for text
        text_frame = QFrame()
        text_frame.setObjectName("helpTextFrame")
        frame_layout = QVBoxLayout(text_frame)
        frame_layout.setContentsMargins(15, 15, 15, 15)

        # Add help text
        help_text = f"""A cross-platform password management tool using system keyring.

        GUI Usage:
        1. Select function (Set, Get, Delete)
        2. Fill required fields
        3. Click Go to execute

        Bulk CSV Processing:
        - Use the "Process Bulk CSV" button
        - CSV format: function,system,username,password
        - Results are saved to a text file

        CLI Usage:
        $ python keyring_tools.py [set|get|del] [args]
        $ python keyring_tools.py --bulk file.csv

        Limitations:
        - System Name: ≤{MAX_SYSTEM_LENGTH} characters
        - Username: ≤{MAX_USERNAME_LENGTH} characters
        - Password: ≤{MAX_PASSWORD_LENGTH} characters
        """
        help_label = QLabel(help_text)
        help_label.setWordWrap(True)
        help_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)

        frame_layout.addWidget(help_label)
        container_layout.addWidget(text_frame)
        scroll_area.setWidget(container)
        layout.addWidget(scroll_area)

        # Add close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button, alignment=Qt.AlignmentFlag.AlignCenter)

        # Apply theme
        self.apply_theme()

    def apply_theme(self) -> None:
        """Apply dark or light theme based on current mode."""
        app = QApplication.instance()

        if self.dark_mode:
            # Set dark palette
            dark_palette = QPalette()
            dark_palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
            dark_palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
            dark_palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
            dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
            self.setPalette(dark_palette)

            # Custom styling for the frame and other elements
            self.setStyleSheet("""
                QFrame#helpTextFrame {
                    background-color: #252525;
                    border: 2px solid #555;
                    border-radius: 5px;
                }
                QLabel {
                    color: white;
                    background-color: transparent;
                }
                QScrollArea {
                    background-color: #353535;
                    border: none;
                }
                QPushButton {
                    background-color: #353535;
                    color: white;
                    border: 1px solid #555;
                    padding: 5px 15px;
                    border-radius: 3px;
                    min-width: 80px;
                }
                QPushButton:hover {
                    background-color: #454545;
                }
            """)
        else:
            # Reset to default light theme
            self.setPalette(QApplication.style().standardPalette())
            self.setStyleSheet("")

class KeyringApp(QWidget):
    """Main application window for the Keyring Tools GUI."""

    def __init__(self) -> None:
        """Initialize the KeyringApp GUI with default settings."""
        super().__init__()
        self.dark_mode = DARK_MODE_STATE  # Set the dark mode state
        self.init_ui()
        self.setWindowTitle(f"Keyring Tools")
        self.toggle_theme(self.dark_mode)  # Toggle theme according to DARK_MODE_STATE

    def init_ui(self) -> None:
        """Initialize all UI components and layouts."""
        self.setFixedSize(320, self.height())  # Set a fixed window width of 320 pixels
        layout = QVBoxLayout()

        # Function selection (Set, Get, Delete)
        self.group_func = QGroupBox("Keyring Function")
        self.radio_set = QRadioButton("Set Password")
        self.radio_get = QRadioButton("Get Password")
        self.radio_del = QRadioButton("Delete Password")
        self.radio_get.setChecked(True)  # Get is selected by default

        func_layout = QVBoxLayout()
        func_layout.addWidget(self.radio_set)
        func_layout.addWidget(self.radio_get)
        func_layout.addWidget(self.radio_del)
        self.group_func.setLayout(func_layout)

        # Input fields
        self.system = QLineEdit()
        self.username = QLineEdit()
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)  # Start masked
        self.copy_btn = QPushButton("Copy")

        # Add reveal button
        self.reveal_btn = QPushButton()
        self.reveal_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_FileDialogContentsView))
        self.reveal_btn.setCheckable(True)
        self.reveal_btn.setChecked(False)
        self.reveal_btn.toggled.connect(self.toggle_password_visibility)
        self.reveal_btn.setStyleSheet("""
            QPushButton {
                border: none;
                padding: 2px;
            }
            QPushButton:hover {
                background-color: #f0f0f0;
            }
        """)

        self.copy_btn = QPushButton("Copy")

        # Password field setup
        self.password.setEnabled(False)
        self.copy_btn.setEnabled(False)

        # Input layout
        input_layout = QVBoxLayout()
        input_layout.addWidget(QLabel("System Name:"))
        input_layout.addWidget(self.system)
        input_layout.addWidget(QLabel("Username:"))
        input_layout.addWidget(self.username)
        input_layout.addWidget(QLabel("Password:"))

        # Password layout
        pass_layout = QHBoxLayout()
        pass_layout.addWidget(self.password)
        pass_layout.addWidget(self.reveal_btn)
        pass_layout.addWidget(self.copy_btn)
        input_layout.addLayout(pass_layout)

        # Result display
        self.result_group = QGroupBox("Results")
        self.result_display = QLabel()
        self.result_display.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.result_display.setWordWrap(True)
        self.result_display.setStyleSheet("color: #999; font-style: italic; font-weight: bold; padding: 5px;")

        result_layout = QVBoxLayout()
        result_layout.addWidget(self.result_display)
        self.result_group.setLayout(result_layout)

        # Bulk CSV button
        self.bulk_btn = QPushButton("Process Bulk CSV")
        self.bulk_btn.clicked.connect(self.process_bulk_csv)

        # Buttons
        self.go_btn = QPushButton("Go")
        self.help_btn = QPushButton("Help")
        self.quit_btn = QPushButton("Quit")

        # Button layout
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.help_btn)
        btn_layout.addWidget(self.quit_btn)
        btn_layout.addWidget(self.go_btn)

        # Assemble main layout
        layout.addWidget(self.group_func)
        layout.addLayout(input_layout)
        layout.addWidget(self.result_group)
        layout.addWidget(self.bulk_btn)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

        # Connect signals
        self.radio_set.toggled.connect(lambda: [self.clear_sensitive_fields(), self.update_ui_state()])
        self.radio_get.toggled.connect(lambda: [self.clear_sensitive_fields(), self.update_ui_state()])
        self.radio_del.toggled.connect(lambda: [self.clear_sensitive_fields(), self.update_ui_state()])
        self.system.textChanged.connect(self.validate_inputs)
        self.username.textChanged.connect(self.validate_inputs)
        self.password.textChanged.connect(self.validate_inputs)
        self.copy_btn.clicked.connect(self.copy_to_clipboard)
        self.go_btn.clicked.connect(self.execute_action)
        self.help_btn.clicked.connect(self.show_help)
        self.quit_btn.clicked.connect(self.secure_exit)

        self.update_ui_state()
        self.validate_inputs()

    def toggle_theme(self, dark_mode: bool = True) -> None:
        """
        Toggle between dark and light themes.

        Args:
            dark_mode: Whether to enable dark mode (True) or light mode (False)
        """
        app = QApplication.instance()
        app.setStyle("Fusion")  # Use Fusion style as base for dark theme

        if dark_mode:
            dark_palette = QPalette()
            dark_palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
            dark_palette.setColor(QPalette.ColorRole.Base, QColor(35, 35, 35))
            dark_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(25, 25, 25))
            dark_palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
            dark_palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
            dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
            dark_palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
            dark_palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
            dark_palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
            dark_palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
            app.setPalette(dark_palette)

            # Additional dark mode styling
            app.setStyleSheet("""
                QToolTip {
                    color: #ffffff;
                    background-color: #2a82da;
                    border: 1px solid white;
                }
                QGroupBox {
                    border: 1px solid gray;
                    border-radius: 3px;
                    margin-top: 0.5em;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0 3px;
                }
            """)
            self.dark_mode = True
        else:
            app.setPalette(app.style().standardPalette())
            app.setStyleSheet("")
            self.dark_mode = False

    def update_ui_state(self) -> None:
        """Update UI state based on selected function and clear sensitive data."""
        is_set = self.radio_set.isChecked()
        self.password.setEnabled(is_set)
        self.copy_btn.setEnabled(False)  # Always disable until new password is retrieved
        self.password.clear()  # Extra safeguard

    def validate_inputs(self) -> None:
        """
        Validate input fields and enable/disable the Go button accordingly.

        Also provides feedback about invalid inputs in the results display.
        """
        system = self.system.text().strip()
        username = self.username.text().strip()
        password = self.password.text().strip()

        # Check length constraints using constants
        system_valid = bool(system) and len(system) <= MAX_SYSTEM_LENGTH
        username_valid = bool(username) and len(username) <= MAX_USERNAME_LENGTH
        password_valid = True  # Default to True unless we're setting a password

        if self.radio_set.isChecked():
            password_valid = bool(password) and len(password) <= MAX_PASSWORD_LENGTH

        # Update UI based on validation
        if system_valid and username_valid and password_valid:
            self.go_btn.setEnabled(True)
            self.go_btn.setStyleSheet(
                """
                QPushButton {
                    background-color: green;
                    color: white;
                    border-radius: 5px;
                    padding: 1px;
                    border: 1px solid darkgreen;
                }
                QPushButton:hover {
                    background-color: darkgreen;
                }
                QPushButton:pressed {
                    background-color: limegreen;
                }
                """
            )
        else:
            self.go_btn.setEnabled(False)
            self.go_btn.setStyleSheet("")  # Reset to default style

        # Provide feedback about invalid inputs
        feedback = []
        if system and not system_valid:
            feedback.append(f"System name must be ≤{MAX_SYSTEM_LENGTH} chars")
        if username and not username_valid:
            feedback.append(f"Username must be ≤{MAX_USERNAME_LENGTH} chars")
        if self.radio_set.isChecked() and password and not password_valid:
            feedback.append(f"Password must be ≤{MAX_PASSWORD_LENGTH} chars")

        if feedback:
            self.result_display.setText("\n".join(feedback))
            self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: red;")
        elif not all([system, username]) or (self.radio_set.isChecked() and not password):
            self.result_display.setText("Fill all required fields")
            self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: #999;")
        else:
            self.result_display.setText("Ready")
            self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: green;")

    def toggle_password_visibility(self, checked: bool) -> None:
        """
        Toggle password visibility in the password field.

        Args:
            checked: Whether to show (True) or hide (False) the password
        """
        if not hasattr(self, 'password'):
            return

        # Only allow showing if password exists and isn't being cleared
        if checked and not self.password.text():
            self.reveal_btn.setChecked(False)
            return

        if checked:
            self.password.setEchoMode(QLineEdit.Normal)
            self.reveal_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        else:
            self.password.setEchoMode(QLineEdit.Password)
            self.reveal_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_FileDialogContentsView))

    def copy_to_clipboard(self) -> None:
        """Copy password to clipboard and schedule auto-clear."""
        if not self.password.text():
            self.result_display.setText("No password to copy")
            self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: red;")
            return

        try:
            pyperclip.copy(self.password.text())
            self.result_display.setText(f"Password copied\n(auto clear in 30s)")
            self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: green;")
            clear_clipboard_after(30)

        except Exception as e:
            self.result_display.setText(f"Clipboard error: {str(e)}")
            self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: red;")

    def clear_sensitive_fields(self) -> None:
        """Clear password field, clipboard, and reset UI states."""
        self.password.clear()
        self.password.setEchoMode(QLineEdit.Password)  # Force hide password
        if hasattr(self, 'reveal_btn'):
            self.reveal_btn.setChecked(False)  # Unpress the button
            self.reveal_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_FileDialogContentsView))  # Reset icon
        try:
            pyperclip.copy("")
        except Exception as e:
            print(f"Clipboard error: {e}", file=sys.stderr)
        self.copy_btn.setEnabled(False)
        self.result_display.clear()
        self.result_display.setStyleSheet("")

    def execute_action(self) -> None:
        """Execute the selected keyring action (Set, Get, Delete)."""
        system = self.system.text().strip()
        username = self.username.text().strip()
        password = self.password.text().strip()

        try:
            if self.radio_get.isChecked():
                DEBUG and print(f"(get) System: {system}, User: {username}, Pass: {password}")
                result = keyring.get_password(system, username)
                if result:
                    self.password.setText(result)
                    self.result_display.setText("Password retrieved successfully")
                    self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: green;")
                    self.copy_btn.setEnabled(True)  # Enable the Copy button
                else:
                    self.result_display.setText("No password found")
                    self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: red;")
                    self.copy_btn.setEnabled(False)

            elif self.radio_del.isChecked():
                DEBUG and print(f"(del) System: {system}, User: {username}, Pass: {password}")
                keyring.delete_password(system, username)
                self.result_display.setText("Password deleted successfully")
                self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: green;")
                self.copy_btn.setEnabled(False)

            elif self.radio_set.isChecked():
                DEBUG and print(f"(set) System: {system}, User: {username}, Pass: {password}")
                if not password:  # Explicit check for empty password
                    self.result_display.setText("Error: Password cannot be empty")
                    self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: red;")
                    return
                keyring.set_password(system, username, password)
                self.result_display.setText("Password set successfully")
                self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: green;")
                self.copy_btn.setEnabled(False)

        except Exception as e:
            self.result_display.setText(f"Error: {str(e)}")
            self.result_display.setStyleSheet("font-style: italic; font-weight: bold; color: red;")

    def secure_exit(self) -> None:
        """Clean up sensitive data and exit the application securely."""
        if hasattr(self, 'clipboard_timer'):
            self.clipboard_timer.cancel()  # Cancel any pending clipboard clears

        self.clear_sensitive_fields()  # Securely clear all sensitive data before exiting
        self.close()  # Close the window
        QApplication.quit()  # Ensure full application exit

    def update_results(self, msg: str, is_error: bool) -> None:
        """
        Update the results display with a message.

        Args:
            msg: Message to display
            is_error: Whether the message is an error (affects styling)
        """
        self.result_display.setText(msg)
        self.result_display.setStyleSheet(
            "color: red;" if is_error else "color: green;")

    def process_bulk_csv(self) -> None:
        """Open file dialog and process selected CSV file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select CSV File", "", "CSV Files (*.csv)")
        if file_path:
            process_bulk_csv(file_path, result_display=self.result_display)

    def show_help(self) -> None:
        """Create and show the custom help dialog."""
        help_dialog = HelpDialog(APP_NAME, APP_VERSION, VERSION_DATE, self)
        help_dialog.exec()

# --- Main Execution ---
if __name__ == "__main__":
    if len(sys.argv) > 1:
        cli_handler()
    else:
        app = QApplication(sys.argv)
        window = KeyringApp()
        window.show()
        sys.exit(app.exec())
