#!/usr/bin/env python3
'''
          Script :: nddu.py
         Version :: v1.1.0 (01-24-2026)
          Author :: jason.thomaschefsky@cdw.com
         Purpose :: Document network devices using "show" commands, processed with concurrent threads.
     Information :: See 'README.md'

MIT License

Copyright (c) 2026 Jason Thomaschefsky

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
import concurrent.futures
import getpass
import ipaddress
import keyring
import logging
import os
import platform
import subprocess
import sys
import json
import urllib.request
import urllib.error
from packaging import version
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List, Set, Tuple, Union, Any, NoReturn
from netmiko import ConnectHandler, NetmikoTimeoutException, NetmikoAuthenticationException
from PySide6.QtCore import QObject, QTimer, Qt, QThread, QRect, QSize, Signal
from PySide6.QtGui import QPalette, QPixmap, QPainter, QTextFormat, QColor, QPixmap, QTextCursor
from PySide6.QtWidgets import (
    QApplication, QFrame, QWidget, QLabel, QPushButton, QLineEdit, QCheckBox, QVBoxLayout, QHBoxLayout,
    QFileDialog, QGroupBox, QMessageBox, QRadioButton, QProgressBar, QScrollArea, QDialog, QTextEdit,
    QPlainTextEdit
)
from urllib.error import URLError, HTTPError
from typing import Optional, Dict

# --- Silence Paramiko and Netmiko logs ---
logging.getLogger("paramiko").setLevel(logging.WARNING)  # Suppresses SSH connection details
logging.getLogger("netmiko").setLevel(logging.WARNING)   # Suppresses Netmiko output

# --- Application Metadata ---
APP_NAME = "Network Device Documentation Utility"
APP_VERSION = "v1.1.0"
VERSION_DATE = "(01-24-2026)"
GITHUB_API_LATEST_RELEASE = "https://api.github.com/repos/RacerJay/nddu/releases/latest"
REPO_URL = "https://github.com/RacerJay/nddu"

# --- Dark mode state ---
DARK_MODE_STATE = True  # Start with dark mode enabled
# DARK_MODE_STATE = False  # Start with dark mode disabled

# --- File Paths (all as Path objects) ---
SCRIPT_DIR = Path(__file__).parent
DEFAULT_INPUT_FOLDER = SCRIPT_DIR / "input"
DEFAULT_OUTPUT_FOLDER = SCRIPT_DIR / "output"
DEFAULT_DEVICE_FILE = DEFAULT_INPUT_FOLDER / "Devices.txt"
DEFAULT_COMMAND_FILE = DEFAULT_INPUT_FOLDER / "Commands.txt"
LOGO_PATH = SCRIPT_DIR / "images" / "nddu.png"
COMBINED_OUTPUT_FILENAME = "Combined.txt"
KEYRING_TOOLS_SCRIPT = "keyring_tools.py"

# --- Logging & Formatting ---
MAX_COMMAND_LENGTH = 256
DIVIDER = '=' * 80
FILLER1 = '!' * 20
FILLER2 = '#' * 20

# --- Command Validation ---
ALLOWED_COMMAND_PREFIXES = {"dir", "mor", "sho", "who"}

# --- Concurrency ---
# Calculate the maximum number of threads to execute concurrently
max_workers = min(32, os.cpu_count() + 8)  # Changed from default of + 4

# --- Add supports for VERBOSE logging level 15 ---
VERBOSE_LEVEL_NUM = 15  # Between INFO(20) and DEBUG(10)
logging.addLevelName(VERBOSE_LEVEL_NUM, "VERBOSE")

def verbose(self, message: str, *args: Any, **kwargs: Any) -> None:
    """Log a message with severity 'VERBOSE' (level 15)."""
    if self.isEnabledFor(VERBOSE_LEVEL_NUM):
        self._log(VERBOSE_LEVEL_NUM, message, args, **kwargs)

logging.Logger.verbose = verbose

def format_time(dt: Optional[datetime] = None) -> str:
    """
    Format a datetime object into a standardized string format.
    
    Args:
        dt: Optional datetime object to format. If None, uses current time.
        
    Returns:
        Formatted datetime string in 'Day MM/DD/YYYY - HH:MM:SS AM/PM' format
    """
    if dt is None:
        dt = datetime.now()
    return dt.strftime('%a %m/%d/%Y - %I:%M:%S %p')

class AllowedCommands:
    """Class to validate and manage allowed command prefixes."""
    
    def __init__(self) -> None:
        """Initialize with default allowed command prefixes."""
        self.allowed_prefixes: Set[str] = ALLOWED_COMMAND_PREFIXES

    def is_command_allowed(self, command: str) -> bool:
        """
        Check if a command is allowed based on its prefix and length.
        
        Args:
            command: The command string to validate
            
        Returns:
            True if command is allowed, False otherwise
        """
        command = command.strip()
        
        # Check length first
        if len(command) > MAX_COMMAND_LENGTH:
            return False
            
        # Then check prefix
        prefix = command[:3].lower()  # Extract the first 3 characters of the command
        return prefix in self.allowed_prefixes

class CustomFormatter(logging.Formatter):
    """Custom log formatter with different formats for file and GUI output."""
    
    def __init__(self, fmt: Optional[str] = None, datefmt: Optional[str] = None, style: str = '%') -> None:
        """Initialize the formatter with format styles for different destinations."""
        super().__init__(fmt, datefmt, style)
        # Format styles for different destinations and levels
        self.file_formats: Dict[int, str] = {
            VERBOSE_LEVEL_NUM: "%(levelname)s: %(message)s",
            logging.INFO: "%(message)s",
            logging.DEBUG: "%(levelname)s: %(message)s",
            logging.WARNING: "%(levelname)s: %(message)s",
            logging.ERROR: "%(levelname)s: %(message)s",
            logging.CRITICAL: "%(levelname)s: %(message)s"
        }
        self.gui_formats: Dict[int, str] = {
            VERBOSE_LEVEL_NUM: "%(levelname)s: %(message)s",
            logging.INFO: "%(message)s",
            logging.DEBUG: "%(levelname)s: %(message)s",
            logging.WARNING: "%(levelname)s: %(message)s",
            logging.ERROR: "%(levelname)s: %(message)s",
            logging.CRITICAL: "%(levelname)s: %(message)s"
        }

    def format(self, record: logging.LogRecord) -> str:
        """
        Format the specified record according to destination.
        
        Args:
            record: Log record to format
            
        Returns:
            Formatted log message
        """
        # Store original format
        original_format = self._style._fmt
        
        # Choose format based on handler type
        if getattr(record, 'for_gui', False):
            self._style._fmt = self.gui_formats.get(record.levelno, "%(message)s")
        else:
            self._style._fmt = self.file_formats.get(record.levelno, "%(levelname)s: %(message)s")
        
        # Format the message
        result = super().format(record)
        
        # Restore original format
        self._style._fmt = original_format
        
        return result

class UnifiedLogger:
    """Centralized logging configuration for both GUI and CLI modes."""
    
    @staticmethod
    def configure_logging(verbose: bool = False, log_file: Optional[str] = None, 
                         gui_handler: Optional[logging.Handler] = None) -> logging.Logger:
        """
        Configure logging with unified formatting.
        
        Args:
            verbose: Enable verbose logging
            log_file: Path to log file
            gui_handler: GUI log handler instance
            
        Returns:
            Configured logger instance
        """
        # Configure base logger
        logger = logging.getLogger()
        
        # Set level based on verbose flag
        if verbose:
            logger.setLevel(VERBOSE_LEVEL_NUM)
        else:
            logger.setLevel(logging.INFO)

        # Clear existing handlers
        logger.handlers = []

        # Common file handler
        if log_file:
            file_handler = logging.FileHandler(
                log_file,
                mode='w',
                encoding='utf-8'
            )
            file_handler.setFormatter(CustomFormatter())
            logger.addHandler(file_handler)

        # Add GUI handler if provided
        if gui_handler:
            logger.addHandler(gui_handler)

        # CLI-specific console handler
        if not gui_handler:
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(CustomFormatter())
            logger.addHandler(console_handler)

        return logger

class LogEmitter(QObject):
    """Qt signal emitter for logging messages to the GUI."""
    log_signal = Signal(str, str)  # (message, levelname)

class GUILogHandler(logging.Handler):
    """Custom logging handler that emits signals for GUI display."""
    
    def __init__(self) -> None:
        """Initialize the handler with a signal emitter."""
        super().__init__()
        self.emitter = LogEmitter()
        self.setFormatter(CustomFormatter())

    def emit(self, record: logging.LogRecord) -> None:
        """
        Emit a log record as a signal for GUI display.
        
        Args:
            record: Log record to emit
        """
        try:
            record.for_gui = True  # Mark record for GUI formatting
            msg = self.format(record)
            self.emitter.log_signal.emit(msg, record.levelname)
        except Exception as e:
            print(f"Logging error: {e}")

class Worker(QThread):
    """Worker thread for executing device documentation tasks."""
    
    progress_signal = Signal(int)  # Progress percentage
    log_signal = Signal(str)       # Log messages
    completion_signal = Signal(str)  # Completion signal with output folder
    stopped_signal = Signal()      # Signal for clean stop

    def __init__(self, device_file: str, command_file: str, credentials: Dict[str, str], 
                 enable_password: str, output_folder: Path, verbose_enabled: bool, 
                 create_combined_output: bool = False) -> None:
        """
        Initialize the worker thread.
        
        Args:
            device_file: Path to device list file
            command_file: Path to command list file
            credentials: Dictionary with 'username' and 'password'
            enable_password: Enable/privileged password
            output_folder: Directory for output files
            verbose_enabled: Whether verbose logging is enabled
            create_combined_output: Whether to create combined output file
        """
        super().__init__()
        self.device_file = device_file
        self.command_file = command_file
        self.credentials = credentials
        self.enable_password = enable_password
        self.output_folder = output_folder
        self.verbose_enabled = verbose_enabled
        self.create_combined_output = create_combined_output
        self._is_cancelled = False  # Cancellation flag
        self.active_connections: List[Any] = []  # Track active connections

        # Configure logging
        self.logger = logging.getLogger(__name__)
        if verbose_enabled:
            self.logger.setLevel("VERBOSE")
        else:
            self.logger.setLevel("INFO")

    def validate_credentials(self, devices: List[str]) -> bool:
        """
        Validate credentials by attempting to connect to the first reachable device.
        
        Args:
            devices: List of device IPs to try
            
        Returns:
            True if credentials are valid, False otherwise
        """
        if not devices:
            self.logger.error(f"No valid devices found in the Devices input file.")
            return False

        # Try to connect to each device until a reachable one is found
        for device in devices:
            self.logger.info(f"Testing credentials on device: {device}")
            try:
                # Attempt to connect to the device
                connection = ConnectHandler(
                    device_type='cisco_ios',
                    host=device,
                    username=self.credentials['username'],
                    password=self.credentials['password'],
                    secret=self.enable_password,
                    banner_timeout=60,  # Set a timeout to wait for the SSH banner, 0 to skip banner
                    conn_timeout=5,  # Connection timeout (seconds)
                    read_timeout_override=5,  # Read timeout (seconds)
                    global_delay_factor=0.5  # Reduce delay factor for faster response
                )
                connection.enable()
                connection.disconnect()
                self.logger.info(f"Credentials validated successfully on device: {device}")
                return True
            except NetmikoTimeoutException:
                self.logger.warning(f"Device unreachable: {device}")
                continue  # Skip to the next device
            except NetmikoAuthenticationException:
                self.logger.error(f"Invalid credentials for device: {device}")
                return False  # Credentials are invalid, no need to try other devices
            except Exception as e:
                self.logger.error(f"Unexpected error validating credentials on device {device}: {e}")
                continue  # Skip to the next device

        # If no reachable devices were found
        self.logger.error(f"No reachable devices found in the Device(s) input file.")
        return False

    def run(self) -> None:
        """Main execution method for the worker thread."""
        try:
            start_time = datetime.now()
            start_time_str = format_time(start_time)
            self.logger.info(f"***** Script started - {start_time_str} *****")
            self.logger.info(f"{DIVIDER}")
            self.logger.info(f"{Path(__file__).name} {APP_VERSION} {VERSION_DATE}")
            self.logger.info(f"{DIVIDER}")

            # Initialize list to track output files
            self.output_files: List[Path] = []

            # Read and validate devices
            devices = self.read_devices(self.device_file)
            
            # Read and validate commands
            commands = self.read_commands(self.command_file)

            # Check if we have at least one valid device and one valid command
            if not devices or not commands:
                self.logger.error(f"Script cannot proceed - no valid devices or commands found")
                end_time = datetime.now()
                end_time_str = format_time(end_time)
                total_time = end_time - start_time
                
                self.logger.info(f"{DIVIDER}")
                self.logger.info(f"Script Summary:")
                self.logger.info(f"  Valid devices found: {len(devices)}")
                self.logger.info(f"  Valid commands found: {len(commands)}")
                self.logger.info(f"  Script run time (h:mm:ss.ms): {total_time}")
                self.logger.info(f"{DIVIDER}")
                self.logger.info(f"***** Script ended - {end_time_str} *****")
                self.completion_signal.emit(str(self.output_folder))
                return

            total_devices = len(devices)
            successful_devices = 0
            failed_devices = 0
            total_commands = len(commands)
            successful_commands = 0
            failed_commands = 0

            # Validate credentials before proceeding
            if not self.validate_credentials(devices):
                self.logger.error(f"Script terminated due to credential validation failure.")
                end_time = datetime.now()
                end_time_str = format_time(end_time)
                total_time = end_time - start_time

                self.logger.info(f"{DIVIDER}")
                self.logger.info(f"Script Summary:")
                self.logger.info(f"  Script run time (h:mm:ss.ms): {total_time}")
                self.logger.info(f"{DIVIDER}")
                self.logger.info(f"***** Script ended - {end_time_str} *****")
                self.completion_signal.emit(str(self.output_folder))  # Emit the completion signal
                return

            # Log Verbose output
            if self.verbose_enabled:
                self.logger.verbose(f"Verbose Status: {self.verbose_enabled}")
                self.logger.verbose(f"Credentials - Username: {self.credentials['username']}")
                self.logger.verbose(f"Credentials - Password: {self.credentials['password']}")
                self.logger.verbose(f"Credentials - Enable Password: {self.enable_password}")
                self.logger.verbose(f'Input File - Devices: "{self.device_file}"')
                self.logger.verbose(f'Input File - Commands: "{self.command_file}"')
                self.logger.verbose(f"CPU Count: {os.cpu_count()}, max_workers: {max_workers}")

            with concurrent.futures.ThreadPoolExecutor(max_workers) as executor:
                futures = {}
                for i, device in enumerate(devices, start=1):  # Track device count
                    future = executor.submit(self.process_device, device, commands, i, total_devices)
                    futures[future] = device

                for future in concurrent.futures.as_completed(futures):
                    try:
                        device, success, commands_executed = future.result()
                        if success:
                            successful_devices += 1
                            successful_commands += commands_executed
                        else:
                            failed_devices += 1
                            failed_commands += commands_executed
                        self.progress_signal.emit(int((successful_devices + failed_devices) / total_devices * 100))
                    except Exception as e:
                        self.logger.error(f"Error processing device: {e}")
                        failed_devices += 1

            # Only create combined output if enabled
            if self.create_combined_output and self.output_files:
                self.create_combined_output_file()

            end_time = datetime.now()
            end_time_str = format_time(end_time)
            total_time = end_time - start_time

            self.logger.info(f"{DIVIDER}")
            self.logger.info(f"Script Summary:")
            self.logger.info(f"  Total commands per device: {total_commands}")
            self.logger.info(f"  Successful devices: {successful_devices}")
            self.logger.info(f"  Failed devices: {failed_devices}")
            self.logger.info(f"  Script run time (h:mm:ss.ms): {total_time}")
            self.logger.info(f"{DIVIDER}")
            self.logger.info(f"***** Script ended - {end_time_str} *****")

            self.completion_signal.emit(str(self.output_folder))  # Emit the completion signal

        except Exception as e:
            self.logger.error(f"{e}")
            self.completion_signal.emit(str(self.output_folder))  # Emit the completion signal even if there's an error

    def read_devices(self, device_file: str) -> List[str]:
        """
        Read and validate the Device(s) input file.
        
        Args:
            device_file: Path to the device list file
            
        Returns:
            List of unique, valid device IPs
        """
        devices: List[str] = []
        seen_devices: Set[str] = set()
        valid_devices_found = False

        try:
            with open(device_file, 'r', encoding='utf-8') as file:
                for line_number, line in enumerate(file, start=1):
                    line = line.strip()
                    if '#' in line:
                        line = line.split('#', 1)[0]
                    
                    line = line.strip()
                    if not line:  # Skip empty lines
                        continue

                    # Check for valid IP address
                    try:
                        ipaddress.ip_address(line)
                        # Check for duplicate IP addresses
                        if line in seen_devices:
                            self.logger.warning(f"Line {line_number}: Duplicate IP address - {line} (ignored)")
                            continue
                        seen_devices.add(line)
                        # If valid, add to the list of devices
                        devices.append(line)
                        valid_devices_found = True
                    except ValueError:
                        self.logger.warning(f"Line {line_number}: Invalid IP address - {line} (ignored)")

            if not valid_devices_found:
                self.logger.error(f"No valid device IP addresses found in the Device(s) input file")
            return devices

        except Exception as e:
            self.logger.error(f"Error reading device file: {e}")
            return []

    def read_commands(self, command_file: str) -> List[str]:
        """
        Read and validate the Command(s) input file.
        
        Args:
            command_file: Path to the command list file
            
        Returns:
            List of unique, valid commands
        """
        commands: List[str] = []
        seen_commands: Set[str] = set()
        valid_commands_found = False
        allowed_commands = AllowedCommands()

        try:
            with open(command_file, 'r', encoding='utf-8') as file:
                for line_number, line in enumerate(file, start=1):
                    line = line.strip()
                    if '#' in line:
                        line = line.split('#', 1)[0]
                    
                    line = line.strip()
                    if not line:  # Skip empty lines
                        continue

                    # Check command length
                    if len(line) > MAX_COMMAND_LENGTH:
                        self.logger.warning(f"Line {line_number}: Command too long (>{MAX_COMMAND_LENGTH} chars) - {line[:20]}... (ignored)")
                        continue

                    # Check for valid command
                    if not allowed_commands.is_command_allowed(line):
                        self.logger.warning(f"Line {line_number}: Invalid command - {line} (ignored)")
                        continue

                    # Check for duplicate commands
                    if line in seen_commands:
                        self.logger.warning(f"Line {line_number}: Duplicate command - {line} (ignored)")
                        continue

                    seen_commands.add(line)
                    commands.append(line)
                    valid_commands_found = True

            if not valid_commands_found:
                self.logger.error(f"No valid commands found in the Command(s) input file")
            return commands

        except Exception as e:
            self.logger.error(f"Error reading command file: {e}")
            return []

    def cancel(self) -> None:
        """Signal the worker to stop execution and clean up connections."""
        self._is_cancelled = True
        self.logger.info(f"Cancellation requested - cleaning up, please wait...")
        self.logger.warning(f"Script execution cancelled by user - output may be incomplete")

        # Disconnect all active connections
        for conn in self.active_connections:
            try:
                if conn.is_alive():
                    conn.disconnect()
            except Exception as e:
                self.logger.warning(f"Error disconnecting: {e}")
        self.active_connections.clear()

    def process_device(self, device: str, commands: List[str], 
                      device_count: int, total_devices: int) -> Tuple[str, bool, int]:
        """
        Process a single device with cancellation support.
        
        Args:
            device: Device IP to process
            commands: List of commands to execute
            device_count: Current device number (for progress tracking)
            total_devices: Total number of devices (for progress tracking)
            
        Returns:
            Tuple of (device_ip, success_flag, commands_executed)
        """
        if self._is_cancelled:
            return device, False, 0

        try:
            connection = ConnectHandler(
                device_type='cisco_ios',
                host=device,
                username=self.credentials['username'],
                password=self.credentials['password'],
                secret=self.enable_password,
                banner_timeout=60,  # Set a timeout to wait for the SSH banner, 0 to skip banner
                conn_timeout=10,  # Connection timeout (seconds) - Default: 10
                read_timeout_override=40,  # Read timeout (seconds) - Default: none
                global_delay_factor=1  # Delay factor (seconds) - Default: 1
            )
            connection.enable()
            self.active_connections.append(connection)  # Track connection

            # Try to get the device's hostname
            devicename = self.get_device_hostname(connection)
            output_file = self.generate_output_filename(device, devicename)

            # Track the output file in order
            self.output_files.append(output_file)

            # Log the connection with the device count
            self.logger.info(f"Connected ({device_count} of {total_devices}) to {devicename} - {device}" if devicename 
                             else f"Connected ({device_count} of {total_devices}) to {device}")

            with open(output_file, 'w', encoding='utf-8') as file:
                # Device output file
                file.write(f"***** DOCUMENTATION STARTED - {format_time()} *****")
                commands_executed = 0
                successful_commands = 0
                failed_commands = 0

                for command in commands:
                    if self._is_cancelled:  # Check for cancellation before each command
                        raise Exception("Execution cancelled by user")

                    try:
                        output = connection.send_command(command)
                        # Check for invalid command output
                        if "Invalid input detected at '^' marker" in output:
                            output = str(f"Command not valid (on this platform): {command}")
                            self.logger.warning(f"Command not valid (on this platform): {command}")

                            failed_commands += 1
                        else:
                            successful_commands += 1

                        file.write(f"\n\n\n!  {format_time()}  {FILLER1}  {command}  {FILLER1}\n")
                        commands_executed += 1
                        file.write(f"\n{output}\n")
                        self.logger.verbose(f"Executed command '{command}' on {devicename} - {device}" if devicename 
                                          else f"Executed command '{command}' on {device}")

                    except Exception as e:
                        self.logger.error(f"Failed to execute command '{command}' on {devicename} - {device}: {e}" 
                                          if devicename else f"Failed to execute command '{command}' on {device}: {e}")
                        failed_commands += 1

                file.write(f"\n\n\n***** DOCUMENTATION ENDED - {format_time()} *****\n")

            connection.disconnect()
            self.active_connections.remove(connection)  # Remove from tracking

            # Log the number of successful and failed commands
            self.logger.info(f"Disconnected ({device_count} of {total_devices}) from {devicename} - {device}" 
                             if devicename else f"Disconnected ({device_count} of {total_devices}) from {device}")
            self.logger.info(f"  {successful_commands} of {len(commands)} command(s) successful, {failed_commands} failed")

            # Log the path to the output file
            try:
                relative_output_path = output_file.relative_to(Path(__file__).parent)
                self.logger.info(f'  Output file: "./{relative_output_path.as_posix()}"')
            except ValueError:
                # If the file is not in a subpath of the script's directory, log the absolute path
                self.logger.info(f'  Output file: "{output_file}"')

            return device, True, commands_executed

        except (NetmikoTimeoutException, NetmikoAuthenticationException) as e:
            self.logger.error(f"({device_count} of {total_devices}) Failed to connect to {device}")
            return device, False, 0

        except Exception as e:
            if str(e) == "Execution cancelled by user":
                self.logger.info(f" Execution stopped while processing {device}")
            else:
                self.logger.error(f"Unexpected error processing {device}: {e}")
            
            # Ensure connection is cleaned up
            if 'connection' in locals() and connection.is_alive():
                try:
                    connection.disconnect()
                    if connection in self.active_connections:
                        self.active_connections.remove(connection)
                except Exception as e:
                    self.logger.warning(f"Error during disconnect: {e}")

            return device, False, 0

    def get_device_hostname(self, connection: Any) -> Optional[str]:
        """
        Retrieve the device's hostname from the running configuration.
        
        Args:
            connection: Active Netmiko connection
            
        Returns:
            Device hostname if found, None otherwise
        """
        try:
            # Send command to get the full hostname from running configuration
            output = connection.send_command('show running-config | include hostname')
            if output:
                hostname_line = output.strip()
                if hostname_line.startswith('hostname'):
                    hostname = hostname_line.split()[1]
                    return hostname
            hostname = connection.base_prompt  # Fallback: Directly use Netmiko's detected prompt
            return hostname.strip()
        except Exception as e:
            self.logger.warning(f"Failed to retrieve hostname: {e}")
            return None

    def generate_output_filename(self, device: str, devicename: Optional[str]) -> Path:
        """
        Generate the output filename for the device.
        
        Args:
            device: Device IP address
            devicename: Optional device hostname
            
        Returns:
            Path object for the output file
        """
        if devicename:
            filename = self.output_folder / f"{devicename} - {device}.txt"
        else:
            filename = self.output_folder / f"{device}.txt"

        # Try to return the relative path, but fall back to the absolute path if it fails
        try:
            relative_path = filename.relative_to(Path(__file__).parent)
            return Path(f"./{relative_path.as_posix()}")
        except ValueError:
            # If the file is not in a subpath of the script's directory, return the absolute path
            return filename

    def create_combined_output_file(self) -> None:
        """Combine all individual output files into one master file in order of processing."""
        if not self.output_files:
            self.logger.warning(f"No output files to combine")
            return

        # Try to use the relative path, but fall back to the absolute path if it fails
        try:
            combined_file = Path(self.output_folder / COMBINED_OUTPUT_FILENAME).relative_to(Path(__file__).parent)
        except ValueError:
            # If the file is not in a subpath of the script's directory, return the absolute path
            combined_file = self.output_folder / COMBINED_OUTPUT_FILENAME

        try:
            with open(combined_file, 'w', encoding='utf-8') as outfile:
                # Write header
                outfile.write(f"***** COMBINED DEVICE OUTPUT - {format_time()} *****\n")
                outfile.write(f"Combined output from {len(self.output_files)} devices\n")
                outfile.write(f"{DIVIDER}\n")

                # Append each device's output file
                for output_file in self.output_files:
                    try:
                        with open(output_file, 'r', encoding='utf-8') as infile:
                            # Write device header
                            outfile.write(f"\n\n{FILLER2} {output_file.stem} {FILLER2}\n\n")
                            # Copy contents
                            outfile.write(infile.read())
                            outfile.write("\n\n")
                    except Exception as e:
                        self.logger.error(f"Failed to combine file {output_file}: {e}")

                # Write footer
                outfile.write(f"{DIVIDER}\n")
                outfile.write(f"***** END OF COMBINED OUTPUT - {format_time()} *****\n")

            self.logger.info(f'Created combined output file: "{combined_file}"')
        except Exception as e:
            self.logger.error(f"Failed to create combined output file: {e}")

class LineNumberArea(QWidget):
    """Widget that displays line numbers for a CodeEditor."""
    
    def __init__(self, editor: 'CodeEditor') -> None:
        """Initialize with a reference to the parent editor."""
        super().__init__(editor)
        self.editor = editor

    def sizeHint(self) -> QSize:
        """Return the recommended size for the line number area."""
        return QSize(self.editor.line_number_area_width(), 0)

    def paintEvent(self, event: Any) -> None:
        """Handle paint events by delegating to the editor."""
        self.editor.line_number_area_paint_event(event)

class CodeEditor(QPlainTextEdit):
    """Enhanced text editor with line numbers and syntax highlighting."""
    
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        """Initialize the editor with line number support."""
        super().__init__(parent)
        self.line_number_area = LineNumberArea(self)

        # Define padding between line numbers and text
        self.TEXT_WINDOW_PADDING = 10

        # Connect signals
        self.blockCountChanged.connect(self.update_line_number_area_width)
        self.updateRequest.connect(self.update_line_number_area)
        self.cursorPositionChanged.connect(self.highlight_current_line)

        # Set up the line number area
        self.update_line_number_area_width()
        self.highlight_current_line()

    def line_number_area_width(self) -> int:
        """Calculate the width required for the line number area."""
        digits = 1
        max_lines = max(1, self.blockCount())
        while max_lines >= 10:
            max_lines /= 10
            digits += 1
        # Add some extra space for padding
        space = 10 + self.fontMetrics().horizontalAdvance('9') * digits
        return space

    def update_line_number_area_width(self) -> None:
        """Update the width of the line number area and add padding to the text window."""
        self.setViewportMargins(
            self.line_number_area_width() + self.TEXT_WINDOW_PADDING,  # Left margin (line numbers + padding)
            0,  # Top margin
            0,  # Right margin
            0   # Bottom margin
        )
        self.line_number_area.update()  # Force a redraw of the line number area

    def update_line_number_area(self, rect: QRect, dy: int) -> None:
        """Update the line number area when the text changes."""
        if dy:
            self.line_number_area.scroll(0, dy)
        else:
            self.line_number_area.update(0, rect.y(), self.line_number_area.width(), rect.height())
        if rect.contains(self.viewport().rect()):
            self.update_line_number_area_width()

    def resizeEvent(self, event: Any) -> None:
        """Handle resize events to adjust the line number area."""
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.line_number_area.setGeometry(
            QRect(
                cr.left(),  # X position
                cr.top(),  # Y position
                self.line_number_area_width(),  # Width
                cr.height()  # Height
            )
        )

    def line_number_area_paint_event(self, event: Any) -> None:
        """Paint the line numbers in the line number area."""
        painter = QPainter(self.line_number_area)
        painter.fillRect(event.rect(), QColor(240, 240, 240))  # Light gray background

        block = self.firstVisibleBlock()
        block_number = block.blockNumber()
        top = self.blockBoundingGeometry(block).translated(self.contentOffset()).top()
        bottom = top + self.blockBoundingRect(block).height()

        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                number = str(block_number + 1)
                painter.setPen(Qt.GlobalColor.black)
                # Calculate the right-aligned position for the line number
                text_width = self.fontMetrics().horizontalAdvance(number)
                x = self.line_number_area.width() - text_width - 5  # Right-align with 5px padding
                painter.drawText(
                    x,  # Right-aligned X position
                    int(top),
                    text_width,  # Width of the text
                    self.fontMetrics().height(),
                    Qt.AlignmentFlag.AlignRight,
                    number
                )

            block = block.next()
            top = bottom
            bottom = top + self.blockBoundingRect(block).height()
            block_number += 1

    def highlight_current_line(self) -> None:
        """Highlight the current line in the editor."""
        extra_selections = []
        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()
            line_color = QColor(255, 255, 0, 50)  # Light yellow highlight
            selection.format.setBackground(line_color)
            selection.format.setProperty(QTextFormat.Property.FullWidthSelection, True)
            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()
            extra_selections.append(selection)
        self.setExtraSelections(extra_selections)

class FileEditorDialog(QDialog):
    """Dialog for editing input files with syntax highlighting."""
    
    file_path_updated = Signal(Path, str)  # Signal to emit when the file path changes (Path, file_type)

    def __init__(self, file_path: Path, file_type: str, parent: Optional[QWidget] = None) -> None:
        """
        Initialize the file editor dialog.
        
        Args:
            file_path: Path to the file to edit
            file_type: Type of file ('device' or 'command')
            parent: Optional parent widget
        """
        super().__init__(parent)
        self.file_path = file_path
        self.file_type = file_type  # Track whether this is a "device" or "command" file
        self.setWindowTitle(f"Editing: {file_path.name}")  # Display only the file name
        self.setModal(True)  # Make the dialog modal
        self.resize(800, 600)
        self.unsaved_changes = False  # Track unsaved changes

        # Create the main layout
        layout = QVBoxLayout(self)

        # Create a CodeEditor for editing the file (with line numbers)
        self.text_edit = CodeEditor(self)
        self.text_edit.textChanged.connect(self.mark_unsaved_changes)  # Connect to textChanged signal
        layout.addWidget(self.text_edit)

        # Create a status label to show messages
        self.status_label = QLabel("Ready", self)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("font-weight: bold; color: black;")
        layout.addWidget(self.status_label)

        # Create a horizontal layout for the buttons
        button_layout = QHBoxLayout()

        # Add Save button
        self.save_button = QPushButton("Save", self)
        self.save_button.clicked.connect(self.save_file)
        self.save_button.setEnabled(False)  # Disabled by default
        button_layout.addWidget(self.save_button)

        # Add Save As button
        self.save_as_button = QPushButton("Save As...", self)
        self.save_as_button.clicked.connect(self.save_file_as)
        self.save_as_button.setEnabled(True)  # Always enabled
        button_layout.addWidget(self.save_as_button)

        # Add Close button
        self.close_button = QPushButton("Close", self)
        self.close_button.clicked.connect(self.close_editor)
        button_layout.addWidget(self.close_button)

        # Add the button layout to the main layout
        layout.addLayout(button_layout)

        # Set the layout
        self.setLayout(layout)

        # Load the file content after setting up the UI
        self.load_file(file_path)

    def load_file(self, file_path: Path) -> None:
        """Load the file contents into the editor."""
        try:
            # Block the textChanged signal while loading the file
            self.text_edit.blockSignals(True)  # Suppress textChanged signal
            with open(file_path, "r") as file:
                self.file_content = file.read()
            self.text_edit.setPlainText(self.file_content)  # Set the text in the CodeEditor
            self.text_edit.blockSignals(False)  # Re-enable textChanged signal

            # Move the cursor to the start of the document
            cursor = self.text_edit.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            self.text_edit.setTextCursor(cursor)

            # Highlight the first line
            self.text_edit.highlight_current_line()

            # Ensure the first line is visible
            self.text_edit.ensureCursorVisible()

            # Ensure the line number area and padding are updated
            self.text_edit.update_line_number_area_width()
            self.text_edit.line_number_area.update()  # Force a redraw of the line number area

            self.status_label.setText("File loaded successfully.")
            self.status_label.setStyleSheet("font-weight: bold; color: green;")
        except Exception as e:
            self.status_label.setText(f"Error reading file: {e}")
            self.status_label.setStyleSheet("font-weight: bold; color: red;")
            self.close()

    def save_file(self) -> None:
        """Save changes to the original file."""
        try:
            with open(self.file_path, "w") as file:
                file.write(self.text_edit.toPlainText())
            self.unsaved_changes = False  # Reset unsaved changes flag
            self.setWindowTitle(f"Editing: {self.file_path.name}")  # Update title
            self.save_button.setEnabled(False)  # Disable Save button after saving
            self.status_label.setText("File saved successfully.")
            self.status_label.setStyleSheet("font-weight: bold; color: green;")
        except Exception as e:
            self.status_label.setText(f"Failed to save file: {e}")
            self.status_label.setStyleSheet("font-weight: bold; color: red;")

    def save_file_as(self) -> None:
        """Save changes to a new file."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save File As", str(self.file_path.parent), "Text Files (*.txt);;All Files (*)"
        )
        if file_path:
            try:
                with open(file_path, "w") as file:
                    file.write(self.text_edit.toPlainText())
                self.unsaved_changes = False  # Reset unsaved changes flag
                self.file_path = Path(file_path)  # Update the file path
                self.setWindowTitle(f"Editing: {self.file_path.name}")  # Update title
                self.save_button.setEnabled(False)  # Disable Save button after saving
                self.status_label.setText("File saved successfully.")
                self.status_label.setStyleSheet("font-weight: bold; color: green;")
                # Emit the new file path and file type
                self.file_path_updated.emit(self.file_path, self.file_type)
            except Exception as e:
                self.status_label.setText(f"Failed to save file: {e}")
                self.status_label.setStyleSheet("font-weight: bold; color: red;")

    def close_editor(self) -> None:
        """Prompt to save changes before closing if there are unsaved changes."""
        if self.unsaved_changes:
            reply = QMessageBox.question(
                self, "Save Changes", "Do you want to save changes before closing?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.save_file()
                self.accept()  # Close the dialog and return the new file path
            elif reply == QMessageBox.StandardButton.No:
                self.status_label.setText("Changes discarded.")
                self.status_label.setStyleSheet("font-weight: bold; color: orange;")
                self.reject()  # Close the dialog without saving
            # If Cancel, do nothing
        else:
            self.reject()  # Close the dialog without saving

    def mark_unsaved_changes(self) -> None:
        """Mark that the file has unsaved changes."""
        if not self.unsaved_changes:  # Only update if changes are not already marked
            self.unsaved_changes = True
            self.setWindowTitle(f"Editing: {self.file_path.name} *")  # Add asterisk to indicate unsaved changes
            self.save_button.setEnabled(True)  # Enable Save button
            self.status_label.setText("Unsaved changes.")
            self.status_label.setStyleSheet("font-weight: bold; color: orange;")

    def closeEvent(self, event: Any) -> None:
        """Override the close event to prompt for saving changes."""
        self.close_editor()
        event.ignore()  # Prevent the dialog from closing immediately

class HelpDialog(QDialog):
    """Custom dialog for displaying help information."""
    
    def __init__(self, title: str, version: str, version_date: str, 
                 parent: Optional[QWidget] = None) -> None:
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
        
        # Set Help window size (width, height)
        self.setFixedSize(620, 800)

        # Get repo URL and callback from parent if available
        self.repo_url = REPO_URL
        # self.check_update_callback = None
        # if parent and hasattr(parent, 'check_for_updates'):
        #     self.check_update_callback = parent.check_for_updates
        
        # Create main layout with margins
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)  # Window margins

        # Add the title and version
        title_label = QLabel(f"<h1>{APP_NAME}</h1><h3>{APP_VERSION} {VERSION_DATE}</h3>")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Create a widget for repo link and update check button
        repo_widget = QWidget()
        repo_layout = QHBoxLayout(repo_widget)
        repo_layout.setContentsMargins(0, 10, 0, 10)  # Add some top margin

        # Repo link
        repo_label = QLabel(f'<a href="{self.repo_url}" style="color: #4CAF50;">Visit Website</a>')
        repo_label.setOpenExternalLinks(True)
        repo_layout.addWidget(repo_label)
        repo_layout.addStretch()

        # Add title and repo widget
        layout.addWidget(title_label)
        layout.addWidget(repo_widget, alignment=Qt.AlignmentFlag.AlignCenter)

        # Create scroll area (no frame)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)  # Remove default border

        # Create container widget
        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)

        # Create framed content area
        content_frame = QFrame()
        content_frame.setObjectName("contentFrame")
        frame_layout = QVBoxLayout(content_frame)
        frame_layout.setContentsMargins(15, 15, 15, 15)  # Padding inside frame

        # Add the help text
        help_text = f"""
        This tool helps you document network devices by processing:
        1. A list of device IPs (Device List).
        2. A list of commands to run on each device (Command List).
        3. Credentials for accessing the devices.

        Usage:
        1. Select the Device List and Command List files.
        2. Choose a credential option (Manual or Keyring).
        3. Click 'Go' to start the process.
        4. Use 'Quit' to exit the application.

        CLI Usage:
        $ python nddu.py -cli [options]

        Options:
        -h, --help\t\tShow this help message and exit
        -v, --version\t\tShow version information and check for updates
        -cli, --cli-mode\tRun in CLI mode (required for CLI mode)
        -d, --device-file\tPath to the device list file (default: ./input/Devices.txt)
        -c, --command-file\tPath to the command list file (default: ./input/Commands.txt)
        -ks, --keyring-system\tKeyring system name for keyring credentials (requires -ku)
        -ku, --keyring-user\tKeyring user name for keyring credentials (requires -ks)
        --verbose\t\tEnable verbose output
        --combined\t\tEnable creation of combined output file

        CLI Behavior:
        - The GUI will only start if no arguments are provided.
        - If -ks or -ku are used, both must be provided.
        - If neither -ks nor -ku are used, the script will prompt for credentials.
        - If -d and -c are used, they will be used as specified.
        - If -d is used without -c, -c will use the default Commands file.
        - If -c is used without -d, -c will use the default Devices file.
        - If neither -d nor -c are used, both default files will be used.
        """

        help_label = QLabel(help_text)
        help_label.setWordWrap(True)
        help_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        frame_layout.addWidget(help_label)
        container_layout.addWidget(content_frame)
        scroll_area.setWidget(container)
        layout.addWidget(scroll_area)

        # Create a widget for the Close button (centered)
        close_widget = QWidget()
        close_layout = QHBoxLayout(close_widget)
        close_layout.setContentsMargins(0, 10, 0, 10)
        
        # Add stretch to center the button
        close_layout.addStretch()
        
        # Add Close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.close)
        close_button.setMinimumWidth(100)
        close_layout.addWidget(close_button)
        
        # Add stretch to center the button
        close_layout.addStretch()
        
        # Add close widget to main layout
        layout.addWidget(close_widget)

        # Apply theme
        self.apply_theme()

    def apply_theme(self) -> None:
        """Apply dark or light theme based on current mode."""
        if self.dark_mode:
            # Dark theme palette
            dark_palette = QPalette()
            dark_palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
            dark_palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
            dark_palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
            dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
            dark_palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
            self.setPalette(dark_palette)
            
            # Dark theme stylesheet
            self.setStyleSheet("""
                QFrame#contentFrame {
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
                    padding: 8px 20px;
                    border-radius: 4px;
                    min-width: 100px;
                }
                QPushButton:hover {
                    background-color: #454545;
                }
                QPushButton:disabled {
                    background-color: #555;
                    color: #999;
                }
            """)
        else:
            # Reset to default light theme
            self.setPalette(QApplication.style().standardPalette())
            self.setStyleSheet("""
                QPushButton:disabled {
                    background-color: #e0e0e0;
                    color: #999;
                }
            """)

class VersionChecker(QObject):
    """Check for updates using GitHub Releases API."""
    update_found = Signal(str, str)  # (new_version, release_url)
    check_complete = Signal(bool)  # Whether check was successful
    
    def check(self) -> None:
        """Check for updates in a non-blocking way."""
        try:
            # Create request with headers (GitHub API likes User-Agent)
            headers = {
                'User-Agent': f'{APP_NAME}/{APP_VERSION}',
                'Accept': 'application/vnd.github.v3+json'
            }
            req = urllib.request.Request(GITHUB_API_LATEST_RELEASE, headers=headers)
            
            with urllib.request.urlopen(req, timeout=3) as response:
                data = json.loads(response.read().decode())
                latest_tag = data.get('tag_name', '')  # e.g., "v1.1.0"
                current_version = APP_VERSION  # e.g., "v1.0.0"
                
                # Remove 'v' prefix for comparison
                latest_version_str = latest_tag.lstrip('v')
                current_version_str = current_version.lstrip('v')
                
                # Compare versions
                latest_ver = version.parse(latest_version_str)
                current_ver = version.parse(current_version_str)
                
                if latest_ver > current_ver:
                    release_url = data.get('html_url', REPO_URL)
                    self.update_found.emit(latest_version_str, release_url)
                
                self.check_complete.emit(True)
                
        except Exception:
            # Silently fail - don't interrupt user
            self.check_complete.emit(False)

class MyWindow(QWidget):
    """Main application window for the Network Device Documentation Utility."""
    
    def __init__(self) -> None:
        """Initialize the main window with default settings."""
        super().__init__()
        self.dark_mode = DARK_MODE_STATE  # Set the dark mode state
        self.enable_was_enabled = False
        self.verbose_was_enabled = False
        self.combined_output_was_enabled = False
        self.stop_requested = False
        self.update_available = False
        self.new_version = ""
        self.release_url = ""
        self.toggle_theme(self.dark_mode)  # Toggle theme according to DARK_MODE_STATE
        self.init_ui()

        # Start update check after UI is shown
        QTimer.singleShot(1000, self.check_for_updates)

    def configure_logging(self, log_file: str) -> None:
        """
        Configure logging for the application.
        
        Args:
            log_file: Path to the log file
        """
        # Configure root logger
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)
        
        # Clear existing handlers
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)

        # Create formatter
        formatter = CustomFormatter()
        
        # Configure file handler
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(formatter)  # Apply to file handler
        
        # Configure GUI handler
        self.gui_handler = GUILogHandler()
        self.gui_handler.setFormatter(logging.Formatter('%(message)s'))  # Raw for GUI
        
        # Add both handlers
        self.logger.addHandler(self.gui_handler)
        self.logger.addHandler(file_handler)
        
        # Connect the signal from the emitter to our handler method
        self.gui_handler.emitter.log_signal.connect(self.handle_log_message)

    def handle_log_message(self, message: str, level_name: str) -> None:
        """
        Handle messages from the logging system.
        
        Args:
            message: Log message content
            level_name: Log level name (e.g., "INFO", "ERROR")
        """
        self.append_colored_message(message, level_name)

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

    def init_ui(self) -> None:
        """Initialize all UI components and layouts."""
        # Set window properties
        self.setWindowTitle(f"nddu")
        self.setMaximumSize(730, 1000)  # Width, Height

        # Define the relative path to the "input" folder
        self.input_folder = DEFAULT_INPUT_FOLDER
        self.default_device_list = DEFAULT_DEVICE_FILE
        self.default_command_list = DEFAULT_COMMAND_FILE

        # Create the logo and title section
        logo_title_layout = QHBoxLayout()
        logo_title_layout.setSpacing(5)

        # Load and resize the logo
        self.logo = QLabel(self)
        logo_path = LOGO_PATH
        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            pixmap = pixmap.scaledToHeight(80, Qt.TransformationMode.SmoothTransformation)
            self.logo.setPixmap(pixmap)
        else:
            self.logo.setText("Logo not found")
            self.logo.setStyleSheet("color: red;")
        logo_title_layout.addSpacing(10)
        logo_title_layout.addWidget(self.logo)
        logo_title_layout.addSpacing(20)

        # Create a container widget with fixed height
        title_container = QWidget()
        title_container.setFixedHeight(60)  # Enough for title + update indicator
        title_layout = QVBoxLayout(title_container)
        title_layout.setContentsMargins(0, 0, 0, 0)  # No margins
        title_layout.setSpacing(2)

        # Add the script name and version
        self.title_label = QLabel(
            f"<span style='font-size: 18px; font-weight: bold;'>{APP_NAME}</span><br>"
            f"<span style='font-size: 12px;'>{APP_VERSION}</span>", 
            self
        )
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        title_layout.addWidget(self.title_label)

        # Add update indicator (initially empty)
        self.update_indicator = QLabel("", self)
        self.update_indicator.setStyleSheet("color: #4CAF50; font-size: 11px;")
        self.update_indicator.setOpenExternalLinks(True)
        self.update_indicator.setCursor(Qt.PointingHandCursor)
        title_layout.addWidget(self.update_indicator)

        # Now add the container to the logo_title_layout
        logo_title_layout.addWidget(title_container)
        logo_title_layout.addStretch()

        # Create the Input Files section
        input_files_group = QGroupBox("Input Files")
        input_files_layout = QVBoxLayout()
        input_files_layout.setSpacing(5)

        # Device List
        self.device_label = QLabel("Device(s):", self)
        self.device_input = QLineEdit(self)
        self.device_input.setFixedWidth(280)
        self.device_input.setReadOnly(True)
        self.device_button = QPushButton("Browse", self)
        self.device_button.clicked.connect(self.browse_device_list)
        self.device_edit_button = QPushButton("Edit", self)
        self.device_edit_button.clicked.connect(self.edit_device_file)
        device_layout = QHBoxLayout()
        device_layout.setSpacing(5)
        device_layout.addWidget(self.device_input)
        device_layout.addWidget(self.device_edit_button)
        device_layout.addWidget(self.device_button)
        input_files_layout.addWidget(self.device_label)
        input_files_layout.addLayout(device_layout)

        # Command List
        self.command_label = QLabel("Command(s):", self)
        self.command_input = QLineEdit(self)
        self.command_input.setFixedWidth(280)
        self.command_input.setReadOnly(True)
        self.command_button = QPushButton("Browse", self)
        self.command_button.clicked.connect(self.browse_command_list)
        self.command_edit_button = QPushButton("Edit", self)
        self.command_edit_button.clicked.connect(self.edit_command_file)
        command_layout = QHBoxLayout()
        command_layout.setSpacing(5)
        command_layout.addWidget(self.command_input)
        command_layout.addWidget(self.command_edit_button)
        command_layout.addWidget(self.command_button)
        input_files_layout.addWidget(self.command_label)
        input_files_layout.addLayout(command_layout)
        input_files_group.setLayout(input_files_layout)

        # Create the Credential Options section
        credentials_group = QGroupBox("Credential Options")
        credentials_layout = QVBoxLayout()
        credentials_layout.setSpacing(5)

        # Radio buttons for credential options
        self.manual_radio = QRadioButton("Manual Credentials", self)
        self.keyring_radio = QRadioButton("Keyring Credentials", self)
        self.manual_radio.setChecked(True)
        self.manual_radio.toggled.connect(self.toggle_credential_options)
        self.keyring_radio.toggled.connect(self.toggle_credential_options)

        # Manual Credentials section
        self.manual_credentials_group = QGroupBox()
        manual_credentials_layout = QVBoxLayout()
        manual_credentials_layout.setSpacing(5)

        # Username
        username_layout = QHBoxLayout()
        username_layout.setSpacing(5)
        self.username_label = QLabel("Username:", self)
        self.username_input = QLineEdit(self)
        self.username_input.setFixedWidth(280)
        self.username_input.textChanged.connect(self.validate_fields)
        username_layout.addWidget(self.username_label)
        username_layout.addWidget(self.username_input)
        manual_credentials_layout.addLayout(username_layout)

        # Password
        password_layout = QHBoxLayout()
        password_layout.setSpacing(5)
        self.password_label = QLabel("Password:", self)
        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setFixedWidth(280)
        self.password_input.textChanged.connect(self.validate_fields)
        self.password_input.textChanged.connect(self.sync_enable_password)
        password_layout.addWidget(self.password_label)
        password_layout.addWidget(self.password_input)
        manual_credentials_layout.addLayout(password_layout)

        # Enable Checkbox
        self.enable_checkbox = QCheckBox("Enable:", self)
        self.enable_checkbox.setChecked(False)
        self.enable_checkbox.stateChanged.connect(self.toggle_enable_field)

        # Enable
        enable_layout = QHBoxLayout()
        enable_layout.setSpacing(5)
        self.enable_input = QLineEdit(self)
        self.enable_input.setDisabled(True)
        self.enable_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.enable_input.setFixedWidth(280)
        self.enable_input.textChanged.connect(self.validate_fields)
        enable_layout.addWidget(self.enable_checkbox)
        enable_layout.addWidget(self.enable_input)
        manual_credentials_layout.addLayout(enable_layout)

        self.manual_credentials_group.setLayout(manual_credentials_layout)

        # Keyring Credentials section
        self.keyring_credentials_group = QGroupBox()
        keyring_credentials_layout = QVBoxLayout()
        keyring_credentials_layout.setSpacing(5)

        # Keyring System Name
        keyring_system_layout = QHBoxLayout()
        keyring_system_layout.setSpacing(5)
        self.keyring_system_label = QLabel("System Name:", self)
        self.keyring_system_input = QLineEdit(self)
        self.keyring_system_input.setFixedWidth(280)
        self.keyring_system_input.textChanged.connect(self.validate_fields)
        keyring_system_layout.addWidget(self.keyring_system_label)
        keyring_system_layout.addWidget(self.keyring_system_input)
        keyring_credentials_layout.addLayout(keyring_system_layout)

        # Keyring User Name
        keyring_user_layout = QHBoxLayout()
        keyring_user_layout.setSpacing(5)
        self.keyring_user_label = QLabel("Username:", self)
        self.keyring_user_input = QLineEdit(self)
        self.keyring_user_input.setFixedWidth(280)
        self.keyring_user_input.textChanged.connect(self.validate_fields)
        keyring_user_layout.addWidget(self.keyring_user_label)
        keyring_user_layout.addWidget(self.keyring_user_input)
        keyring_credentials_layout.addLayout(keyring_user_layout)

        self.keyring_credentials_group.setLayout(keyring_credentials_layout)

        # Add radio buttons and credential sections to the main credentials layout
        credentials_layout.addWidget(self.manual_radio)
        credentials_layout.addWidget(self.manual_credentials_group)
        credentials_layout.addWidget(self.keyring_radio)
        credentials_layout.addWidget(self.keyring_credentials_group)
        credentials_group.setLayout(credentials_layout)

        # Add the Script Options section
        options_group = QGroupBox("Script Options")
        options_layout = QHBoxLayout()
        options_layout.setSpacing(40)   # Add some spacing between checkboxes

        # Verbose Output checkbox
        self.verbose_checkbox = QCheckBox("Verbose Output", self)
        self.verbose_checkbox.setChecked(False)
        options_layout.addWidget(self.verbose_checkbox)

        # Combined Output File checkbox
        self.combined_output_checkbox = QCheckBox("Combined Output File", self)
        self.combined_output_checkbox.setChecked(False)
        options_layout.addWidget(self.combined_output_checkbox)

        # Add stretch to push checkboxes to the left
        options_layout.addStretch()
        options_group.setLayout(options_layout)

        # Actions section
        actions_group = QGroupBox("Actions")
        actions_layout = QVBoxLayout()
        actions_layout.setSpacing(5)

        # First row of Actions buttons
        top_button_layout = QHBoxLayout()
        top_button_layout.setSpacing(5)
        self.help_button = QPushButton("Help", self)
        self.help_button.clicked.connect(self.show_help)
        self.quit_button = QPushButton("Quit", self)
        self.quit_button.clicked.connect(self.close)
        self.go_button = QPushButton("Go", self)
        self.go_button.clicked.connect(self.on_go)
        self.go_button.setEnabled(False)
        self.update_go_button_style()  # New method to handle button styling
        top_button_layout.addWidget(self.help_button)
        top_button_layout.addWidget(self.quit_button)
        top_button_layout.addWidget(self.go_button)

        # Second row of Actions buttons
        bottom_button_layout = QHBoxLayout()
        bottom_button_layout.setSpacing(5)
        self.keyring_tools_button = QPushButton("Keyring Tools", self)
        self.keyring_tools_button.clicked.connect(self.open_keyring_tools)
        self.show_output_button = QPushButton("Open Output Folder", self)
        self.show_output_button.clicked.connect(self.show_output_folder)
        bottom_button_layout.addWidget(self.keyring_tools_button)
        bottom_button_layout.addWidget(self.show_output_button)

        actions_layout.addLayout(top_button_layout)
        actions_layout.addLayout(bottom_button_layout)
        actions_group.setLayout(actions_layout)

        # Add Output section with scroll bars
        output_group = QGroupBox("Output")
        output_layout = QVBoxLayout()
        output_layout.setSpacing(5)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        output_layout.addWidget(self.progress_bar)

        # Create a scroll area for the log output
        self.log_output = QTextEdit(self)
        self.log_output.setReadOnly(True)
        self.log_output.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.log_output.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)  # Disable text wrapping
        self.log_output.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)  # Enable horizontal scrollbar
        self.log_output.horizontalScrollBar().setValue(0)  # Set horizontal scrollbar to the left initially
        # Connect the textChanged signal to a custom slot to control scrolling behavior
        self.log_output.textChanged.connect(self.keep_horizontal_scroll_left)

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(self.log_output)
        scroll_area.setMinimumHeight(150)
        output_layout.addWidget(scroll_area)
        output_group.setLayout(output_layout)

        # Set up the main layout
        self.main_layout = QVBoxLayout()
        self.main_layout.setSpacing(5)
        self.main_layout.addLayout(logo_title_layout)
        self.main_layout.addSpacing(10)
        self.main_layout.addWidget(input_files_group)
        self.main_layout.addWidget(credentials_group)
        self.main_layout.addWidget(options_group)
        self.main_layout.addWidget(actions_group)
        self.main_layout.addWidget(output_group)
        self.setLayout(self.main_layout)

        # Set default files
        self.set_default_files()

        # Initialize credential options
        self.toggle_credential_options()

        # Reduce margins for the main layout
        self.main_layout.setContentsMargins(15, 5, 15, 15)  # Left, Top, Right, Bottom margins
        input_files_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins for the Input Files layout
        credentials_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins for the Credentials layout
        options_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins for the Options layout
        actions_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins for the Actions layout
        output_layout.setContentsMargins(5, 5, 5, 5)  # Reduce margins for the Output layout

    def check_for_updates(self) -> None:
        """Start background update check."""
        self.version_checker = VersionChecker()
        self.version_checker.update_found.connect(self.on_update_found)
        self.version_checker.check_complete.connect(self.on_check_complete)
        
        # Run in a thread to avoid blocking UI
        self.check_thread = QThread()
        self.version_checker.moveToThread(self.check_thread)
        self.check_thread.started.connect(self.version_checker.check)
        self.check_thread.start()

    def on_update_found(self, new_version: str, release_url: str) -> None:
        """Handle when an update is found."""
        self.update_available = True
        self.new_version = new_version
        self.release_url = release_url
        
        # Show update indicator
        update_text = f'<a href="{release_url}" style="color: #4CAF50; text-decoration: none;">'
        update_text += f'Update available: v{new_version} ↗</a>'
        self.update_indicator.setText(update_text)
        
        # Also log to output
        # self.append_colored_message(f"Update available: v{new_version} (current: v{APP_VERSION.lstrip('v')})", "INFO")

    def on_check_complete(self, success: bool) -> None:
        """Clean up after update check."""
        if hasattr(self, 'check_thread'):
            self.check_thread.quit()
            self.check_thread.wait()
            
        if not success and not self.update_available:
            # Check failed but that's OK - we don't show errors
            pass

    def keep_horizontal_scroll_left(self) -> None:
        """Ensure the horizontal scrollbar stays on the left side when new text is added."""
        self.log_output.horizontalScrollBar().setValue(0)

    def showEvent(self, event: Any) -> None:
        """Override the showEvent to center the window after it is fully laid out."""
        super().showEvent(event)  # Call the base class implementation
        self.center()  # Center the window after it is shown

    def center(self) -> None:
        """Center the window exactly in the middle of the screen."""
        # Get the screen's geometry
        screen_geometry = QApplication.primaryScreen().geometry()

        # Calculate the center position
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2

        # Move the window to the calculated position
        self.move(x, y)

    def set_default_files(self) -> None:
        """Set the default input files in the UI."""
        # Set the default file for Device(s) input file
        if self.default_device_list.exists():
            try:
                relative_path = self.default_device_list.relative_to(Path(__file__).parent)
                self.device_input.setText(f"./{relative_path.as_posix()}")
                self.device_input.setStyleSheet("")
            except ValueError:
                # If the file is not in a subpath of the script's directory, use the absolute path
                self.device_input.setText(str(self.default_device_list))
                self.device_input.setStyleSheet("")
        else:
            self.device_input.setText(f"Default file not found: ./{self.default_device_list.relative_to(Path(__file__).parent).as_posix()}")
            self.device_input.setStyleSheet("color: red;")

        # Set the default file for Command(s) input file
        if self.default_command_list.exists():
            try:
                relative_path = self.default_command_list.relative_to(Path(__file__).parent)
                self.command_input.setText(f"./{relative_path.as_posix()}")
                self.command_input.setStyleSheet("")
            except ValueError:
                # If the file is not in a subpath of the script's directory, use the absolute path
                self.command_input.setText(str(self.default_command_list))
                self.command_input.setStyleSheet("")
        else:
            self.command_input.setText(f"Default file not found: ./{self.default_command_list.relative_to(Path(__file__).parent).as_posix()}")
            self.command_input.setStyleSheet("color: red;")

    def browse_device_list(self) -> None:
        """Open a file dialog to browse for the Device(s) input file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Device List File", str(self.input_folder), "Text Files (*.txt);;All Files (*)"
        )
        if file_path:
            file_path = Path(file_path)
            try:
                relative_path = file_path.relative_to(Path(__file__).parent)
                self.device_input.setText(f"./{relative_path.as_posix()}")
                self.device_input.setStyleSheet("")  # Reset to default color
            except ValueError:
                # If the file is not in a subpath of the script's directory, use the absolute path
                self.device_input.setText(str(file_path))
                self.device_input.setStyleSheet("")  # Reset to default color
            self.validate_fields()

    def browse_command_list(self) -> None:
        """Open a file dialog to browse for the Command(s) input file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Command List File", str(self.input_folder), "Text Files (*.txt);;All Files (*)"
        )
        if file_path:
            file_path = Path(file_path)
            try:
                relative_path = file_path.relative_to(Path(__file__).parent)
                self.command_input.setText(f"./{relative_path.as_posix()}")
                self.command_input.setStyleSheet("")  # Reset to default color
            except ValueError:
                # If the file is not in a subpath of the script's directory, use the absolute path
                self.command_input.setText(str(file_path))
                self.command_input.setStyleSheet("")  # Reset to default color
            self.validate_fields()

    def toggle_enable_field(self) -> None:
        """Enable or disable the Enable field based on the checkbox state."""
        if self.enable_checkbox.isChecked():
            self.enable_input.setDisabled(False)
            self.enable_input.clear()
        else:
            self.enable_input.setDisabled(True)
            self.sync_enable_password()
        self.validate_fields()

    def sync_enable_password(self) -> None:
        """Sync the Enable password with the Manual or Keyring password."""
        if not self.enable_checkbox.isChecked():
            if self.manual_radio.isChecked():
                self.enable_input.setText(self.password_input.text())
            else:
                self.enable_input.setText("")  # Clear if Keyring is selected

    def toggle_credential_options(self) -> None:
        """Enable/disable credential sections based on the selected radio button."""
        if self.manual_radio.isChecked():
            self.manual_credentials_group.setEnabled(True)
            self.keyring_credentials_group.setEnabled(False)
            self.reset_credentials()  # Reset all credentials when switching to Manual
            self.sync_enable_password()  # Sync Enable password when switching to Manual
        else:
            self.manual_credentials_group.setEnabled(False)
            self.keyring_credentials_group.setEnabled(True)
            self.reset_credentials()  # Reset all credentials when switching to Keyring
            self.sync_enable_password()  # Sync Enable password when switching to Keyring
        self.validate_fields()

    def reset_credentials(self) -> None:
        """Reset all credential fields to their default values."""
        self.username_input.clear()
        self.password_input.clear()
        self.enable_input.clear()
        self.keyring_system_input.clear()
        self.keyring_user_input.clear()

    def disable_input_controls(self) -> None:
        """Disable all input controls while the script is running."""
        self.device_button.setEnabled(False)
        self.device_edit_button.setEnabled(False)
        self.command_button.setEnabled(False)
        self.command_edit_button.setEnabled(False)
        # self.go_button.setEnabled(False)  # Remove this line
        self.go_button.setStyleSheet("")  # Reset the button's style to default
        self.quit_button.setEnabled(False)
        self.manual_radio.setEnabled(False)
        self.keyring_radio.setEnabled(False)
        self.username_input.setEnabled(False)
        self.password_input.setEnabled(False)
        
        # Store current enable state before disabling
        self.enable_was_enabled = self.enable_checkbox.isEnabled()
        self.enable_checkbox.setEnabled(False)
        self.enable_input.setEnabled(False)
        
        # Store and disable Options checkboxes
        self.verbose_was_enabled = self.verbose_checkbox.isEnabled()
        self.combined_output_was_enabled = self.combined_output_checkbox.isEnabled()
        self.verbose_checkbox.setEnabled(False)
        self.combined_output_checkbox.setEnabled(False)
        
        self.keyring_system_input.setEnabled(False)
        self.keyring_user_input.setEnabled(False)

    def enable_input_controls(self) -> None:
        """Enable all input controls after the script completes."""
        self.device_button.setEnabled(True)
        self.device_edit_button.setEnabled(True)
        self.command_button.setEnabled(True)
        self.command_edit_button.setEnabled(True)
        self.go_button.setEnabled(True)
        self.quit_button.setEnabled(True)
        self.manual_radio.setEnabled(True)
        self.keyring_radio.setEnabled(True)
        self.username_input.setEnabled(True)
        self.password_input.setEnabled(True)
        
        # Restore enable checkbox and input based on previous state
        if self.enable_was_enabled:
            self.enable_checkbox.setEnabled(True)
            if self.enable_checkbox.isChecked():
                self.enable_input.setEnabled(True)
        
        # Restore Options checkboxes
        if self.verbose_was_enabled:
            self.verbose_checkbox.setEnabled(True)
        if self.combined_output_was_enabled:
            self.combined_output_checkbox.setEnabled(True)
        
        self.keyring_system_input.setEnabled(True)
        self.keyring_user_input.setEnabled(True)
        self.update_go_button_style(is_stop=False)
        self.validate_fields()  # Revalidate fields

    def show_output_folder(self) -> None:
        """Open the output folder using the OS's native file explorer."""
        output_folder = DEFAULT_OUTPUT_FOLDER

        if not output_folder.exists():
            # Log an error message in the Output section
            self.append_colored_message("Output folder does not exist yet.", "ERROR")
            return

        if platform.system() == "Windows":
            os.startfile(output_folder)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", output_folder])
        elif platform.system() == "Linux":
            subprocess.run(["xdg-open", output_folder])
        else:
            # Log an error message in the Output section
            self.append_colored_message("Unsupported operating system.", "ERROR")

    def validate_fields(self) -> None:
        """Validate all required fields and enable/disable the Go button."""
        device_file = self.device_input.text().strip()
        command_file = self.command_input.text().strip()

        if self.manual_radio.isChecked():
            username = self.username_input.text().strip()
            password = self.password_input.text().strip()
            enable = self.enable_input.text().strip()
            is_valid = bool(
                device_file
                and command_file
                and username
                and password
                and (not self.enable_checkbox.isChecked() or enable)
            )
        else:
            keyring_system = self.keyring_system_input.text().strip()
            keyring_user = self.keyring_user_input.text().strip()
            is_valid = bool(
                device_file
                and command_file
                and keyring_system
                and keyring_user
            )

        # Enable/disable the Go button and change its color
        self.go_button.setEnabled(is_valid)
        if is_valid:
            self.go_button.setStyleSheet("")  # Reset to default style
            self.go_button.setStyleSheet(
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
            self.go_button.setStyleSheet("")  # Reset to default style

    def update_go_button_style(self, is_stop: bool = False) -> None:
        """
        Update the Go/Stop button appearance.
        
        Args:
            is_stop: Whether to show the button in "Stop" mode
        """
        if is_stop:
            self.go_button.setText("Stop")
            self.go_button.setStyleSheet("""
                QPushButton {
                    background-color: #d9534f;
                    color: white;
                    border-radius: 5px;
                    padding: 1px;
                    border: 1px solid #d43f3a;
                }
                QPushButton:hover {
                    background-color: #c9302c;
                }
                QPushButton:pressed {
                    background-color: #ac2925;
                }
            """)
        else:
            self.go_button.setText("Go")
            if self.go_button.isEnabled():
                self.go_button.setStyleSheet("""
                    QPushButton {
                        background-color: #5cb85c;
                        color: white;
                        border-radius: 5px;
                        padding: 1px;
                        border: 1px solid #4cae4c;
                    }
                    QPushButton:hover {
                        background-color: #449d44;
                    }
                    QPushButton:pressed {
                        background-color: #398439;
                    }
                """)
            else:
                self.go_button.setStyleSheet("")  # Default disabled style

    def show_help(self) -> None:
        """Create and show the custom help dialog."""
        help_dialog = HelpDialog(APP_NAME, APP_VERSION, VERSION_DATE, self)
        help_dialog.exec()

    def on_go(self) -> None:
        """Handle Go/Stop button click."""
        # Clear the log output when starting a new run
        if not (hasattr(self, 'worker') and self.worker and self.worker.isRunning()):
            self.log_output.clear()
            # self.log_output.setText("")
        
        if hasattr(self, 'worker') and self.worker and self.worker.isRunning():
            # If worker is running, this click means stop
            self.stop_script()
        else:
            # Normal execution
            self.start_script()

    def start_script(self) -> None:
        """Start script execution."""
        # Clear previous worker if exists
        if hasattr(self, 'worker'):
            self.worker.quit()
            self.worker.wait(100)
            del self.worker

        # Reset state
        self.stop_requested = False
        self.progress_bar.setValue(0)

        # Disable controls and update button, Update UI immediately
        self.disable_input_controls()
        self.update_go_button_style(is_stop=True)
        QApplication.processEvents()  # Force UI update

        # Create new output folder and log file for this run
        output_folder = DEFAULT_OUTPUT_FOLDER / datetime.now().strftime("%m-%d-%Y - %I_%M_%p")
        output_folder.mkdir(parents=True, exist_ok=True)
        log_file = str(output_folder / Path(__file__).stem) + ".log"

        # Configure logging for this run
        self.configure_logging(log_file)

        # Get input values
        device_file = self.device_input.text().strip()
        command_file = self.command_input.text().strip()

        if self.manual_radio.isChecked():
            credentials = {
                'username': self.username_input.text().strip(),
                'password': self.password_input.text().strip()
            }
            enable_password = self.enable_input.text().strip()
        else:
            keyring_system = self.keyring_system_input.text().strip()
            keyring_user = self.keyring_user_input.text().strip()
            try:
                password = keyring.get_password(keyring_system, keyring_user)
                if password:
                    credentials = {
                        'username': keyring_user,
                        'password': password
                    }
                    enable_password = password
                else:
                    self.append_colored_message("No credentials found in the keyring.", "ERROR")
                    self.enable_input_controls()  # Re-enable controls on error
                    return
            except Exception as e:
                self.append_colored_message(f"Failed to fetch credentials from the keyring: {e}", "ERROR")
                self.enable_input_controls()  # Re-enable controls on error
                return

        # Ensure the output folder is always a subfolder of the script's directory
        output_folder = DEFAULT_OUTPUT_FOLDER / datetime.now().strftime("%m-%d-%Y - %I_%M_%p")
        output_folder.mkdir(parents=True, exist_ok=True)

        # Create and start worker
        self.worker = Worker(
            device_file=device_file,
            command_file=command_file,
            credentials=credentials,
            enable_password=enable_password,
            output_folder=output_folder,
            verbose_enabled=self.verbose_checkbox.isChecked(),
            create_combined_output=self.combined_output_checkbox.isChecked()
        )
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.log_signal.connect(self.update_log)
        self.worker.completion_signal.connect(self.on_script_complete)
        self.worker.start()

    def update_progress(self, value: int) -> None:
        """Update the progress bar with the current value."""
        self.progress_bar.setValue(value)

    def append_colored_message(self, message: str, level_name: str = "INFO") -> None:
        """
        Append a colored message to the log output.
        
        Args:
            message: The message text to append
            level_name: Log level name (e.g., "INFO", "ERROR")
        """
        # Only use this for non-logger GUI messages
        cursor = self.log_output.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        
        color = self.get_color_for_level(level_name)
        formatted = self.format_message(message, level_name)
        
        cursor.insertHtml(f'<span style="color:{color}">{formatted}</span><br>')
        self.log_output.setTextCursor(cursor)
        self.log_output.ensureCursorVisible()

    def get_color_for_level(self, level_name: str) -> str:
        """
        Get the color associated with a log level.
        
        Args:
            level_name: Log level name
            
        Returns:
            Color string for the log level
        """
        color_map = {
            "DEBUG": "blue",
            "VERBOSE": "green", 
            "INFO": "",
            "WARNING": "orange",
            "ERROR": "red",
            "CRITICAL": "darkred"
        }
        return color_map.get(level_name, "")

    def format_message(self, message: str, level_name: str) -> str:
        """
        Format a message for display with proper HTML escaping.
        
        Args:
            message: The message text
            level_name: Log level name
            
        Returns:
            Formatted HTML string
        """
        # Convert message to HTML with proper line breaks and spacing
        message = (message
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\n", "<br>")
            .replace(" ", "&nbsp;")
            .replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;"))
        
        if level_name != "INFO":
            return f"{level_name}: {message}"
        return message

    def update_log(self, message: str) -> None:
        """Append a colored message to the log output."""
        self.append_colored_message(message)
        self.log_output.ensureCursorVisible()  # Scroll to the bottom
        self.keep_horizontal_scroll_left()  # Reset horizontal scrollbar to the left

    def edit_device_file(self) -> None:
        """Open the device file for editing."""
        file_path = self.device_input.text().strip()
        if file_path:
            file_path = Path(file_path)
            if file_path.exists() and file_path.is_file():
                self.show_file_editor(file_path, "device")  # Pass "device" as the file type
            else:
                # Log an error message in the Output section
                self.append_colored_message("The selected file does not exist or is not a valid file.", "ERROR")
        else:
            # Log an error message in the Output section
            self.append_colored_message("No file selected.", "ERROR")

    def edit_command_file(self) -> None:
        """Open the command file for editing."""
        file_path = self.command_input.text().strip()
        if file_path:
            file_path = Path(file_path)
            if file_path.exists() and file_path.is_file():
                self.show_file_editor(file_path, "command")  # Pass "command" as the file type
            else:
                # Log an error message in the Output section
                self.append_colored_message("The selected file does not exist or is not a valid file.", "ERROR")
        else:
            # Log an error message in the Output section
            self.append_colored_message("No file selected.", "ERROR")

    def show_file_editor(self, file_path: Path, file_type: str) -> None:
        """Create and show the file editor dialog."""
        # Create and show the file editor dialog
        self.file_editor_dialog = FileEditorDialog(file_path, file_type, self)
        # Connect the file_path_updated signal to update the input file path
        self.file_editor_dialog.file_path_updated.connect(self.update_input_file)
        self.file_editor_dialog.exec()

    def update_input_file(self, new_file_path: Path, file_type: str) -> None:
        """
        Update the input file path in the main window based on the file type.
        
        Args:
            new_file_path: New path to the file
            file_type: Type of file ('device' or 'command')
        """
        try:
            relative_path = new_file_path.relative_to(Path(__file__).parent)
            display_path = f"./{relative_path.as_posix()}"
        except ValueError:
            display_path = str(new_file_path)
        
        if file_type == "device":
            self.device_input.setText(display_path)
            self.device_input.setStyleSheet("")  # Reset to default color
        elif file_type == "command":
            self.command_input.setText(display_path)
            self.command_input.setStyleSheet("")  # Reset to default color
        self.validate_fields()  # Revalidate fields in case the file path changed

    def open_keyring_tools(self) -> None:
        """Open the Keyring Tools utility in a separate process."""
        # Define the path to the Keyring Tools script
        keyring_tools_path = SCRIPT_DIR / KEYRING_TOOLS_SCRIPT

        # Check if the Keyring Tools script exists
        if not keyring_tools_path.exists():
            # If the script is not found, log an error message in the Output section
            self.append_colored_message(f"'{KEYRING_TOOLS_SCRIPT}' not found.\n- Please ensure it is in the same folder as this script.", "ERROR")
            return

        try:
            # Import the KeyringApp from the external script
            keyring_tools_path = SCRIPT_DIR / KEYRING_TOOLS_SCRIPT
            if platform.system() == "Windows":
                subprocess.Popen(
                    [sys.executable, str(keyring_tools_path)],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
            else:
                subprocess.Popen([sys.executable, str(keyring_tools_path)])
        except Exception as e:
            # If there's an error during import or execution, log it in the Output section
            self.append_colored_message(f"ERROR: Failed to open Keyring Tools - {e}", "ERROR")

    def stop_script(self) -> None:
        """Stop script execution."""
        self.stop_requested = True
        self.logger.warning("Stopping script execution...")  # Will appear once in both places
        self.go_button.setEnabled(False)  # Disable while stopping
        
        if hasattr(self, 'worker'):
            self.worker.cancel()  # Signal the worker to stop

    def on_script_complete(self, output_folder: str) -> None:
        """
        Handle script completion.
        
        Args:
            output_folder: Path to the output folder
        """
        if self.stop_requested:
            self.logger.warning("Script execution was cancelled - collected output may be incomplete!")
        
        # Rest of your existing completion handling
        self.enable_input_controls()  # Re-enable all input controls
        self.update_go_button_style(is_stop=False)
        
def check_cli_updates() -> Optional[Dict[str, str]]:
    """Check for updates in CLI mode using GitHub API."""
    try:
        headers = {
            'User-Agent': f'{APP_NAME}/{APP_VERSION}',
            'Accept': 'application/vnd.github.v3+json'
        }
        req = urllib.request.Request(GITHUB_API_LATEST_RELEASE, headers=headers)
        
        with urllib.request.urlopen(req, timeout=3) as response:
            data = json.loads(response.read().decode())
            latest_tag = data.get('tag_name', '')  # e.g., "v1.1.0"
            current_version = APP_VERSION  # e.g., "v1.0.0"
            
            # Remove 'v' prefix for comparison
            latest_version_str = latest_tag.lstrip('v')
            current_version_str = current_version.lstrip('v')
            
            # Compare versions
            latest_ver = version.parse(latest_version_str)
            current_ver = version.parse(current_version_str)
            
            if latest_ver > current_ver:
                return {
                    'latest_version': latest_version_str,
                    'release_url': data.get('html_url', REPO_URL),
                    'download_url': data.get('zipball_url', ''),
                    'body': data.get('body', '')  # Release notes
                }
    except Exception:
        pass  # Silently fail
    return None

def parse_args() -> argparse.Namespace:
    """
    Parse command line arguments.
    
    Returns:
        Namespace with parsed arguments
    """
    # Custom formatter to show our version format
    class CustomHelpFormatter(argparse.HelpFormatter):
        def add_usage(self, usage, actions, groups, prefix=None):
            if prefix is None:
                prefix = 'Usage: '
            return super().add_usage(usage, actions, groups, prefix)
    
    parser = argparse.ArgumentParser(
        description=f"Network Device Documentation Utility",
        formatter_class=CustomHelpFormatter,
        add_help=False  # We'll add help manually to control order
    )
    
    # Add arguments in logical order
    parser.add_argument("-h", "--help", action="store_true",
                       help=f"Show this help message and exit")
    parser.add_argument("-v", "--version", action="store_true",
                       help=f"Show version information and check for updates")
    parser.add_argument("-cli", "--cli-mode", action="store_true", 
                       help=f"Run in CLI mode (required for CLI mode)")
    parser.add_argument("-d", "--device-file", type=str, 
                       help=f"Path to the device list file (default: %(default)s)", 
                       default=str(DEFAULT_DEVICE_FILE))
    parser.add_argument("-c", "--command-file", type=str, 
                       help=f"Path to the command list file (default: %(default)s)", 
                       default=str(DEFAULT_COMMAND_FILE))
    parser.add_argument("-ks", "--keyring-system", type=str, 
                       help=f"Keyring system name for keyring credentials (requires -ku)")
    parser.add_argument("-ku", "--keyring-user", type=str, 
                       help=f"Keyring user name for keyring credentials (requires -ks)")
    parser.add_argument("--verbose", action="store_true", 
                       help=f"Enable verbose output (default: %(default)s)", default=False)
    parser.add_argument("--combined", action="store_true", 
                       help=f"Enable creation of combined output file (default: %(default)s)", 
                       default=False)
    
    return parser.parse_args()

def run_cli() -> None:
    """Run the script in command line interface mode."""
    args = parse_args()
    
    # Handle help first
    if args.help:
        print(f"nddu - Network Device Documentation Utility")
        print(f"Version: {APP_VERSION} {VERSION_DATE}")
        print()
        print(f"Usage: python nddu.py [options]")
        print()
        print(f"Options:")
        print(f"  -h, --help            Show this help message and exit")
        print(f"  -v, --version         Show version information and check for updates")
        print(f"  -cli, --cli-mode      Run in CLI mode (required for CLI mode)")
        print(f"  -d, --device-file     Path to the device list file (default: ./input/Devices.txt)")
        print(f"  -c, --command-file    Path to the command list file (default: ./input/Commands.txt)")
        print(f"  -ks, --keyring-system Keyring system name for keyring credentials (requires -ku)")
        print(f"  -ku, --keyring-user   Keyring user name for keyring credentials (requires -ks)")
        print(f"  --verbose             Enable verbose output")
        print(f"  --combined            Enable creation of combined output file")
        print()
        print(f"Examples:")
        print(f"  python nddu.py                    # Launch GUI")
        print(f"  python nddu.py -cli -v            # Show version and check updates")
        print(f"  python nddu.py -cli -d devices.txt -c commands.txt")
        print()
        
        # Check for updates silently when showing help
        update_info = check_cli_updates()
        if update_info:
            current = APP_VERSION.lstrip('v')
            latest = update_info.get('latest_version', '').lstrip('v')
            print(f"Note: Update available! v{latest} (current: v{current})")
            print(f"      Run with -v to see update details.")
        
        sys.exit(0)
    
    # Handle version/update check
    if args.version:
        print(f"nddu {APP_VERSION} {VERSION_DATE}")
        
        update_info = check_cli_updates()
        if update_info:
            current = APP_VERSION.lstrip('v')
            latest = update_info.get('latest_version', '').lstrip('v')
            release_url = update_info.get('release_url', REPO_URL)
            
            print(f"\n{DIVIDER}")
            print(f"Update available!")
            print(f"Current version: v{current}")
            print(f"Latest version:  v{latest}")
            print(f"Release URL: {release_url}")
            
            # Show release notes
            release_notes = update_info.get('body', '').strip()
            if release_notes:
                print(f"\nRelease notes:")
                lines = release_notes.split('\n')
                for line in lines[:5]:  # Show first 5 lines
                    if line.strip():
                        print(f"  {line[:120]}{'...' if len(line) > 120 else ''}")
                if len(lines) > 5:
                    print(f"  ... (see full release notes at {release_url})")

            print(f"{DIVIDER}\n")
        else:
            print(f"No updates found.")
        
        sys.exit(0)
    
    # If we get here, it's a normal CLI run
    # Show update notification at start of normal run
    update_info = check_cli_updates()
    if update_info:
        current = APP_VERSION.lstrip('v')
        latest = update_info.get('latest_version', '').lstrip('v')
        print(f"{DIVIDER}")
        print(f"[UPDATE] New version available: v{latest} (current: v{current})")
        print(f"[UPDATE] Release: {update_info.get('release_url', REPO_URL)}")
        print(f"{DIVIDER}")
    
    # Validate CLI mode flag
    if not args.cli_mode:
        print(f"ERROR: CLI mode requires the -cli argument.")
        print(f"Use -h for help.")
        sys.exit(1)
    
    # Validate -ks and -ku: they must be used together
    if (args.keyring_system and not args.keyring_user) or (args.keyring_user and not args.keyring_system):
        print("ERROR: -ks and -ku must be used together.")
        sys.exit(1)

    # Set default input folder
    input_folder = DEFAULT_INPUT_FOLDER

    # Handle -d and -c arguments using config variables
    if args.device_file and args.command_file:
        # Use both specified files
        device_file = Path(args.device_file)
        command_file = Path(args.command_file)
    elif args.device_file:
        # Use specified device file and default command file
        device_file = Path(args.device_file)
        command_file = DEFAULT_COMMAND_FILE
    elif args.command_file:
        # Use specified command file and default device file
        command_file = Path(args.command_file)
        device_file = DEFAULT_DEVICE_FILE
    else:
        # Use both default files
        device_file = DEFAULT_DEVICE_FILE
        command_file = DEFAULT_COMMAND_FILE

    # Validate input files
    if not device_file.exists():
        print(f"ERROR: Device file not found: {device_file}")
        sys.exit(1)
    if not command_file.exists():
        print(f"ERROR: Command file not found: {command_file}")
        sys.exit(1)

    # Define and create Output folder
    output_folder = DEFAULT_OUTPUT_FOLDER / datetime.now().strftime("%m-%d-%Y - %I_%M_%p")
    output_folder.mkdir(parents=True, exist_ok=True)
    log_file = str(output_folder / Path(__file__).stem) + ".log"

    logger = UnifiedLogger.configure_logging(
        verbose=args.verbose,
        log_file=log_file
    )

    # Read and validate devices
    worker = Worker(device_file, command_file, {}, None, output_folder, args.verbose)

    # Get credentials
    if args.keyring_system and args.keyring_user:
        # Use Keyring credentials
        try:
            password = keyring.get_password(args.keyring_system, args.keyring_user)
            if password:
                credentials = {
                    'username': args.keyring_user,
                    'password': password
                }
                enable_password = password
                logger.verbose("Using Keyring credentials.")
            else:
                logger.error(f'No credentials found in the keyring "{args.keyring_system}"')
                sys.exit(1)
        except Exception as e:
            logger.error(f"Failed to fetch credentials from the keyring: {e}")
            sys.exit(1)
    else:
        # Prompt for manual credentials
        logger.verbose(f"No Keyring options provided.")
        logger.info(f"Please enter credentials:")
        while True:
            username = input("Username: ").strip()
            if username:
                break
            logger.error("Username cannot be blank.")

        while True:
            password = getpass.getpass("Password: ").strip()
            if password:
                break
            logger.error("Password cannot be blank.")

        enable_password = getpass.getpass("Enable Password (leave blank to use the same as Password): ").strip()
        if not enable_password:
            enable_password = password

        credentials = {
            'username': username,
            'password': password
        }
        logger.verbose(f"Using manual credentials.")

    # Run the worker
    worker = Worker(device_file, command_file, credentials, enable_password, output_folder, args.verbose, args.combined)
    worker.run()

# --- Main Execution ---
if __name__ == "__main__":
    # Quick check for version/help flags (don't parse fully yet)
    if len(sys.argv) == 1:
        # No arguments, launch GUI
        app = QApplication(sys.argv)
        window = MyWindow()
        window.show()
        sys.exit(app.exec())
    elif '-h' in sys.argv or '--help' in sys.argv:
        # Show help immediately
        run_cli()  # This will exit after showing help
    elif '-v' in sys.argv or '--version' in sys.argv:
        # Show version immediately
        run_cli()  # This will exit after showing version
    else:
        # Parse all args for normal execution
        args = parse_args()
        run_cli()
