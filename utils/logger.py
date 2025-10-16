"""
utils/logger.py

Purpose:
---------
Centralized logging utility for the entire project.

Responsibilities:
-----------------
- Configure consistent logging output across all modules.
- Create a log file ('logs/activity.log') and a console stream simultaneously.
- Format messages with timestamps, log levels, and module names.
- Ensure idempotency: multiple calls to `setup_logger()` return the same logger.

Example:
---------
    from utils.logger import setup_logger
    logger = setup_logger("excel_handler")
    logger.info("Excel file processed successfully.")
"""

import logging          # Standard library for configurable logging
import os               # For directory and path handling
from typing import Optional  # For optional type hinting

# Constants for directory and file path setup
LOG_DIR = "logs"
LOG_FILE = os.path.join(LOG_DIR, "activity.log")


def setup_logger(name: Optional[str] = None) -> logging.Logger:
    """
    Configure and return a logger object for the given module name.

    Args:
        name (str | None): Optional name for the logger (usually the module name).

    Returns:
        logging.Logger: A fully configured logger instance.

    Design Notes:
    -------------
    - Idempotent: If the logger already has handlers, it won’t add duplicates.
    - Dual output: Logs go to both a file ('logs/activity.log') and the console.
    - Log format includes timestamp, log level, logger name, and message.
    """
    # Ensure logs directory exists
    os.makedirs(LOG_DIR, exist_ok=True)

    # Create or retrieve a logger with the given name
    logger = logging.getLogger(name)

    # Prevent duplicate handlers if setup_logger() is called multiple times
    if logger.handlers:
        return logger

    # Set default logging level
    logger.setLevel(logging.INFO)

    # Define log message format:
    # Example: [2025-10-16 10:24:05] [INFO] [excel_handler] - Excel file loaded
    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] [%(name)s] - %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )

    # ------------------------------
    # 1️⃣ FILE HANDLER CONFIGURATION
    # ------------------------------
    # Writes logs to logs/activity.log
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.INFO)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # ------------------------------
    # 2️⃣ CONSOLE (STREAM) HANDLER
    # ------------------------------
    # Prints logs to terminal output
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    # Return the configured logger instance
    return logger
