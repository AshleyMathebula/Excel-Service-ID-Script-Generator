"""
utils/file_writer.py

Purpose
-------
Provides a lightweight utility class for safely writing text files to disk.
Ensures that the specified output directory exists before writing any file.

Responsibilities
----------------
- Automatically create an output directory (default: 'output') if it doesn’t exist.
- Write iterable lines of text to UTF-8 encoded files.
- Return the final absolute file path for logging, validation, or downstream use.
"""

from pathlib import Path
from typing import Iterable

from utils.logger import setup_logger

# ---------------------------------------------------------------------------
# Module-Level Logger
# ---------------------------------------------------------------------------
# Create a dedicated logger instance for this module to capture file operations
logger = setup_logger("file_writer")


class FileWriter:
    """
    Encapsulates safe file-writing operations and output directory management.
    """

    def __init__(self, output_dir: str = "output"):
        """
        Initialize a FileWriter instance.

        Args:
            output_dir (str, optional):
                Directory where files will be written. Defaults to 'output'.

        Notes:
            - The directory and all parent directories will be created automatically.
            - Logs the resolved directory path for transparency.
        """
        # Store and prepare the output directory as a Path object
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        logger.info("Output directory initialized: %s", self.output_dir.resolve())

    # -----------------------------------------------------------------------

    def write_text_file(self, path: Path, lines: Iterable[str]) -> Path:
        """
        Write a list or iterable of text lines to a UTF-8 encoded file.

        Args:
            path (Path | str):
                Destination file path. Can be absolute or relative to the output directory.
            lines (Iterable[str]):
                Iterable containing strings to write (e.g., list[str] or generator).

        Returns:
            Path: The absolute path of the successfully written file.

        Workflow:
            1. Normalize the file path (convert str → Path if necessary).
            2. If path is relative, resolve it inside the configured output directory.
            3. Ensure parent directories exist.
            4. Join text lines with newline characters and append a final newline.
            5. Write the file to disk with UTF-8 encoding.
            6. Log success or record any exceptions raised.

        Raises:
            Exception: Propagates any file I/O or OS-level errors after logging them.
        """
        # Step 1: Normalize input path
        if not isinstance(path, Path):
            path = Path(path)

        # Step 2: Ensure relative paths are anchored to output directory
        if not path.is_absolute():
            path = self.output_dir / path

        # Step 3: Guarantee that parent directories exist
        path.parent.mkdir(parents=True, exist_ok=True)

        # Step 4: Prepare text content
        # Add a trailing newline for proper formatting in text editors
        content = "\n".join(lines) + ("\n" if lines else "")

        try:
            # Step 5: Write to disk using UTF-8 encoding
            path.write_text(content, encoding="utf-8")
            logger.info("Successfully wrote file: %s", path.resolve())
        except Exception as e:
            # Step 6: Capture and re-raise exceptions with traceback
            logger.exception("Failed to write file %s: %s", path, e)
            raise

        # Step 7: Return absolute file path for caller reference
        return path.resolve()
