"""
utils/file_writer.py

Purpose:
---------
Small utility class that handles writing text files to disk, ensuring that
the output directory exists before writing.

Responsibilities:
-----------------
- Automatically create an 'output' directory if it doesn't exist.
- Write a list of string lines to a text file (UTF-8 encoded).
- Return the final Path object for downstream use (e.g., logging, validation).
"""

from pathlib import Path       # Built-in library for object-oriented filesystem paths
from typing import Iterable    # For type hinting collections like list[str]

from utils.logger import setup_logger  # Custom logging setup utility

# Initialize a dedicated logger for this module
logger = setup_logger("file_writer")


class FileWriter:
    """Encapsulates file writing operations and directory management."""

    def __init__(self, output_dir: str = "output"):
        """
        Constructor — initializes a FileWriter instance.

        Args:
            output_dir (str): Directory where files will be written (default = 'output').

        Responsibilities:
        -----------------
        - Store output directory as a Path object.
        - Create the directory (and parent directories) if they do not exist.
        - Log the resolved absolute path for transparency.
        """
        self.output_dir = Path(output_dir)

        # Create the directory if missing (parents=True handles nested paths)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        logger.info("Output directory: %s", self.output_dir.resolve())

    def write_text_file(self, path: Path, lines: Iterable[str]) -> Path:
        """
        Write the provided lines into a UTF-8 encoded text file.

        Args:
            path (Path | str): Destination file path (can be absolute or relative).
            lines (Iterable[str]): Iterable of strings representing lines of text.

        Returns:
            Path: The absolute path of the file successfully written.

        Workflow:
        ----------
        1. Convert path to a Path object (if it’s a string).
        2. If the path is relative, place it inside the output directory.
        3. Ensure the parent directories exist before writing.
        4. Join lines with newline characters and ensure file ends with a newline.
        5. Write to disk using UTF-8 encoding.
        6. Log success or catch and log any exceptions.
        """
        # Normalize path argument to Path type
        if not isinstance(path, Path):
            path = Path(path)

        # If user provided a relative path, prepend the output directory
        if not path.is_absolute():
            path = self.output_dir / path

        # Make sure parent directories exist
        path.parent.mkdir(parents=True, exist_ok=True)

        # Join all text lines with newline characters
        # Ensures there's a trailing newline at the end of the file
        content = "\n".join(lines) + ("\n" if lines else "")

        try:
            # Write file using UTF-8 encoding (safe for text data)
            path.write_text(content, encoding="utf-8")
            logger.info("Wrote file: %s", path.resolve())
        except Exception as e:
            # Log full traceback and re-raise exception
            logger.exception("Failed to write file %s: %s", path, e)
            raise

        # Return fully resolved (absolute) file path
        return path.resolve()
