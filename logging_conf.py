"""
logging_conf.py
=================

This module defines a helper function to configure logging for the application.
It writes logs both to the console and to rotating log files located in the
application's log directory. The log files rotate at 10MB and keep up to 5
backups. The configuration can be invoked from the application entry point
to ensure consistent logging across modules.
"""

from __future__ import annotations

import logging
import logging.handlers
from pathlib import Path

from constants import APP_NAME
# Import paths module. When executed as a script, relative import will fail. Use
# dynamic loading as a fallback.
try:
    from . import paths  # type: ignore
except ImportError:
    import importlib.util
    import os
    _paths_spec = importlib.util.spec_from_file_location(
        "paths", os.path.join(os.path.dirname(__file__), "paths.py")
    )
    assert _paths_spec and _paths_spec.loader
    paths = importlib.util.module_from_spec(_paths_spec)
    _paths_spec.loader.exec_module(paths)  # type: ignore


def setup_logging(level: int = logging.INFO) -> None:
    """Configure logging for the application.

    This function sets up two handlers:

    - A RotatingFileHandler that writes to ``<log_dir>/app.log``, rotating at
      10MB with up to 5 backups.
    - A StreamHandler that writes to stderr (console).

    Args:
        level: Logging level for the root logger. Defaults to INFO.
    """
    log_dir: Path = paths.get_log_dir()
    log_file = log_dir / "app.log"
    # Define formatters
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    # Rotating file handler
    file_handler = logging.handlers.RotatingFileHandler(
        log_file, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    # Clear existing handlers to avoid duplicate logs when reconfiguring
    for h in list(root_logger.handlers):
        root_logger.removeHandler(h)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)
    # Reduce noise from third-party libraries
    logging.getLogger("google").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
