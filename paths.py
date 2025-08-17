"""
paths.py
==================

This module centralizes determination of application-specific paths for configuration,
data and logs. It uses sensible defaults on Windows, macOS and Linux to place
application data in user-specific directories instead of the working directory.

The functions defined here ensure that directories exist before returning
paths. They rely on the APP_NAME constant from the ``constants`` module.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional

from constants import APP_NAME


def _get_windows_base_dir() -> Path:
    """Return the base directory for storing application data on Windows.

    Uses the %APPDATA% environment variable if it exists, falling back to the
    user's home directory if not.
    """
    appdata = os.environ.get("APPDATA")
    if appdata:
        return Path(appdata) / APP_NAME
    # fall back to HOME directory
    return Path.home() / APP_NAME


def _get_macos_base_dir() -> Path:
    """Return the base directory for storing application data on macOS.

    On macOS, use ``~/Library/Application Support/<APP_NAME>``.
    """
    return Path.home() / "Library" / "Application Support" / APP_NAME


def _get_linux_base_dir() -> Path:
    """Return the base directory for storing application data on Linux.

    Respects the XDG Base Directory Specification by using $XDG_CONFIG_HOME
    if set, otherwise falls back to ``~/.config/<APP_NAME>``.
    """
    xdg_config_home = os.environ.get("XDG_CONFIG_HOME")
    if xdg_config_home:
        return Path(xdg_config_home) / APP_NAME
    return Path.home() / ".config" / APP_NAME


def get_base_dir() -> Path:
    """Return the base directory for storing application-specific data.

    The directory is created if it does not already exist. On Windows this
    corresponds to %APPDATA%\<APP_NAME>, on macOS to
    ~/Library/Application Support/<APP_NAME>, and on Linux to
    ~/.config/<APP_NAME> (or $XDG_CONFIG_HOME/<APP_NAME> if set).
    """
    if sys.platform.startswith("win"):
        base_dir = _get_windows_base_dir()
    elif sys.platform == "darwin":
        base_dir = _get_macos_base_dir()
    else:
        base_dir = _get_linux_base_dir()
    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir


def get_config_file_path(filename: Optional[str] = None) -> Path:
    """Return the path to the configuration file.

    If *filename* is provided, it is used as the filename; otherwise the
    default ``config.json`` is used.
    """
    base_dir = get_base_dir()
    if filename is None:
        filename = "config.json"
    return base_dir / filename


def get_log_dir() -> Path:
    """Return the directory for log files.

    The directory is created if it does not already exist.
    """
    base_dir = get_base_dir()
    log_dir = base_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir


def get_data_dir() -> Path:
    """Return the directory for miscellaneous data (e.g. history, attachments).

    The directory is created if it does not already exist.
    """
    base_dir = get_base_dir()
    data_dir = base_dir / "data"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir
