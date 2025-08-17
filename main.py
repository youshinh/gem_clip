# main.py
from app import ClipboardToolApp
# Import logging_conf dynamically to support running as a script (no package context)
try:
    # When run as a package (python -m agent.main) logging_conf will be found as a sibling module
    from .logging_conf import setup_logging  # type: ignore[attr-defined]
except Exception:
    # Fallback: load logging_conf from this directory
    import importlib.util
    import os
    _log_conf_spec = importlib.util.spec_from_file_location(
        "logging_conf", os.path.join(os.path.dirname(__file__), "logging_conf.py")
    )
    assert _log_conf_spec and _log_conf_spec.loader
    _logging_conf_module = importlib.util.module_from_spec(_log_conf_spec)
    _log_conf_spec.loader.exec_module(_logging_conf_module)  # type: ignore
    setup_logging = _logging_conf_module.setup_logging

if __name__ == "__main__":
    # Initialize logging before starting the application. This ensures that
    # messages emitted during app startup are captured.
    setup_logging()
    app_instance = ClipboardToolApp()
    app_instance.run()
