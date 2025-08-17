# Distribution & Packaging Guide

This document provides practical tips for distributing this application as
source or as binaries.

## Source Distribution (GitHub)
- Include: `LICENSE`, `THIRD_PARTY_LICENSES.md`, `README*`, `PRIVACY.md`,
  `SECURITY.md`, `SUPPORT.md`.
- Exclude user‑specific files: `.gitignore` (use `.gitignore.example`),
  `config.json`, `prompt_set/session.json`, `prompt_set/sessions/`.
- Ensure screenshots in `@img/` are cleared for redistribution.

## Binary Distribution (PyInstaller or similar)
- Bundle third‑party license files and notices. At minimum include:
  - `LICENSE`, `THIRD_PARTY_LICENSES.md` (or `NOTICE`)
- Provide an “About” or “Settings” link to these documents, or include them
  beside the executable.
- Ensure OS keyring backend is available for the target platform.

## Platform Notes
- Windows: Global hotkeys use Windows API where possible. On some systems,
  administrator privileges are not required but certain key combos may be
  reserved. Tray functionality relies on system tray availability.
- macOS: Tray icon support depends on menu‑bar availability. Tk on macOS should
  be installed via the Python.org build or a distribution that includes Tk.
- Linux: Tray icons may require libappindicator or compatible desktop support.
  Under Wayland, behavior can differ by compositor.

> Tested OS Note: As of now, the application has only been functionally
> validated on Windows. macOS/Linux are unverified; community reports and PRs
> are welcome.

## Versioning & Releases
- Tag releases with semantic versions (e.g., v1.2.0).
- Provide release notes highlighting features, bugfixes, and breaking changes.
- Consider attaching platform‑specific builds in GitHub Releases.

## Branding/Trademarks
- This application is unofficial and not affiliated with Google. Avoid using
  Google logos or marks unless permitted. Include the disclaimer from README.

## Data & Privacy
- Review `PRIVACY.md`. If distributing to end users, ensure the documentation
  explains what is sent to the LLM service and how to disable it.
