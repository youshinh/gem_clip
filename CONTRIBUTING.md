# Contributing

Thanks for your interest in contributing! This project welcomes issues and pull
requests. Please read the guidelines below to help keep changes easy to review
and maintain.

## Getting Started
- Python 3.10+
- Create venv and install deps:
  ```bash
  python -m venv .venv
  source .venv/bin/activate  # Windows: .venv\Scripts\activate
  pip install -r requirements.txt
  pip install -r requirements-dev.txt  # optional
  ```
- Run app: `python main.py`

## Branch & Commit Style
- Branch names: `feature/...`, `fix/...`, `chore/...`
- Commits: small, imperative (e.g., "Add i18n for matrix previews")
- Submit focused PRs with:
  - Summary/rationale/scope
  - Screenshots/GIFs for UI changes
  - Testing steps
  - Any config/migration notes

## Code Style
- PEP 8 with type hints (mypy-friendly)
- Use `tr("key")` for all user-facing strings (see `locales/`)
- Avoid blocking the UI thread; long work should run in the agent worker loop

## Tests
- Prefer tests for pure helpers (e.g., config/pure utilities)
- GUI is typically validated manually (screenshots/notes welcomed)
- Optional tools: `pytest`, `pytest-cov`, `ruff`, `mypy`

## i18n (Internationalization)
- Language files live under `locales/<lang>.json`
- English (`locales/en.json`) is the baseline; add new keys there first
- Reference strings via `from i18n import tr` and `tr("key", **kwargs)`
- Keep keys consistent and avoid duplications

## Security & Privacy
- Never commit secrets. API keys are stored in OS keyring
- Do not commit `config.json` or user data (see `.gitignore.example`)
- Review `PRIVACY.md` and `SECURITY.md` for guidelines

## Licensing
- Project license is MIT (see `LICENSE`)
- Include appropriate notices for third-party code (see `THIRD_PARTY_LICENSES.md`)

## Windows-First Note
- The app is currently tested on Windows only; macOS/Linux are unverified
- CI attempts syntax checks; platform quirks may require adjustments

