# Repository Guidelines

## Project Structure & Module Organization
- Core entry: `main.py` (CLI/GUI), `app.py` (CustomTkinter shell), `agent.py` (clipboard + LLM agent, tray, hotkeys, queue).
- Models/Helpers: `common_models.py`, `config_manager.py`.
- UI: `ui_components.py`, `styles.py`.
- Batch: `matrix_batch_processor.py`.
- Assets in repo root: `icon.ico`, `completion.mp3`, `delete_icon.png`, `api_price.json`.
- Config: `config.json` is created on first run (user-specific; not committed).

## Build, Test, and Development Commands
- Create env: `python -m venv .venv && source .venv/bin/activate` (Windows: `\.venv\\Scripts\\activate`).
- Install deps: `pip install -r requirements.txt`.
- Run app: `python main.py` (initializes `config.json`).
- Set API key: use in-app Settings (stored via OS `keyring`).
- Batch runs: `python matrix_batch_processor.py` for matrix validations.

## Coding Style & Naming Conventions
- Python 3.10+, 4-space indentation, PEP 8 with type hints.
- Files/functions: `snake_case`; classes: `CapWords`.
- Prompt IDs: lower_snake_case derived from names (see `app.py`).
- Centralize UI text/colors in `styles.py`; keep assets small and cross‑platform.

## Testing Guidelines
- Framework: `pytest` (optional, preferred for pure helpers).
- Naming: place tests as `test_*.py` near targets or in root.
- Run: `pytest` from repo root.
- Focus: `config_manager.py` and pure helpers; manual checks for GUI, tray, hotkeys, first‑run setup, clipboard flows, and batch processor.

## Commit & Pull Request Guidelines
- Commits: small, focused, imperative (e.g., "Add prompt editor validation").
- Branches: `feature/...`, `fix/...`, `chore/...`.
- PRs: include summary, rationale, scope, screenshots/GIFs for UI changes, steps to reproduce, test notes, and any config/migration notes.

## Security & Configuration Tips
- Never commit secrets. API keys live in `keyring`.
- Do not commit user-specific `config.json`.
- Validate prompts before broad changes; use `matrix_batch_processor.py` for safe comparisons.

## Architecture Overview & Agent Tips
- The agent manages clipboard polling, LLM calls, tray menu, hotkeys, and a job queue.
- Performance: avoid blocking the UI thread; route long work through the agent’s worker loop.

## Internationalization (i18n)
- Use `from i18n import tr` and call `tr("key", **kwargs)` for all user-facing strings.
- Language files live under `locales/<lang>.json`. Maintain keys across languages; English (`en.json`) is the canonical baseline.
- Default resolution order is current language → English → the key itself.
- When adding UI text, add corresponding keys to `locales/en.json` and (ideally) `locales/ja.json`.
