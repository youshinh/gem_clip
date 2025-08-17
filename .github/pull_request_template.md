# Pull Request

## Summary
- What does this change do and why?

## Linked Issues
- Closes #

## Scope
- Affected modules/files:
  - `main.py`, `app.py`, `agent.py`
  - `common_models.py`, `config_manager.py`
  - `ui_components.py`, `styles.py`
  - `matrix_batch_processor.py`

## Screenshots / GIFs (UI)
- If UI-related, include before/after.

## i18n Notes
- If user-facing text changed, include locale updates in `locales/*`.

## How to Test
- Create env: `python -m venv .venv && source .venv/bin/activate`
- Install: `pip install -r requirements.txt`
- Run app: `python main.py`
- Optional tests: `pytest`
- Optional batch: `python matrix_batch_processor.py`

## Risk / Impact
- Notes on UX, performance, or compatibility.

## Config / Migration Notes
- `config.json` is created on first run (do not commit).
- API keys are stored via OS `keyring`.

## Checklist
- [ ] Small, focused, imperative commits
- [ ] Branch named `feature/...`, `fix/...`, or `chore/...`
- [ ] No secrets committed; uses `keyring`
- [ ] App starts and tray icon appears
- [ ] Hotkeys register and respond
- [ ] First run creates `config.json`
- [ ] Clipboard text/image flows through agent
- [ ] Batch processor runs as expected
- [ ] Updated docs (e.g., `AGENTS.md`) if needed
- [ ] Added/updated tests (`pytest`) where applicable
- [ ] i18n keys added/updated in `locales/en.json` and others as needed
