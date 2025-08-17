# Privacy & Data Handling

This application runs locally and does not collect analytics or telemetry.
Below is a summary of how data is handled at runtime.

## Data Sources
- Clipboard: Text, images, or file references may be read locally to execute the
  selected prompts. You control when to run.
- File Attachments: When you explicitly attach files, their content may be read
  locally in order to send to the LLM API (if applicable to the chosen prompt).

## Where Data Goes
- Local: Input previews, results, and summaries are displayed in the UI and kept
  in memory. Optional session/set files (prompt presets, session state) are
  stored under `prompt_set/` in your working directory.
- Cloud (LLM calls): When you run a prompt that calls the Gemini API, the input
  you selected (text/images/files) is sent to the Google Generative AI service
  according to its API Terms and Safety Policies.

## Secrets & Configuration
- API Key: Stored in the OS keyring (not in plain text files). You can delete or
  rotate the key from Settings at any time.
- User Config: `config.json` is created on first run in your working directory
  and is ignored by git. It contains non‑secret preferences (language, hotkeys,
  history size, etc.).

## Logging
- The app produces minimal logs for troubleshooting. No clipboard content is
  intentionally logged. If you share logs for support, review/redact sensitive
  information first.

## Your Control
- Use the History selector and Matrix inputs to decide exactly which content to
  process. Close the app or revoke API key to prevent further requests.
- To remove local session/preset data, delete the files under `prompt_set/`.

## Compliance Notes
- This project is a desktop helper. Compliance with data‑handling requirements
  (e.g., internal policies, PII rules) depends on how you use it and what you
  submit to the LLM. Avoid sending confidential or regulated data unless
  contractually permitted and compliant with applicable policies and laws.

