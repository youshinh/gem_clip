# Security Policy

## Supported Versions
We maintain the main branch. For reported issues, please include the commit
hash or release tag you are using.

## Reporting a Vulnerability
- Do not open public issues for sensitive security reports.
- Please send a private report (security@your-domain.example or via GitHub
  Security Advisories if enabled). Include:
  - Reproduction steps, impact, and scope
  - Version/commit, OS, and environment details
  - Proof‑of‑concept if possible
- We aim to acknowledge within 72 hours and provide a remediation plan or
  timeline after triage.

## Handling Secrets
- API keys are stored in the OS keyring. Do not hardcode or commit secrets in
  config files.
- If a secret is accidentally disclosed, rotate it immediately and open a
  private report if repository actions are required.

## Supply‑chain Considerations
- Lock runtime dependencies where possible (see `requirements.txt`).
- Review transitive dependencies periodically. Consider pinning for releases.

