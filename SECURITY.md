# Security Policy

## Supported versions

Only the latest release receives security fixes.

| Version | Supported |
|---------|-----------|
| 3.6.x   | ✅ |
| < 3.6   | ❌ |

## Reporting a vulnerability

Please report security issues **privately**, not in public issues:

- Preferably via **GitHub → Security → Report a vulnerability** (private
  advisory) on this repository, or
- by email to the maintainer (**50bvd**).

Include what you found, steps to reproduce, and the impact. We aim to
acknowledge reports within a few days and to fix confirmed issues promptly.
Please give us reasonable time to release a fix before public disclosure.

## Good to know

- The app runs **elevated** (administrator) because Windows/Office activation
  requires it. Review what it runs (`slmgr.vbs`, `ospp.vbs`) before use.
- Releases are built by GitHub Actions and carry an Authenticode signature.
  Until a trusted certificate is configured (see `CODE_SIGNING.md`), the
  signature is **self-signed**: verify the SHA-256 published in `RELEASE_INFO.txt`
  against your download.
- The built-in update check only performs a read-only HTTPS request to the
  GitHub Releases API and never downloads or executes anything automatically —
  it just opens the release page in your browser.
