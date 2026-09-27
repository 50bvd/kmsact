# Code signing

## What the build does today (free, no setup)

`Compile-CI.ps1` (run by the release workflow) already:

- Embeds **version metadata** in `KMS_Activator.exe` (`CompanyName = 50bvd`,
  `ProductName = MS KMS Activator`, description, version). Your name shows in
  *File → Properties → Details* and in the UAC prompt, with no certificate.
- Applies a **self-signed** Authenticode signature (subject `50bvd`) with a
  timestamp. This is *not* trusted on other machines, so Windows still shows
  "Unknown publisher" elsewhere — it only proves integrity.

## Optional: trusted signature via SignPath Foundation (free for OSS)

The release workflow can submit the built EXE to
[SignPath Foundation](https://signpath.org/) for a publicly-trusted signature.
The steps are **gated**: they run only when `SIGNPATH_API_TOKEN` is set, so
until then the self-signed build is published unchanged.

One-time setup:

1. Apply for OSS signing at <https://signpath.org/apply> using this repo's URL.
   Describe the tool honestly (a GUI KMS client for sysadmins that targets your
   own KMS host by default). Approval is manual and case-by-case.
2. In the SignPath console: install the GitHub App on this repo, create a
   **Project** (slug `kmsact`), an **Artifact Configuration** that signs
   `KMS_Activator.exe`, and a **Signing Policy** (slug `release-signing`); create
   a CI API token.
3. In GitHub **Settings → Secrets and variables → Actions**: add secret
   `SIGNPATH_API_TOKEN` and variable `SIGNPATH_ORGANIZATION_ID` (and optionally
   `SIGNPATH_PROJECT_SLUG`, `SIGNPATH_SIGNING_POLICY_SLUG`).

The `signpath/...@v1` action should be pinned to a commit SHA to match a strict
supply-chain policy.

## Honest note on Windows Defender

Signing changes the *publisher*, not the *classification*. Windows Defender flags
KMS activation tooling as `HackTool`/PUA based on **behavior** (installing a GVLK,
`slmgr /skms`, `ospp.vbs /act`) — the same actions a legitimate KMS client uses.
So a signed build may still be flagged. The only legitimate remedy is a
false-positive submission to Microsoft (<https://www.microsoft.com/wdsi/filesubmission>);
it may be declined because the behavior is intentional. Do not attempt to hide
or evade detection.
