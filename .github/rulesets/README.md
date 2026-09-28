# Branch rulesets

Versioned copies of the branch protection for this repo, so it is documented and
easy to re-apply.

- `main.json` — protects the default branch: no deletion, no force-push, linear
  history, requires a pull request (0 approvals — solo friendly) and the
  `PSScriptAnalyzer` check to pass. Repository admins can bypass.
- `dev.json` — lighter: no deletion, no force-push.

## Apply them (choose one)

### A. Import in the UI (easiest)
**Settings → Rules → Rulesets → New ruleset → Import a ruleset**, then select the
JSON file. Repeat for each file. Review and **Save**.

### B. PowerShell + GitHub CLI
PowerShell has no `<<'JSON'` heredoc — pipe a here-string instead:

```powershell
Get-Content .github/rulesets/main.json -Raw | gh api --method POST repos/50bvd/kmsact/rulesets --input -
Get-Content .github/rulesets/dev.json  -Raw | gh api --method POST repos/50bvd/kmsact/rulesets --input -
```

## Note on required checks
`main.json` requires the `PSScriptAnalyzer` check. If your check appears under a
different name in the Actions tab, update the `context` value (or pick it from the
dropdown when importing) so merges are not blocked waiting on a non-existent check.
