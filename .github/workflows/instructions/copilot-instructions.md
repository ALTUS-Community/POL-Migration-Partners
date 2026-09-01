---
applyTo: '**'
---

# POL-Migration Debugging Reference Instructions

## Repository boundary

This repository is a debugging-only reference for Project Online to Altus migration support. It contains readable source, product schemas, mapping examples, and sanitized diagnostic fixtures.

Every operational PowerShell entry point and the Visual Basic helper is fail-closed by `DebugOnlyGuard.ps1`. Do not remove, bypass, or move the guard below any command that could authenticate, access an API, invoke Microsoft Project, write files, install modules, or start a process.

Working migration scripts and separately approved migration tools are obtained from the [Altus Partner Portal](https://partners.altus.pro). Do not attempt operational work from this checkout.

## Data and configuration

- Never add customer records, exported project files, MPP/XLSX files, logs, tokens, secrets, connection strings, private URLs, client IDs, role IDs, plugin-step IDs, or generated output.
- Keep only approved, sanitized XML and JSON reference material. Do not add a blanket XML ignore rule; product schema XML is intentionally trackable.
- `export.config.json` contains product `sensei_*` mappings and must not be replaced with customer data.
- `appSettings.example.json` is a format reference only. Local `appSettings.json` files must remain untracked.
- The DataMigration executable, dependencies, and other compiled artifacts are never stored or packaged here.

## Source changes

- Preserve the numbered stage structure and readable historical source.
- Keep script changes minimal and safety-focused. Do not restore execution paths, package builders, runtime copies, or release automation.
- Use generic examples such as `https://contoso.sharepoint.com/sites/pwa`, `https://orgname.crm.dynamics.com/`, and `YOUR-ENTRA-CLIENT-ID`.
- Do not add customer-shaped names or environment-specific identifiers to comments, examples, fixtures, or tests.

## Validation

Run the cross-platform repository check from the repository root:

```powershell
pwsh -NoProfile -File .\4.SharePoint-Migration\Tests\Validate-DebugOnlyRepository.ps1
```

The check must pass before a change is handed off. It verifies tracked artifact policy, private-value hygiene, PowerShell parsing, retained JSON/XML parsing, and guard ordering. Pester tests may be run when the Pester module is available, but the repository validation must not require the migration runtime.

## Support handoff

Use the approved working tools in the Partner Portal to reproduce an issue. Bring back only a redacted command, exact error text, sanitized diagnostic fixtures, and a focused question. Redact tokens, customer names, URLs, identifiers, file contents, and connection details before adding anything here.

CI is validation-only. It must not stage the whole repository, create archives, create releases, upload assets, call Power Automate, or publish migration tools.
