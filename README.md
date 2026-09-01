# POL-Migration

## Debugging Reference Only

This repository contains readable source, schemas, mapping examples, and sanitized diagnostic fixtures for Project Online to Altus migration support.

The scripts and Visual Basic helper are intentionally fail-closed. They do not export, import, authenticate, modify files, automate Microsoft Project, or create packages from this repository. Do not add customer data, tokens, secrets, connection strings, binaries, packages, generated output, or release archives.

Working migration scripts and separately approved migration tools are provided through the [Altus Partner Portal](https://partners.altus.pro). Use the portal's current instructions for operational work.

## Support Loop

1. Reproduce the issue with the approved working tools in the Partner Portal.
2. Record the exact command, redacting tokens, URLs, customer names, identifiers, and file contents.
3. Bring the redacted command, exact error text, relevant sanitized fixture, and a focused question to this repository.

## Support Agent

Use `@pol-migration-support` in VS Code Chat for a single guided support experience across setup, configuration, migration process questions, error diagnosis, UAT, aftercare, and remigration.

The agent searches the `References` catalogue, explains the likely cause and available workaround, and works through one diagnostic check at a time. Customer-specific details may be included when they are needed for diagnosis, but remove them from any `Support Ready` escalation: replace customer names, URLs, identifiers, file contents, and other identifying values with placeholders, and never include tokens or secrets. The agent does not run this checkout's scripts or authenticate to an environment; use the approved Altus Partner Portal for operational reproduction.

The numbered folders preserve the historical migration stages. Their README files describe source behavior for diagnosis; they are not execution instructions for this repository.
