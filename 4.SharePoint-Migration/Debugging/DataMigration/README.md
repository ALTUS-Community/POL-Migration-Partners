# DataMigration Reference

This directory documents the configuration and data shape used by the separately approved Altus DataMigration tool. It does not contain the executable, dependencies, runtime configuration, plugin-step mappings, role identifiers, or customer data.

Obtain the current signed migration tool and its operational instructions from the [Altus Partner Portal](https://partners.altus.pro). The scripts in this repository remain blocked by `DebugOnlyGuard.ps1`.

## Configuration Example

`appSettings.example.json` is a strict JSON format reference. It uses interactive OAuth authentication, generic URLs and identifiers, a sample data path, and plugin bypass disabled. Copy it to a working location only when the approved tool's documentation requires this shape; do not add the working copy to this repository.

The example deliberately omits plugin-step mappings, administrator role identifiers, client secrets, telemetry connection strings, and other environment-specific settings.

## Data XML

The migration tool consumes CMT-style `Data.xml` files. The retained schema and mapping files in this repository describe the product entities and fields used by the historical stages. Diagnostic fixtures must be sanitized and must not contain customer records, access tokens, environment-specific identifiers, or private URLs.
