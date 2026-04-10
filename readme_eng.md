PRTG to Zabbix Migrator
This script deliberately migrates only what can be cleanly and safely automated between PRTG and Zabbix:

PRTG device inventory

Host group assignment

Zabbix interfaces

Template assignment via mapping

Host macros and tags

The following are not automatically migrated:

Historical data

PRTG sensor logic 1:1

Notifications, dependencies, maps, dashboards

Sensitive SNMP/credential details without explicit mapping

Files
prtg-to-zabbix.ps1: Main script

mapping.example.json: Example for mapping rules

Prerequisites
PowerShell 5.1 or newer

PRTG API token with at least read permissions

Zabbix API token with permissions to read, create, and update hosts

Best Practices
First execute -Action ExportPrtg or -Action FullSync -DryRun.

Use a dedicated temporary host group in Zabbix, e.g., Migrated from PRTG.

Never blindly derive Zabbix agent monitoring from PRTG WMI/SSH/HTTP sensors; explicitly maintain the mapping for this.

Do not implicitly adopt SNMP credentials from PRTG, but consciously set them as macros or interface details in the mapping.

Only shut down PRTG once hosts, interfaces, templates, and reachability have been validated in Zabbix.

Typical Usage
1. Export from PRTG only
PowerShell
$env:PRTG_URL = "https://prtg.example.local"
$env:PRTG_API_TOKEN = "REDACTED"

powershell -NoProfile -ExecutionPolicy Bypass -File .\prtg-to-zabbix.ps1 `
  -Action ExportPrtg `
  -MappingFile .\mapping.example.json `
  -OutputJson .\out\prtg_inventory.json `
  -OutputCsv .\out\prtg_inventory.csv
2. Dry-run for complete transfer
PowerShell
$env:PRTG_URL = "https://prtg.example.local"
$env:PRTG_API_TOKEN = "REDACTED"
$env:ZABBIX_URL = "https://zabbix.example.local/zabbix"
$env:ZABBIX_API_TOKEN = "REDACTED"

powershell -NoProfile -ExecutionPolicy Bypass -File .\prtg-to-zabbix.ps1 `
  -Action FullSync `
  -MappingFile .\mapping.example.json `
  -Mode SafeUpsert `
  -DryRun `
  -OutputJson .\out\prtg_inventory.json `
  -OutputCsv .\out\prtg_inventory.csv `
  -ReportFile .\out\prtg_zabbix_report.json
3. Import from existing export to Zabbix
PowerShell
$env:ZABBIX_URL = "https://zabbix.example.local/zabbix"
$env:ZABBIX_API_TOKEN = "REDACTED"

powershell -NoProfile -ExecutionPolicy Bypass -File .\prtg-to-zabbix.ps1 `
  -Action SyncZabbix `
  -InputJson .\out\prtg_inventory.json `
  -Mode SafeUpsert `
  -ReportFile .\out\prtg_zabbix_report.json
Modes
CreateOnly: existing hosts in Zabbix are not touched.

SafeUpsert: adds missing groups, templates, tags, macros, and interfaces without aggressively overwriting.

ForceMerge: deliberately updates existing fields more aggressively.

Mapping Logic
mapping.example.json controls:

defaults

groupRules

interfaceRules

templateRules

macroRules

Supported Match Fields
deviceRegex

hostRegex

groupRegex

probeRegex

commentsRegex

locationRegex

deviceTagRegex

deviceTagsAnyOf

deviceTagsAllOf

sensorNameRegex

sensorTypeRegex

sensorTagRegex

sensorTagsAnyOf

sensorTagsAllOf

Important Notes on Mapping
The example configuration is only a starting point and must be adapted to your template names.

By default, the script carefully derives SNMP, JMX, and IPMI from sensor types.

A Zabbix agent interface is only created by default if you explicitly tag it, e.g., zbx-agent.

Output
prtg_inventory.json: normalized migration data

prtg_inventory.csv: compact audit for review

prtg_zabbix_report.json: result report with create/update/missing template numbers

Limitations
Zabbix expects interfaces for host.create. Therefore, the script skips PRTG objects without cleanly derivable Zabbix interfaces by default.

Historical data is intentionally not included because it does not cleanly fit 1:1 between the two platforms, neither functionally nor from an API perspective.

Template names are searched for exactly in Zabbix; non-existent templates end up in the report.