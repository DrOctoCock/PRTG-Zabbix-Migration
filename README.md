# PRTG nach Zabbix Migrator

Dieses Skript migriert bewusst nur das, was zwischen PRTG und Zabbix sauber und risikoarm automatisierbar ist:

- PRTG-Geraeteinventar
- Host-Gruppen-Zuordnung
- Zabbix-Interfaces
- Template-Zuordnung per Mapping
- Host-Makros und Tags

Nicht automatisiert migriert werden:

- historische Messwerte
- PRTG-Sensorlogik 1:1
- Notifications, Dependencies, Maps, Dashboards
- sensible SNMP-/Credential-Details ohne explizites Mapping

## Dateien

- `prtg-to-zabbix.ps1`: Hauptskript
- `mapping.example.json`: Beispiel fuer Mapping-Regeln

## Voraussetzungen

- PowerShell 5.1 oder neuer
- PRTG API Token mit mindestens Leserechten
- Zabbix API Token mit Rechten zum Lesen, Anlegen und Aktualisieren von Hosts

## Best Practices

- Erst `-Action ExportPrtg` oder `-Action FullSync -DryRun` ausfuehren.
- Eine dedizierte temporäre Hostgruppe in Zabbix verwenden, z. B. `Migrated from PRTG`.
- Zabbix-Agent-Monitoring nie blind aus PRTG-WMI/SSH/HTTP-Sensoren ableiten; dafuer das Mapping explizit pflegen.
- SNMP-Credentials nicht implizit aus PRTG uebernehmen, sondern bewusst als Makros oder Interface-Details im Mapping setzen.
- PRTG erst abschalten, wenn Hosts, Interfaces, Templates und Erreichbarkeit in Zabbix validiert wurden.

## Typische Nutzung

### 1. Nur Export aus PRTG

```powershell
$env:PRTG_URL = "https://prtg.example.local"
$env:PRTG_API_TOKEN = "REDACTED"

powershell -NoProfile -ExecutionPolicy Bypass -File .\prtg-to-zabbix.ps1 `
  -Action ExportPrtg `
  -MappingFile .\mapping.example.json `
  -OutputJson .\out\prtg_inventory.json `
  -OutputCsv .\out\prtg_inventory.csv
```

### 2. Dry-Run fuer kompletten Transfer

```powershell
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
```

### 3. Import aus vorhandenem Export nach Zabbix

```powershell
$env:ZABBIX_URL = "https://zabbix.example.local/zabbix"
$env:ZABBIX_API_TOKEN = "REDACTED"

powershell -NoProfile -ExecutionPolicy Bypass -File .\prtg-to-zabbix.ps1 `
  -Action SyncZabbix `
  -InputJson .\out\prtg_inventory.json `
  -Mode SafeUpsert `
  -ReportFile .\out\prtg_zabbix_report.json
```

## Modi

- `CreateOnly`: existierende Hosts in Zabbix werden nicht angefasst.
- `SafeUpsert`: fuegt fehlende Gruppen, Templates, Tags, Makros und Interfaces hinzu, ohne aggressiv zu ueberschreiben.
- `ForceMerge`: aktualisiert bestehende Felder bewusst staerker.

## Mapping-Logik

`mapping.example.json` steuert:

- `defaults`
- `groupRules`
- `interfaceRules`
- `templateRules`
- `macroRules`

### Unterstuetzte Match-Felder

- `deviceRegex`
- `hostRegex`
- `groupRegex`
- `probeRegex`
- `commentsRegex`
- `locationRegex`
- `deviceTagRegex`
- `deviceTagsAnyOf`
- `deviceTagsAllOf`
- `sensorNameRegex`
- `sensorTypeRegex`
- `sensorTagRegex`
- `sensorTagsAnyOf`
- `sensorTagsAllOf`

### Wichtige Hinweise zum Mapping

- Die Beispielkonfiguration ist nur ein Startpunkt und muss an eure Template-Namen angepasst werden.
- Das Skript leitet standardmaessig SNMP, JMX und IPMI vorsichtig aus Sensortypen ab.
- Ein Zabbix-Agent-Interface wird standardmaessig nur erzeugt, wenn ihr es explizit taggt, z. B. `zbx-agent`.

## Ausgabe

- `prtg_inventory.json`: normalisierte Migrationsdaten
- `prtg_inventory.csv`: kompaktes Audit fuer Review
- `prtg_zabbix_report.json`: Ergebnisbericht mit Create/Update/Missing-Template-Zahlen

## Einschraenkungen

- Zabbix erwartet fuer `host.create` Interfaces. Deshalb ueberspringt das Skript standardmaessig PRTG-Objekte ohne sauber ableitbare Zabbix-Schnittstelle.
- Historische Daten sind absichtlich nicht enthalten, weil das zwischen beiden Plattformen weder fachlich noch API-seitig sauber 1:1 passt.
- Template-Namen werden in Zabbix exakt gesucht; nicht vorhandene Templates landen im Report.
