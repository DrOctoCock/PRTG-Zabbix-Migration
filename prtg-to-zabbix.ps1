[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('ExportPrtg', 'SyncZabbix', 'FullSync')]
    [string]$Action,

    [string]$PrtgUrl = $env:PRTG_URL,
    [string]$PrtgToken = $env:PRTG_API_TOKEN,
    [ValidateSet('Header', 'Query')]
    [string]$PrtgAuthMode = 'Header',

    [string]$ZabbixUrl = $env:ZABBIX_URL,
    [string]$ZabbixToken = $env:ZABBIX_API_TOKEN,
    [ValidateSet('Auto', 'Header', 'Auth')]
    [string]$ZabbixAuthMode = 'Auto',

    [string]$MappingFile,
    [string]$InputJson,
    [string]$OutputJson = '.\out\prtg_inventory.json',
    [string]$OutputCsv = '.\out\prtg_inventory.csv',
    [string]$ReportFile = '.\out\prtg_zabbix_report.json',

    [ValidateSet('CreateOnly', 'SafeUpsert', 'ForceMerge')]
    [string]$Mode = 'SafeUpsert',

    [int]$PageSize = 1000,
    [int]$TimeoutSec = 60,
    [int]$MaxRetries = 4,
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RequestCounter = 1
$script:ZabbixState = @{
    Endpoint      = $null
    ApiVersion    = $null
    AuthMode      = $null
    GroupCache    = @{}
    TemplateCache = @{}
}

function Write-Log {
    param(
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO',
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Write-Host ("[{0}] [{1}] {2}" -f $timestamp, $Level, $Message)
}

function Ensure-DirectoryForFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    $directory = Split-Path -Path $Path -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
}

function Save-JsonFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Data
    )

    Ensure-DirectoryForFile -Path $Path
    $json = $Data | ConvertTo-Json -Depth 50
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    [System.IO.File]::WriteAllText($fullPath, $json, [System.Text.Encoding]::UTF8)
}

function Load-JsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Datei nicht gefunden: $Path"
    }

    $content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($content)) {
        throw "JSON-Datei ist leer: $Path"
    }

    return ConvertFrom-Json -InputObject $content
}

function ConvertTo-Hashtable {
    param([Parameter(ValueFromPipeline = $true)]$InputObject)

    process {
        if ($null -eq $InputObject) {
            return $null
        }

        if ($InputObject -is [System.Collections.IDictionary]) {
            $result = @{}
            foreach ($key in $InputObject.Keys) {
                $result[$key] = ConvertTo-Hashtable -InputObject $InputObject[$key]
            }
            return $result
        }

        if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
            $result = @{}
            foreach ($property in $InputObject.PSObject.Properties) {
                $result[$property.Name] = ConvertTo-Hashtable -InputObject $property.Value
            }
            return $result
        }

        if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
            $items = @()
            foreach ($item in $InputObject) {
                $items += , (ConvertTo-Hashtable -InputObject $item)
            }
            return $items
        }

        return $InputObject
    }
}

function Merge-ConfigValues {
    param(
        $Base,
        $Override
    )

    if ($null -eq $Base) {
        return $Override
    }

    if ($null -eq $Override) {
        return $Base
    }

    if ($Base -is [System.Collections.IDictionary] -and $Override -is [System.Collections.IDictionary]) {
        $result = @{}
        foreach ($key in $Base.Keys) {
            $result[$key] = $Base[$key]
        }
        foreach ($key in $Override.Keys) {
            if ($result.ContainsKey($key)) {
                $result[$key] = Merge-ConfigValues -Base $result[$key] -Override $Override[$key]
            } else {
                $result[$key] = $Override[$key]
            }
        }
        return $result
    }

    if (
        $Base -is [System.Collections.IEnumerable] -and -not ($Base -is [string]) -and
        $Override -is [System.Collections.IEnumerable] -and -not ($Override -is [string])
    ) {
        return @($Base) + @($Override)
    }

    return $Override
}

function Ensure-Array {
    param($Value)

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [string]) {
        return @($Value)
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        return @($Value)
    }

    return @($Value)
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $true)]$InputObject,
        [Parameter(Mandatory = $true)][string[]]$Names,
        $Default = $null
    )

    foreach ($name in $Names) {
        if ($null -eq $InputObject) {
            break
        }

        if ($InputObject -is [System.Collections.IDictionary] -and $InputObject.ContainsKey($name)) {
            return $InputObject[$name]
        }

        if ($InputObject.PSObject.Properties.Match($name).Count -gt 0) {
            return $InputObject.$name
        }
    }

    return $Default
}

function Normalize-Text {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    $text = [string]$Value
    $text = $text.Replace('<br/>', ' ').Replace('<br />', ' ').Replace('&nbsp;', ' ')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    $text = ($text -replace '\s+', ' ').Trim()
    return $text
}

function Split-TagList {
    param($Value)

    $text = Normalize-Text -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return @()
    }

    return @($text -split '[,\s;|]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function ConvertTo-Slug {
    param([string]$Text)

    $normalized = Normalize-Text -Value $Text
    $slug = $normalized.ToLowerInvariant() -replace '[^a-z0-9._-]+', '-'
    $slug = $slug.Trim('-')
    if ([string]::IsNullOrWhiteSpace($slug)) {
        return 'host'
    }

    if ($slug.Length -gt 48) {
        return $slug.Substring(0, 48).Trim('-')
    }

    return $slug
}

function Test-IsIpAddress {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $parsed = $null
    return [System.Net.IPAddress]::TryParse($Value, [ref]$parsed)
}

function New-QueryString {
    param([hashtable]$Parameters)

    $pairs = @()
    foreach ($key in $Parameters.Keys) {
        foreach ($value in Ensure-Array -Value $Parameters[$key]) {
            if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
                continue
            }
            $pairs += ('{0}={1}' -f [System.Uri]::EscapeDataString([string]$key), [System.Uri]::EscapeDataString([string]$value))
        }
    }

    return ($pairs -join '&')
}

function Invoke-WebJson {
    param(
        [ValidateSet('Get', 'Post')]
        [string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [hashtable]$Headers = @{},
        $Body = $null,
        [int]$Timeout = 60,
        [int]$Retries = 4,
        [string]$ContentType = 'application/json'
    )

    $attempt = 0
    while ($attempt -lt $Retries) {
        try {
            $invokeParams = @{
                Method      = $Method
                Uri         = $Uri
                Headers     = $Headers
                TimeoutSec  = $Timeout
                ErrorAction = 'Stop'
            }

            if ($null -ne $Body) {
                $invokeParams['Body'] = ($Body | ConvertTo-Json -Depth 50 -Compress)
                $invokeParams['ContentType'] = $ContentType
            }

            return Invoke-RestMethod @invokeParams
        } catch {
            $attempt++
            $statusCode = $null
            try {
                $statusCode = [int]$_.Exception.Response.StatusCode.value__
            } catch {
                $statusCode = $null
            }

            $isRetryable = ($attempt -lt $Retries) -and (
                $null -eq $statusCode -or $statusCode -in 408, 429, 500, 502, 503, 504
            )

            if (-not $isRetryable) {
                throw
            }

            $sleepSeconds = [Math]::Min(15, [Math]::Pow(2, $attempt))
            Write-Log -Level 'WARN' -Message ("HTTP-Aufruf fehlgeschlagen (Versuch {0}/{1}), erneuter Versuch in {2}s: {3}" -f $attempt, $Retries, $sleepSeconds, $Uri)
            Start-Sleep -Seconds $sleepSeconds
        }
    }
}

function Get-DefaultMapping {
    return @{
        defaults       = @{
            defaultHostGroup                    = 'Migrated from PRTG'
            technicalNameStrategy              = 'prtg_objid'
            visibleNameStrategy                = 'device'
            hostStatus                         = 'enabled'
            skipDevicesWithoutSupportedInterface = $true
            preservePrtgTagsAsZabbixTags       = $true
            addProbeAndGroupAsTags             = $true
            includeDisabledPrtgDevices         = $false
        }
        groupRules      = @()
        interfaceRules  = @(
            @{
                name      = 'SNMP sensors'
                match     = @{
                    sensorTypeRegex = '^snmp'
                }
                interface = @{
                    type    = 'SNMP'
                    port    = '161'
                    main    = $true
                    address = 'host'
                    details = @{
                        version = 2
                        bulk    = 1
                    }
                }
            },
            @{
                name      = 'JMX sensors'
                match     = @{
                    sensorTypeRegex = '^jmx'
                }
                interface = @{
                    type    = 'JMX'
                    port    = '12345'
                    main    = $true
                    address = 'host'
                }
            },
            @{
                name      = 'IPMI sensors'
                match     = @{
                    sensorTypeRegex = '^ipmi'
                }
                interface = @{
                    type    = 'IPMI'
                    port    = '623'
                    main    = $true
                    address = 'host'
                }
            },
            @{
                name      = 'Explicit Zabbix agent tag'
                match     = @{
                    deviceTagRegex = '(^|[\s,;])zbx[-_ ]?agent($|[\s,;])'
                }
                interface = @{
                    type    = 'AGENT'
                    port    = '10050'
                    main    = $true
                    address = 'host'
                }
            }
        )
        templateRules   = @()
        macroRules      = @()
    }
}

function Load-Mapping {
    param([string]$Path)

    $defaults = Get-DefaultMapping
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $defaults
    }

    $fileData = ConvertTo-Hashtable -InputObject (Load-JsonFile -Path $Path)
    return Merge-ConfigValues -Base $defaults -Override $fileData
}

function Get-InterfaceTypeCode {
    param([string]$Type)

    switch ($Type.ToUpperInvariant()) {
        'AGENT' { return 1 }
        'SNMP' { return 2 }
        'IPMI' { return 3 }
        'JMX' { return 4 }
        default { throw "Nicht unterstützter Interface-Typ: $Type" }
    }
}

function Get-DefaultPort {
    param([string]$Type)

    switch ($Type.ToUpperInvariant()) {
        'AGENT' { return '10050' }
        'SNMP' { return '161' }
        'IPMI' { return '623' }
        'JMX' { return '12345' }
        default { return '0' }
    }
}

function Test-RegexField {
    param(
        [hashtable]$Match,
        [string]$FieldName,
        [string]$Value
    )

    if (-not $Match.ContainsKey($FieldName)) {
        return $true
    }

    $pattern = Normalize-Text -Value $Match[$FieldName]
    if ([string]::IsNullOrWhiteSpace($pattern)) {
        return $true
    }

    return ($Value -match $pattern)
}

function Test-TagCriteria {
    param(
        [hashtable]$Match,
        [string]$AnyKey,
        [string]$AllKey,
        [string[]]$Tags
    )

    $tagSet = @($Tags | ForEach-Object { $_.ToLowerInvariant() })

    if ($Match.ContainsKey($AnyKey)) {
        $requiredAny = @((Ensure-Array -Value $Match[$AnyKey]) | ForEach-Object { [string]$_ })
        $matchedAny = $false
        foreach ($tag in $requiredAny) {
            if ($tagSet -contains $tag.ToLowerInvariant()) {
                $matchedAny = $true
                break
            }
        }
        if (-not $matchedAny) {
            return $false
        }
    }

    if ($Match.ContainsKey($AllKey)) {
        $requiredAll = @((Ensure-Array -Value $Match[$AllKey]) | ForEach-Object { [string]$_ })
        foreach ($tag in $requiredAll) {
            if (-not ($tagSet -contains $tag.ToLowerInvariant())) {
                return $false
            }
        }
    }

    return $true
}

function Test-SensorCriteria {
    param(
        [hashtable]$Match,
        [hashtable]$Sensor
    )

    if (-not (Test-RegexField -Match $Match -FieldName 'sensorNameRegex' -Value (Normalize-Text -Value $Sensor.name))) {
        return $false
    }
    if (-not (Test-RegexField -Match $Match -FieldName 'sensorTypeRegex' -Value (Normalize-Text -Value $Sensor.type))) {
        return $false
    }
    if (-not (Test-RegexField -Match $Match -FieldName 'sensorTagRegex' -Value (($Sensor.tags | ForEach-Object { $_.ToLowerInvariant() }) -join ' '))) {
        return $false
    }
    if (-not (Test-TagCriteria -Match $Match -AnyKey 'sensorTagsAnyOf' -AllKey 'sensorTagsAllOf' -Tags $Sensor.tags)) {
        return $false
    }

    return $true
}

function Test-RuleMatch {
    param(
        [hashtable]$Rule,
        [hashtable]$Device
    )

    $match = ConvertTo-Hashtable -InputObject (Get-PropertyValue -InputObject $Rule -Names @('match') -Default @{})
    if ($match.Count -eq 0) {
        return $true
    }

    $deviceTagsJoined = (($Device.tags | ForEach-Object { $_.ToLowerInvariant() }) -join ' ')

    if (-not (Test-RegexField -Match $match -FieldName 'deviceRegex' -Value (Normalize-Text -Value $Device.device))) {
        return $false
    }
    if (-not (Test-RegexField -Match $match -FieldName 'hostRegex' -Value (Normalize-Text -Value $Device.host))) {
        return $false
    }
    if (-not (Test-RegexField -Match $match -FieldName 'groupRegex' -Value (Normalize-Text -Value $Device.group))) {
        return $false
    }
    if (-not (Test-RegexField -Match $match -FieldName 'probeRegex' -Value (Normalize-Text -Value $Device.probe))) {
        return $false
    }
    if (-not (Test-RegexField -Match $match -FieldName 'commentsRegex' -Value (Normalize-Text -Value $Device.comments))) {
        return $false
    }
    if (-not (Test-RegexField -Match $match -FieldName 'locationRegex' -Value (Normalize-Text -Value $Device.location))) {
        return $false
    }
    if (-not (Test-RegexField -Match $match -FieldName 'deviceTagRegex' -Value $deviceTagsJoined)) {
        return $false
    }
    if (-not (Test-TagCriteria -Match $match -AnyKey 'deviceTagsAnyOf' -AllKey 'deviceTagsAllOf' -Tags $Device.tags)) {
        return $false
    }

    $sensorCriteriaPresent = @(
        'sensorNameRegex',
        'sensorTypeRegex',
        'sensorTagRegex',
        'sensorTagsAnyOf',
        'sensorTagsAllOf'
    ) | Where-Object { $match.ContainsKey($_) }

    if ($sensorCriteriaPresent.Count -eq 0) {
        return $true
    }

    foreach ($sensor in $Device.sensors) {
        if (Test-SensorCriteria -Match $match -Sensor $sensor) {
            return $true
        }
    }

    return $false
}

function Resolve-InterfaceAddress {
    param(
        [hashtable]$Device,
        [hashtable]$InterfaceDefinition
    )

    $strategy = Normalize-Text -Value (Get-PropertyValue -InputObject $InterfaceDefinition -Names @('address') -Default 'host')
    $explicitValue = Normalize-Text -Value (Get-PropertyValue -InputObject $InterfaceDefinition -Names @('addressValue') -Default '')
    $candidate = ''

    switch ($strategy.ToLowerInvariant()) {
        'literal' {
            $candidate = $explicitValue
        }
        default {
            $candidate = Normalize-Text -Value $Device.host
        }
    }

    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return $null
    }

    switch ($strategy.ToLowerInvariant()) {
        'dns' {
            return @{
                useip = 0
                ip    = ''
                dns   = $candidate
            }
        }
        'ip' {
            if (-not (Test-IsIpAddress -Value $candidate)) {
                return $null
            }
            return @{
                useip = 1
                ip    = $candidate
                dns   = ''
            }
        }
        default {
            if (Test-IsIpAddress -Value $candidate) {
                return @{
                    useip = 1
                    ip    = $candidate
                    dns   = ''
                }
            }
            return @{
                useip = 0
                ip    = ''
                dns   = $candidate
            }
        }
    }
}

function Add-UniqueString {
    param(
        [System.Collections.Generic.List[string]]$Target,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return
    }

    if (-not $Target.Contains($Value)) {
        $Target.Add($Value)
    }
}

function Merge-TagObjects {
    param([array]$Tags)

    $seen = @{}
    $merged = @()

    foreach ($tag in Ensure-Array -Value $Tags) {
        $entry = ConvertTo-Hashtable -InputObject $tag
        $name = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('tag') -Default '')
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }
        $value = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('value') -Default '')
        $identity = '{0}|{1}' -f $name, $value
        if ($seen.ContainsKey($identity)) {
            continue
        }
        $seen[$identity] = $true
        $merged += @{
            tag   = $name
            value = $value
        }
    }

    return $merged
}

function Merge-MacroObjects {
    param([array]$Macros)

    $seen = @{}
    $merged = @()

    foreach ($macro in Ensure-Array -Value $Macros) {
        $entry = ConvertTo-Hashtable -InputObject $macro
        $name = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('macro') -Default '')
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($seen.ContainsKey($name)) {
            continue
        }

        $seen[$name] = $true
        $merged += @{
            macro       = $name
            value       = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('value') -Default '')
            description = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('description') -Default '')
        }
    }

    return $merged
}

function Merge-InterfaceObjects {
    param([array]$Interfaces)

    $seen = @{}
    $merged = @()

    foreach ($interface in Ensure-Array -Value $Interfaces) {
        $entry = ConvertTo-Hashtable -InputObject $interface
        $identity = '{0}|{1}|{2}|{3}|{4}' -f
            (Normalize-Text -Value $entry.type).ToUpperInvariant(),
            [string](Get-PropertyValue -InputObject $entry -Names @('useip') -Default 1),
            (Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('ip') -Default '')),
            (Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('dns') -Default '')),
            (Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('port') -Default ''))

        if ($seen.ContainsKey($identity)) {
            continue
        }
        $seen[$identity] = $true
        $merged += $entry
    }

    return $merged
}

function Get-HostStatusCode {
    param([string]$Status)

    if ($Status -eq 'disabled') {
        return 1
    }

    return 0
}

function Get-TechnicalHostName {
    param(
        [hashtable]$Device,
        [string]$Strategy
    )

    switch ($Strategy) {
        'device_name' {
            return ConvertTo-Slug -Text $Device.device
        }
        'host_or_ip' {
            $basis = Normalize-Text -Value $Device.host
            if ([string]::IsNullOrWhiteSpace($basis)) {
                $basis = Normalize-Text -Value $Device.device
            }
            return ConvertTo-Slug -Text $basis
        }
        'prtg_objid_slug' {
            return ('prtg-{0}-{1}' -f $Device.prtgId, (ConvertTo-Slug -Text $Device.device))
        }
        default {
            return ('prtg-{0}' -f $Device.prtgId)
        }
    }
}

function Get-VisibleHostName {
    param(
        [hashtable]$Device,
        [string]$Strategy
    )

    switch ($Strategy) {
        'host' {
            $value = Normalize-Text -Value $Device.host
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
            return Normalize-Text -Value $Device.device
        }
        default {
            $value = Normalize-Text -Value $Device.device
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
            return Normalize-Text -Value $Device.host
        }
    }
}

function Convert-RuleToInterface {
    param(
        [hashtable]$Rule,
        [hashtable]$Device
    )

    $definition = ConvertTo-Hashtable -InputObject (Get-PropertyValue -InputObject $Rule -Names @('interface') -Default $null)
    if ($null -eq $definition) {
        return $null
    }

    $type = (Normalize-Text -Value (Get-PropertyValue -InputObject $definition -Names @('type') -Default '')).ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($type)) {
        return $null
    }

    $address = Resolve-InterfaceAddress -Device $Device -InterfaceDefinition $definition
    if ($null -eq $address) {
        return $null
    }

    return @{
        type    = $type
        main    = if ([bool](Get-PropertyValue -InputObject $definition -Names @('main') -Default $true)) { 1 } else { 0 }
        useip   = $address.useip
        ip      = $address.ip
        dns     = $address.dns
        port    = Normalize-Text -Value (Get-PropertyValue -InputObject $definition -Names @('port') -Default (Get-DefaultPort -Type $type))
        details = ConvertTo-Hashtable -InputObject (Get-PropertyValue -InputObject $definition -Names @('details') -Default @{})
    }
}

function New-BaseTagsForDevice {
    param(
        [hashtable]$Device,
        [hashtable]$Defaults
    )

    $tags = @(
        @{
            tag   = 'source'
            value = 'PRTG'
        },
        @{
            tag   = 'prtg_objid'
            value = [string]$Device.prtgId
        }
    )

    if ([bool](Get-PropertyValue -InputObject $Defaults -Names @('addProbeAndGroupAsTags') -Default $true)) {
        if (-not [string]::IsNullOrWhiteSpace($Device.probe)) {
            $tags += @{
                tag   = 'prtg_probe'
                value = $Device.probe
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($Device.group)) {
            $tags += @{
                tag   = 'prtg_group'
                value = $Device.group
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($Device.host)) {
            $tags += @{
                tag   = 'prtg_host'
                value = $Device.host
            }
        }
    }

    if ([bool](Get-PropertyValue -InputObject $Defaults -Names @('preservePrtgTagsAsZabbixTags') -Default $true)) {
        foreach ($tag in $Device.tags) {
            $tags += @{
                tag   = 'prtg_tag'
                value = $tag
            }
        }
    }

    return $tags
}

function Convert-PrtgDeviceToPlan {
    param(
        [hashtable]$Device,
        [hashtable]$Mapping
    )

    $defaults = ConvertTo-Hashtable -InputObject $Mapping.defaults
    $notes = New-Object System.Collections.Generic.List[string]

    $groups = New-Object System.Collections.Generic.List[string]
    Add-UniqueString -Target $groups -Value (Normalize-Text -Value (Get-PropertyValue -InputObject $defaults -Names @('defaultHostGroup') -Default 'Migrated from PRTG'))

    foreach ($rule in Ensure-Array -Value $Mapping.groupRules) {
        $nativeRule = ConvertTo-Hashtable -InputObject $rule
        if (-not (Test-RuleMatch -Rule $nativeRule -Device $Device)) {
            continue
        }
        foreach ($groupName in Ensure-Array -Value (Get-PropertyValue -InputObject $nativeRule -Names @('targetGroups', 'groups') -Default @())) {
            Add-UniqueString -Target $groups -Value (Normalize-Text -Value $groupName)
        }
    }

    $interfaces = @()
    $templates = New-Object System.Collections.Generic.List[string]
    $macros = @()
    $tags = New-BaseTagsForDevice -Device $Device -Defaults $defaults

    foreach ($rule in Ensure-Array -Value $Mapping.interfaceRules) {
        $nativeRule = ConvertTo-Hashtable -InputObject $rule
        if (-not (Test-RuleMatch -Rule $nativeRule -Device $Device)) {
            continue
        }

        $interface = Convert-RuleToInterface -Rule $nativeRule -Device $Device
        if ($null -eq $interface) {
            Add-UniqueString -Target $notes -Value ("Interface-Regel '{0}' konnte keine gueltige Adresse aufloesen." -f (Normalize-Text -Value (Get-PropertyValue -InputObject $nativeRule -Names @('name') -Default 'unnamed')))
            continue
        }
        $interfaces += $interface

        foreach ($template in Ensure-Array -Value (Get-PropertyValue -InputObject $nativeRule -Names @('templates') -Default @())) {
            Add-UniqueString -Target $templates -Value (Normalize-Text -Value $template)
        }
        $macros += Ensure-Array -Value (Get-PropertyValue -InputObject $nativeRule -Names @('macros') -Default @())
        $tags += Ensure-Array -Value (Get-PropertyValue -InputObject $nativeRule -Names @('tags') -Default @())
    }

    foreach ($rule in Ensure-Array -Value $Mapping.templateRules) {
        $nativeRule = ConvertTo-Hashtable -InputObject $rule
        if (-not (Test-RuleMatch -Rule $nativeRule -Device $Device)) {
            continue
        }
        foreach ($template in Ensure-Array -Value (Get-PropertyValue -InputObject $nativeRule -Names @('templates') -Default @())) {
            Add-UniqueString -Target $templates -Value (Normalize-Text -Value $template)
        }
    }

    foreach ($rule in Ensure-Array -Value $Mapping.macroRules) {
        $nativeRule = ConvertTo-Hashtable -InputObject $rule
        if (-not (Test-RuleMatch -Rule $nativeRule -Device $Device)) {
            continue
        }
        $macros += Ensure-Array -Value (Get-PropertyValue -InputObject $nativeRule -Names @('macros') -Default @())
    }

    $interfaces = Merge-InterfaceObjects -Interfaces $interfaces
    $macros = Merge-MacroObjects -Macros $macros
    $tags = Merge-TagObjects -Tags $tags

    $sensorTypes = @($Device.sensors | ForEach-Object { Normalize-Text -Value $_.type } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($interfaces.Count -eq 0 -and $sensorTypes.Count -gt 0) {
        Add-UniqueString -Target $notes -Value ("Keine Zabbix-Schnittstelle abgeleitet. Sensor-Typen in PRTG: {0}" -f ($sensorTypes -join ', '))
    }

    $hasSnmp = @($interfaces | Where-Object { $_.type -eq 'SNMP' }).Count -gt 0
    $hasSnmpCommunity = @($macros | Where-Object { $_.macro -eq '{$SNMP_COMMUNITY}' }).Count -gt 0
    if ($hasSnmp -and -not $hasSnmpCommunity) {
        Add-UniqueString -Target $notes -Value 'SNMP erkannt, aber keine {$SNMP_COMMUNITY}- oder SNMPv3-Makros gesetzt. Zugangsdaten im Mapping pruefen.'
    }

    $inventoryNotes = @(
        ("Imported from PRTG object {0}" -f $Device.prtgId),
        ("PRTG path: {0} / {1}" -f $Device.probe, $Device.group)
    )
    if (-not [string]::IsNullOrWhiteSpace($Device.comments)) {
        $inventoryNotes += ("PRTG comments: {0}" -f $Device.comments)
    }
    foreach ($note in $notes) {
        $inventoryNotes += ("Migration note: {0}" -f $note)
    }

    $skipWithoutInterface = [bool](Get-PropertyValue -InputObject $defaults -Names @('skipDevicesWithoutSupportedInterface') -Default $true)
    $skip = $skipWithoutInterface -and $interfaces.Count -eq 0

    return @{
        skipped     = $skip
        skipReason  = if ($skip) { 'No supported Zabbix interface inferred' } else { '' }
        source      = @{
            system      = 'PRTG'
            prtgId      = [string]$Device.prtgId
            probe       = $Device.probe
            group       = $Device.group
            device      = $Device.device
            host        = $Device.host
            tags        = $Device.tags
            comments    = $Device.comments
            location    = $Device.location
            sensorCount = @($Device.sensors).Count
            sensorTypes = $sensorTypes
        }
        zabbix      = @{
            host       = Get-TechnicalHostName -Device $Device -Strategy (Normalize-Text -Value (Get-PropertyValue -InputObject $defaults -Names @('technicalNameStrategy') -Default 'prtg_objid'))
            name       = Get-VisibleHostName -Device $Device -Strategy (Normalize-Text -Value (Get-PropertyValue -InputObject $defaults -Names @('visibleNameStrategy') -Default 'device'))
            status     = Normalize-Text -Value (Get-PropertyValue -InputObject $defaults -Names @('hostStatus') -Default 'enabled')
            groups     = @($groups | Select-Object -Unique)
            interfaces = $interfaces
            templates  = @($templates | Select-Object -Unique)
            macros     = $macros
            tags       = $tags
            inventory  = @{
                location = $Device.location
                notes    = ($inventoryNotes -join "`n")
            }
        }
        notes       = @($notes | Select-Object -Unique)
    }
}

function Resolve-PrtgRows {
    param(
        $Payload,
        [string]$Content
    )

    $native = ConvertTo-Hashtable -InputObject $Payload
    $candidates = @($Content, $Content.ToLowerInvariant(), 'items', 'records', 'data')
    foreach ($candidate in $candidates) {
        if ($native.ContainsKey($candidate)) {
            return @(Ensure-Array -Value $native[$candidate])
        }
    }

    $arrayProperties = @()
    foreach ($key in $native.Keys) {
        $value = $native[$key]
        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string]) -and -not ($value -is [System.Collections.IDictionary])) {
            $arrayProperties += $key
        }
    }

    if ($arrayProperties.Count -eq 1) {
        return @(Ensure-Array -Value $native[$arrayProperties[0]])
    }

    throw "Konnte PRTG-Antwort fuer '$Content' nicht interpretieren."
}

function Invoke-PrtgTable {
    param(
        [string]$BaseUrl,
        [string]$Token,
        [string]$AuthMode,
        [string]$Content,
        [string[]]$Columns,
        [int]$Start,
        [int]$Count,
        [int]$Timeout,
        [int]$Retries
    )

    $headers = @{}
    $query = @{
        content = $Content
        columns = ($Columns -join ',')
        start   = $Start
        count   = $Count
    }

    if ($AuthMode -eq 'Query') {
        $query['apitoken'] = $Token
    } else {
        $headers['Authorization'] = "Bearer $Token"
    }

    $uri = ('{0}/api/table.json?{1}' -f $BaseUrl.TrimEnd('/'), (New-QueryString -Parameters $query))
    return Invoke-WebJson -Method 'Get' -Uri $uri -Headers $headers -Timeout $Timeout -Retries $Retries
}

function Get-PrtgInventory {
    param(
        [string]$BaseUrl,
        [string]$Token,
        [string]$AuthMode,
        [hashtable]$Mapping,
        [int]$Count,
        [int]$Timeout,
        [int]$Retries
    )

    Write-Log -Level 'INFO' -Message 'Lese Geraete aus PRTG...'
    $deviceColumns = @('objid', 'probe', 'group', 'device', 'host', 'tags', 'comments', 'location', 'active', 'status', 'parentid')
    $devices = @()
    $start = 0
    $prtgVersion = $null
    while ($true) {
        $payload = Invoke-PrtgTable -BaseUrl $BaseUrl -Token $Token -AuthMode $AuthMode -Content 'devices' -Columns $deviceColumns -Start $start -Count $Count -Timeout $Timeout -Retries $Retries
        if ($null -eq $prtgVersion) {
            $prtgVersion = Normalize-Text -Value (Get-PropertyValue -InputObject $payload -Names @('version', 'prtg-version') -Default '')
        }
        $rows = Resolve-PrtgRows -Payload $payload -Content 'devices'
        if ($rows.Count -eq 0) {
            break
        }
        $devices += $rows
        if ($rows.Count -lt $Count) {
            break
        }
        $start += $Count
    }

    Write-Log -Level 'INFO' -Message 'Lese Sensoren aus PRTG...'
    $sensorColumns = @('objid', 'parentid', 'device', 'sensor', 'type', 'tags', 'status', 'message', 'active')
    $sensors = @()
    $start = 0
    while ($true) {
        $payload = Invoke-PrtgTable -BaseUrl $BaseUrl -Token $Token -AuthMode $AuthMode -Content 'sensors' -Columns $sensorColumns -Start $start -Count $Count -Timeout $Timeout -Retries $Retries
        if ($null -eq $prtgVersion) {
            $prtgVersion = Normalize-Text -Value (Get-PropertyValue -InputObject $payload -Names @('version', 'prtg-version') -Default '')
        }
        $rows = Resolve-PrtgRows -Payload $payload -Content 'sensors'
        if ($rows.Count -eq 0) {
            break
        }
        $sensors += $rows
        if ($rows.Count -lt $Count) {
            break
        }
        $start += $Count
    }

    $deviceMap = @{}
    foreach ($rawDevice in $devices) {
        $entry = ConvertTo-Hashtable -InputObject $rawDevice
        $prtgId = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('objid', 'objid_RAW') -Default '')
        if ([string]::IsNullOrWhiteSpace($prtgId)) {
            continue
        }

        $deviceMap[$prtgId] = @{
            prtgId   = $prtgId
            probe    = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('probe') -Default '')
            group    = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('group') -Default '')
            device   = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('device', 'name') -Default '')
            host     = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('host') -Default '')
            tags     = Split-TagList -Value (Get-PropertyValue -InputObject $entry -Names @('tags') -Default '')
            comments = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('comments') -Default '')
            location = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('location') -Default '')
            active   = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('active') -Default '')
            status   = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('status') -Default '')
            sensors  = @()
        }
    }

    foreach ($rawSensor in $sensors) {
        $entry = ConvertTo-Hashtable -InputObject $rawSensor
        $parentId = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('parentid', 'parentid_RAW') -Default '')
        if ([string]::IsNullOrWhiteSpace($parentId) -or -not $deviceMap.ContainsKey($parentId)) {
            continue
        }

        $deviceMap[$parentId].sensors += @{
            prtgId   = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('objid', 'objid_RAW') -Default '')
            parentId = $parentId
            name     = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('sensor', 'name') -Default '')
            type     = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('type') -Default '')
            tags     = Split-TagList -Value (Get-PropertyValue -InputObject $entry -Names @('tags') -Default '')
            status   = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('status') -Default '')
            message  = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('message') -Default '')
            active   = Normalize-Text -Value (Get-PropertyValue -InputObject $entry -Names @('active') -Default '')
        }
    }

    $includeDisabled = [bool](Get-PropertyValue -InputObject $Mapping.defaults -Names @('includeDisabledPrtgDevices') -Default $false)
    $inventory = @()
    $skipped = @()

    foreach ($device in $deviceMap.Values | Sort-Object { [int]$_.prtgId }) {
        if (-not $includeDisabled -and ($device.status -match '^Paused' -or $device.active -eq 'false')) {
            continue
        }

        $plan = Convert-PrtgDeviceToPlan -Device $device -Mapping $Mapping
        if ($plan.skipped) {
            $skipped += $plan
        } else {
            $inventory += $plan
        }
    }

    return @{
        generatedAt = (Get-Date).ToString('o')
        generator   = 'prtg-to-zabbix.ps1'
        prtg        = @{
            url         = $BaseUrl
            version     = $prtgVersion
            deviceCount = @($deviceMap.Values).Count
            sensorCount = @($sensors).Count
        }
        inventory   = $inventory
        skipped     = $skipped
    }
}

function Export-InventoryCsv {
    param(
        [Parameter(Mandatory = $true)]$ExportData,
        [Parameter(Mandatory = $true)][string]$Path
    )

    Ensure-DirectoryForFile -Path $Path
    $rows = foreach ($entry in Ensure-Array -Value $ExportData.inventory) {
        $source = ConvertTo-Hashtable -InputObject $entry.source
        $zabbix = ConvertTo-Hashtable -InputObject $entry.zabbix

        [PSCustomObject]@{
            prtg_objid        = $source.prtgId
            prtg_probe        = $source.probe
            prtg_group        = $source.group
            prtg_device       = $source.device
            prtg_host         = $source.host
            sensor_types      = (@($source.sensorTypes) -join '; ')
            zabbix_host       = $zabbix.host
            zabbix_name       = $zabbix.name
            zabbix_groups     = (@($zabbix.groups) -join '; ')
            zabbix_interfaces = (@($zabbix.interfaces | ForEach-Object {
                    if ($_.useip -eq 1) {
                        '{0}:{1}/{2}' -f $_.type, $_.ip, $_.port
                    } else {
                        '{0}:{1}/{2}' -f $_.type, $_.dns, $_.port
                    }
                }) -join '; ')
            templates         = (@($zabbix.templates) -join '; ')
            macros            = (@($zabbix.macros | ForEach-Object { $_.macro }) -join '; ')
            notes             = (@($entry.notes) -join '; ')
        }
    }

    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Get-ZabbixEndpointCandidates {
    param([string]$BaseUrl)

    $url = $BaseUrl.TrimEnd('/')
    if ($url -match 'api_jsonrpc\.php$') {
        return @($url)
    }
    if ($url -match '/ui$') {
        return @(
            "$url/api_jsonrpc.php",
            ($url -replace '/ui$', '/api_jsonrpc.php')
        )
    }
    return @(
        "$url/ui/api_jsonrpc.php",
        "$url/api_jsonrpc.php"
    )
}

function Invoke-ZabbixRpc {
    param(
        [Parameter(Mandatory = $true)][string]$Endpoint,
        [Parameter(Mandatory = $true)][string]$Method,
        $Params = @{},
        [ValidateSet('None', 'Header', 'Auth')]
        [string]$AuthMode = 'None',
        [string]$Token = '',
        [int]$Timeout = 60,
        [int]$Retries = 4
    )

    $headers = @{}
    $payload = [ordered]@{
        jsonrpc = '2.0'
        method  = $Method
        params  = $Params
        id      = $script:RequestCounter
    }
    $script:RequestCounter++

    if ($AuthMode -eq 'Header') {
        $headers['Authorization'] = "Bearer $Token"
    } elseif ($AuthMode -eq 'Auth') {
        $payload['auth'] = $Token
    }

    $response = Invoke-WebJson -Method 'Post' -Uri $Endpoint -Headers $headers -Body $payload -Timeout $Timeout -Retries $Retries -ContentType 'application/json-rpc'
    $nativeResponse = ConvertTo-Hashtable -InputObject $response

    if ($nativeResponse.ContainsKey('error') -and $null -ne $nativeResponse.error) {
        $error = ConvertTo-Hashtable -InputObject $nativeResponse.error
        $message = Normalize-Text -Value (Get-PropertyValue -InputObject $error -Names @('message') -Default 'Zabbix RPC error')
        $data = Normalize-Text -Value (Get-PropertyValue -InputObject $error -Names @('data') -Default '')
        throw ("Zabbix API Fehler bei {0}: {1} {2}" -f $Method, $message, $data)
    }

    return $nativeResponse.result
}

function Initialize-ZabbixConnection {
    param(
        [string]$BaseUrl,
        [string]$Token,
        [string]$PreferredAuthMode,
        [int]$Timeout,
        [int]$Retries
    )

    if ($script:ZabbixState.Endpoint) {
        return
    }

    $candidates = Get-ZabbixEndpointCandidates -BaseUrl $BaseUrl
    foreach ($endpoint in $candidates) {
        try {
            $version = Invoke-ZabbixRpc -Endpoint $endpoint -Method 'apiinfo.version' -Params @{} -AuthMode 'None' -Timeout $Timeout -Retries $Retries
            $script:ZabbixState.Endpoint = $endpoint
            $script:ZabbixState.ApiVersion = [string]$version
            break
        } catch {
            continue
        }
    }

    if (-not $script:ZabbixState.Endpoint) {
        throw 'Konnte keinen gueltigen Zabbix API-Endpunkt finden.'
    }

    $candidateAuthModes = switch ($PreferredAuthMode) {
        'Header' { @('Header') }
        'Auth' { @('Auth') }
        default { @('Header', 'Auth') }
    }

    foreach ($authMode in $candidateAuthModes) {
        try {
            Invoke-ZabbixRpc -Endpoint $script:ZabbixState.Endpoint -Method 'hostgroup.get' -Params @{
                output = @('groupid')
                limit  = 1
            } -AuthMode $authMode -Token $Token -Timeout $Timeout -Retries $Retries | Out-Null
            $script:ZabbixState.AuthMode = $authMode
            return
        } catch {
            continue
        }
    }

    throw 'Konnte Zabbix-Authentifizierung mit dem bereitgestellten Token nicht verifizieren.'
}

function Invoke-ZabbixApi {
    param(
        [string]$Method,
        $Params = @{},
        [string]$Token,
        [int]$Timeout,
        [int]$Retries
    )

    return Invoke-ZabbixRpc -Endpoint $script:ZabbixState.Endpoint -Method $Method -Params $Params -AuthMode $script:ZabbixState.AuthMode -Token $Token -Timeout $Timeout -Retries $Retries
}

function Ensure-ZabbixGroupIds {
    param(
        [string[]]$GroupNames,
        [string]$Token,
        [int]$Timeout,
        [int]$Retries,
        [switch]$WhatIfOnly,
        [hashtable]$Summary
    )

    $ids = @()
    foreach ($groupNameRaw in $GroupNames) {
        $groupName = Normalize-Text -Value $groupNameRaw
        if ([string]::IsNullOrWhiteSpace($groupName)) {
            continue
        }

        if ($script:ZabbixState.GroupCache.ContainsKey($groupName)) {
            $ids += [string]$script:ZabbixState.GroupCache[$groupName]
            continue
        }

        $existing = @(Invoke-ZabbixApi -Method 'hostgroup.get' -Params @{
                filter = @{
                    name = @($groupName)
                }
                output = @('groupid', 'name')
            } -Token $Token -Timeout $Timeout -Retries $Retries)

        if ($existing.Count -gt 0) {
            $groupId = [string](ConvertTo-Hashtable -InputObject $existing[0]).groupid
            $script:ZabbixState.GroupCache[$groupName] = $groupId
            $ids += $groupId
            continue
        }

        if ($WhatIfOnly) {
            $fakeId = "DRYRUN_GROUP_$groupName"
            $script:ZabbixState.GroupCache[$groupName] = $fakeId
            $ids += $fakeId
            $Summary.zabbix.groupsPlanned += 1
            continue
        }

        $created = Invoke-ZabbixApi -Method 'hostgroup.create' -Params @{
            name = $groupName
        } -Token $Token -Timeout $Timeout -Retries $Retries
        $groupId = [string](Ensure-Array -Value (Get-PropertyValue -InputObject $created -Names @('groupids') -Default @()))[0]
        $script:ZabbixState.GroupCache[$groupName] = $groupId
        $ids += $groupId
        $Summary.zabbix.groupsCreated += 1
    }

    return @($ids | Select-Object -Unique)
}

function Resolve-TemplateIds {
    param(
        [string[]]$TemplateNames,
        [string]$Token,
        [int]$Timeout,
        [int]$Retries,
        [hashtable]$Summary
    )

    $templateIds = @()
    foreach ($templateNameRaw in $TemplateNames) {
        $templateName = Normalize-Text -Value $templateNameRaw
        if ([string]::IsNullOrWhiteSpace($templateName)) {
            continue
        }

        if ($script:ZabbixState.TemplateCache.ContainsKey($templateName)) {
            $cached = $script:ZabbixState.TemplateCache[$templateName]
            if ($cached) {
                $templateIds += [string]$cached
            } elseif (-not ($Summary.zabbix.templatesMissing -contains $templateName)) {
                $Summary.zabbix.templatesMissing += $templateName
            }
            continue
        }

        $matches = @(Invoke-ZabbixApi -Method 'template.get' -Params @{
                output = @('templateid', 'host', 'name')
                filter = @{
                    host = @($templateName)
                }
            } -Token $Token -Timeout $Timeout -Retries $Retries)

        if ($matches.Count -eq 0) {
            $matches = @(Invoke-ZabbixApi -Method 'template.get' -Params @{
                    output      = @('templateid', 'host', 'name')
                    search      = @{
                        name = $templateName
                    }
                    searchByAny = $true
                } -Token $Token -Timeout $Timeout -Retries $Retries)
            $matches = @($matches | ForEach-Object { ConvertTo-Hashtable -InputObject $_ } | Where-Object {
                    $_.host -eq $templateName -or $_.name -eq $templateName
                })
        } else {
            $matches = @($matches | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })
        }

        if ($matches.Count -eq 1) {
            $templateId = [string]$matches[0].templateid
            $script:ZabbixState.TemplateCache[$templateName] = $templateId
            $templateIds += $templateId
            continue
        }

        $script:ZabbixState.TemplateCache[$templateName] = $null
        if (-not ($Summary.zabbix.templatesMissing -contains $templateName)) {
            $Summary.zabbix.templatesMissing += $templateName
        }
    }

    return @($templateIds | Select-Object -Unique)
}

function Get-ZabbixHostState {
    param(
        [string]$TechnicalName,
        [string]$Token,
        [int]$Timeout,
        [int]$Retries
    )

    $hostResult = @(Invoke-ZabbixApi -Method 'host.get' -Params @{
            filter           = @{
                host = @($TechnicalName)
            }
            output           = @('hostid', 'host', 'name', 'status', 'inventory_mode')
            selectInterfaces = 'extend'
            selectTags       = 'extend'
            selectInventory  = 'extend'
        } -Token $Token -Timeout $Timeout -Retries $Retries)

    if ($hostResult.Count -eq 0) {
        return $null
    }

    $host = ConvertTo-Hashtable -InputObject $hostResult[0]
    $host['groups'] = @(Invoke-ZabbixApi -Method 'hostgroup.get' -Params @{
            hostids = @([string]$host.hostid)
            output  = @('groupid', 'name')
        } -Token $Token -Timeout $Timeout -Retries $Retries | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })
    $host['parentTemplates'] = @(Invoke-ZabbixApi -Method 'template.get' -Params @{
            hostids = @([string]$host.hostid)
            output  = @('templateid', 'host', 'name')
        } -Token $Token -Timeout $Timeout -Retries $Retries | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })
    $host['macros'] = @(Invoke-ZabbixApi -Method 'usermacro.get' -Params @{
            hostids = @([string]$host.hostid)
            output  = 'extend'
        } -Token $Token -Timeout $Timeout -Retries $Retries | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })

    return $host
}

function Convert-ToZabbixInterfacePayload {
    param(
        [hashtable]$Interface,
        [string]$HostId = ''
    )

    $payload = @{
        type  = Get-InterfaceTypeCode -Type $Interface.type
        main  = [int]$Interface.main
        useip = [int]$Interface.useip
        ip    = [string]$Interface.ip
        dns   = [string]$Interface.dns
        port  = [string]$Interface.port
    }
    if (-not [string]::IsNullOrWhiteSpace($HostId)) {
        $payload['hostid'] = [string]$HostId
    }
    if ($Interface.ContainsKey('details') -and $Interface.details.Count -gt 0) {
        $payload['details'] = $Interface.details
    }
    return $payload
}

function Compare-StringSets {
    param(
        [string[]]$Left,
        [string[]]$Right
    )

    $leftSet = @($Left | Sort-Object -Unique)
    $rightSet = @($Right | Sort-Object -Unique)

    if ($leftSet.Count -ne $rightSet.Count) {
        return $false
    }

    for ($i = 0; $i -lt $leftSet.Count; $i++) {
        if ($leftSet[$i] -ne $rightSet[$i]) {
            return $false
        }
    }

    return $true
}

function Merge-InventoryFields {
    param(
        [hashtable]$Current,
        [hashtable]$Desired,
        [string]$Mode
    )

    $result = @{}
    $keys = @($Current.Keys + $Desired.Keys | Select-Object -Unique)
    foreach ($key in $keys) {
        $currentValue = Normalize-Text -Value (Get-PropertyValue -InputObject $Current -Names @($key) -Default '')
        $desiredValue = Normalize-Text -Value (Get-PropertyValue -InputObject $Desired -Names @($key) -Default '')

        if ([string]::IsNullOrWhiteSpace($desiredValue)) {
            if (-not [string]::IsNullOrWhiteSpace($currentValue)) {
                $result[$key] = $currentValue
            }
            continue
        }

        if ($Mode -eq 'ForceMerge') {
            $result[$key] = $desiredValue
            continue
        }

        if ([string]::IsNullOrWhiteSpace($currentValue)) {
            $result[$key] = $desiredValue
        } else {
            $result[$key] = $currentValue
        }
    }

    return $result
}

function Merge-GroupIds {
    param(
        [string[]]$Current,
        [string[]]$Desired
    )

    return @($Current + $Desired | Select-Object -Unique)
}

function Merge-TemplateIds {
    param(
        [string[]]$Current,
        [string[]]$Desired
    )

    return @($Current + $Desired | Select-Object -Unique)
}

function Sync-ZabbixMacros {
    param(
        [string]$HostId,
        [array]$CurrentMacros,
        [array]$DesiredMacros,
        [string]$Mode,
        [switch]$WhatIfOnly,
        [string]$Token,
        [int]$Timeout,
        [int]$Retries,
        [hashtable]$Summary
    )

    $currentMap = @{}
    foreach ($macro in Ensure-Array -Value $CurrentMacros) {
        $currentMap[$macro.macro] = $macro
    }

    foreach ($macro in Ensure-Array -Value $DesiredMacros) {
        $entry = ConvertTo-Hashtable -InputObject $macro
        $name = Normalize-Text -Value $entry.macro
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if (-not $currentMap.ContainsKey($name)) {
            if ($WhatIfOnly) {
                $Summary.zabbix.macrosPlanned += 1
            } else {
                Invoke-ZabbixApi -Method 'usermacro.create' -Params @{
                    hostid      = $HostId
                    macro       = $name
                    value       = [string]$entry.value
                    description = [string]$entry.description
                } -Token $Token -Timeout $Timeout -Retries $Retries | Out-Null
                $Summary.zabbix.macrosCreated += 1
            }
            continue
        }

        $existing = $currentMap[$name]
        $needsUpdate = (
            (Normalize-Text -Value $existing.value) -ne (Normalize-Text -Value $entry.value) -or
            (Normalize-Text -Value $existing.description) -ne (Normalize-Text -Value $entry.description)
        )

        if ($needsUpdate -and $Mode -eq 'ForceMerge') {
            if ($WhatIfOnly) {
                $Summary.zabbix.macrosPlanned += 1
            } else {
                Invoke-ZabbixApi -Method 'usermacro.update' -Params @{
                    hostmacroid = [string]$existing.hostmacroid
                    macro       = $name
                    value       = [string]$entry.value
                    description = [string]$entry.description
                } -Token $Token -Timeout $Timeout -Retries $Retries | Out-Null
                $Summary.zabbix.macrosUpdated += 1
            }
        }
    }
}

function Compare-InterfacesLoosely {
    param(
        [hashtable]$Current,
        [hashtable]$Desired
    )

    if ([int]$Current.type -ne (Get-InterfaceTypeCode -Type $Desired.type)) {
        return $false
    }
    if ([int]$Current.useip -ne [int]$Desired.useip) {
        return $false
    }
    if ((Normalize-Text -Value $Current.ip) -ne (Normalize-Text -Value $Desired.ip)) {
        return $false
    }
    if ((Normalize-Text -Value $Current.dns) -ne (Normalize-Text -Value $Desired.dns)) {
        return $false
    }

    return $true
}

function Sync-ZabbixInterfaces {
    param(
        [string]$HostId,
        [array]$CurrentInterfaces,
        [array]$DesiredInterfaces,
        [string]$Mode,
        [switch]$WhatIfOnly,
        [string]$Token,
        [int]$Timeout,
        [int]$Retries,
        [hashtable]$Summary
    )

    $currentList = @($CurrentInterfaces | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })

    foreach ($desired in Ensure-Array -Value $DesiredInterfaces) {
        $desiredInterface = ConvertTo-Hashtable -InputObject $desired
        $exact = $null
        $loose = $null

        foreach ($current in $currentList) {
            if (
                (Compare-InterfacesLoosely -Current $current -Desired $desiredInterface) -and
                (Normalize-Text -Value $current.port) -eq (Normalize-Text -Value $desiredInterface.port)
            ) {
                $exact = $current
                break
            }

            if ((Compare-InterfacesLoosely -Current $current -Desired $desiredInterface) -and $null -eq $loose) {
                $loose = $current
            }
        }

        if ($exact) {
            continue
        }

        if ($loose -and $Mode -eq 'ForceMerge') {
            if ($WhatIfOnly) {
                $Summary.zabbix.interfacesPlanned += 1
            } else {
                $payload = Convert-ToZabbixInterfacePayload -Interface $desiredInterface
                $payload['interfaceid'] = [string]$loose.interfaceid
                Invoke-ZabbixApi -Method 'hostinterface.update' -Params $payload -Token $Token -Timeout $Timeout -Retries $Retries | Out-Null
                $Summary.zabbix.interfacesUpdated += 1
            }
            continue
        }

        $sameTypeExists = @($currentList | Where-Object { [int]$_.type -eq (Get-InterfaceTypeCode -Type $desiredInterface.type) -and [int]$_.main -eq 1 }).Count -gt 0
        if ($sameTypeExists -and [int]$desiredInterface.main -eq 1) {
            $desiredInterface.main = 0
        }

        if ($WhatIfOnly) {
            $Summary.zabbix.interfacesPlanned += 1
        } else {
            $payload = Convert-ToZabbixInterfacePayload -Interface $desiredInterface -HostId $HostId
            Invoke-ZabbixApi -Method 'hostinterface.create' -Params $payload -Token $Token -Timeout $Timeout -Retries $Retries | Out-Null
            $Summary.zabbix.interfacesCreated += 1
        }
        $currentList += $desiredInterface
    }
}

function Sync-ZabbixHost {
    param(
        [hashtable]$Entry,
        [string]$Token,
        [int]$Timeout,
        [int]$Retries,
        [string]$Mode,
        [switch]$WhatIfOnly,
        [hashtable]$Summary
    )

    $desired = ConvertTo-Hashtable -InputObject $Entry.zabbix
    $groupIds = Ensure-ZabbixGroupIds -GroupNames $desired.groups -Token $Token -Timeout $Timeout -Retries $Retries -WhatIfOnly:$WhatIfOnly -Summary $Summary
    $templateIds = Resolve-TemplateIds -TemplateNames $desired.templates -Token $Token -Timeout $Timeout -Retries $Retries -Summary $Summary

    $current = Get-ZabbixHostState -TechnicalName $desired.host -Token $Token -Timeout $Timeout -Retries $Retries
    if ($null -eq $current) {
        if ($WhatIfOnly) {
            $Summary.zabbix.hostsPlannedCreate += 1
            return
        }

        $payload = @{
            host           = $desired.host
            name           = $desired.name
            status         = Get-HostStatusCode -Status $desired.status
            groups         = @($groupIds | ForEach-Object { @{ groupid = [string]$_ } })
            interfaces     = @($desired.interfaces | ForEach-Object { Convert-ToZabbixInterfacePayload -Interface (ConvertTo-Hashtable -InputObject $_) })
            tags           = $desired.tags
            inventory_mode = 0
            inventory      = $desired.inventory
        }
        if ($templateIds.Count -gt 0) {
            $payload['templates'] = @($templateIds | ForEach-Object { @{ templateid = [string]$_ } })
        }
        if (@($desired.macros).Count -gt 0) {
            $payload['macros'] = $desired.macros
        }

        Invoke-ZabbixApi -Method 'host.create' -Params $payload -Token $Token -Timeout $Timeout -Retries $Retries | Out-Null
        $Summary.zabbix.hostsCreated += 1
        return
    }

    if ($Mode -eq 'CreateOnly') {
        $Summary.zabbix.hostsSkippedExisting += 1
        return
    }

    $currentGroupIds = @($current.groups | ForEach-Object { [string]$_.groupid })
    $desiredGroupIds = @($groupIds | ForEach-Object { [string]$_ })
    $mergedGroupIds = Merge-GroupIds -Current $currentGroupIds -Desired $desiredGroupIds

    $currentTemplateIds = @($current.parentTemplates | ForEach-Object { [string]$_.templateid })
    $desiredTemplateIds = @($templateIds | ForEach-Object { [string]$_ })
    $mergedTemplateIds = Merge-TemplateIds -Current $currentTemplateIds -Desired $desiredTemplateIds

    $mergedTags = Merge-TagObjects -Tags (@($current.tags) + @($desired.tags))
    $currentTagIds = @($current.tags | ForEach-Object { '{0}|{1}' -f $_.tag, $_.value })
    $mergedTagIds = @($mergedTags | ForEach-Object { '{0}|{1}' -f $_.tag, $_.value })

    $mergedInventory = Merge-InventoryFields -Current (ConvertTo-Hashtable -InputObject $current.inventory) -Desired (ConvertTo-Hashtable -InputObject $desired.inventory) -Mode $Mode
    $currentInventoryComparable = ConvertTo-Hashtable -InputObject $current.inventory
    $inventoryChanged = (($mergedInventory | ConvertTo-Json -Depth 20 -Compress) -ne ($currentInventoryComparable | ConvertTo-Json -Depth 20 -Compress))

    $updatePayload = @{
        hostid = [string]$current.hostid
    }
    $hasHostUpdate = $false

    if (-not (Compare-StringSets -Left $currentGroupIds -Right $mergedGroupIds)) {
        $updatePayload['groups'] = @($mergedGroupIds | ForEach-Object { @{ groupid = [string]$_ } })
        $hasHostUpdate = $true
    }
    if (-not (Compare-StringSets -Left $currentTemplateIds -Right $mergedTemplateIds)) {
        $updatePayload['templates'] = @($mergedTemplateIds | ForEach-Object { @{ templateid = [string]$_ } })
        $hasHostUpdate = $true
    }
    if (-not (Compare-StringSets -Left $currentTagIds -Right $mergedTagIds)) {
        $updatePayload['tags'] = $mergedTags
        $hasHostUpdate = $true
    }
    if ($inventoryChanged) {
        $updatePayload['inventory_mode'] = 0
        $updatePayload['inventory'] = $mergedInventory
        $hasHostUpdate = $true
    }
    if ($Mode -eq 'ForceMerge' -and (Normalize-Text -Value $current.name) -ne (Normalize-Text -Value $desired.name)) {
        $updatePayload['name'] = $desired.name
        $hasHostUpdate = $true
    }
    if ($Mode -eq 'ForceMerge' -and [int]$current.status -ne (Get-HostStatusCode -Status $desired.status)) {
        $updatePayload['status'] = Get-HostStatusCode -Status $desired.status
        $hasHostUpdate = $true
    }

    if ($hasHostUpdate) {
        if ($WhatIfOnly) {
            $Summary.zabbix.hostsPlannedUpdate += 1
        } else {
            Invoke-ZabbixApi -Method 'host.update' -Params $updatePayload -Token $Token -Timeout $Timeout -Retries $Retries | Out-Null
            $Summary.zabbix.hostsUpdated += 1
        }
    }

    Sync-ZabbixMacros -HostId ([string]$current.hostid) -CurrentMacros $current.macros -DesiredMacros $desired.macros -Mode $Mode -WhatIfOnly:$WhatIfOnly -Token $Token -Timeout $Timeout -Retries $Retries -Summary $Summary
    Sync-ZabbixInterfaces -HostId ([string]$current.hostid) -CurrentInterfaces $current.interfaces -DesiredInterfaces $desired.interfaces -Mode $Mode -WhatIfOnly:$WhatIfOnly -Token $Token -Timeout $Timeout -Retries $Retries -Summary $Summary
}

function Sync-ZabbixInventory {
    param(
        $ExportData,
        [string]$BaseUrl,
        [string]$Token,
        [string]$PreferredAuthMode,
        [string]$Mode,
        [switch]$WhatIfOnly,
        [int]$Timeout,
        [int]$Retries
    )

    Initialize-ZabbixConnection -BaseUrl $BaseUrl -Token $Token -PreferredAuthMode $PreferredAuthMode -Timeout $Timeout -Retries $Retries

    $summary = @{
        generatedAt = (Get-Date).ToString('o')
        dryRun      = [bool]$WhatIfOnly
        mode        = $Mode
        source      = @{
            plannedHosts  = @(Ensure-Array -Value $ExportData.inventory).Count
            skippedInFile = @(Ensure-Array -Value $ExportData.skipped).Count
        }
        zabbix      = @{
            endpoint             = $script:ZabbixState.Endpoint
            apiVersion           = $script:ZabbixState.ApiVersion
            authMode             = $script:ZabbixState.AuthMode
            groupsCreated        = 0
            groupsPlanned        = 0
            hostsCreated         = 0
            hostsUpdated         = 0
            hostsPlannedCreate   = 0
            hostsPlannedUpdate   = 0
            hostsSkippedExisting = 0
            interfacesCreated    = 0
            interfacesUpdated    = 0
            interfacesPlanned    = 0
            macrosCreated        = 0
            macrosUpdated        = 0
            macrosPlanned        = 0
            templatesMissing     = @()
        }
    }

    $entries = @(Ensure-Array -Value $ExportData.inventory | ForEach-Object { ConvertTo-Hashtable -InputObject $_ })
    foreach ($entry in $entries) {
        Sync-ZabbixHost -Entry $entry -Token $Token -Timeout $Timeout -Retries $Retries -Mode $Mode -WhatIfOnly:$WhatIfOnly -Summary $summary
    }

    return $summary
}

function Validate-RequiredSettings {
    param(
        [string]$ActionName,
        [string]$PrtgUrl,
        [string]$PrtgToken,
        [string]$ZabbixUrl,
        [string]$ZabbixToken,
        [string]$InputJsonPath
    )

    switch ($ActionName) {
        'ExportPrtg' {
            if ([string]::IsNullOrWhiteSpace($PrtgUrl) -or [string]::IsNullOrWhiteSpace($PrtgToken)) {
                throw 'Fuer ExportPrtg werden PRTG-URL und PRTG-Token benoetigt.'
            }
        }
        'SyncZabbix' {
            if ([string]::IsNullOrWhiteSpace($ZabbixUrl) -or [string]::IsNullOrWhiteSpace($ZabbixToken)) {
                throw 'Fuer SyncZabbix werden Zabbix-URL und Zabbix-Token benoetigt.'
            }
            if ([string]::IsNullOrWhiteSpace($InputJsonPath)) {
                throw 'Fuer SyncZabbix wird -InputJson benoetigt.'
            }
        }
        'FullSync' {
            if ([string]::IsNullOrWhiteSpace($PrtgUrl) -or [string]::IsNullOrWhiteSpace($PrtgToken)) {
                throw 'Fuer FullSync werden PRTG-URL und PRTG-Token benoetigt.'
            }
            if ([string]::IsNullOrWhiteSpace($ZabbixUrl) -or [string]::IsNullOrWhiteSpace($ZabbixToken)) {
                throw 'Fuer FullSync werden Zabbix-URL und Zabbix-Token benoetigt.'
            }
        }
    }
}

try {
    Validate-RequiredSettings -ActionName $Action -PrtgUrl $PrtgUrl -PrtgToken $PrtgToken -ZabbixUrl $ZabbixUrl -ZabbixToken $ZabbixToken -InputJsonPath $InputJson
    $mapping = Load-Mapping -Path $MappingFile

    switch ($Action) {
        'ExportPrtg' {
            $export = Get-PrtgInventory -BaseUrl $PrtgUrl -Token $PrtgToken -AuthMode $PrtgAuthMode -Mapping $mapping -Count $PageSize -Timeout $TimeoutSec -Retries $MaxRetries
            Save-JsonFile -Path $OutputJson -Data $export
            Export-InventoryCsv -ExportData $export -Path $OutputCsv
            Write-Log -Level 'INFO' -Message ("Export abgeschlossen. Hosts: {0}, uebersprungen: {1}" -f @($export.inventory).Count, @($export.skipped).Count)
        }
        'SyncZabbix' {
            $export = Load-JsonFile -Path $InputJson
            $summary = Sync-ZabbixInventory -ExportData $export -BaseUrl $ZabbixUrl -Token $ZabbixToken -PreferredAuthMode $ZabbixAuthMode -Mode $Mode -WhatIfOnly:$DryRun -Timeout $TimeoutSec -Retries $MaxRetries
            Save-JsonFile -Path $ReportFile -Data $summary
            Write-Log -Level 'INFO' -Message ("Zabbix-Sync abgeschlossen. Erstellt: {0}, aktualisiert: {1}" -f $summary.zabbix.hostsCreated, $summary.zabbix.hostsUpdated)
        }
        'FullSync' {
            $export = Get-PrtgInventory -BaseUrl $PrtgUrl -Token $PrtgToken -AuthMode $PrtgAuthMode -Mapping $mapping -Count $PageSize -Timeout $TimeoutSec -Retries $MaxRetries
            Save-JsonFile -Path $OutputJson -Data $export
            Export-InventoryCsv -ExportData $export -Path $OutputCsv
            $summary = Sync-ZabbixInventory -ExportData $export -BaseUrl $ZabbixUrl -Token $ZabbixToken -PreferredAuthMode $ZabbixAuthMode -Mode $Mode -WhatIfOnly:$DryRun -Timeout $TimeoutSec -Retries $MaxRetries
            Save-JsonFile -Path $ReportFile -Data $summary
            Write-Log -Level 'INFO' -Message ("FullSync abgeschlossen. Exportierte Hosts: {0}, erstellt: {1}, aktualisiert: {2}" -f @($export.inventory).Count, $summary.zabbix.hostsCreated, $summary.zabbix.hostsUpdated)
        }
    }
} catch {
    Write-Log -Level 'ERROR' -Message $_.Exception.Message
    exit 1
}
