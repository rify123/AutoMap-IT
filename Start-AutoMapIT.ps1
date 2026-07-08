# AutoMap IT
# Copyright (c) 2026 Aziz Arbiine
# Todos los derechos reservados.
# Autor: Aziz Arbiine
# LinkedIn: https://www.linkedin.com/in/aziz-arbiine
# Queda prohibida la copia, distribución, modificación o atribución falsa sin autorización escrita.

param(
    [int]$Port = 8765,
    [string[]]$ComputerName = @("localhost"),
    [switch]$NoBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$PortalPath = Join-Path $Root "index.html"
$DataDir = Join-Path $Root "data"
$StatePath = Join-Path $DataDir "automap-state.json"
$ReportPath = Join-Path $DataDir "automap-report.html"
$SmtpConfigPath = Join-Path $DataDir "smtp-config.json"

if (-not (Test-Path -LiteralPath $PortalPath)) {
    throw "No se encontro index.html junto a Start-AutoMapIT.ps1."
}
if (-not (Test-Path -LiteralPath $DataDir)) {
    New-Item -ItemType Directory -Path $DataDir | Out-Null
}

function New-AutoMapState {
    [ordered]@{
        generatedAt = (Get-Date).ToString("s")
        currentUser = [ordered]@{
            name = [Environment]::UserName
            domain = [Environment]::UserDomainName
            fullName = "$([Environment]::UserDomainName)\$([Environment]::UserName)"
            machine = [Environment]::MachineName
        }
        servers = @()
        automations = @()
        scripts = @()
        tasks = @()
        services = @()
        credentials = @()
        alerts = @()
        reports = @()
        activity = @(
            [ordered]@{ type = "success"; title = "Portal iniciado"; detail = "AutoMap IT corriendo desde PowerShell"; at = (Get-Date).ToString("s") }
        )
        scanSummary = [ordered]@{
            lastScan = "Pendiente"
            targets = @()
            servers = 0
            scripts = 0
            tasks = 0
            services = 0
            automations = 0
            alerts = 0
            message = "Aun no se ha ejecutado ningun escaneo."
        }
        configSections = [ordered]@{}
    }
}

function Get-State {
    if (-not (Test-Path -LiteralPath $StatePath)) {
        New-AutoMapState | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $StatePath -Encoding UTF8
    }
    $state = Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json
    $defaults = New-AutoMapState
    foreach ($key in $defaults.Keys) {
        if (-not ($state.PSObject.Properties.Name -contains $key)) {
            $state | Add-Member -MemberType NoteProperty -Name $key -Value $defaults[$key]
        }
    }
    $oldDemoNames = @("SRV-DC-01", "SRV-FS-01", "SRV-APP-01", "SRV-SQL-01")
    $serverNames = @($state.servers | ForEach-Object { $_.name })
    if ($serverNames.Count -eq 4 -and @(Compare-Object $oldDemoNames $serverNames -SyncWindow 0).Count -eq 0) {
        $state.servers = @()
        $state.automations = @()
        $state.scripts = @()
        $state.tasks = @()
        $state.services = @()
        $state.credentials = @()
        $state.alerts = @()
        $state.reports = @()
        $state.scanSummary = $defaults.scanSummary
        Save-State $state
    }
    $state.currentUser = [pscustomobject]@{
        name = [Environment]::UserName
        domain = [Environment]::UserDomainName
        fullName = "$([Environment]::UserDomainName)\$([Environment]::UserName)"
        machine = [Environment]::MachineName
    }
    $state
}

function Save-State($State) {
    $State.generatedAt = (Get-Date).ToString("s")
    $State | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $StatePath -Encoding UTF8
}

function Get-StatusText([int]$StatusCode) {
    switch ($StatusCode) {
        200 { "OK" }
        404 { "Not Found" }
        405 { "Method Not Allowed" }
        500 { "Internal Server Error" }
        default { "OK" }
    }
}

function Send-Bytes($Context, [byte[]]$Bytes, [string]$ContentType, [int]$StatusCode = 200, [hashtable]$Headers = @{}) {
    $statusText = Get-StatusText $StatusCode
    $headerLines = @(
        "HTTP/1.1 $StatusCode $statusText",
        "Content-Type: $ContentType",
        "Content-Length: $($Bytes.Length)",
        "Connection: close",
        "Cache-Control: no-store"
    )
    foreach ($key in $Headers.Keys) {
        $headerLines += "${key}: $($Headers[$key])"
    }
    $headerText = ($headerLines -join "`r`n") + "`r`n`r`n"
    $headerBytes = [Text.Encoding]::UTF8.GetBytes($headerText)
    $Context.Stream.Write($headerBytes, 0, $headerBytes.Length)
    $Context.Stream.Write($Bytes, 0, $Bytes.Length)
    $Context.Stream.Flush()
}

function Send-Text($Context, [string]$Text, [string]$ContentType = "text/plain; charset=utf-8", [int]$StatusCode = 200) {
    Send-Bytes $Context ([Text.Encoding]::UTF8.GetBytes($Text)) $ContentType $StatusCode
}

function Send-Json($Context, $Value, [int]$StatusCode = 200) {
    Send-Text $Context ($Value | ConvertTo-Json -Depth 10) "application/json; charset=utf-8" $StatusCode
}

function ConvertTo-CsvText($Rows) {
    [object[]]$rowsArray = @($Rows)
    if (($rowsArray | Measure-Object).Count -eq 0) { return "" }
    $translatedRows = foreach ($row in $rowsArray) {
        $item = [ordered]@{}
        foreach ($prop in $row.PSObject.Properties) {
            $item[(Get-ReportDisplayName $prop.Name)] = $prop.Value
        }
        [pscustomobject]$item
    }
    ($translatedRows | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
}

function Get-AuthorMark {
    "Autor: Aziz Arbiine | Copyright (c) 2026 Aziz Arbiine | Todos los derechos reservados | LinkedIn: https://www.linkedin.com/in/aziz-arbiine"
}

function ConvertTo-TxtText([string]$Title, $Rows) {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("AutoMap IT - $Title")
    $lines.Add("Generado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add((Get-AuthorMark))
    $lines.Add("")
    foreach ($row in @($Rows)) {
        $pairs = foreach ($prop in $row.PSObject.Properties) {
            "$(Get-ReportDisplayName $prop.Name): $($prop.Value)"
        }
        $lines.Add(($pairs -join " | "))
    }
    if ($lines.Count -le 3) { $lines.Add("Sin registros") }
    $lines.Add("")
    $lines.Add((Get-AuthorMark))
    $lines.Add("Queda prohibida la atribucion falsa de autoria sin autorizacion escrita.")
    $lines -join "`r`n"
}

function Get-ReportDisplayName([string]$Name) {
    switch ($Name) {
        "name" { "Nombre" }
        "type" { "Tipo" }
        "status" { "Estado" }
        "criticality" { "Criticidad" }
        "owner" { "Propietario" }
        "extension" { "Extension" }
        "path" { "Ruta" }
        "server" { "Servidor" }
        "lastUse" { "Ultimo uso" }
        "lastError" { "Ultimo error" }
        "risk" { "Riesgo" }
        "lastRun" { "Ultima ejecucion" }
        "nextRun" { "Proxima ejecucion" }
        "account" { "Cuenta" }
        "schedule" { "Programacion" }
        "result" { "Resultado" }
        "system" { "Sistema" }
        "startup" { "Modo de inicio" }
        "systems" { "Sistemas" }
        "script" { "Script" }
        "source" { "Fuente" }
        "location" { "Ubicacion" }
        "severity" { "Severidad" }
        "origin" { "Origen" }
        "asset" { "Activo" }
        "createdAt" { "Fecha" }
        "template" { "Plantilla" }
        "format" { "Formato" }
        "size" { "Tamano" }
        "Seccion" { "Seccion" }
        "Metrica" { "Metrica" }
        "Valor" { "Valor" }
        "Detalle" { "Detalle" }
        default { $Name }
    }
}

function Get-ReportHeaders($Rows, [string[]]$Preferred = @()) {
    [object[]]$rowsArray = @($Rows)
    [string[]]$preferredArray = @($Preferred)
    if (($preferredArray | Measure-Object).Count -gt 0) { return $preferredArray }
    if (($rowsArray | Measure-Object).Count -gt 0) { return @($rowsArray[0].PSObject.Properties.Name) }
    @()
}

function Get-ReportBadgeClass([object]$Value) {
    $text = [string]$Value
    if ($text -match "Critic|Critica|Alto|Alta|Error|Fall|Abierto") { return "badge high" }
    if ($text -match "Medio|Media|Advert|Pendiente") { return "badge medium" }
    if ($text -match "Bajo|Baja|Exito|Detectado|Habilitado|Ejecutando|En linea|Generado") { return "badge low" }
    "badge neutral"
}

function New-StyledReportTable([string[]]$Headers, $Rows, [string[]]$BadgeColumns = @("criticality", "risk", "Riesgo", "Estado", "status", "severity", "Valor")) {
    [string[]]$headersArray = @($Headers)
    [object[]]$rowsArray = @($Rows)
    if (($headersArray | Measure-Object).Count -eq 0 -or ($rowsArray | Measure-Object).Count -eq 0) {
        return '<div class="empty">Sin registros</div>'
    }
    $thead = ($headersArray | ForEach-Object { "<th>$(HtmlEncode (Get-ReportDisplayName $_))</th>" }) -join ""
    $bodyRows = foreach ($row in $rowsArray) {
        $cells = foreach ($h in $headersArray) {
            $value = $row.$h
            if ($BadgeColumns -contains $h) {
                "<td><span class=""$(Get-ReportBadgeClass $value)"">$(HtmlEncode $value)</span></td>"
            } else {
                "<td>$(HtmlEncode $value)</td>"
            }
        }
        "<tr>$($cells -join '')</tr>"
    }
    "<table><thead><tr>$thead</tr></thead><tbody>$($bodyRows -join "`r`n")</tbody></table>"
}

function Get-ReportCss {
@"
body{font-family:Segoe UI,Arial,sans-serif;color:#102047;margin:0;background:#eef2f7}
.page{max-width:1120px;margin:0 auto;background:#fff;min-height:100vh;padding:34px 42px}
.reportTop{display:flex;justify-content:space-between;gap:18px;align-items:flex-start;border-bottom:1px solid #e6edf7;padding-bottom:18px;margin-bottom:22px}
h1{font-size:24px;line-height:1.2;margin:0;color:#16295a}h2{font-size:17px;margin:28px 0 10px;color:#16295a}
.muted{color:#64748b;font-size:12px}.grid{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin:18px 0 24px}
.card{border:1px solid #e6edf7;border-radius:8px;padding:14px;background:#fbfdff}.num{font-size:26px;font-weight:900;color:#4338ca}
.summary{border:1px solid #e6edf7;border-radius:8px;background:#f8fafc;padding:14px;margin-bottom:14px}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e6edf7;border-radius:8px;overflow:visible;table-layout:auto}
th{background:#f8fafc;color:#35476b;font-size:12px;text-align:left;padding:13px 12px;border-bottom:1px solid #e6edf7}
td{font-size:12px;color:#102047;padding:14px 12px;border-bottom:1px solid #e6edf7;vertical-align:top;white-space:normal;overflow-wrap:anywhere;word-break:break-word}
tbody tr:nth-child(even) td{background:#fbfdff}td:first-child{font-weight:800;color:#4338ca}.badge{display:inline-block;border-radius:7px;padding:5px 9px;font-weight:800;font-size:12px}
.badge.high{background:#fee2e2;color:#dc2626}.badge.medium{background:#fff4d9;color:#a16207}.badge.low{background:#dcfce7;color:#15803d}.badge.neutral{background:#eef2f7;color:#64748b}
.empty{border:1px dashed #cbd5e1;border-radius:8px;color:#64748b;padding:22px;text-align:center;background:#f8fafc}
.printBtn{border:1px solid #dbe3f0;background:#4338ca;color:#fff;border-radius:8px;padding:10px 14px;font-weight:800;cursor:pointer}
.authorMark{margin:14px 0 18px;border:1px solid #dbe3f0;border-left:4px solid #4338ca;border-radius:8px;background:#fbfdff;color:#334155;padding:10px 12px;font-size:12px;font-weight:700}
.reportFooter{margin-top:26px;border-top:1px solid #e6edf7;padding-top:12px;color:#475569;font-size:12px;font-weight:700}
@media print{body{background:#fff}.page{padding:18px;max-width:none}.printBtn{display:none}.grid{grid-template-columns:repeat(3,1fr)}}
"@
}

function ConvertTo-ExcelHtml([string]$Title, $Rows) {
    [object[]]$rowsArray = @($Rows)
    $headers = Get-ReportHeaders -Rows $rowsArray
    $table = New-StyledReportTable -Headers $headers -Rows $rowsArray
    @"
<!doctype html>
<html lang="es">
<head><meta charset="utf-8"><style>$(Get-ReportCss)</style></head>
<body><main class="page"><div class="reportTop"><div><h1>AutoMap IT - $(HtmlEncode $Title)</h1><div class="muted">Generado: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div></div></div><div class="authorMark">$(HtmlEncode (Get-AuthorMark))</div>$table<div class="reportFooter">$(HtmlEncode (Get-AuthorMark))<br>Queda prohibida la atribucion falsa de autoria sin autorizacion escrita.</div></main></body>
</html>
"@
}

function Escape-PdfText([object]$Value) {
    $text = [string]$Value
    $text = $text.Replace("\", "\\").Replace("(", "\(").Replace(")", "\)")
    $text = $text -replace "[\r\n\t]+", " "
    $text
}

function Split-PdfCellLines([object]$Value, [int]$MaxChars) {
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return @("-") }
    $max = [Math]::Max(6, $MaxChars)
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($text -split "\s+")) {
        if ($part.Length -gt $max) {
            if ($lines.Count -eq 0 -or $lines[$lines.Count - 1].Length -gt 0) { $lines.Add("") }
            $chunk = $part
            while ($chunk.Length -gt $max) {
                $lines.Add($chunk.Substring(0, $max))
                $chunk = $chunk.Substring($max)
            }
            if ($chunk.Length -gt 0) { $lines.Add($chunk) }
            continue
        }
        if ($lines.Count -eq 0) {
            $lines.Add($part)
            continue
        }
        $current = $lines[$lines.Count - 1]
        if (($current.Length + 1 + $part.Length) -le $max) {
            $lines[$lines.Count - 1] = ($current + " " + $part).Trim()
        } else {
            $lines.Add($part)
        }
    }
    @($lines)
}

function Get-PdfColumnWidths([string[]]$Headers, [int]$AvailableWidth) {
    if ((@($Headers) | Measure-Object).Count -eq 0) { return @() }
    $weights = foreach ($h in @($Headers)) {
        switch -Regex ($h) {
            "path|location|Detalle|Ruta|Ubicacion" { 3.2; break }
            "name|Nombre|owner|Propietario|account|Cuenta" { 2.1; break }
            "server|Servidor|system|Sistema|asset|Activo" { 1.7; break }
            "last|created|Fecha|Ultima|Proxima|schedule|Programacion" { 1.45; break }
            "result|Resultado|status|Estado" { 1.35; break }
            "extension|type|risk|Riesgo|severity|Severidad|format|size" { 1.0; break }
            default { 1.25 }
        }
    }
    $sum = ($weights | Measure-Object -Sum).Sum
    if (-not $sum -or $sum -le 0) { $sum = 1 }
    [object[]]$widths = @(foreach ($w in $weights) { [Math]::Max(38, [int][Math]::Floor($AvailableWidth * ($w / $sum))) })
    $diff = $AvailableWidth - (($widths | Measure-Object -Sum).Sum)
    if (($widths | Measure-Object).Count -gt 0) { $widths[$widths.Count - 1] = [int]$widths[$widths.Count - 1] + $diff }
    @($widths)
}

function ConvertTo-PdfBytes([string]$Title, $Rows) {
    [object[]]$rowsArray = @($Rows)
    [string[]]$headers = @(Get-ReportHeaders -Rows $rowsArray)
    $pageContents = New-Object System.Collections.Generic.List[string]
    $pageWidth = 842
    $pageHeight = 612
    $left = 28
    $availableWidth = 786
    $top = 548
    $bottom = 54
    $fontSize = 7
    $lineHeight = 8
    $headerHeight = 26
    $generated = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $colWidths = Get-PdfColumnWidths -Headers $headers -AvailableWidth $availableWidth

    $newPage = {
        $c = ""
        $c += "q 0.94 0.96 0.99 rg 0 0 $pageWidth $pageHeight re f Q`n"
        $c += "q 1 1 1 rg 18 22 806 568 re f Q`n"
        $c += "BT /F2 17 Tf 28 574 Td (AutoMap IT - $(Escape-PdfText $Title)) Tj ET`n"
        $c += "BT /F1 9 Tf 28 558 Td (Generado: $generated) Tj ET`n"
        $c += "BT /F1 8 Tf 28 42 Td ($(Escape-PdfText (Get-AuthorMark))) Tj ET`n"
        $c += "BT /F1 7 Tf 28 30 Td (Queda prohibida la atribucion falsa de autoria sin autorizacion escrita.) Tj ET`n"
        if ((@($headers) | Measure-Object).Count -gt 0) {
            $c += "q 0.97 0.98 1 rg $left $top $availableWidth $headerHeight re f Q`n"
            $cx = $left
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $label = Get-ReportDisplayName $headers[$i]
                $labelLines = Split-PdfCellLines $label ([Math]::Max(6, [int](($colWidths[$i] - 8) / 4.2)))
                $ly = $top + 15
                foreach ($ll in @($labelLines | Select-Object -First 2)) {
                    $c += "BT /F2 7 Tf $($cx + 4) $ly Td ($(Escape-PdfText $ll)) Tj ET`n"
                    $ly -= 8
                }
                $cx += $colWidths[$i]
            }
        }
        $c
    }

    if ((@($headers) | Measure-Object).Count -eq 0 -or ($rowsArray | Measure-Object).Count -eq 0) {
        $content = & $newPage
        $content += "q 0.97 0.98 1 rg 28 500 786 34 re f Q`n"
        $content += "BT /F1 10 Tf 388 512 Td (Sin registros) Tj ET`n"
        $pageContents.Add($content)
    } else {
        $content = & $newPage
        $y = $top - 24
        foreach ($row in $rowsArray) {
            $cellLines = @()
            $maxLines = 1
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $h = $headers[$i]
                $maxChars = [Math]::Max(6, [int](($colWidths[$i] - 8) / 3.8))
                $lines = @(Split-PdfCellLines $row.$h $maxChars)
                $cellLines += ,$lines
                if ($lines.Count -gt $maxLines) { $maxLines = $lines.Count }
            }
            $rowH = [Math]::Max(22, 10 + ($maxLines * $lineHeight))
            if (($y - $rowH) -lt $bottom) {
                $pageContents.Add($content)
                $content = & $newPage
                $y = $top - 24
            }
            $content += "q 1 1 1 rg $left $($y - $rowH + 4) $availableWidth $rowH re f Q`n"
            $content += "q 0.90 0.93 0.97 rg $left $($y - $rowH + 4) $availableWidth 0.6 re f Q`n"
            $cx = $left
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $font = if ($i -eq 0) { "/F2" } else { "/F1" }
                $ly = $y - 8
                foreach ($line in @($cellLines[$i])) {
                    $content += "BT $font $fontSize Tf $($cx + 4) $ly Td ($(Escape-PdfText $line)) Tj ET`n"
                    $ly -= $lineHeight
                }
                $cx += $colWidths[$i]
            }
            $y -= $rowH
        }
        $pageContents.Add($content)
    }

    $kids = New-Object System.Collections.Generic.List[string]
    $objects = New-Object System.Collections.Generic.List[string]
    $objects.Add("<< /Type /Catalog /Pages 2 0 R >>")
    $objects.Add("__PAGES__")
    $objects.Add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    $objects.Add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
    for ($p = 0; $p -lt $pageContents.Count; $p++) {
        $pageObj = $objects.Count + 1
        $contentObj = $objects.Count + 2
        $kids.Add("$pageObj 0 R")
        $objects.Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 $pageWidth $pageHeight] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> /Contents $contentObj 0 R >>")
        $content = $pageContents[$p]
        $objects.Add("<< /Length $([Text.Encoding]::GetEncoding(1252).GetByteCount($content)) >>`nstream`n$content`nendstream")
    }
    $objects[1] = "<< /Type /Pages /Kids [$($kids -join ' ')] /Count $($pageContents.Count) >>"
    $encoding = [Text.Encoding]::GetEncoding(1252)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("%PDF-1.4`n")
    $offsets = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $objects.Count; $i++) {
        $offsets.Add($encoding.GetByteCount($sb.ToString()))
        [void]$sb.Append("$($i + 1) 0 obj`n$($objects[$i])`nendobj`n")
    }
    $xref = $encoding.GetByteCount($sb.ToString())
    [void]$sb.Append("xref`n0 $($objects.Count + 1)`n0000000000 65535 f `n")
    foreach ($off in $offsets) { [void]$sb.Append(("{0:0000000000} 00000 n `n" -f $off)) }
    [void]$sb.Append("trailer`n<< /Size $($objects.Count + 1) /Root 1 0 R >>`nstartxref`n$xref`n%%EOF")
    $encoding.GetBytes($sb.ToString())
}

function New-InventoryExportRows($State) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($State.servers)) {
        $rows.Add([ordered]@{ Tipo="Servidor"; Nombre=(Get-ExportProp $r "name"); Estado=(Get-ExportProp $r "status"); SistemaIP=(Get-ExportProp $r "ip"); Detalle=(Get-ExportProp $r "os"); PropietarioCuenta=(Get-ExportProp $r "owner"); UltimaActividad="-"; ResultadoSeveridad="-"; Riesgo=(Get-ExportProp $r "risk"); HallazgosImpacto=(Get-ExportProp $r "findings" 0); ProximaModo=(Get-ExportProp $r "environment" "Produccion") })
    }
    foreach ($r in @($State.automations)) {
        $rows.Add([ordered]@{ Tipo="Automatizacion"; Nombre=(Get-ExportProp $r "name"); Estado=(Get-ExportProp $r "status"); SistemaIP=(Get-ExportProp $r "system"); Detalle=(Get-ExportProp $r "type"); PropietarioCuenta=(Get-ExportProp $r "owner"); UltimaActividad=(Get-ExportProp $r "lastRun"); ResultadoSeveridad=(Get-ExportProp $r "result"); Riesgo=(Get-ExportProp $r "criticality"); HallazgosImpacto="-"; ProximaModo="-" })
    }
    foreach ($r in @($State.scripts)) {
        $rows.Add([ordered]@{ Tipo="Script"; Nombre=(Get-ExportProp $r "name"); Estado="Detectado"; SistemaIP=(Get-ExportProp $r "server"); Detalle=(Get-ExportProp $r "path"); PropietarioCuenta=(Get-ExportProp $r "owner"); UltimaActividad=(Get-ExportProp $r "lastUse"); ResultadoSeveridad=(Get-ExportProp $r "lastError"); Riesgo=(Get-ExportProp $r "risk"); HallazgosImpacto="-"; ProximaModo="-" })
    }
    foreach ($r in @($State.tasks)) {
        $rows.Add([ordered]@{ Tipo="Tarea programada"; Nombre=(Get-ExportProp $r "name"); Estado=(Get-ExportProp $r "status"); SistemaIP=(Get-ExportProp $r "server"); Detalle=(Get-ExportProp $r "schedule"); PropietarioCuenta=(Get-ExportProp $r "account"); UltimaActividad=(Get-ExportProp $r "lastRun"); ResultadoSeveridad=(Get-ExportProp $r "result"); Riesgo=(Get-ExportProp $r "risk"); HallazgosImpacto="-"; ProximaModo=(Get-ExportProp $r "nextRun") })
    }
    foreach ($r in @($State.services)) {
        $rows.Add([ordered]@{ Tipo="Servicio"; Nombre=(Get-ExportProp $r "name"); Estado=(Get-ExportProp $r "status"); SistemaIP=(Get-ExportProp $r "server"); Detalle=(Get-ExportProp $r "script"); PropietarioCuenta=(Get-ExportProp $r "account"); UltimaActividad="-"; ResultadoSeveridad="-"; Riesgo=(Get-ExportProp $r "risk"); HallazgosImpacto=(Get-ExportProp $r "systems" 0); ProximaModo=(Get-ExportProp $r "startup") })
    }
    foreach ($r in @($State.credentials)) {
        $rows.Add([ordered]@{ Tipo="Credencial"; Nombre=(Get-ExportProp $r "name"); Estado=(Get-ExportProp $r "status"); SistemaIP=(Get-ExportProp $r "location"); Detalle=(Get-ExportProp $r "source"); PropietarioCuenta=(Get-ExportProp $r "account"); UltimaActividad="-"; ResultadoSeveridad="-"; Riesgo=(Get-ExportProp $r "risk"); HallazgosImpacto="-"; ProximaModo="-" })
    }
    foreach ($r in @($State.alerts)) {
        $rows.Add([ordered]@{ Tipo="Alerta"; Nombre=(Get-ExportProp $r "name"); Estado=(Get-ExportProp $r "status"); SistemaIP=(Get-ExportProp $r "asset"); Detalle=(Get-ExportProp $r "origin"); PropietarioCuenta=(Get-ExportProp $r "owner"); UltimaActividad=(Get-ExportProp $r "createdAt"); ResultadoSeveridad=(Get-ExportProp $r "severity"); Riesgo=(Get-ExportProp $r "severity"); HallazgosImpacto="-"; ProximaModo="-" })
    }
    $rows
}

function New-ExecutiveReportRows($State) {
    @(
        [pscustomobject]@{ Seccion="Resumen"; Metrica="Servidores"; Valor=@($State.servers).Count; Detalle="Servidores descubiertos" }
        [pscustomobject]@{ Seccion="Resumen"; Metrica="Automatizaciones"; Valor=@($State.automations).Count; Detalle="Automatizaciones detectadas" }
        [pscustomobject]@{ Seccion="Resumen"; Metrica="Scripts"; Valor=@($State.scripts).Count; Detalle="PowerShell, BAT, CMD y VBS" }
        [pscustomobject]@{ Seccion="Resumen"; Metrica="Tareas programadas"; Valor=@($State.tasks).Count; Detalle="Tareas de Windows" }
        [pscustomobject]@{ Seccion="Resumen"; Metrica="Servicios"; Valor=@($State.services).Count; Detalle="Servicios que lanzan scripts" }
        [pscustomobject]@{ Seccion="Riesgo"; Metrica="Alertas activas"; Valor=@($State.alerts | Where-Object { $_.status -eq "Activa" }).Count; Detalle="Pendientes de reconocer" }
        [pscustomobject]@{ Seccion="Riesgo"; Metrica="Credenciales"; Valor=@($State.credentials).Count; Detalle="Cuentas y secretos detectados" }
    )
}

function Get-ExportProp($Object, [string]$Name, [object]$Fallback = "-") {
    if ($null -eq $Object) { return $Fallback }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value -or [string]::IsNullOrWhiteSpace([string]$prop.Value)) { return $Fallback }
    $prop.Value
}

function Get-ExportRows($State, [string]$Name) {
    switch ($Name) {
        "servers" { $State.servers }
        "automations" { $State.automations }
        "scripts" { $State.scripts }
        "tasks" { $State.tasks }
        "services" { $State.services }
        "credentials" { $State.credentials }
        "alerts" { $State.alerts }
        "reports" { $State.reports }
        "activity" { $State.activity }
        "inventory" { New-InventoryExportRows $State }
        "executive-report" { New-ExecutiveReportRows $State }
        default { @() }
    }
}

function Get-CsvHeader([string]$Name) {
    switch ($Name) {
        "servers" { '"Nombre","IP","SO","Estado","Riesgo","Hallazgos","Propietario"' }
        "automations" { '"Nombre","Tipo","Estado","Criticidad","Propietario","Ultima ejecucion","Resultado","Sistema"' }
        "scripts" { '"Nombre","Extension","Ruta","Servidor","Propietario","Ultimo uso","Ultimo error","Riesgo"' }
        "tasks" { '"Nombre","Servidor","Cuenta","Programacion","Ultima ejecucion","Proxima ejecucion","Resultado","Riesgo","Estado"' }
        "services" { '"Nombre","Estado","Modo de inicio","Riesgo","Sistemas","Script","Servidor","Cuenta"' }
        "credentials" { '"Nombre","Cuenta","Fuente","Ubicacion","Riesgo","Estado"' }
        "alerts" { '"Severidad","Origen","Nombre","Activo","Estado","Propietario","Fecha"' }
        "reports" { '"Nombre","Plantilla","Formato","Tamano","Estado","Fecha"' }
        default { '"Nombre","Valor"' }
    }
}

function HtmlEncode([object]$Value) {
    [Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-ReportHtml($State) {
    $servers = @($State.servers)
    $scripts = @($State.scripts)
    $tasks = @($State.tasks)
    $services = @($State.services)
    $automations = @($State.automations)
    $alerts = @($State.alerts)
    $highRisk = @($automations | Where-Object { $_.criticality -match "Alto|Alta|Critic|Critica|Cr.tic" })
    $errors = @($tasks | Where-Object { $_.result -match "Error|Fall" })
    $automationRows = @($automations | Select-Object -First 250)
    $automationTable = New-StyledReportTable -Headers @("name", "type", "status", "criticality", "owner", "lastRun", "result", "system") -Rows $automationRows -BadgeColumns @("criticality", "status", "result")
    $css = Get-ReportCss
    @"
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>AutoMap IT - Informe ejecutivo</title>
<style>
$css
</style>
</head>
<body>
<main class="page">
<div class="reportTop">
<div>
<h1>AutoMap IT - Informe ejecutivo</h1>
<div class="muted">Generado: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</div>
</div>
<button class="printBtn" onclick="window.print()">Imprimir / Guardar como PDF</button>
</div>
<div class="authorMark">$(HtmlEncode (Get-AuthorMark))</div>
<div class="grid">
<div class="card"><div class="num">$($servers.Count)</div><div>Servidores</div></div>
<div class="card"><div class="num">$($automations.Count)</div><div>Automatizaciones</div></div>
<div class="card"><div class="num">$($scripts.Count)</div><div>Scripts</div></div>
<div class="card"><div class="num">$($tasks.Count)</div><div>Tareas programadas</div></div>
<div class="card"><div class="num">$($services.Count)</div><div>Servicios con scripts</div></div>
<div class="card"><div class="num">$($highRisk.Count)</div><div>Riesgo alto</div></div>
</div>
<h2>Resumen de riesgo</h2>
<div class="summary">Alertas activas: <b>$($alerts.Count)</b>. Tareas con error: <b>$($errors.Count)</b>. Automatizaciones de riesgo alto: <b>$($highRisk.Count)</b>.</div>
<h2>Automatizaciones detectadas</h2>
$automationTable
<div class="reportFooter">$(HtmlEncode (Get-AuthorMark))<br>Queda prohibida la atribucion falsa de autoria sin autorizacion escrita.</div>
</main>
</body>
</html>
"@
}

function Get-Risk {
    param(
        [string]$Account,
        [string]$Path,
        [object]$LastResult,
        [string]$Status
    )
    $text = "$Account $Path $LastResult $Status"
    if ($text -match "LocalSystem|SYSTEM|Administrator|Admin|Temp|Users\\Public|AppData|Error|Fall|Cr.tic|1067|1|2") { return "Alto" }
    if ($text -match "Advert|Manual|Deshabil|Disabled|Unknown|-") { return "Medio" }
    "Bajo"
}

function Convert-TaskResult {
    param([object]$Code)
    if ($null -eq $Code -or "$Code" -eq "") { return "Sin datos" }
    if ([int]$Code -eq 0) { return "Exito" }
    "Error $Code"
}

function Get-FirstPropertyValue {
    param(
        [object]$Object,
        [string[]]$Names
    )
    foreach ($name in $Names) {
        $prop = $Object.PSObject.Properties | Where-Object { $_.Name -eq $name } | Select-Object -First 1
        if ($prop -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
            return [string]$prop.Value
        }
    }
    ""
}

function Get-ServerIPv4 {
    param([string]$Computer)

    $candidates = @()
    $isLocal = [string]::IsNullOrWhiteSpace($Computer) -or $Computer -in @("localhost", ".", $env:COMPUTERNAME)
    try {
        if ($isLocal) {
            $defaultRoute = Get-NetRoute -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue |
                Sort-Object RouteMetric, InterfaceMetric |
                Select-Object -First 1
            if ($defaultRoute) {
                $candidates += Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $defaultRoute.InterfaceIndex -ErrorAction SilentlyContinue |
                    ForEach-Object { $_.IPAddress } |
                    Where-Object { $_ -match "^\d{1,3}(\.\d{1,3}){3}$" -and $_ -notmatch "^169\.254\." -and $_ -ne "127.0.0.1" }
            }
            try {
                $blocks = (ipconfig.exe 2>$null) -join "`n" -split "(\r?\n){2,}"
                foreach ($block in $blocks) {
                    $ipv4 = [regex]::Match($block, "IPv4.*?:\s*([0-9]{1,3}(\.[0-9]{1,3}){3})")
                    $gateway = [regex]::Match($block, "(Puerta de enlace predeterminada|Default Gateway).*?:\s*([0-9]{1,3}(\.[0-9]{1,3}){3})")
                    if ($ipv4.Success -and $gateway.Success) { $candidates += $ipv4.Groups[1].Value }
                }
                foreach ($block in $blocks) {
                    if ($block -match "VMware|vEthernet|Virtual|Loopback") { continue }
                    $ipv4 = [regex]::Match($block, "IPv4.*?:\s*([0-9]{1,3}(\.[0-9]{1,3}){3})")
                    if ($ipv4.Success) { $candidates += $ipv4.Groups[1].Value }
                }
            } catch {}
            $candidates += Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop |
                ForEach-Object { $_.IPAddress } |
                Where-Object { $_ -match "^\d{1,3}(\.\d{1,3}){3}$" -and $_ -notmatch "^169\.254\." -and $_ -ne "127.0.0.1" }
        }
    } catch {}

    try {
        $names = if ($isLocal) { @($env:COMPUTERNAME, [System.Net.Dns]::GetHostName()) } else { @($Computer) }
        foreach ($name in @($names | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)) {
            try {
                $candidates += Resolve-DnsName -Name $name -Type A -ErrorAction SilentlyContinue |
                    ForEach-Object { $_.IPAddress } |
                    Where-Object { $_ -and $_ -notmatch "^169\.254\." -and $_ -ne "127.0.0.1" }
            } catch {}
            $dns = [System.Net.Dns]::GetHostAddresses($name) |
                Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
                ForEach-Object { $_.IPAddressToString } |
                Where-Object { $_ -and $_ -notmatch "^169\.254\." -and $_ -ne "127.0.0.1" }
            $candidates += $dns
        }
    } catch {}

    try {
        if (-not $isLocal -and -not [string]::IsNullOrWhiteSpace($Computer)) {
            $candidates += Test-Connection -ComputerName $Computer -Count 1 -ErrorAction SilentlyContinue |
                ForEach-Object {
                    if ($_.PSObject.Properties["IPV4Address"] -and $_.IPV4Address) { $_.IPV4Address.IPAddressToString }
                    elseif ($_.PSObject.Properties["Address"] -and $_.Address) { [string]$_.Address }
                } |
                Where-Object { $_ -and $_ -match "^\d{1,3}(\.\d{1,3}){3}$" -and $_ -notmatch "^169\.254\." -and $_ -ne "127.0.0.1" }
        }
    } catch {}

    try {
        if ($isLocal) {
            $ipconfig = ipconfig.exe 2>$null
            $candidates += @($ipconfig | Select-String -Pattern "IPv4.*?:\s*([0-9]{1,3}(\.[0-9]{1,3}){3})" | ForEach-Object { $_.Matches[0].Groups[1].Value }) |
                Where-Object { $_ -and $_ -notmatch "^169\.254\." -and $_ -ne "127.0.0.1" }
        }
    } catch {}

    $ip = @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique | Select-Object -First 1)
    if ($ip.Count -gt 0) { return [string]$ip[0] }
    "-"
}

function Format-WindowsServerOS {
    param([string]$Name)

    $value = [string]$Name
    $value = $value -replace "^Microsoft\s+", ""
    $value = $value -replace "^Windows\s+", ""
    if ([string]::IsNullOrWhiteSpace($value)) { return "Windows" }
    $value
}

function Get-ServerResources {
    $cpu = $null
    $memory = $null
    $disk = $null

    try {
        $cpuRows = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop)
        if ($cpuRows.Count -gt 0) {
            $cpu = [math]::Round((($cpuRows | Measure-Object -Property LoadPercentage -Average).Average), 0)
        }
    } catch {
        try {
            $counter = Get-Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
            $cpu = [math]::Round($counter.CounterSamples[0].CookedValue, 0)
        } catch {
            try {
                $counter = Get-Counter "\Información del procesador(_Total)\% de utilidad del procesador" -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
                $cpu = [math]::Round($counter.CounterSamples[0].CookedValue, 0)
            } catch {
                try {
                    $counterPath = Get-Counter -ListSet *procesador* -ErrorAction SilentlyContinue |
                        ForEach-Object { $_.Counter } |
                        Where-Object { $_ -match "utilidad.*procesador" } |
                        Select-Object -First 1
                    if ($counterPath) {
                        $counterPath = [regex]::Replace([string]$counterPath, "\(\*\)", "(_Total)")
                        $counter = Get-Counter $counterPath -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
                        $cpu = [math]::Round($counter.CounterSamples[0].CookedValue, 0)
                    }
                } catch {}
            }
        }
    }

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $total = [double]$os.TotalVisibleMemorySize
        $free = [double]$os.FreePhysicalMemory
        if ($total -gt 0) { $memory = [math]::Round((($total - $free) / $total) * 100, 0) }
    } catch {
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
            $computerInfo = [Microsoft.VisualBasic.Devices.ComputerInfo]::new()
            $total = [double]$computerInfo.TotalPhysicalMemory
            $free = [double]$computerInfo.AvailablePhysicalMemory
            if ($total -gt 0) { $memory = [math]::Round((($total - $free) / $total) * 100, 0) }
        } catch {}
    }

    try {
        $disks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop | Where-Object { $_.Size -gt 0 })
        $size = [double](($disks | Measure-Object -Property Size -Sum).Sum)
        $free = [double](($disks | Measure-Object -Property FreeSpace -Sum).Sum)
        if ($size -gt 0) { $disk = [math]::Round((($size - $free) / $size) * 100, 0) }
    } catch {
        try {
            $drives = @(Get-PSDrive -PSProvider FileSystem -ErrorAction Stop | Where-Object { ($_.Used + $_.Free) -gt 0 })
            $used = [double](($drives | Measure-Object -Property Used -Sum).Sum)
            $free = [double](($drives | Measure-Object -Property Free -Sum).Sum)
            if (($used + $free) -gt 0) { $disk = [math]::Round(($used / ($used + $free)) * 100, 0) }
        } catch {}
    }
    if ($null -eq $disk) {
        try {
            $drives = @([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady -and $_.DriveType -eq [System.IO.DriveType]::Fixed -and $_.TotalSize -gt 0 })
            $size = [double](($drives | Measure-Object -Property TotalSize -Sum).Sum)
            $free = [double](($drives | Measure-Object -Property AvailableFreeSpace -Sum).Sum)
            if ($size -gt 0) { $disk = [math]::Round((($size - $free) / $size) * 100, 0) }
        } catch {}
    }

    [pscustomobject]@{
        cpu = if ($null -ne $cpu) { [int]$cpu } else { $null }
        memory = if ($null -ne $memory) { [int]$memory } else { $null }
        disk = if ($null -ne $disk) { [int]$disk } else { $null }
    }
}

function Get-LocalAutomationSnapshot {
    param([string]$Computer)

    $scanPaths = @(
        "C:\Scripts",
        "C:\Deploy",
        "C:\Reports",
        "C:\Users",
        "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup",
        "C:\Program Files (x86)\Microsoft Intune Management Extension",
        "C:\Windows\SYSVOL",
        "C:\Windows\System32\GroupPolicy"
    )
    $scriptExtensions = @("*.ps1", "*.bat", "*.cmd", "*.vbs")
    $server = $null
    $scripts = @()
    $tasks = @()
    $services = @()
    $alerts = @()

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $resources = Get-ServerResources
        $server = [pscustomobject]@{
            name = $env:COMPUTERNAME
            ip = Get-ServerIPv4 $Computer
            os = Format-WindowsServerOS $os.Caption
            status = "En linea"
            risk = "Bajo"
            findings = 0
            owner = if ($cs.Domain) { $cs.Domain } else { "Local" }
            cpu = $resources.cpu
            memory = $resources.memory
            disk = $resources.disk
        }
    } catch {
        $resources = Get-ServerResources
        $server = [pscustomobject]@{ name = $Computer; ip = (Get-ServerIPv4 $Computer); os = "Windows"; status = "Advertencia"; risk = "Medio"; findings = 0; owner = "Desconocido"; cpu = $resources.cpu; memory = $resources.memory; disk = $resources.disk }
        $alerts += [pscustomobject]@{ severity = "Alta"; origin = "Servidor"; name = "No se pudo consultar sistema operativo"; asset = $Computer; status = "Activa"; owner = "Infraestructura"; createdAt = (Get-Date).ToString("s") }
    }

    foreach ($path in $scanPaths) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        foreach ($ext in $scriptExtensions) {
            try {
                $found = Get-ChildItem -LiteralPath $path -Filter $ext -File -Recurse -Depth 5 -ErrorAction SilentlyContinue | Select-Object -First 250
                foreach ($file in $found) {
                    $owner = "Sin propietario"
                    try { $owner = (Get-Acl -LiteralPath $file.FullName).Owner } catch {}
                    $risk = Get-Risk -Account $owner -Path $file.FullName -LastResult "-" -Status "Detectado"
                    $scripts += [pscustomobject]@{
                        name = $file.Name
                        extension = $file.Extension
                        path = $file.FullName
                        server = $env:COMPUTERNAME
                        owner = $owner
                        lastUse = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                        lastError = "-"
                        risk = $risk
                    }
                }
            } catch {}
        }
    }

    try {
        foreach ($task in (Get-ScheduledTask -ErrorAction SilentlyContinue)) {
            $info = $null
            try { $info = $task | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue } catch {}
            $actions = @($task.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join " "
            $isAutomation = $actions -match "\.ps1|\.bat|\.cmd|\.vbs|powershell|pwsh|wscript|cscript|cmd.exe|sqlcmd|robocopy|xcopy"
            $lastResult = if ($info) { Convert-TaskResult $info.LastTaskResult } else { "Sin datos" }
            $risk = if ($isAutomation) { Get-Risk -Account $task.Principal.UserId -Path $actions -LastResult $lastResult -Status $task.State } else { "Bajo" }
            $tasks += [pscustomobject]@{
                name = $task.TaskName
                server = $env:COMPUTERNAME
                account = $task.Principal.UserId
                schedule = if ($isAutomation) { "Automatizacion: $actions" } else { (@($task.Triggers | ForEach-Object { $_.ToString() }) -join "; ") }
                lastRun = if ($info -and $info.LastRunTime) { $info.LastRunTime.ToString("yyyy-MM-dd HH:mm") } else { "Nunca" }
                nextRun = if ($info -and $info.NextRunTime) { $info.NextRunTime.ToString("yyyy-MM-dd HH:mm") } else { "-" }
                result = $lastResult
                risk = $risk
                status = "$($task.State)"
            }
        }
    } catch {
        $alerts += [pscustomobject]@{ severity = "Alta"; origin = "Tarea programada"; name = "No se pudieron leer tareas programadas"; asset = $Computer; status = "Activa"; owner = "Infraestructura"; createdAt = (Get-Date).ToString("s") }
    }

    if (@($tasks).Count -eq 0) {
        try {
            $rawTasks = schtasks.exe /query /fo csv /v 2>$null | ConvertFrom-Csv
            foreach ($row in @($rawTasks | Select-Object -First 400)) {
                $taskName = Get-FirstPropertyValue $row @("TaskName", "Nombre de tarea", "Nombre de la tarea")
                if ([string]::IsNullOrWhiteSpace($taskName)) { continue }
                $taskRun = Get-FirstPropertyValue $row @("Task To Run", "Tarea que se va a ejecutar", "Tarea para ejecutar")
                $account = Get-FirstPropertyValue $row @("Run As User", "Ejecutar como usuario", "Ejecutar como")
                $status = Get-FirstPropertyValue $row @("Status", "Estado")
                $lastRun = Get-FirstPropertyValue $row @("Last Run Time", "Hora de ultima ejecucion", "Hora de última ejecución")
                $nextRun = Get-FirstPropertyValue $row @("Next Run Time", "Hora de proxima ejecucion", "Hora de próxima ejecución")
                $lastResult = Get-FirstPropertyValue $row @("Last Result", "Resultado de ultima ejecucion", "Resultado de última ejecución")
                $schedule = Get-FirstPropertyValue $row @("Schedule Type", "Tipo de programacion", "Tipo de programación")
                $isAutomation = $taskRun -match "\.ps1|\.bat|\.cmd|\.vbs|powershell|pwsh|wscript|cscript|cmd.exe|sqlcmd|robocopy|xcopy"
                $risk = if ($isAutomation) { Get-Risk -Account $account -Path $taskRun -LastResult $lastResult -Status $status } else { "Bajo" }
                $tasks += [pscustomobject]@{
                    name = $taskName
                    server = $env:COMPUTERNAME
                    account = $account
                    schedule = if ($isAutomation) { "Automatizacion: $taskRun" } elseif ($schedule) { $schedule } else { $taskRun }
                    lastRun = if ($lastRun) { $lastRun } else { "Nunca" }
                    nextRun = if ($nextRun) { $nextRun } else { "-" }
                    result = if ($lastResult) { $lastResult } else { "Sin datos" }
                    risk = $risk
                    status = if ($status) { $status } else { "Detectada" }
                }
            }
        } catch {
            $alerts += [pscustomobject]@{ severity = "Media"; origin = "Tarea programada"; name = "schtasks.exe no devolvio tareas: $($_.Exception.Message)"; asset = $Computer; status = "Activa"; owner = "Infraestructura"; createdAt = (Get-Date).ToString("s") }
        }
    }

    try {
        $svcRows = Get-CimInstance -ClassName Win32_Service -ErrorAction SilentlyContinue | Where-Object {
            $_.PathName -match "\.ps1|\.bat|\.cmd|\.vbs|powershell|pwsh|wscript|cscript|cmd.exe"
        }
        foreach ($svc in $svcRows) {
            $risk = Get-Risk -Account $svc.StartName -Path $svc.PathName -LastResult "-" -Status $svc.State
            $services += [pscustomobject]@{
                name = $svc.Name
                status = if ($svc.State -eq "Running") { "Ejecutando" } else { "Detenido" }
                startup = $svc.StartMode
                risk = $risk
                systems = 1
                script = $svc.PathName
                server = $env:COMPUTERNAME
                account = $svc.StartName
            }
        }
    } catch {
        $alerts += [pscustomobject]@{ severity = "Alta"; origin = "Servicio"; name = "No se pudieron leer servicios"; asset = $Computer; status = "Activa"; owner = "Infraestructura"; createdAt = (Get-Date).ToString("s") }
    }

    $findings = @($scripts).Count + @($tasks).Count + @($services).Count
    $server.findings = $findings
    $server.risk = if (($scripts + $tasks + $services | Where-Object { $_.risk -eq "Alto" } | Select-Object -First 1)) { "Alto" } elseif ($findings -gt 0) { "Medio" } else { "Bajo" }

    [pscustomobject]@{
        server = $server
        scripts = @($scripts)
        tasks = @($tasks)
        services = @($services)
        alerts = @($alerts)
    }
}

function Invoke-WindowsAutomationScan {
    param([string[]]$Targets)
    $results = @()
    foreach ($target in $Targets) {
        if ([string]::IsNullOrWhiteSpace($target)) { continue }
        $name = $target.Trim()
        try {
            if ($name -in @("localhost", ".", $env:COMPUTERNAME)) {
                $results += Get-LocalAutomationSnapshot -Computer $name
            } else {
                $remoteSource = @(
                    "function Get-Risk { $(${function:Get-Risk}.ToString()) }",
                    "function Convert-TaskResult { $(${function:Convert-TaskResult}.ToString()) }",
                    "function Get-FirstPropertyValue { $(${function:Get-FirstPropertyValue}.ToString()) }",
                    "function Get-ServerIPv4 { $(${function:Get-ServerIPv4}.ToString()) }",
                    "function Format-WindowsServerOS { $(${function:Format-WindowsServerOS}.ToString()) }",
                    "function Get-ServerResources { $(${function:Get-ServerResources}.ToString()) }",
                    "function Get-LocalAutomationSnapshot { $(${function:Get-LocalAutomationSnapshot}.ToString()) }",
                    'Get-LocalAutomationSnapshot -Computer $args[0]'
                ) -join "`n`n"
                $results += Invoke-Command -ComputerName $name -ScriptBlock ([scriptblock]::Create($remoteSource)) -ArgumentList $name -ErrorAction Stop
            }
        } catch {
            $results += [pscustomobject]@{
                server = [pscustomobject]@{ name = $name; ip = (Get-ServerIPv4 $name); os = "Windows"; status = "Sin contacto"; risk = "Alto"; findings = 1; owner = "Desconocido" }
                scripts = @()
                tasks = @()
                services = @()
                alerts = @([pscustomobject]@{ severity = "Alta"; origin = "Servidor"; name = "No se pudo escanear: $($_.Exception.Message)"; asset = $name; status = "Activa"; owner = "Infraestructura"; createdAt = (Get-Date).ToString("s") })
            }
        }
    }
    $results
}

function Build-AutomationsFromScan($Scripts, $Tasks, $Services) {
    $items = @()
    foreach ($s in @($Scripts)) {
        $items += [pscustomobject]@{ name = $s.name; type = "Script"; status = "Detectado"; criticality = $s.risk; owner = $s.owner; lastRun = $s.lastUse; result = $s.lastError; system = $s.server }
    }
    foreach ($t in @($Tasks)) {
        $items += [pscustomobject]@{ name = $t.name; type = "Tarea programada"; status = $t.status; criticality = $t.risk; owner = $t.account; lastRun = $t.lastRun; result = $t.result; system = $t.server }
    }
    foreach ($svc in @($Services)) {
        $items += [pscustomobject]@{ name = $svc.name; type = "Servicio"; status = $svc.status; criticality = $svc.risk; owner = $svc.account; lastRun = "-"; result = $svc.script; system = $svc.server }
    }
    $items
}

function Add-Activity($State, [string]$Title, [string]$Detail) {
    $item = [pscustomobject]@{
        type = "success"
        title = $Title
        detail = $Detail
        at = (Get-Date).ToString("s")
    }
    $State.activity = @($item) + @($State.activity)
}

function Resolve-ScanServerIPs {
    param([object[]]$Servers)

    foreach ($server in @($Servers)) {
        if (-not $server) { continue }
        $current = [string]$server.ip
        if ([string]::IsNullOrWhiteSpace($current) -or $current -eq "-") {
            $resolved = Get-ServerIPv4 ([string]$server.name)
            if (-not [string]::IsNullOrWhiteSpace($resolved) -and $resolved -ne "-") {
                $server.ip = $resolved
            }
        }
    }
    @($Servers)
}

function Invoke-Scan {
    param([string[]]$Targets = $ComputerName)
    $state = Get-State
    $scan = Invoke-WindowsAutomationScan -Targets $Targets
    $state.servers = @(Resolve-ScanServerIPs @($scan | ForEach-Object { $_.server }))
    $state.scripts = @($scan | ForEach-Object { $_.scripts } | Where-Object { $null -ne $_ })
    $state.tasks = @($scan | ForEach-Object { $_.tasks } | Where-Object { $null -ne $_ })
    $state.services = @($scan | ForEach-Object { $_.services } | Where-Object { $null -ne $_ })
    $state.automations = @(Build-AutomationsFromScan $state.scripts $state.tasks $state.services)
    $state.alerts = @($scan | ForEach-Object { $_.alerts } | Where-Object { $null -ne $_ })
    $state.scanSummary = [pscustomobject]@{
        lastScan = (Get-Date).ToString("s")
        targets = @($Targets)
        servers = @($state.servers).Count
        scripts = @($state.scripts).Count
        tasks = @($state.tasks).Count
        services = @($state.services).Count
        automations = @($state.automations).Count
        alerts = @($state.alerts).Count
        message = "Escaneo completado. Servidores: $(@($state.servers).Count), scripts: $(@($state.scripts).Count), tareas: $(@($state.tasks).Count), servicios: $(@($state.services).Count), alertas: $(@($state.alerts).Count)."
    }
    Add-Activity $state "Escaneo Windows ejecutado" $state.scanSummary.message
    Save-State $state
    $state
}

function New-Report {
    param(
        [string]$Format = "PDF",
        [string]$Template = "Resumen ejecutivo"
    )
    $state = Get-State
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
    if ([string]::IsNullOrWhiteSpace($Format)) { $Format = "PDF" }
    if ([string]::IsNullOrWhiteSpace($Template)) { $Template = "Resumen ejecutivo" }
    $Format = $Format.ToUpperInvariant()
    ConvertTo-ReportHtml $state | Set-Content -LiteralPath $ReportPath -Encoding UTF8
    $report = [pscustomobject]@{
        name = "$Template - $stamp"
        template = $Template
        format = $Format
        size = "$([math]::Round((Get-Item -LiteralPath $ReportPath).Length / 1KB, 1)) KB"
        status = "Generado"
        createdAt = (Get-Date).ToString("s")
    }
    $state.reports = @($report) + @($state.reports)
    Add-Activity $state "Reporte generado" $report.name
    Save-State $state
    $state
}

function Resolve-Alert {
    $state = Get-State
    $active = @($state.alerts | Where-Object { $_.status -eq "Activa" } | Select-Object -First 1)
    if ($active.Count -gt 0) {
        $active[0].status = "Reconocida"
        Add-Activity $state "Alerta reconocida" $active[0].name
        Save-State $state
    }
    $state
}

function Set-Config {
    param(
        [string]$Tab,
        [string]$Data
    )
    $state = Get-State
    if ([string]::IsNullOrWhiteSpace($Tab)) { $Tab = "General" }
    if (-not $state.PSObject.Properties["configSections"] -or $null -eq $state.configSections) {
        $state | Add-Member -NotePropertyName configSections -NotePropertyValue ([pscustomobject]@{}) -Force
    }
    $values = [pscustomobject]@{}
    if (-not [string]::IsNullOrWhiteSpace($Data)) {
        try { $values = $Data | ConvertFrom-Json } catch { $values = [pscustomobject]@{ raw = $Data } }
    }
    $state.configSections | Add-Member -NotePropertyName $Tab -NotePropertyValue $values -Force
    if ($Tab -eq "General") {
        foreach ($p in @($values.PSObject.Properties)) {
            if ($p.Name -match "organizacion|organization|org" -and $state.config) { $state.config.org = [string]$p.Value }
            if ($p.Name -match "dominio|domain|tenant" -and $state.config) { $state.config.domain = [string]$p.Value }
        }
    }
    Add-Activity $state "Configuracion guardada" $Tab
    Save-State $state
    $state
}

function Get-ObjValue {
    param(
        $Object,
        [string]$Name,
        $Default = $null
    )
    if ($null -eq $Object) { return $Default }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $Default }
    $prop.Value
}

function Get-SmtpConfig {
    if (-not (Test-Path -LiteralPath $SmtpConfigPath)) {
        return [pscustomobject]@{
            enabled = $false
            server = ""
            port = 587
            security = "STARTTLS"
            from = ""
            user = ""
            passwordProtected = ""
            passwordSaved = $false
            savedAt = ""
        }
    }
    try {
        $cfg = Get-Content -LiteralPath $SmtpConfigPath -Raw | ConvertFrom-Json
    } catch {
        $cfg = [pscustomobject]@{}
    }
    $passwordProtected = [string](Get-ObjValue $cfg "passwordProtected" "")
    [pscustomobject]@{
        enabled = [bool](Get-ObjValue $cfg "enabled" $false)
        server = [string](Get-ObjValue $cfg "server" "")
        port = [int](Get-ObjValue $cfg "port" 587)
        security = [string](Get-ObjValue $cfg "security" "STARTTLS")
        from = [string](Get-ObjValue $cfg "from" "")
        user = [string](Get-ObjValue $cfg "user" "")
        passwordProtected = $passwordProtected
        passwordSaved = (-not [string]::IsNullOrWhiteSpace($passwordProtected))
        savedAt = [string](Get-ObjValue $cfg "savedAt" "")
    }
}

function Get-SmtpPublicConfig {
    $cfg = Get-SmtpConfig
    [pscustomobject]@{
        enabled = [bool]$cfg.enabled
        server = [string]$cfg.server
        port = [int]$cfg.port
        security = [string]$cfg.security
        from = [string]$cfg.from
        user = [string]$cfg.user
        passwordSaved = [bool]$cfg.passwordSaved
        savedAt = [string]$cfg.savedAt
        protectedBy = "$([Environment]::UserDomainName)\$([Environment]::UserName)"
    }
}

function Save-SmtpConfig {
    param(
        [string]$Enabled,
        [string]$Server,
        [string]$Port,
        [string]$Security,
        [string]$From,
        [string]$User,
        [string]$Password
    )
    $existing = Get-SmtpConfig
    $protected = [string]$existing.passwordProtected
    if (-not [string]::IsNullOrWhiteSpace($Password)) {
        $secure = ConvertTo-SecureString $Password -AsPlainText -Force
        $protected = ConvertFrom-SecureString $secure
    }
    $portValue = 587
    if (-not [int]::TryParse($Port, [ref]$portValue)) { $portValue = 587 }
    $enabledValue = $false
    if ($Enabled -match "^(1|true|si|sí|yes|on)$") { $enabledValue = $true }
    $cfg = [ordered]@{
        enabled = $enabledValue
        server = [string]$Server
        port = $portValue
        security = if ([string]::IsNullOrWhiteSpace($Security)) { "STARTTLS" } else { [string]$Security }
        from = [string]$From
        user = [string]$User
        passwordProtected = $protected
        savedAt = (Get-Date).ToString("s")
        savedBy = "$([Environment]::UserDomainName)\$([Environment]::UserName)"
    }
    $cfg | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $SmtpConfigPath -Encoding UTF8
    $state = Get-State
    Add-Activity $state "SMTP guardado" "Configuracion SMTP protegida localmente"
    Save-State $state
    Get-SmtpPublicConfig
}

function Test-SmtpConfig {
    param(
        [string]$Server,
        [string]$Port
    )
    $cfg = Get-SmtpConfig
    if ([string]::IsNullOrWhiteSpace($Server)) { $Server = [string]$cfg.server }
    $portValue = [int]$cfg.port
    if (-not [string]::IsNullOrWhiteSpace($Port)) {
        [void][int]::TryParse($Port, [ref]$portValue)
    }
    if ([string]::IsNullOrWhiteSpace($Server)) {
        return [pscustomobject]@{ ok = $false; message = "Configura primero el servidor SMTP."; server = ""; port = $portValue; checkedAt = (Get-Date).ToString("s") }
    }
    $client = [Net.Sockets.TcpClient]::new()
    try {
        $async = $client.BeginConnect($Server, $portValue, $null, $null)
        $connected = $async.AsyncWaitHandle.WaitOne(5000, $false)
        if (-not $connected) { throw "Tiempo de espera agotado." }
        $client.EndConnect($async)
        [pscustomobject]@{
            ok = $true
            message = "Conexion TCP correcta con $Server`:$portValue. La clave SMTP queda protegida con DPAPI."
            server = $Server
            port = $portValue
            checkedAt = (Get-Date).ToString("s")
        }
    } catch {
        [pscustomobject]@{
            ok = $false
            message = "No se pudo conectar con $Server`:$portValue. $($_.Exception.Message)"
            server = $Server
            port = $portValue
            checkedAt = (Get-Date).ToString("s")
        }
    } finally {
        try { $client.Close() } catch {}
    }
}

function Add-Item {
    param(
        [string]$Type,
        [string]$Name,
        [string]$Owner,
        [string]$Risk,
        [string]$Server
    )
    $state = Get-State
    if ([string]::IsNullOrWhiteSpace($Name)) { $Name = "Nuevo_$Type" }
    if ([string]::IsNullOrWhiteSpace($Owner)) { $Owner = "Admin Demo" }
    if ([string]::IsNullOrWhiteSpace($Risk)) { $Risk = "Bajo" }
    if ([string]::IsNullOrWhiteSpace($Server)) { $Server = "SRV-APP-01" }

    switch -Wildcard ($Type) {
        "*Script*" {
            $scriptName = if ($Name -match "\.(ps1|bat|cmd|vbs)$") { $Name } else { "$Name.ps1" }
            $extension = [IO.Path]::GetExtension($scriptName)
            $item = [pscustomobject]@{ name = $scriptName; extension = $extension; path = "C:\Scripts\$scriptName"; server = $Server; owner = $Owner; lastUse = "Ahora"; lastError = "-"; risk = $Risk }
            $state.scripts = @($item) + @($state.scripts)
        }
        "*Tarea*" {
            $item = [pscustomobject]@{ name = $Name; server = $Server; account = "svc_automap"; schedule = "Diaria 02:00"; lastRun = "Nunca"; nextRun = "Manana 02:00"; result = "Pendiente"; risk = $Risk; status = "Habilitada" }
            $state.tasks = @($item) + @($state.tasks)
        }
        "*Credencial*" {
            $item = [pscustomobject]@{ name = $Name; account = "********"; source = "Manual"; location = $Server; risk = $Risk; status = "Abierto" }
            $state.credentials = @($item) + @($state.credentials)
        }
        default {
            $item = [pscustomobject]@{ name = $Name; type = $Type; status = "Habilitado"; criticality = $Risk; owner = $Owner; lastRun = "Ahora"; result = "Pendiente"; system = $Server }
            $state.automations = @($item) + @($state.automations)
        }
    }
    Add-Activity $state "$Type creado" $Name
    Save-State $state
    $state
}

function Get-QueryParams([string]$RawPath) {
    $queryIndex = $RawPath.IndexOf("?")
    $result = @{}
    if ($queryIndex -lt 0) { return $result }
    $query = $RawPath.Substring($queryIndex + 1)
    foreach ($pair in $query.Split("&")) {
        if ([string]::IsNullOrWhiteSpace($pair)) { continue }
        $parts = $pair.Split("=", 2)
        $key = [Uri]::UnescapeDataString($parts[0])
        $value = if ($parts.Count -gt 1) { [Uri]::UnescapeDataString($parts[1].Replace("+", " ")) } else { "" }
        $result[$key] = $value
    }
    $result
}

New-AutoMapState | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $StatePath -Encoding UTF8

$listener = $null
$actualPort = $Port
$lastError = $null
for ($candidate = $Port; $candidate -le ($Port + 30); $candidate++) {
    try {
        $candidateListener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $candidate)
        $candidateListener.Start()
        $listener = $candidateListener
        $actualPort = $candidate
        break
    } catch {
        $lastError = $_.Exception.Message
        if ($null -ne $candidateListener) {
            try { $candidateListener.Stop() } catch {}
        }
    }
}

if ($null -eq $listener) {
    throw "No se pudo iniciar el portal entre los puertos $Port y $($Port + 30). Ultimo error: $lastError"
}

$prefix = "http://localhost:$actualPort/"

Write-Host ""
Write-Host "AutoMap IT iniciado desde PowerShell" -ForegroundColor Cyan
Write-Host "Portal: $prefix" -ForegroundColor Green
Write-Host "API:    $($prefix)api/state" -ForegroundColor DarkCyan
Write-Host "Pulsa Ctrl+C para detener." -ForegroundColor Yellow
Write-Host ""

if (-not $NoBrowser) {
    try { Start-Process $prefix } catch { Write-Host "Abre manualmente $prefix" -ForegroundColor Yellow }
}

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()
        try {
            $stream = $client.GetStream()
            $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $false, 4096, $true)
            $requestLine = $reader.ReadLine()
            if ([string]::IsNullOrWhiteSpace($requestLine)) {
                $client.Close()
                continue
            }
            do { $line = $reader.ReadLine() } while ($null -ne $line -and $line.Length -gt 0)

            $parts = $requestLine.Split(" ")
            $method = $parts[0].ToUpperInvariant()
            $rawPath = $parts[1]
            $query = Get-QueryParams $rawPath
            $path = ([Uri]("http://localhost$rawPath")).AbsolutePath.TrimEnd("/")
            if ([string]::IsNullOrWhiteSpace($path)) { $path = "/" }
            $context = [pscustomobject]@{ Stream = $stream; Method = $method; Path = $path }

            switch -Regex ($path) {
                "^/$|^/portal$|^/index\.html$" {
                    $html = [IO.File]::ReadAllText($PortalPath, [Text.Encoding]::UTF8)
                    Send-Text $context $html "text/html; charset=utf-8"
                    break
                }
                "^/api/health$" {
                    Send-Json $context ([ordered]@{ ok = $true; portal = $prefix; time = (Get-Date).ToString("s") })
                    break
                }
                "^/api/state$" {
                    Send-Json $context (Get-State)
                    break
                }
                "^/api/scan$" {
                    if ($method -ne "POST") { Send-Text $context "Metodo no permitido" "text/plain; charset=utf-8" 405; break }
                    $targets = if ($query.ContainsKey("targets") -and -not [string]::IsNullOrWhiteSpace($query["targets"])) { $query["targets"].Split(",") } else { $ComputerName }
                    Send-Json $context (Invoke-Scan -Targets $targets)
                    break
                }
                "^/api/reports$" {
                    if ($method -ne "POST") { Send-Text $context "Metodo no permitido" "text/plain; charset=utf-8" 405; break }
                    Send-Json $context (New-Report -Format $query["format"] -Template $query["template"])
                    break
                }
                "^/api/report/html$" {
                    $state = Get-State
                    $html = ConvertTo-ReportHtml $state
                    $html | Set-Content -LiteralPath $ReportPath -Encoding UTF8
                    Send-Text $context $html "text/html; charset=utf-8"
                    break
                }
                "^/api/create$" {
                    if ($method -ne "POST") { Send-Text $context "Metodo no permitido" "text/plain; charset=utf-8" 405; break }
                    Send-Json $context (Add-Item -Type $query["type"] -Name $query["name"] -Owner $query["owner"] -Risk $query["risk"] -Server $query["server"])
                    break
                }
                "^/api/config$" {
                    if ($method -ne "POST") { Send-Text $context "Metodo no permitido" "text/plain; charset=utf-8" 405; break }
                    Send-Json $context (Set-Config -Tab $query["tab"] -Data $query["data"])
                    break
                }
                "^/api/smtp$" {
                    Send-Json $context (Get-SmtpPublicConfig)
                    break
                }
                "^/api/smtp/save$" {
                    if ($method -ne "POST") { Send-Text $context "Metodo no permitido" "text/plain; charset=utf-8" 405; break }
                    Send-Json $context (Save-SmtpConfig -Enabled $query["enabled"] -Server $query["server"] -Port $query["port"] -Security $query["security"] -From $query["from"] -User $query["user"] -Password $query["password"])
                    break
                }
                "^/api/smtp/test$" {
                    if ($method -ne "POST") { Send-Text $context "Metodo no permitido" "text/plain; charset=utf-8" 405; break }
                    Send-Json $context (Test-SmtpConfig -Server $query["server"] -Port $query["port"])
                    break
                }
                "^/api/alerts/resolve$" {
                    if ($method -ne "POST") { Send-Text $context "Metodo no permitido" "text/plain; charset=utf-8" 405; break }
                    Send-Json $context (Resolve-Alert)
                    break
                }
                "^/api/export/(?<name>[^/]+)(?:/(?<format>txt|csv|excel|pdf))?$" {
                    $state = Get-State
                    $name = $Matches.name.ToLowerInvariant()
                    $format = if ($Matches.format) { $Matches.format.ToLowerInvariant() } else { "csv" }
                    $rows = @(Get-ExportRows $state $name)
                    if ($format -eq "txt") {
                        $txt = ConvertTo-TxtText $name $rows
                        Send-Bytes $context ([Text.Encoding]::UTF8.GetBytes($txt)) "text/plain; charset=utf-8" 200 @{ "Content-Disposition" = "attachment; filename=automap-$name.txt" }
                        break
                    }
                    if ($format -eq "excel") {
                        $excel = ConvertTo-ExcelHtml $name $rows
                        Send-Bytes $context ([Text.Encoding]::UTF8.GetBytes($excel)) "application/vnd.ms-excel; charset=utf-8" 200 @{ "Content-Disposition" = "attachment; filename=automap-$name.xls" }
                        break
                    }
                    if ($format -eq "pdf") {
                        $pdf = ConvertTo-PdfBytes $name $rows
                        Send-Bytes $context $pdf "application/pdf" 200 @{ "Content-Disposition" = "attachment; filename=automap-$name.pdf" }
                        break
                    }
                    $csv = if (@($rows).Count -gt 0) { ConvertTo-CsvText $rows } else { Get-CsvHeader $name }
                    Send-Bytes $context ([Text.Encoding]::UTF8.GetBytes($csv)) "text/csv; charset=utf-8" 200 @{ "Content-Disposition" = "attachment; filename=automap-$name.csv" }
                    break
                }
                default {
                    Send-Json $context ([ordered]@{ error = "Ruta no encontrada"; path = $path }) 404
                }
            }
        } catch {
            if ($null -ne $context) {
                Send-Json $context ([ordered]@{ error = $_.Exception.Message }) 500
            }
        } finally {
            $client.Close()
        }
    }
} finally {
    $listener.Stop()
}
