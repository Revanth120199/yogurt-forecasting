param(
    [string]$InputDir = ".",
    [int]$ForecastYear = 2026,
    [string]$CacheDir = ".cache_extracted_history",
    [string]$HistoryCsv = "purchase_history_4cols.csv",
    [string]$ForecastCsv = "two_step_forecast_2026.csv",
    [string]$ReportHtml = "two_step_forecast_2026.html",
    [string]$SummaryJson = "two_step_forecast_2026_summary.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-NormalizedHeader {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    return (($Value -replace "[^A-Za-z0-9]+", " ").Trim().ToUpperInvariant() -replace "\s+", " ")
}

function Get-ColumnLetters {
    param([string]$CellReference)
    if ($CellReference -match "^[A-Z]+") { return $matches[0] }
    return $null
}

function Get-XmlDocumentFromZip {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { throw "Missing ZIP entry: $EntryName" }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { $content = $reader.ReadToEnd() } finally { $reader.Dispose() }
    return [xml]$content
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)
    $entry = $Zip.GetEntry("xl/sharedStrings.xml")
    if (-not $entry) { return @() }
    $doc = Get-XmlDocumentFromZip -Zip $Zip -EntryName "xl/sharedStrings.xml"
    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $values = New-Object System.Collections.Generic.List[string]
    foreach ($si in $doc.SelectNodes("//x:si", $ns)) {
        $parts = $si.SelectNodes(".//x:t", $ns)
        if ($parts.Count -gt 0) {
            $values.Add(($parts | ForEach-Object { $_.InnerText }) -join "")
        } else {
            $values.Add($si.InnerText)
        }
    }
    return ,$values.ToArray()
}

function Get-DateStyleIndexes {
    param([System.IO.Compression.ZipArchive]$Zip)
    $doc = Get-XmlDocumentFromZip -Zip $Zip -EntryName "xl/styles.xml"
    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    $customNumFmtMap = @{}
    foreach ($numFmt in $doc.SelectNodes("//x:numFmts/x:numFmt", $ns)) {
        $customNumFmtMap[[int]$numFmt.numFmtId] = [string]$numFmt.formatCode
    }

    $dateStyleIndexes = New-Object System.Collections.Generic.HashSet[int]
    $builtinDateIds = [System.Collections.Generic.HashSet[int]]::new()
    @(14,15,16,17,18,19,20,21,22,27,28,29,30,31,32,33,34,35,36,45,46,47,50,51,52,53,54,55,56,57,58) | ForEach-Object {
        [void]$builtinDateIds.Add($_)
    }

    $cellXfs = $doc.SelectNodes("//x:cellXfs/x:xf", $ns)
    for ($i = 0; $i -lt $cellXfs.Count; $i++) {
        $xf = $cellXfs[$i]
        $numFmtId = [int]$xf.numFmtId
        $isDate = $builtinDateIds.Contains($numFmtId)
        if (-not $isDate -and $customNumFmtMap.ContainsKey($numFmtId)) {
            $formatCode = $customNumFmtMap[$numFmtId].ToLowerInvariant()
            $sanitized = $formatCode -replace '".*?"', "" -replace '\\.', ""
            if ($sanitized -match "[ymdhis]") { $isDate = $true }
        }
        if ($isDate) { [void]$dateStyleIndexes.Add($i) }
    }

    return ,$dateStyleIndexes
}

function Get-SheetEntryPath {
    param([System.IO.Compression.ZipArchive]$Zip)
    $workbook = Get-XmlDocumentFromZip -Zip $Zip -EntryName "xl/workbook.xml"
    $rels = Get-XmlDocumentFromZip -Zip $Zip -EntryName "xl/_rels/workbook.xml.rels"

    $workbookNs = [System.Xml.XmlNamespaceManager]::new($workbook.NameTable)
    $workbookNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $workbookNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $relsNs = [System.Xml.XmlNamespaceManager]::new($rels.NameTable)
    $relsNs.AddNamespace("pr", "http://schemas.openxmlformats.org/package/2006/relationships")

    $sheetNodes = $workbook.SelectNodes("//x:sheets/x:sheet", $workbookNs)
    $targetSheet = $sheetNodes | Where-Object { $_.name -eq "EXPORT" } | Select-Object -First 1
    if (-not $targetSheet) { $targetSheet = $sheetNodes | Select-Object -First 1 }
    if (-not $targetSheet) { throw "No worksheets found in workbook." }

    $relationshipId = $targetSheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $relationship = $rels.SelectSingleNode("//pr:Relationship[@Id='$relationshipId']", $relsNs)
    if (-not $relationship) { throw "Could not resolve sheet relationship id $relationshipId." }
    return "xl/" + $relationship.Target
}

function Convert-CellValue {
    param(
        [string]$RawValue,
        [string]$CellType,
        [int]$StyleIndex,
        [string[]]$SharedStrings,
        [System.Collections.Generic.HashSet[int]]$DateStyleIndexes
    )
    if ([string]::IsNullOrEmpty($RawValue)) { return $null }
    switch ($CellType) {
        "s" { return $SharedStrings[[int]$RawValue] }
        "b" { return ([int]$RawValue -eq 1).ToString() }
        default {
            if ($DateStyleIndexes.Contains($StyleIndex)) {
                try { return [DateTime]::FromOADate([double]$RawValue).ToString("yyyy-MM-dd") } catch { return $RawValue }
            }
            return $RawValue
        }
    }
}

function Get-RequiredColumnMap {
    return @{
        "UNIT" = "UNIT"
        "NDA ITEM DESCRIPTION" = "NDA ITEM DESCRIPTION"
        "INVOICE DATE" = "INVOICE DATE"
        "ORDERED" = "ORDERED"
    }
}

function Get-WeekOfMonth {
    param([datetime]$DateValue)
    return [int][Math]::Floor(($DateValue.Day - 1) / 7) + 1
}

function Get-StoreCode {
    param([string]$Unit)
    if ([string]::IsNullOrWhiteSpace($Unit)) { return "" }
    if ($Unit -match "(?<!\d)(\d{3,5})(?!\d)") { return $matches[1] }
    if ($Unit -match "(?<![A-Z0-9])([A-Z]{1,3}\d{2,5}[A-Z0-9]*)(?![A-Z0-9])") { return $matches[1] }
    return ""
}

function Get-SafeFileStem {
    param([string]$Name)
    return (($Name -replace '[\\/:*?"<>|]+', "_") -replace "\s+", " ").Trim()
}

function Import-PurchaseRows {
    param([string]$WorkbookPath)

    $requiredColumns = Get-RequiredColumnMap
    $zip = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)

    try {
        $sharedStrings = Get-SharedStrings -Zip $zip
        $dateStyleIndexes = Get-DateStyleIndexes -Zip $zip
        $sheetEntryPath = Get-SheetEntryPath -Zip $zip
        $sheetEntry = $zip.GetEntry($sheetEntryPath)
        if (-not $sheetEntry) { throw "Worksheet entry not found: $sheetEntryPath" }

        $rows = New-Object System.Collections.Generic.List[object]
        $headerMap = @{}
        $headerRowNumber = $null

        $settings = [System.Xml.XmlReaderSettings]::new()
        $settings.IgnoreWhitespace = $true
        $reader = [System.Xml.XmlReader]::Create($sheetEntry.Open(), $settings)
        try {
            while ($reader.Read()) {
                if ($reader.NodeType -ne [System.Xml.XmlNodeType]::Element -or $reader.Name -ne "row") { continue }

                $rowNumber = [int]$reader.GetAttribute("r")
                $rowReader = $reader.ReadSubtree()
                $cells = @{}

                try {
                    while ($rowReader.Read()) {
                        if ($rowReader.NodeType -ne [System.Xml.XmlNodeType]::Element -or $rowReader.Name -ne "c") { continue }
                        $cellReference = $rowReader.GetAttribute("r")
                        $columnLetters = Get-ColumnLetters -CellReference $cellReference
                        $cellType = $rowReader.GetAttribute("t")
                        $styleIndexValue = $rowReader.GetAttribute("s")
                        $styleIndex = if ($null -ne $styleIndexValue -and $styleIndexValue -ne "") { [int]$styleIndexValue } else { -1 }
                        $cellSubtree = $rowReader.ReadSubtree()
                        $rawValue = $null
                        try {
                            while ($cellSubtree.Read()) {
                                if ($cellSubtree.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                                if ($cellSubtree.Name -eq "v") {
                                    $rawValue = $cellSubtree.ReadElementContentAsString()
                                } elseif ($cellSubtree.Name -eq "t" -and $cellType -eq "inlineStr") {
                                    $rawValue = $cellSubtree.ReadElementContentAsString()
                                }
                            }
                        } finally {
                            $cellSubtree.Dispose()
                        }
                        $cells[$columnLetters] = Convert-CellValue -RawValue $rawValue -CellType $cellType -StyleIndex $styleIndex -SharedStrings $sharedStrings -DateStyleIndexes $dateStyleIndexes
                    }
                } finally {
                    $rowReader.Dispose()
                }

                if ($cells.Count -eq 0) { continue }

                if (-not $headerRowNumber) {
                    $candidateHeaderMap = @{}
                    foreach ($column in $cells.Keys) {
                        $normalized = Get-NormalizedHeader -Value ([string]$cells[$column])
                        if ($requiredColumns.ContainsKey($normalized)) {
                            $candidateHeaderMap[$requiredColumns[$normalized]] = $column
                        }
                    }
                    if ($candidateHeaderMap.Count -eq $requiredColumns.Count) {
                        $headerMap = $candidateHeaderMap
                        $headerRowNumber = $rowNumber
                    }
                    continue
                }

                if ($rowNumber -le $headerRowNumber) { continue }

                $unit = [string]$cells[$headerMap["UNIT"]]
                $itemDescription = [string]$cells[$headerMap["NDA ITEM DESCRIPTION"]]
                $invoiceDateRaw = [string]$cells[$headerMap["INVOICE DATE"]]
                $orderedRaw = [string]$cells[$headerMap["ORDERED"]]

                if ([string]::IsNullOrWhiteSpace($unit) -and [string]::IsNullOrWhiteSpace($itemDescription) -and [string]::IsNullOrWhiteSpace($invoiceDateRaw) -and [string]::IsNullOrWhiteSpace($orderedRaw)) {
                    continue
                }

                try {
                    $invoiceDate = [DateTime]::Parse($invoiceDateRaw, [System.Globalization.CultureInfo]::InvariantCulture)
                } catch {
                    continue
                }

                try {
                    $ordered = [double]::Parse(($orderedRaw -replace ","), [System.Globalization.CultureInfo]::InvariantCulture)
                } catch {
                    continue
                }

                $rows.Add([pscustomobject]@{
                    UNIT = $unit.Trim()
                    STORE_CODE = Get-StoreCode -Unit $unit
                    "NDA ITEM DESCRIPTION" = $itemDescription.Trim()
                    "INVOICE DATE" = $invoiceDate.ToString("yyyy-MM-dd")
                    ORDERED = [math]::Round($ordered, 4)
                    YEAR = $invoiceDate.Year
                    MONTH = $invoiceDate.Month
                    WEEK_OF_MONTH = Get-WeekOfMonth -DateValue $invoiceDate
                })
            }
        } finally {
            $reader.Dispose()
        }

        return $rows
    } finally {
        $zip.Dispose()
    }
}

function New-HtmlReport {
    param(
        [object[]]$ForecastRows,
        [string]$OutputPath,
        [int]$ForecastYearValue
    )

    $jsRows = $ForecastRows | ForEach-Object {
        @(
            $_.unit,
            $_.store_code,
            $_.month_key,
            $_.flavor,
            $_.monthly_forecast_cases,
            $_.history_years_used,
            $_.week_1_cases,
            $_.week_2_cases,
            $_.week_3_cases,
            $_.week_4_cases,
            $_.week_5_cases
        )
    } | ConvertTo-Json -Compress -Depth 4

    $html = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Two-Step Forecast __YEAR__</title>
<style>
:root {
  --ink: #17313b;
  --muted: #5b6e76;
  --line: #d5e3e8;
  --panel: #f8fbfc;
  --accent: #0f766e;
  --accent-2: #1d4ed8;
}
* { box-sizing: border-box; }
body { margin: 0; font-family: Segoe UI, Arial, sans-serif; color: var(--ink); background: #fff; }
header { padding: 24px 28px 16px; border-bottom: 1px solid var(--line); }
h1 { margin: 0 0 6px; font-size: 26px; }
p { margin: 4px 0; color: var(--muted); font-size: 13px; }
main { padding: 22px 28px 36px; }
.filters { display: grid; grid-template-columns: repeat(2, minmax(220px, 320px)); gap: 12px; margin-bottom: 18px; }
label { display: block; font-size: 11px; color: var(--muted); text-transform: uppercase; margin-bottom: 4px; }
select { width: 100%; padding: 9px 10px; border: 1px solid var(--line); border-radius: 6px; background: #fff; }
.grid { display: grid; grid-template-columns: repeat(3, minmax(180px, 1fr)); gap: 12px; margin-bottom: 18px; }
.card { border: 1px solid var(--line); border-radius: 8px; background: var(--panel); padding: 12px; }
.card .label { color: var(--muted); font-size: 11px; text-transform: uppercase; }
.card .value { margin-top: 5px; font-size: 22px; font-weight: 700; }
.table-wrap { border: 1px solid var(--line); border-radius: 8px; overflow: auto; margin-top: 14px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { border-bottom: 1px solid var(--line); padding: 8px 10px; text-align: right; white-space: nowrap; }
th:first-child, td:first-child, th:nth-child(2), td:nth-child(2) { text-align: left; }
th { background: #eef6f8; font-size: 12px; }
.section-title { margin: 20px 0 10px; font-size: 18px; }
@media (max-width: 900px) {
  .filters, .grid { grid-template-columns: 1fr; }
  header, main { padding-left: 16px; padding-right: 16px; }
}
</style>
</head>
<body>
<header>
  <h1>Two-Step Flavor Forecast for __YEAR__</h1>
  <p>Step 1: monthly forecast = average of the same month from the previous 5 years.</p>
  <p>Step 2: weekly forecast = monthly forecast split by the 5-year week-of-month sales pattern for that same flavor and store.</p>
  <p>Week-of-month buckets: days 1-7, 8-14, 15-21, 22-28, 29-end.</p>
</header>
<main>
  <div class="filters">
    <div><label for="storeFilter">Store / Unit</label><select id="storeFilter"></select></div>
    <div><label for="monthFilter">Month</label><select id="monthFilter"></select></div>
  </div>
  <div class="grid">
    <div class="card"><div class="label">Flavor Count</div><div class="value" id="flavorCount">0</div></div>
    <div class="card"><div class="label">Monthly Cases</div><div class="value" id="monthlyCases">0</div></div>
    <div class="card"><div class="label">History Years Used</div><div class="value" id="historyYears">0</div></div>
  </div>

  <h2 class="section-title">Flavor-Wise Monthly Forecast</h2>
  <div class="table-wrap">
    <table>
      <thead>
        <tr><th>Store</th><th>Flavor</th><th>Monthly Forecast</th><th>Years Used</th></tr>
      </thead>
      <tbody id="flavorRows"></tbody>
    </table>
  </div>

  <h2 class="section-title">Week-Wise Distribution</h2>
  <div class="table-wrap">
    <table>
      <thead>
        <tr><th>Store</th><th>Flavor</th><th>Week 1</th><th>Week 2</th><th>Week 3</th><th>Week 4</th><th>Week 5</th><th>Total</th></tr>
      </thead>
      <tbody id="weeklyRows"></tbody>
    </table>
  </div>
</main>
<script>
const rows = __ROWS__;
const stores = [...new Set(rows.map(r => r[0]))].sort();
const months = [...new Set(rows.map(r => r[2]))].sort();

function fillSelect(id, values) {
  const el = document.getElementById(id);
  el.innerHTML = values.map(v => `<option value="${v}">${v}</option>`).join('');
}

function fmt(n) {
  return Number(n || 0).toFixed(2);
}

function render() {
  const store = document.getElementById('storeFilter').value;
  const month = document.getElementById('monthFilter').value;
  const filtered = rows.filter(r => r[0] === store && r[2] === month).sort((a, b) => b[4] - a[4]);

  document.getElementById('flavorCount').textContent = filtered.length.toLocaleString();
  document.getElementById('monthlyCases').textContent = fmt(filtered.reduce((s, r) => s + Number(r[4] || 0), 0));
  document.getElementById('historyYears').textContent = filtered.length ? Math.max(...filtered.map(r => Number(r[5] || 0))) : 0;

  document.getElementById('flavorRows').innerHTML = filtered.map(r =>
    `<tr><td>${r[0]}</td><td>${r[3]}</td><td>${fmt(r[4])}</td><td>${r[5]}</td></tr>`
  ).join('');

  document.getElementById('weeklyRows').innerHTML = filtered.map(r => {
    const total = Number(r[6]||0) + Number(r[7]||0) + Number(r[8]||0) + Number(r[9]||0) + Number(r[10]||0);
    return `<tr><td>${r[0]}</td><td>${r[3]}</td><td>${fmt(r[6])}</td><td>${fmt(r[7])}</td><td>${fmt(r[8])}</td><td>${fmt(r[9])}</td><td>${fmt(r[10])}</td><td>${fmt(total)}</td></tr>`;
  }).join('');
}

fillSelect('storeFilter', stores);
fillSelect('monthFilter', months);
document.getElementById('storeFilter').addEventListener('change', render);
document.getElementById('monthFilter').addEventListener('change', render);
render();
</script>
</body>
</html>
'@

    $html = $html.Replace("__YEAR__", [string]$ForecastYearValue).Replace("__ROWS__", $jsRows)

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
}

$resolvedInputDir = (Resolve-Path -LiteralPath $InputDir).Path
$cacheDirPath = Join-Path $resolvedInputDir $CacheDir
$historyCsvPath = Join-Path $resolvedInputDir $HistoryCsv
$forecastCsvPath = Join-Path $resolvedInputDir $ForecastCsv
$reportHtmlPath = Join-Path $resolvedInputDir $ReportHtml
$summaryJsonPath = Join-Path $resolvedInputDir $SummaryJson
New-Item -ItemType Directory -Force -Path $cacheDirPath | Out-Null

$workbooks = Get-ChildItem -LiteralPath $resolvedInputDir -File -Filter "*.xlsx" | Sort-Object Name
if (-not $workbooks) { throw "No .xlsx files found in $resolvedInputDir" }

$historyRows = New-Object System.Collections.Generic.List[object]
$fileSummaries = New-Object System.Collections.Generic.List[object]

foreach ($workbook in $workbooks) {
    $cacheFile = Join-Path $cacheDirPath ((Get-SafeFileStem -Name $workbook.BaseName) + ".csv")
    if (Test-Path -LiteralPath $cacheFile) {
        Write-Host "Using cache for $($workbook.Name)..."
        $rows = @(Import-Csv -LiteralPath $cacheFile | ForEach-Object {
            [pscustomobject]@{
                UNIT = [string]$_.UNIT
                STORE_CODE = [string]$_.STORE_CODE
                "NDA ITEM DESCRIPTION" = [string]$_."NDA ITEM DESCRIPTION"
                "INVOICE DATE" = [string]$_."INVOICE DATE"
                ORDERED = [double]$_.ORDERED
                YEAR = [int]$_.YEAR
                MONTH = [int]$_.MONTH
                WEEK_OF_MONTH = [int]$_.WEEK_OF_MONTH
            }
        })
    } else {
        Write-Host "Reading $($workbook.Name)..."
        $rows = Import-PurchaseRows -WorkbookPath $workbook.FullName
        $rows | Export-Csv -LiteralPath $cacheFile -NoTypeInformation -Encoding UTF8
    }
    foreach ($row in $rows) { $historyRows.Add($row) }
    $fileSummaries.Add([pscustomobject]@{ file = $workbook.Name; rows = $rows.Count })
}

$historyRows | Export-Csv -LiteralPath $historyCsvPath -NoTypeInformation -Encoding UTF8

$forecastRows = New-Object System.Collections.Generic.List[object]
$forecastMonths = 1..12
$lookbackYears = @(($ForecastYear - 5)..($ForecastYear - 1))

$targetUnits = $historyRows | Select-Object -ExpandProperty UNIT -Unique | Sort-Object
$monthlyLookup = @{}
$weeklyLookup = @{}

foreach ($row in $historyRows) {
    $monthlyKey = "$($row.UNIT)||$($row.'NDA ITEM DESCRIPTION')||$($row.YEAR)||$($row.MONTH)"
    if (-not $monthlyLookup.ContainsKey($monthlyKey)) { $monthlyLookup[$monthlyKey] = 0.0 }
    $monthlyLookup[$monthlyKey] += [double]$row.ORDERED

    $weeklyKey = "$($row.UNIT)||$($row.'NDA ITEM DESCRIPTION')||$($row.YEAR)||$($row.MONTH)||$($row.WEEK_OF_MONTH)"
    if (-not $weeklyLookup.ContainsKey($weeklyKey)) { $weeklyLookup[$weeklyKey] = 0.0 }
    $weeklyLookup[$weeklyKey] += [double]$row.ORDERED
}

$unitFlavorPairs = $historyRows |
    Where-Object { $lookbackYears -contains $_.YEAR } |
    Group-Object UNIT, "NDA ITEM DESCRIPTION" |
    ForEach-Object {
        [pscustomobject]@{
            UNIT = $_.Group[0].UNIT
            STORE_CODE = $_.Group[0].STORE_CODE
            FLAVOR = $_.Group[0]."NDA ITEM DESCRIPTION"
        }
    } |
    Sort-Object UNIT, FLAVOR

foreach ($pair in $unitFlavorPairs) {
    foreach ($month in $forecastMonths) {
        $historyYears = $lookbackYears
        $yearlyTotals = New-Object System.Collections.Generic.List[double]
        foreach ($year in $historyYears) {
            $key = "$($pair.UNIT)||$($pair.FLAVOR)||$year||$month"
            $value = 0.0
            if ($monthlyLookup.ContainsKey($key)) { $value = [double]$monthlyLookup[$key] }
            $yearlyTotals.Add($value)
        }

        $monthlyForecast = if ($yearlyTotals.Count -gt 0) { ($yearlyTotals | Measure-Object -Sum).Sum / $yearlyTotals.Count } else { 0.0 }

        $weekTotals = @()
        foreach ($week in 1..5) {
            $sum = 0.0
            foreach ($year in $historyYears) {
                $key = "$($pair.UNIT)||$($pair.FLAVOR)||$year||$month||$week"
                if ($weeklyLookup.ContainsKey($key)) { $sum += [double]$weeklyLookup[$key] }
            }
            $weekTotals += $sum
        }

        $weekPatternTotal = ($weekTotals | Measure-Object -Sum).Sum
        if ($weekPatternTotal -le 0) {
            $shares = @(0.25, 0.25, 0.25, 0.25, 0.0)
        } else {
            $shares = $weekTotals | ForEach-Object { $_ / $weekPatternTotal }
        }

        $weekValues = $shares | ForEach-Object { [math]::Round($monthlyForecast * $_, 4) }
        $monthKey = "{0:D4}-{1:D2}" -f $ForecastYear, $month

        $forecastRows.Add([pscustomobject]@{
            unit = $pair.UNIT
            store_code = $pair.STORE_CODE
            month_key = $monthKey
            flavor = $pair.FLAVOR
            monthly_forecast_cases = [math]::Round($monthlyForecast, 4)
            history_years_used = 5
            week_1_cases = $weekValues[0]
            week_2_cases = $weekValues[1]
            week_3_cases = $weekValues[2]
            week_4_cases = $weekValues[3]
            week_5_cases = $weekValues[4]
        })
    }
}

$forecastRows | Export-Csv -LiteralPath $forecastCsvPath -NoTypeInformation -Encoding UTF8
New-HtmlReport -ForecastRows $forecastRows.ToArray() -OutputPath $reportHtmlPath -ForecastYearValue $ForecastYear

$summary = [pscustomobject]@{
    generated_at = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssK")
    forecast_year = $ForecastYear
    files = $fileSummaries
    cache_directory = $cacheDirPath
    history_rows = $historyRows.Count
    forecast_rows = $forecastRows.Count
    history_csv = $historyCsvPath
    forecast_csv = $forecastCsvPath
    report_html = $reportHtmlPath
    method = @{
        step_1 = "Monthly forecast is the average of the same month across the previous 5 years, with missing years treated as zero."
        step_2 = "Weekly forecast uses the week-of-month share from the same store, flavor, and month across the previous 5 years."
        week_definition = "Week 1 = days 1-7, Week 2 = days 8-14, Week 3 = days 15-21, Week 4 = days 22-28, Week 5 = days 29-end."
    }
}

$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $summaryJsonPath -Encoding UTF8

Write-Host "Wrote history: $historyCsvPath"
Write-Host "Wrote forecast: $forecastCsvPath"
Write-Host "Wrote report: $reportHtmlPath"
Write-Host "Wrote summary: $summaryJsonPath"
