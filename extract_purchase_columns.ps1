param(
    [string]$InputDir = ".",
    [string]$OutputCsv = "combined_purchase_columns.csv",
    [string]$OutputJson = "combined_purchase_columns_summary.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-NormalizedHeader {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    return (($Value -replace "[^A-Za-z0-9]+", " ").Trim().ToUpperInvariant() -replace "\s+", " ")
}

function Get-ColumnLetters {
    param([string]$CellReference)

    if ($CellReference -match "^[A-Z]+") {
        return $matches[0]
    }

    return $null
}

function Get-XmlDocumentFromZip {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) {
        throw "Missing ZIP entry: $EntryName"
    }

    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
        $content = $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }

    return [xml]$content
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip)

    $entry = $Zip.GetEntry("xl/sharedStrings.xml")
    if (-not $entry) {
        return @()
    }

    $doc = Get-XmlDocumentFromZip -Zip $Zip -EntryName "xl/sharedStrings.xml"
    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    $values = New-Object System.Collections.Generic.List[string]
    foreach ($si in $doc.SelectNodes("//x:si", $ns)) {
        $parts = $si.SelectNodes(".//x:t", $ns)
        if ($parts.Count -gt 0) {
            $values.Add(($parts | ForEach-Object { $_.InnerText }) -join "")
        }
        else {
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
            if ($sanitized -match "[ymdhis]") {
                $isDate = $true
            }
        }

        if ($isDate) {
            [void]$dateStyleIndexes.Add($i)
        }
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
    if (-not $targetSheet) {
        $targetSheet = $sheetNodes | Select-Object -First 1
    }

    if (-not $targetSheet) {
        throw "No worksheets found in workbook."
    }

    $relationshipId = $targetSheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $relationship = $rels.SelectSingleNode("//pr:Relationship[@Id='$relationshipId']", $relsNs)
    if (-not $relationship) {
        throw "Could not resolve sheet relationship id $relationshipId."
    }

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

    if ([string]::IsNullOrEmpty($RawValue)) {
        return $null
    }

    switch ($CellType) {
        "s" {
            return $SharedStrings[[int]$RawValue]
        }
        "b" {
            return ([int]$RawValue -eq 1).ToString()
        }
        default {
            if ($DateStyleIndexes.Contains($StyleIndex)) {
                try {
                    return [DateTime]::FromOADate([double]$RawValue).ToString("yyyy-MM-dd")
                }
                catch {
                    return $RawValue
                }
            }

            return $RawValue
        }
    }
}

function Get-InnerTextFromReader {
    param([System.Xml.XmlReader]$Reader)

    return $Reader.ReadElementContentAsString()
}

function Get-RequiredColumnMap {
    return @{
        "UNIT" = "UNIT"
        "NDA ITEM DESCRIPTION" = "NDA ITEM DESCRIPTION"
        "INVOICE DATE" = "INVOICE DATE"
        "ORDERED" = "ORDERED"
    }
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
        if (-not $sheetEntry) {
            throw "Worksheet entry not found: $sheetEntryPath"
        }

        $rows = New-Object System.Collections.Generic.List[object]
        $headerMap = @{}
        $headerRowNumber = $null

        $settings = [System.Xml.XmlReaderSettings]::new()
        $settings.IgnoreWhitespace = $true
        $reader = [System.Xml.XmlReader]::Create($sheetEntry.Open(), $settings)

        try {
            while ($reader.Read()) {
                if ($reader.NodeType -ne [System.Xml.XmlNodeType]::Element -or $reader.Name -ne "row") {
                    continue
                }

                $rowNumber = [int]$reader.GetAttribute("r")
                $rowReader = $reader.ReadSubtree()
                $cells = @{}

                try {
                    while ($rowReader.Read()) {
                        if ($rowReader.NodeType -ne [System.Xml.XmlNodeType]::Element -or $rowReader.Name -ne "c") {
                            continue
                        }

                        $cellReference = $rowReader.GetAttribute("r")
                        $columnLetters = Get-ColumnLetters -CellReference $cellReference
                        $cellType = $rowReader.GetAttribute("t")
                        $styleIndexValue = $rowReader.GetAttribute("s")
                        $styleIndex = if ($null -ne $styleIndexValue -and $styleIndexValue -ne "") { [int]$styleIndexValue } else { -1 }

                        $cellSubtree = $rowReader.ReadSubtree()
                        $rawValue = $null

                        try {
                            while ($cellSubtree.Read()) {
                                if ($cellSubtree.NodeType -ne [System.Xml.XmlNodeType]::Element) {
                                    continue
                                }

                                if ($cellSubtree.Name -eq "v") {
                                    $rawValue = Get-InnerTextFromReader -Reader $cellSubtree
                                }
                                elseif ($cellSubtree.Name -eq "t" -and $cellType -eq "inlineStr") {
                                    $rawValue = Get-InnerTextFromReader -Reader $cellSubtree
                                }
                            }
                        }
                        finally {
                            $cellSubtree.Dispose()
                        }

                        $cells[$columnLetters] = Convert-CellValue -RawValue $rawValue -CellType $cellType -StyleIndex $styleIndex -SharedStrings $sharedStrings -DateStyleIndexes $dateStyleIndexes
                    }
                }
                finally {
                    $rowReader.Dispose()
                }

                if ($cells.Count -eq 0) {
                    continue
                }

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

                if ($rowNumber -le $headerRowNumber) {
                    continue
                }

                $unit = [string]$cells[$headerMap["UNIT"]]
                $itemDescription = [string]$cells[$headerMap["NDA ITEM DESCRIPTION"]]
                $invoiceDate = [string]$cells[$headerMap["INVOICE DATE"]]
                $orderedValue = $cells[$headerMap["ORDERED"]]

                if ([string]::IsNullOrWhiteSpace($unit) -and [string]::IsNullOrWhiteSpace($itemDescription) -and [string]::IsNullOrWhiteSpace($invoiceDate) -and [string]::IsNullOrWhiteSpace([string]$orderedValue)) {
                    continue
                }

                $rows.Add([pscustomobject]@{
                    UNIT = $unit
                    "NDA ITEM DESCRIPTION" = $itemDescription
                    "INVOICE DATE" = $invoiceDate
                    ORDERED = $orderedValue
                })
            }
        }
        finally {
            $reader.Dispose()
        }

        return $rows
    }
    finally {
        $zip.Dispose()
    }
}

$resolvedInputDir = (Resolve-Path -LiteralPath $InputDir).Path
$csvPath = Join-Path $resolvedInputDir $OutputCsv
$jsonPath = Join-Path $resolvedInputDir $OutputJson

$workbooks = Get-ChildItem -LiteralPath $resolvedInputDir -File -Filter "*.xlsx" |
    Sort-Object Name

if (-not $workbooks) {
    throw "No .xlsx files found in $resolvedInputDir"
}

$allRows = New-Object System.Collections.Generic.List[object]
$fileSummaries = New-Object System.Collections.Generic.List[object]

foreach ($workbook in $workbooks) {
    Write-Host "Reading $($workbook.Name)..."
    $rows = Import-PurchaseRows -WorkbookPath $workbook.FullName
    foreach ($row in $rows) {
        $allRows.Add($row)
    }

    $fileSummaries.Add([pscustomobject]@{
        file = $workbook.Name
        rows = $rows.Count
    })
}

$allRows |
    Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

$summary = [pscustomobject]@{
    generated_at = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssK")
    input_directory = $resolvedInputDir
    output_csv = $csvPath
    total_rows = $allRows.Count
    files = $fileSummaries
    columns = @(
        "UNIT",
        "NDA ITEM DESCRIPTION",
        "INVOICE DATE",
        "ORDERED"
    )
    notes = @(
        "UNIT is the location name.",
        "Rows were extracted from the EXPORT sheet when present."
    )
}

$summary | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

Write-Host "Wrote $($allRows.Count) rows to $csvPath"
Write-Host "Wrote summary to $jsonPath"
