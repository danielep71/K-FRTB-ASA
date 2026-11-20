<#
.SYNOPSIS
    Adds the FRTB-SA ribbon tab to a macro-enabled workbook.

.DESCRIPTION
    Excel's object model saves a workbook without a custom UI part, so the
    part is added to the saved file as a zip package: src\ribbon\customUI14.xml
    becomes customUI/customUI14.xml, and a relationship of the package points
    to it. The workbook's other parts, the VBA project included, are not
    touched. Running it twice leaves one part and one relationship.

    Close the workbook first: Excel locks the file while it is open.

.PARAMETER Workbook
    The .xlsm to change.

.PARAMETER Xml
    The custom UI part. Defaults to src\ribbon\customUI14.xml.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\Add-RibbonTab.ps1 -Workbook dist\K-FRTB-ASA.xlsm
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Workbook,
    [string]$Xml
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$partName = 'customUI/customUI14.xml'
$relationshipType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
$relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
$contentTypesNamespace = 'http://schemas.openxmlformats.org/package/2006/content-types'

if (-not $Xml) { $Xml = Join-Path (Split-Path -Parent $PSScriptRoot) 'src\ribbon\customUI14.xml' }
if (-not (Test-Path -LiteralPath $Workbook)) { throw "Workbook not found: $Workbook" }
if (-not (Test-Path -LiteralPath $Xml)) { throw "Custom UI part not found: $Xml" }
$Workbook = (Resolve-Path -LiteralPath $Workbook).Path
$bytes = [System.IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Xml).Path)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-Entry {
    param($Archive, [string]$Name)
    $entry = $Archive.GetEntry($Name)
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open())
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Write-Entry {
    param($Archive, [string]$Name, [byte[]]$Content)
    $existing = $Archive.GetEntry($Name)
    if ($null -ne $existing) { $existing.Delete() }
    $entry = $Archive.CreateEntry($Name)
    $stream = $entry.Open()
    try { $stream.Write($Content, 0, $Content.Length) } finally { $stream.Dispose() }
}

$archive = [System.IO.Compression.ZipFile]::Open($Workbook, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    Write-Entry $archive $partName $bytes

    # The package relationship that points to the part
    $relsText = Read-Entry $archive '_rels/.rels'
    if ($null -eq $relsText) { throw 'The workbook has no _rels/.rels part: it is not an Office package.' }
    $rels = New-Object System.Xml.XmlDocument
    $rels.LoadXml($relsText)
    $known = $false
    foreach ($node in $rels.DocumentElement.ChildNodes) {
        if ($node -is [System.Xml.XmlElement] -and $node.GetAttribute('Type') -eq $relationshipType) { $known = $true }
    }
    if (-not $known) {
        $node = $rels.CreateElement('Relationship', $relationshipNamespace)
        $node.SetAttribute('Id', 'rIdFrtbRibbon')
        $node.SetAttribute('Type', $relationshipType)
        $node.SetAttribute('Target', $partName)
        [void]$rels.DocumentElement.AppendChild($node)
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $buffer = New-Object System.IO.MemoryStream
        $writer = [System.Xml.XmlWriter]::Create($buffer, $settings)
        $rels.Save($writer)
        $writer.Dispose()
        Write-Entry $archive '_rels/.rels' $buffer.ToArray()
    }

    # An xml part needs the default content type of the extension
    $typesText = Read-Entry $archive '[Content_Types].xml'
    if ($null -eq $typesText) { throw 'The workbook has no [Content_Types].xml part: it is not an Office package.' }
    $types = New-Object System.Xml.XmlDocument
    $types.LoadXml($typesText)
    $hasXml = $false
    foreach ($node in $types.DocumentElement.ChildNodes) {
        if ($node -is [System.Xml.XmlElement] -and $node.LocalName -eq 'Default' -and $node.GetAttribute('Extension') -eq 'xml') { $hasXml = $true }
    }
    if (-not $hasXml) {
        $node = $types.CreateElement('Default', $contentTypesNamespace)
        $node.SetAttribute('Extension', 'xml')
        $node.SetAttribute('ContentType', 'application/xml')
        [void]$types.DocumentElement.PrependChild($node)
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $buffer = New-Object System.IO.MemoryStream
        $writer = [System.Xml.XmlWriter]::Create($buffer, $settings)
        $types.Save($writer)
        $writer.Dispose()
        Write-Entry $archive '[Content_Types].xml' $buffer.ToArray()
    }
} finally {
    $archive.Dispose()
}
Write-Host "Added the ribbon tab to $Workbook"
