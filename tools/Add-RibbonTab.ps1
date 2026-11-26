<#
.SYNOPSIS
    Adds the FRTB-ASA ribbon tab to a macro-enabled workbook.

.DESCRIPTION
    Excel's object model saves a workbook without a custom UI part, so the
    part is added to the saved file as a zip package: src\ribbon\customUI14.xml
    becomes customUI/customUI14.xml, and a relationship of the package points
    to it. The workbook's other parts, the VBA project included, are not
    touched. Running it twice leaves one part and one relationship.

    A control whose image attribute names an icon gets that icon from
    src\ribbon\icons\<name>.png: the script writes it to customUI/images,
    adds the relationships of the part that point to it and makes sure the
    package has the default content type of png.

    Close the workbook first: Excel locks the file while it is open.

.PARAMETER Workbook
    The .xlsm to change.

.PARAMETER Xml
    The custom UI part. Defaults to src\ribbon\customUI14.xml. Its icons are
    read from the icons folder next to it.

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
$Xml = (Resolve-Path -LiteralPath $Xml).Path
$bytes = [System.IO.File]::ReadAllBytes($Xml)

# The icons the part names: the value of every image attribute is a
# relationship id and the name of a png in the icons folder next to the part
$partDocument = New-Object System.Xml.XmlDocument
$partDocument.LoadXml([System.Text.Encoding]::UTF8.GetString($bytes))
$icons = @()
$seen = @{}
foreach ($node in $partDocument.SelectNodes('//*[@image]')) {
    $name = [string]$node.GetAttribute('image')
    if ($name -notmatch '^[A-Za-z0-9_-]+$') { throw "Icon name not allowed: $name" }
    $file = Join-Path (Join-Path (Split-Path -Parent $Xml) 'icons') "$name.png"
    if (-not (Test-Path -LiteralPath $file)) { throw "Icon not found: $file" }
    if (-not $seen.ContainsKey($name)) {
        $seen[$name] = $true
        $icons += [pscustomobject]@{ Name = $name; Path = $file }
    }
}

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

    # The icons and the relationships of the part that point to them
    if ($icons.Count -gt 0) {
        $imageType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
        $partRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
                    '<Relationships xmlns="' + $relationshipNamespace + '">'
        foreach ($icon in $icons) {
            Write-Entry $archive "customUI/images/$($icon.Name).png" ([System.IO.File]::ReadAllBytes($icon.Path))
            $partRels += '<Relationship Id="' + $icon.Name + '" Type="' + $imageType + '" Target="images/' + $icon.Name + '.png"/>'
        }
        $partRels += '</Relationships>'
        Write-Entry $archive 'customUI/_rels/customUI14.xml.rels' ((New-Object System.Text.UTF8Encoding($false)).GetBytes($partRels))
    }

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

    # A part needs the default content type of its extension: xml for the part, png for the icons
    $typesText = Read-Entry $archive '[Content_Types].xml'
    if ($null -eq $typesText) { throw 'The workbook has no [Content_Types].xml part: it is not an Office package.' }
    $types = New-Object System.Xml.XmlDocument
    $types.LoadXml($typesText)
    $wanted = @(@{ Extension = 'xml'; ContentType = 'application/xml' })
    if ($icons.Count -gt 0) { $wanted += @{ Extension = 'png'; ContentType = 'image/png' } }
    $changed = $false
    foreach ($default in $wanted) {
        $has = $false
        foreach ($node in $types.DocumentElement.ChildNodes) {
            if ($node -is [System.Xml.XmlElement] -and $node.LocalName -eq 'Default' -and $node.GetAttribute('Extension') -eq $default.Extension) { $has = $true }
        }
        if (-not $has) {
            $node = $types.CreateElement('Default', $contentTypesNamespace)
            $node.SetAttribute('Extension', $default.Extension)
            $node.SetAttribute('ContentType', $default.ContentType)
            [void]$types.DocumentElement.PrependChild($node)
            $changed = $true
        }
    }
    if ($changed) {
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
