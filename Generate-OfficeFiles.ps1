<#
.SYNOPSIS
    Generates Office documents (.docx, .xlsx, .pptx) for demo data.
.DESCRIPTION
    Creates real Office Open XML files from demo content.
    Run once to generate files in data/files/, then upload via Load-DemoData.ps1.
    Requires .NET Framework (System.IO.Packaging) - ships with Windows PS 5.1.
#>

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName WindowsBase

$outDir = Join-Path $PSScriptRoot "data\files"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

# ════════════════════════════════════════════════════════════════════
# Helper: Create an OpenXML package with parts
# ════════════════════════════════════════════════════════════════════

function New-OpenXmlPackage {
    param(
        [string]$FilePath,
        [hashtable[]]$Parts,       # @{ Uri; ContentType; Content; RelType }
        [string]$ContentTypesXml
    )
    if (Test-Path $FilePath) { Remove-Item $FilePath -Force }

    $pkg = [System.IO.Packaging.Package]::Open($FilePath, [System.IO.FileMode]::Create)
    try {
        foreach ($part in $Parts) {
            $uri  = New-Object System.Uri($part.Uri, [System.UriKind]::Relative)
            $p    = $pkg.CreatePart($uri, $part.ContentType, [System.IO.Packaging.CompressionOption]::Normal)
            $stream = $p.GetStream()
            $bytes  = [System.Text.Encoding]::UTF8.GetBytes($part.Content)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Close()
            if ($part.RelType) {
                $pkg.CreateRelationship($uri, [System.IO.Packaging.TargetMode]::Internal, $part.RelType) | Out-Null
            }
        }
    }
    finally { $pkg.Close() }
}

# ════════════════════════════════════════════════════════════════════
# 1. Adatum Corp Briefing Deck.pptx
# ════════════════════════════════════════════════════════════════════

Write-Host "Generating Adatum Corp Briefing Deck.pptx..." -ForegroundColor Cyan

$pptxNs = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

function New-SlideXml {
    param([string]$Title, [string]$Body)
    $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
    $escapedBody  = [System.Security.SecurityElement]::Escape($Body)
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>$escapedTitle</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0" sz="1600"/><a:t>$escapedBody</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
"@
}

$slides = @(
    @{ Title = "Apex Manufacturing x Adatum Corp"; Body = "Strategic Partnership for Manufacturing Automation`nConfidential" },
    @{ Title = "Agenda"; Body = "1. Apex Manufacturing Overview`n2. ProLine X Platform Demo`n3. Adatum-Specific Use Cases`n4. Implementation Approach (8-Week Plan)`n5. ROI Analysis and Proof Points`n6. Reference Customers`n7. Pricing and Next Steps" },
    @{ Title = "Apex Manufacturing at a Glance"; Body = "Founded 2008, headquartered in Chicago`n2,400+ enterprise customers across manufacturing`n`$340M ARR, 28% YoY growth`nKey verticals: Discrete, Process, Food and Beverage`nStrategic partnerships: Microsoft Azure IoT, SAP, Siemens" },
    @{ Title = "Why ProLine X"; Body = "Next-generation manufacturing automation`n`n- 40% faster processing than ProLine S`n- Built-in AI/ML for predictive maintenance`n- Azure IoT Hub native integration`n- 8-week deployment (vs 4-6 months competitors)`n- Modular design - start small, scale fast" },
    @{ Title = "Adatum-Specific Value"; Body = "Current ProLine S results: 18.3% throughput improvement`n3 plants running since Q3 2025, high satisfaction`n`nProLine X Opportunity:`n- Expand to 3 additional plants (Detroit, Phoenix, Toronto)`n- Add predictive maintenance (est. 25% downtime reduction)`n- Azure IoT for real-time monitoring`n- Projected: 30-40% total throughput improvement" },
    @{ Title = "Competitive Comparison"; Body = "vs Fabrikam X200:`n- We deploy in 8 weeks vs their 6 months`n- AI included vs `$150K add-on`n`nvs Northwind NX-Pro:`n- Native AI/ML capabilities (they have none)`n- Azure IoT integration (theirs is custom/fragile)`n`nKey message: Fastest time to value with built-in intelligence" },
    @{ Title = "8-Week Implementation Plan"; Body = "Week 1-2: Discovery and Architecture`nWeek 3-4: Core Platform Deployment`nWeek 5-6: Integration and Customization`nWeek 7: Testing and Validation`nWeek 8: Go-Live and Handoff`n`nDedicated team: Solution Architect + 2 Implementation Engineers`nAdatum requirement: 1 Plant Manager + 1 IT Lead" },
    @{ Title = "ROI Analysis"; Body = "Investment:`n- Platform License: `$1,200,000`n- Implementation Services: `$480,000`n- Annual Support: `$240,000/yr`n- IoT Sensor Package (optional): `$480,000`n- Total Year 1: `$2,400,000`n`nReturns:`n- Projected annual savings: `$2.1M across 3 plants`n- Payback period: 14 months`n- 3-year ROI: 162%" },
    @{ Title = "Reference Customers"; Body = "Contoso Manufacturing`n`"ProLine S transformed our Detroit plant. Expanding to ProLine X.`"`n`nAdventure Works`n`"8-week deployment was accurate. We were live ahead of schedule.`"`n`nTailspin Industries`n`"The AI capabilities caught issues our old system missed completely.`"" },
    @{ Title = "Proposed Next Steps"; Body = "1. [Today] Align on scope and timeline`n2. [This week] Technical assessment of Plant 1`n3. [Next week] Final proposal with SOW`n4. [May 15] Contract review`n5. [June 1] Kick-off implementation`n6. [June 30] Budget commitment deadline (Adatum FY)" }
)

# Build slide parts and relationship entries
$pptParts = @()
$slideRelEntries = ""
$slideListEntries = ""
$contentTypeEntries = ""

for ($i = 0; $i -lt $slides.Count; $i++) {
    $num = $i + 1
    $pptParts += @{
        Uri         = "/ppt/slides/slide${num}.xml"
        ContentType = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
        Content     = (New-SlideXml -Title $slides[$i].Title -Body $slides[$i].Body)
        RelType     = $null
    }
    $slideRelEntries += "  <Relationship Id=`"rId$num`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`" Target=`"slides/slide${num}.xml`"/>`n"
    $slideListEntries += "    <p:sldId id=`"$($256 + $i)`" r:id=`"rId$num`"/>`n"
    $contentTypeEntries += "  <Override PartName=`"/ppt/slides/slide${num}.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.slide+xml`"/>`n"
}

$presentationXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst/>
  <p:sldIdLst>
$slideListEntries  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"@

$presRelXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
$slideRelEntries</Relationships>
"@

$pptParts += @{
    Uri         = "/ppt/presentation.xml"
    ContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    Content     = $presentationXml
    RelType     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
}

$pptParts += @{
    Uri         = "/ppt/_rels/presentation.xml.rels"
    ContentType = "application/vnd.openxmlformats-package.relationships+xml"
    Content     = $presRelXml
    RelType     = $null
}

New-OpenXmlPackage -FilePath (Join-Path $outDir "Adatum Corp Briefing Deck.pptx") -Parts $pptParts
Write-Host "  [OK] Adatum Corp Briefing Deck.pptx" -ForegroundColor Green

# ════════════════════════════════════════════════════════════════════
# 2. ProLine X Launch Plan - Exec Summary.docx
# ════════════════════════════════════════════════════════════════════

Write-Host "Generating ProLine X Launch Plan - Exec Summary.docx..." -ForegroundColor Cyan

$docBody = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Title"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="48"/></w:rPr><w:t>ProLine X Launch Plan</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="28"/><w:color w:val="666666"/></w:rPr><w:t>Executive Summary - DRAFT</w:t></w:r></w:p>
    <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:color w:val="999999"/></w:rPr><w:t>Author: James Bowen, VP of Operations | Status: INCOMPLETE DRAFT</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Overview</w:t></w:r></w:p>
    <w:p><w:r><w:t>ProLine X is Apex Manufacturing's most significant product launch in 2026. Product engineering has confirmed GA readiness for June 9, 2026. This document outlines the cross-functional launch plan including go-to-market strategy, target accounts, competitive positioning, and key milestones.</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Launch Objective</w:t></w:r></w:p>
    <w:p><w:r><w:t>Successfully bring ProLine X to market by June 9 with sales team fully enabled, first 10 target accounts engaged, competitive positioning locked, customer reference program activated, and press/analyst briefings scheduled.</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Key Stakeholders</w:t></w:r></w:p>
    <w:p><w:r><w:t>Molly Clark, Product Manager - Product readiness, features, demo environment</w:t></w:r></w:p>
    <w:p><w:r><w:t>Spencer Low, VP of Sales - Sales enablement, pricing, target accounts</w:t></w:r></w:p>
    <w:p><w:r><w:t>Renee Lo, Marketing Lead - Competitive positioning, battle cards, launch comms</w:t></w:r></w:p>
    <w:p><w:r><w:t>David So, Chief of Staff - Program management, logistics, exec reporting</w:t></w:r></w:p>
    <w:p><w:r><w:t>James Bowen, VP of Operations - Overall launch leadership, exec summary, board update</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Product Highlights</w:t></w:r></w:p>
    <w:p><w:r><w:t>- 40% faster processing than ProLine S</w:t></w:r></w:p>
    <w:p><w:r><w:t>- Built-in AI/ML for predictive maintenance</w:t></w:r></w:p>
    <w:p><w:r><w:t>- Azure IoT Hub native integration</w:t></w:r></w:p>
    <w:p><w:r><w:t>- New modular design - customers can start small and expand</w:t></w:r></w:p>
    <w:p><w:r><w:t>- Starting price: $800K (AI included, unlike competitors)</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Target Accounts (Top 10 by Pipeline Value)</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr>
      <w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="0078D4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Rank</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="0078D4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Account</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="0078D4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Pipeline</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="0078D4"/></w:tcPr><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>Status</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Adatum Corp</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$2.4M</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Active engagement, meeting this week</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Contoso Manufacturing</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$1.8M</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Existing ProLine S customer</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Tailspin Industries</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$1.2M</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Greenfield, strong AI interest</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>4</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Wide World Importers</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$900K</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Competitive displacement</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>5</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Adventure Works</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$850K</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Expanding from pilot</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>6</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Proseware Inc</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$750K</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>New logo prospect</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>7</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Trey Research</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$600K</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R&amp;D use case</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>8</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Litware Inc</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$550K</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Mid-market entry</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>9</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Datum Corp</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$500K</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Partner referral</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>10</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Munson's Pickles</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$450K</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Food manufacturing vertical</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Total addressable pipeline: $10M+</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Competitive Landscape</w:t></w:r></w:p>
    <w:p><w:r><w:t>vs Fabrikam X200: Faster deployment (8wk vs 6mo), AI included (vs $150K add-on), Azure native IoT</w:t></w:r></w:p>
    <w:p><w:r><w:t>vs Northwind NX-Pro: Native AI/ML (they have none), modern architecture, better modular design</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Value proposition: "Fastest time to value with built-in intelligence - start small, scale fast."</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Key Milestones</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>[TODO - Needs completion: Add milestone dates, owners, and dependencies]</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Risks and Mitigations</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>[TODO - Needs completion: Engineering delay risk, Fabrikam pricing counter, reference customer availability]</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Budget Estimate</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>[TODO - Needs completion: Marketing budget, sales enablement costs, event budget]</w:t></w:r></w:p>
    <w:p/>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>
"@

$docRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"@

$docParts = @(
    @{ Uri = "/word/document.xml"; ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"; Content = $docBody; RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" },
    @{ Uri = "/word/_rels/document.xml.rels"; ContentType = "application/vnd.openxmlformats-package.relationships+xml"; Content = $docRelsXml; RelType = $null }
)

New-OpenXmlPackage -FilePath (Join-Path $outDir "ProLine X Launch Plan - Exec Summary.docx") -Parts $docParts
Write-Host "  [OK] ProLine X Launch Plan - Exec Summary.docx" -ForegroundColor Green

# ════════════════════════════════════════════════════════════════════
# 3. ProLine X Pipeline Tracker.xlsx
# ════════════════════════════════════════════════════════════════════

Write-Host "Generating ProLine X Pipeline Tracker.xlsx..." -ForegroundColor Cyan

$sheetData = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <cols>
    <col min="1" max="1" width="5" customWidth="1"/>
    <col min="2" max="2" width="25" customWidth="1"/>
    <col min="3" max="3" width="15" customWidth="1"/>
    <col min="4" max="4" width="18" customWidth="1"/>
    <col min="5" max="5" width="18" customWidth="1"/>
    <col min="6" max="6" width="30" customWidth="1"/>
    <col min="7" max="7" width="20" customWidth="1"/>
    <col min="8" max="8" width="15" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Rank</t></is></c><c r="B1" t="inlineStr"><is><t>Account</t></is></c><c r="C1" t="inlineStr"><is><t>Pipeline Value</t></is></c><c r="D1" t="inlineStr"><is><t>Stage</t></is></c><c r="E1" t="inlineStr"><is><t>Owner</t></is></c><c r="F1" t="inlineStr"><is><t>Notes</t></is></c><c r="G1" t="inlineStr"><is><t>Next Step</t></is></c><c r="H1" t="inlineStr"><is><t>Close Date</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>1</t></is></c><c r="B2" t="inlineStr"><is><t>Adatum Corp</t></is></c><c r="C2" t="inlineStr"><is><t>$2,400,000</t></is></c><c r="D2" t="inlineStr"><is><t>Proposal</t></is></c><c r="E2" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F2" t="inlineStr"><is><t>Customer meeting Thursday. VP Procurement is decision maker.</t></is></c><c r="G2" t="inlineStr"><is><t>Send SOW by May 1</t></is></c><c r="H2" t="inlineStr"><is><t>Jun 30</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>2</t></is></c><c r="B3" t="inlineStr"><is><t>Contoso Manufacturing</t></is></c><c r="C3" t="inlineStr"><is><t>$1,800,000</t></is></c><c r="D3" t="inlineStr"><is><t>Qualified</t></is></c><c r="E3" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F3" t="inlineStr"><is><t>Existing ProLine S customer. Natural upgrade path.</t></is></c><c r="G3" t="inlineStr"><is><t>Schedule demo</t></is></c><c r="H3" t="inlineStr"><is><t>Jul 15</t></is></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>3</t></is></c><c r="B4" t="inlineStr"><is><t>Tailspin Industries</t></is></c><c r="C4" t="inlineStr"><is><t>$1,200,000</t></is></c><c r="D4" t="inlineStr"><is><t>Discovery</t></is></c><c r="E4" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F4" t="inlineStr"><is><t>Greenfield. CTO is pushing AI-first vendor selection.</t></is></c><c r="G4" t="inlineStr"><is><t>Intro call with CTO</t></is></c><c r="H4" t="inlineStr"><is><t>Aug 30</t></is></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>4</t></is></c><c r="B5" t="inlineStr"><is><t>Wide World Importers</t></is></c><c r="C5" t="inlineStr"><is><t>$900,000</t></is></c><c r="D5" t="inlineStr"><is><t>Qualified</t></is></c><c r="E5" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F5" t="inlineStr"><is><t>Currently on Fabrikam. Frustrated with deployment timeline.</t></is></c><c r="G5" t="inlineStr"><is><t>Competitive pitch</t></is></c><c r="H5" t="inlineStr"><is><t>Jul 30</t></is></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>5</t></is></c><c r="B6" t="inlineStr"><is><t>Adventure Works</t></is></c><c r="C6" t="inlineStr"><is><t>$850,000</t></is></c><c r="D6" t="inlineStr"><is><t>Negotiation</t></is></c><c r="E6" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F6" t="inlineStr"><is><t>Expanding from successful pilot. High satisfaction.</t></is></c><c r="G6" t="inlineStr"><is><t>Final pricing review</t></is></c><c r="H6" t="inlineStr"><is><t>Jun 15</t></is></c></row>
    <row r="7"><c r="A7" t="inlineStr"><is><t>6</t></is></c><c r="B7" t="inlineStr"><is><t>Proseware Inc</t></is></c><c r="C7" t="inlineStr"><is><t>$750,000</t></is></c><c r="D7" t="inlineStr"><is><t>Prospecting</t></is></c><c r="E7" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F7" t="inlineStr"><is><t>New logo. Referred by partner channel.</t></is></c><c r="G7" t="inlineStr"><is><t>Partner intro</t></is></c><c r="H7" t="inlineStr"><is><t>Sep 30</t></is></c></row>
    <row r="8"><c r="A8" t="inlineStr"><is><t>7</t></is></c><c r="B8" t="inlineStr"><is><t>Trey Research</t></is></c><c r="C8" t="inlineStr"><is><t>$600,000</t></is></c><c r="D8" t="inlineStr"><is><t>Discovery</t></is></c><c r="E8" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F8" t="inlineStr"><is><t>R&amp;D innovation use case. Interested in AI/ML.</t></is></c><c r="G8" t="inlineStr"><is><t>Technical deep dive</t></is></c><c r="H8" t="inlineStr"><is><t>Sep 15</t></is></c></row>
    <row r="9"><c r="A9" t="inlineStr"><is><t>8</t></is></c><c r="B9" t="inlineStr"><is><t>Litware Inc</t></is></c><c r="C9" t="inlineStr"><is><t>$550,000</t></is></c><c r="D9" t="inlineStr"><is><t>Prospecting</t></is></c><c r="E9" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F9" t="inlineStr"><is><t>Mid-market entry point. Good fit for modular design.</t></is></c><c r="G9" t="inlineStr"><is><t>Initial outreach</t></is></c><c r="H9" t="inlineStr"><is><t>Oct 30</t></is></c></row>
    <row r="10"><c r="A10" t="inlineStr"><is><t>9</t></is></c><c r="B10" t="inlineStr"><is><t>Datum Corp</t></is></c><c r="C10" t="inlineStr"><is><t>$500,000</t></is></c><c r="D10" t="inlineStr"><is><t>Prospecting</t></is></c><c r="E10" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F10" t="inlineStr"><is><t>Partner referral. Manufacturing modernization project.</t></is></c><c r="G10" t="inlineStr"><is><t>Qualification call</t></is></c><c r="H10" t="inlineStr"><is><t>Oct 15</t></is></c></row>
    <row r="11"><c r="A11" t="inlineStr"><is><t>10</t></is></c><c r="B11" t="inlineStr"><is><t>Munson's Pickles</t></is></c><c r="C11" t="inlineStr"><is><t>$450,000</t></is></c><c r="D11" t="inlineStr"><is><t>Discovery</t></is></c><c r="E11" t="inlineStr"><is><t>Spencer Low</t></is></c><c r="F11" t="inlineStr"><is><t>Food manufacturing vertical. FDA compliance is key.</t></is></c><c r="G11" t="inlineStr"><is><t>Compliance review</t></is></c><c r="H11" t="inlineStr"><is><t>Sep 30</t></is></c></row>
  </sheetData>
</worksheet>
"@

$wbXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Pipeline" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"@

$wbRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"@

$xlParts = @(
    @{ Uri = "/xl/worksheets/sheet1.xml"; ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"; Content = $sheetData; RelType = $null },
    @{ Uri = "/xl/workbook.xml"; ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"; Content = $wbXml; RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" },
    @{ Uri = "/xl/_rels/workbook.xml.rels"; ContentType = "application/vnd.openxmlformats-package.relationships+xml"; Content = $wbRelsXml; RelType = $null }
)

New-OpenXmlPackage -FilePath (Join-Path $outDir "ProLine X Pipeline Tracker.xlsx") -Parts $xlParts
Write-Host "  [OK] ProLine X Pipeline Tracker.xlsx" -ForegroundColor Green

# ════════════════════════════════════════════════════════════════════
# 4. ProLine X Competitive Analysis.xlsx
# ════════════════════════════════════════════════════════════════════

Write-Host "Generating ProLine X Competitive Analysis.xlsx..." -ForegroundColor Cyan

$compSheet = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <cols>
    <col min="1" max="1" width="25" customWidth="1"/>
    <col min="2" max="2" width="22" customWidth="1"/>
    <col min="3" max="3" width="22" customWidth="1"/>
    <col min="4" max="4" width="22" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Feature</t></is></c><c r="B1" t="inlineStr"><is><t>Apex ProLine X</t></is></c><c r="C1" t="inlineStr"><is><t>Fabrikam X200</t></is></c><c r="D1" t="inlineStr"><is><t>Northwind NX-Pro</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Processing Speed</t></is></c><c r="B2" t="inlineStr"><is><t>40% faster (BEST)</t></is></c><c r="C2" t="inlineStr"><is><t>15% faster</t></is></c><c r="D2" t="inlineStr"><is><t>Baseline</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>AI/ML Built-in</t></is></c><c r="B3" t="inlineStr"><is><t>Yes - included</t></is></c><c r="C3" t="inlineStr"><is><t>$150K add-on</t></is></c><c r="D3" t="inlineStr"><is><t>No</t></is></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>IoT Integration</t></is></c><c r="B4" t="inlineStr"><is><t>Azure native</t></is></c><c r="C4" t="inlineStr"><is><t>AWS only</t></is></c><c r="D4" t="inlineStr"><is><t>Custom (expensive)</t></is></c></row>
    <row r="5"><c r="A5" t="inlineStr"><is><t>Deployment Time</t></is></c><c r="B5" t="inlineStr"><is><t>8 weeks (BEST)</t></is></c><c r="C5" t="inlineStr"><is><t>6 months</t></is></c><c r="D5" t="inlineStr"><is><t>4 months</t></is></c></row>
    <row r="6"><c r="A6" t="inlineStr"><is><t>Modular Design</t></is></c><c r="B6" t="inlineStr"><is><t>Yes</t></is></c><c r="C6" t="inlineStr"><is><t>No</t></is></c><c r="D6" t="inlineStr"><is><t>Partial</t></is></c></row>
    <row r="7"><c r="A7" t="inlineStr"><is><t>Starting Price</t></is></c><c r="B7" t="inlineStr"><is><t>$800K</t></is></c><c r="C7" t="inlineStr"><is><t>$950K</t></is></c><c r="D7" t="inlineStr"><is><t>$700K</t></is></c></row>
    <row r="8"><c r="A8" t="inlineStr"><is><t>Total Cost (w/ AI)</t></is></c><c r="B8" t="inlineStr"><is><t>$800K (included)</t></is></c><c r="C8" t="inlineStr"><is><t>$1,100K</t></is></c><c r="D8" t="inlineStr"><is><t>N/A - no AI</t></is></c></row>
    <row r="9"><c r="A9" t="inlineStr"><is><t>EU Presence</t></is></c><c r="B9" t="inlineStr"><is><t>Growing</t></is></c><c r="C9" t="inlineStr"><is><t>Strong</t></is></c><c r="D9" t="inlineStr"><is><t>Moderate</t></is></c></row>
    <row r="10"><c r="A10" t="inlineStr"><is><t>SAP Integration</t></is></c><c r="B10" t="inlineStr"><is><t>Good (closing gap)</t></is></c><c r="C10" t="inlineStr"><is><t>Deep</t></is></c><c r="D10" t="inlineStr"><is><t>Basic</t></is></c></row>
    <row r="11"><c r="A11" t="inlineStr"><is><t>Predictive Maintenance</t></is></c><c r="B11" t="inlineStr"><is><t>Built-in ML models</t></is></c><c r="C11" t="inlineStr"><is><t>Add-on only</t></is></c><c r="D11" t="inlineStr"><is><t>Not available</t></is></c></row>
  </sheetData>
</worksheet>
"@

$compWbXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Competitive Analysis" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"@

$compWbRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"@

$compParts = @(
    @{ Uri = "/xl/worksheets/sheet1.xml"; ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"; Content = $compSheet; RelType = $null },
    @{ Uri = "/xl/workbook.xml"; ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"; Content = $compWbXml; RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" },
    @{ Uri = "/xl/_rels/workbook.xml.rels"; ContentType = "application/vnd.openxmlformats-package.relationships+xml"; Content = $compWbRelsXml; RelType = $null }
)

New-OpenXmlPackage -FilePath (Join-Path $outDir "ProLine X Competitive Analysis.xlsx") -Parts $compParts
Write-Host "  [OK] ProLine X Competitive Analysis.xlsx" -ForegroundColor Green

# ════════════════════════════════════════════════════════════════════
# 5. Adatum Corp Meeting Notes.docx (from Jan 15 txt)
# ════════════════════════════════════════════════════════════════════

Write-Host "Generating Adatum Corp Meeting Notes.docx..." -ForegroundColor Cyan

$meetingNotesBody = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>Adatum Corp - Meeting Notes</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:color w:val="666666"/></w:rPr><w:t>January 15, 2026 | 10:00 AM - 11:30 AM | Teams</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:color w:val="666666"/></w:rPr><w:t>Attendees: James Bowen, Spencer Low, David So, Lisa Chen (Adatum VP Procurement), Marcus Webb (Adatum Dir. Operations)</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Key Takeaways</w:t></w:r></w:p>
    <w:p><w:r><w:t>1. Adatum committed to a pilot if we can demonstrate 15% throughput improvement - our ProLine S data shows 18.3%</w:t></w:r></w:p>
    <w:p><w:r><w:t>2. Lisa Chen (VP Procurement) is the decision maker. Marcus Webb influences technical evaluation.</w:t></w:r></w:p>
    <w:p><w:r><w:t>3. Budget cycle: they need to commit by end of Q2 (June 30) or it rolls to next fiscal year</w:t></w:r></w:p>
    <w:p><w:r><w:t>4. They asked about our partnership with Azure IoT - we should highlight that in our next meeting</w:t></w:r></w:p>
    <w:p><w:r><w:t>5. Adatum is also evaluating Fabrikam X200 and Northwind NX-Pro</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Action Items</w:t></w:r></w:p>
    <w:p><w:r><w:t>- [Spencer] Send ProLine S performance data to Marcus Webb</w:t></w:r></w:p>
    <w:p><w:r><w:t>- [James] Prepare briefing deck for next meeting</w:t></w:r></w:p>
    <w:p><w:r><w:t>- [David] Schedule follow-up meeting for April</w:t></w:r></w:p>
    <w:p><w:r><w:t>- [Renee] Prepare competitive analysis (Fabrikam, Northwind)</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Discussion Notes</w:t></w:r></w:p>
    <w:p><w:r><w:t>Lisa opened by saying Adatum is in the middle of a $50M smart manufacturing investment across North American plants. They need a platform that can scale across 6 plants and integrate with their existing SAP ERP.</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:t>Marcus was impressed with the ProLine S demo but had concerns about integration complexity. We walked through our 8-week deployment model and he seemed reassured. He specifically liked the Azure IoT native integration since Adatum is already on Azure.</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:t>Lisa mentioned they're talking to Fabrikam and Northwind. Spencer positioned our speed-to-value advantage well. She agreed that a 6-month deployment (Fabrikam) would be a problem for their timeline.</w:t></w:r></w:p>
    <w:p/>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Next Steps</w:t></w:r></w:p>
    <w:p><w:r><w:t>Schedule a follow-up meeting for late April with a focused demo of ProLine X (pre-GA build). Bring competitive positioning materials and detailed ROI analysis based on their specific plant metrics.</w:t></w:r></w:p>
    <w:p/>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>
"@

$meetingParts = @(
    @{ Uri = "/word/document.xml"; ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"; Content = $meetingNotesBody; RelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" },
    @{ Uri = "/word/_rels/document.xml.rels"; ContentType = "application/vnd.openxmlformats-package.relationships+xml"; Content = $docRelsXml; RelType = $null }
)

New-OpenXmlPackage -FilePath (Join-Path $outDir "Adatum Corp Meeting Notes - Jan 15.docx") -Parts $meetingParts
Write-Host "  [OK] Adatum Corp Meeting Notes - Jan 15.docx" -ForegroundColor Green

# ════════════════════════════════════════════════════════════════════
Write-Host ""
Write-Host "All Office files generated successfully!" -ForegroundColor Green
Write-Host "Files are in: $outDir" -ForegroundColor White
Write-Host ""
Write-Host "New files:" -ForegroundColor Cyan
Get-ChildItem $outDir -Filter "*.docx" | ForEach-Object { Write-Host "  $_" }
Get-ChildItem $outDir -Filter "*.xlsx" | ForEach-Object { Write-Host "  $_" }
Get-ChildItem $outDir -Filter "*.pptx" | ForEach-Object { Write-Host "  $_" }
