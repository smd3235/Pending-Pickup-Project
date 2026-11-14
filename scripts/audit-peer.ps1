# audit-peer.ps1
# Compares the shared, portable parts of this project against a sibling project and
# reports what each side has that the other does not: files, VBA procedures,
# PowerShell/Python functions, and Markdown sections.
#
# It never changes either project. It only reads.
#
#   .\scripts\audit-peer.ps1
#   .\scripts\audit-peer.ps1 -Sections scripts,commands
#   .\scripts\audit-peer.ps1 -PeerPath "C:\path\to\Other Project" -Report documentation\cross-audit.md
#
# The peer is found automatically: a sibling directory that holds a root *.bas and a
# scripts\ folder. Pass -PeerPath when the projects are not siblings.
#
# Exit codes: 0 = audited, 2 = peer not found, 3 = nothing auditable on either side.

[CmdletBinding()]
param(
    [string]   $PeerPath,
    [string[]] $Sections = @("scripts", "documentation", "commands", "vba", "rules"),
    [string]   $Report,
    [int]      $SymbolLimit = 20,
    [switch]   $Quiet
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $PSScriptRoot

# ---------------------------------------------------------------- section specs
# Each section says where to look and which files count as portable.
$SectionSpecs = @{
    "scripts"       = @{ Dir = "scripts";          Include = @("*.ps1", "*.py", "*.cmd", "*.bat"); Recursive = $false }
    "documentation" = @{ Dir = "documentation";    Include = @("*.md", "*.ps1", "*.py");               Recursive = $false }
    "commands"      = @{ Dir = ".kilo\command";    Include = @("*.md");                                Recursive = $false }
    "vba"           = @{ Dir = ".";                Include = @("*.bas");                               Recursive = $false }
    "rules"         = @{ Dir = "documentation";    Include = @("AGENTS.md");                          Recursive = $false }
}

function Resolve-PeerRoot {
    param([string]$Requested)
    if ($Requested) {
        if (-not (Test-Path -LiteralPath $Requested)) {
            Write-Host "Peer not found: $Requested" -ForegroundColor Red
            exit 2
        }
        return (Resolve-Path -LiteralPath $Requested).Path
    }
    $parent = Split-Path -Parent $ProjectRoot
    $candidates = @()
    foreach ($dir in (Get-ChildItem -LiteralPath $parent -Directory -ErrorAction SilentlyContinue)) {
        if ($dir.FullName -eq $ProjectRoot) { continue }
        $hasBas     = @(Get-ChildItem -LiteralPath $dir.FullName -Filter "*.bas" -File -ErrorAction SilentlyContinue).Count -gt 0
        $hasScripts = Test-Path -LiteralPath (Join-Path $dir.FullName "scripts")
        if ($hasBas -and $hasScripts) { $candidates += $dir.FullName }
    }
    if ($candidates.Count -eq 1) { return $candidates[0] }
    if ($candidates.Count -eq 0) {
        Write-Host "No sibling project with a root *.bas and a scripts\ folder was found next to $ProjectRoot" -ForegroundColor Red
        Write-Host "Pass -PeerPath <path to the other project>." -ForegroundColor Yellow
        exit 2
    }
    Write-Host "Several candidate peers found; pick one with -PeerPath:" -ForegroundColor Yellow
    foreach ($c in $candidates) { Write-Host "  $c" }
    exit 2
}

function Get-SectionFiles {
    param([string]$Root, [string[]]$Wanted)

    $files = @()
    foreach ($section in $Wanted) {
        $spec = $SectionSpecs[$section]
        if (-not $spec) {
            Write-Host "Unknown section '$section' (known: $($SectionSpecs.Keys -join ', '))" -ForegroundColor Red
            exit 2
        }
        $dir = Join-Path $Root $spec.Dir
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        $gci = @{ LiteralPath = $dir; File = $true; ErrorAction = "SilentlyContinue" }
        if ($spec.Recursive) { $gci["Recurse"] = $true }
        foreach ($f in (Get-ChildItem @gci)) {
            foreach ($pattern in $spec.Include) {
                if ($f.Name -like $pattern) {
                    $rel = $f.FullName.Substring($Root.Length).TrimStart("\")
                    $files += [pscustomobject]@{
                        Section = $section
                        Rel     = $rel
                        Name    = $f.Name
                        Kind    = $f.Extension.TrimStart(".").ToLower()
                        Path    = $f.FullName
                        Size    = $f.Length
                        Hash    = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash
                        Lines   = 0
                        Symbols = @()
                    }
                    break
                }
            }
        }
    }
    return $files
}

function Get-SymbolList {
    param([string]$Path, [string]$Kind)

    $rx = $null
    if ($Kind -eq "ps1") { $rx = '^\s*function\s+([A-Za-z_][\w-]*)' }
    elseif ($Kind -eq "py") { $rx = '^\s*(?:def|class)\s+(\w+)' }
    elseif ($Kind -eq "bas") { $rx = '^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)' }
    elseif ($Kind -eq "md") { $rx = '^#{1,4}\s+(.+?)\s*#*\s*$' }

    if (-not $rx) { return @() }
    $lines  = [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)
    $found  = @()
    $seen   = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($line in $lines) {
        $m = [regex]::Match($line, $rx)
        if ($m.Success) {
            $name = $m.Groups[1].Value.Trim()
            if ($name.Length -gt 0 -and $seen.Add($name.ToLowerInvariant())) { $found += $name }
        }
    }
    return $found
}

function Get-MatchFor {
    param([object]$File, [object[]]$Pool)

    foreach ($other in $Pool) {
        if ($other.Rel -eq $File.Rel) { return $other }
    }
    foreach ($other in $Pool) {
        if ($other.Name -eq $File.Name) { return $other }
    }
    return $null
}

function Format-List {
    param([string[]]$Items, [int]$Limit, [string]$Prefix = "      ")

    if ($Items.Count -eq 0) { return @() }
    $shown = @()
    for ($i = 0; $i -lt $Items.Count -and $i -lt $Limit; $i++) { $shown += "$Prefix$($Items[$i])" }
    if ($Items.Count -gt $Limit) { $shown += "$Prefix... (+$($Items.Count - $Limit) more)" }
    return $shown
}

# ------------------------------------------------------------------- audit body
$peerRoot = Resolve-PeerRoot -Requested $PeerPath
$peerLeaf = Split-Path -Leaf $peerRoot
$thisLeaf = Split-Path -Leaf $ProjectRoot

if (-not $Quiet) {
    Write-Host ""
    Write-Host "CROSS AUDIT  $thisLeaf  <->  $peerLeaf" -ForegroundColor Cyan
    Write-Host ("  this : {0}" -f $ProjectRoot)
    Write-Host ("  peer : {0}" -f $peerRoot)
}

$hereFiles = @(Get-SectionFiles -Root $ProjectRoot -Wanted $Sections)
$peerFiles = @(Get-SectionFiles -Root $peerRoot    -Wanted $Sections)

if ($hereFiles.Count -eq 0 -and $peerFiles.Count -eq 0) {
    Write-Host "Nothing auditable in either project for sections: $($Sections -join ', ')" -ForegroundColor Red
    exit 3
}

# Fill in size/line/symbol data for both sides (hashes are already in the objects).
foreach ($group in @(@{ List = $hereFiles }, @{ List = $peerFiles })) {
    foreach ($f in $group.List) {
        $f.Lines   = @([System.IO.File]::ReadAllLines($f.Path, [System.Text.Encoding]::UTF8)).Count
        $f.Symbols = @(Get-SymbolList -Path $f.Path -Kind $f.Kind)
    }
}

$rows          = @()
$portCandidates = @()
$portBack       = @()

foreach ($peer in $peerFiles) {
    $match     = Get-MatchFor -File $peer -Pool $hereFiles
    $addedPeer = @()
    $addedHere = @()
    if (-not $match) {
        $status = "peer only"
        $detail = "$($peer.Lines) lines"
        $portCandidates += $peer
    }
    else {
        $addedHere = @()
        $addedPeer = @()
        foreach ($s in $peer.Symbols)  { if ($match.Symbols -notcontains $s) { $addedPeer += $s } }
        foreach ($s in $match.Symbols) { if ($peer.Symbols  -notcontains $s) { $addedHere += $s } }

        if ($peer.Hash -eq $match.Hash) {
            $status = "identical"
            $detail = "$($peer.Lines) lines"
        }
        else {
            $status = "changed"
            $detail = "peer $($peer.Lines) vs here $($match.Lines) lines"
            if ($addedPeer.Count -gt 0) { $detail += " | peer adds: " + (($addedPeer | Select-Object -First 3) -join ", ") }
            if ($addedHere.Count -gt 0) { $detail += " | here adds: " + (($addedHere | Select-Object -First 3) -join ", ") }
        }
    }
    $rows += [pscustomobject]@{
        Section = $peer.Section
        File    = $peer.Rel
        Status  = $status
        Detail  = $detail
        Added   = $addedPeer
        Match   = $match
    }
}

foreach ($here in $hereFiles) {
    $match = Get-MatchFor -File $here -Pool $peerFiles
    if (-not $match) {
        $portBack += $here
        $rows += [pscustomobject]@{
            Section = $here.Section
            File    = $here.Rel
            Status  = "here only"
            Detail  = "$($here.Lines) lines"
            Added   = @()
            Match   = $null
        }
    }
}

$portBackRows = @($rows | Where-Object { $_.Status -eq "here only" })

# ------------------------------------------------------------------ console out
$md = New-Object 'System.Collections.Generic.List[string]'
$md.Add("# Cross audit: $thisLeaf <-> $peerLeaf")
$md.Add("")
$md.Add("Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$md.Add("")
$md.Add("- this: ``$ProjectRoot``")
$md.Add("- peer: ``$peerRoot``")

foreach ($section in $Sections) {
    $sectionRows = @($rows | Where-Object { $_.Section -eq $section })
    if ($sectionRows.Count -eq 0) { continue }

    if (-not $Quiet) {
        Write-Host ""
        Write-Host ("== {0} ==" -f $section.ToUpper()) -ForegroundColor Yellow
    }
    $md.Add("## $section")
    $md.Add("")
    $md.Add("| file | status | detail |")
    $md.Add("| --- | --- | --- |")

    foreach ($r in $sectionRows) {
        $colour = "Gray"
        if ($r.Status -eq "peer only")  { $colour = "Green" }
        if ($r.Status -eq "here only")  { $colour = "Magenta" }
        if ($r.Status -eq "changed")    { $colour = "Yellow" }
        if (-not $Quiet) {
            Write-Host ("  {0,-46} {1,-11} {2}" -f $r.File, $r.Status, $r.Detail) -ForegroundColor $colour
            if ($r.Added.Count -gt 3) { foreach ($line in (Format-List -Items $r.Added -Limit $SymbolLimit)) {
                if ($line.Trim().Length -gt 0) { Write-Host $line -ForegroundColor DarkGray }
            } }
        }
        $md.Add("| ``$($r.File)`` | $($r.Status) | $($r.Detail) |")
        foreach ($a in @($r.Added | Select-Object -First $SymbolLimit)) {
            $md.Add("| | | + $a |")
        }
    }
}

# -------------------------------------------------- VBA procedure cross-diff
# The two modules have different names, so a file-level comparison says nothing.
# Procedure names are the comparable unit: a helper that exists in both is already
# shared design; one that exists only in the peer is a porting candidate.
$vbaHere = @{}
foreach ($f in @($hereFiles | Where-Object { $_.Section -eq "vba" })) {
    foreach ($s in $f.Symbols) { if (-not $vbaHere.ContainsKey($s.ToLowerInvariant())) { $vbaHere[$s.ToLowerInvariant()] = $s } }
}
$vbaPeer = @{}
foreach ($f in @($peerFiles | Where-Object { $_.Section -eq "vba" })) {
    foreach ($s in $f.Symbols) { if (-not $vbaPeer.ContainsKey($s.ToLowerInvariant())) { $vbaPeer[$s.ToLowerInvariant()] = $s } }
}
$vbaPeerOnlyNames = @()
$vbaSharedNames   = @()
if ($vbaPeer.Count -gt 0 -or $vbaHere.Count -gt 0) {
    foreach ($key in $vbaPeer.Keys) {
        if ($vbaHere.ContainsKey($key)) { $vbaSharedNames += $vbaPeer[$key] }
        else                            { $vbaPeerOnlyNames += $vbaPeer[$key] }
    }
    $vbaPeerOnlyNames = @($vbaPeerOnlyNames | Sort-Object)
    $vbaSharedNames   = @($vbaSharedNames | Sort-Object)

    if (-not $Quiet) {
        Write-Host ""
        Write-Host ("== VBA PROCEDURES ==  shared={0}  peer-only={1}  here={2}" -f $vbaSharedNames.Count, $vbaPeerOnlyNames.Count, $vbaHere.Count) -ForegroundColor Yellow
        foreach ($line in (Format-List -Items $vbaPeerOnlyNames -Limit $SymbolLimit -Prefix "  peer only  ")) { Write-Host $line -ForegroundColor Green }
        foreach ($line in (Format-List -Items $vbaSharedNames   -Limit 10 -Prefix "  shared     ")) { Write-Host $line -ForegroundColor DarkGray }
    }
    $md.Add("## VBA procedures")
    $md.Add("")
    $md.Add("- shared: $($vbaSharedNames.Count)")
    $md.Add("- peer only: $($vbaPeerOnlyNames.Count)")
    $md.Add("- this project: $($vbaHere.Count)")
    $md.Add("")
    $md.Add("### Procedures only the peer has (port candidates)")
    $md.Add("")
    foreach ($s in $vbaPeerOnlyNames) { $md.Add("- $s") }
}

# ------------------------------------------------------------- port shortlists
if ($portCandidates.Count -gt 0) {
    $script = @($portCandidates | Where-Object { $_.Section -eq "scripts" -or $_.Section -eq "documentation" })
    if ($script.Count -gt 0) {
        if (-not $Quiet) {
            Write-Host ""
            Write-Host "== WORTH A LOOK (peer has, this project does not) ==" -ForegroundColor Green
        }
        $md.Add("## Port candidates (peer -> this)")
        $md.Add("")
        foreach ($f in $script) {
            $line = "  {0,-46} {1,6} lines" -f $f.Rel, $f.Lines
            if (-not $Quiet) { Write-Host $line -ForegroundColor Green }
            $md.Add("- ``$($f.Rel)`` - $($f.Lines) lines, $([math]::Round($f.Size / 1KB, 1)) KB")
            foreach ($s in @($f.Symbols | Select-Object -First $SymbolLimit)) {
                if (-not $Quiet) { Write-Host "      $s" -ForegroundColor DarkGray }
                $md.Add("  - $s")
            }
        }
    }

    $cmd = @($portCandidates | Where-Object { $_.Section -eq "commands" })
    if ($cmd.Count -gt 0) {
        if (-not $Quiet) { Write-Host "" ; Write-Host "Commands only the peer has: $($cmd.Name -join ', ')" -ForegroundColor Green }
        $md.Add("### Commands only the peer has")
        $md.Add("")
        foreach ($f in $cmd) { $md.Add("- ``$($f.Rel)``") }
    }

    $vba = @($portCandidates | Where-Object { $_.Section -eq "vba" })
    if ($vba.Count -gt 0) {
        if (-not $Quiet) { Write-Host "VBA modules only the peer has: $($vba.Name -join ', ')" -ForegroundColor Green }
        $md.Add("### VBA modules only the peer has")
        $md.Add("")
        foreach ($f in $vba) { $md.Add("- ``$($f.Rel)`` ($($f.Lines) lines, $($f.Symbols.Count) procedures)") }
    }
}

if ($portBackRows.Count -gt 0) {
    if (-not $Quiet) {
        Write-Host ""
        Write-Host "== PEER MAY WANT (this has, peer does not) ==" -ForegroundColor Magenta
        foreach ($r in $portBackRows) { Write-Host ("  {0,-46} {1}" -f $r.File, $r.Detail) -ForegroundColor Magenta }
    }
    $md.Add("## Port candidates (this -> peer)")
    $md.Add("")
    foreach ($r in $portBackRows) { $md.Add("- ``$($r.File)`` - $($r.Detail)") }
}

# --------------------------------------------------------------------- summary
$identical = @($rows | Where-Object { $_.Status -eq "identical" }).Count
$changed   = @($rows | Where-Object { $_.Status -eq "changed" }).Count
$peerOnly  = @($rows | Where-Object { $_.Status -eq "peer only" }).Count
$hereOnly  = @($rows | Where-Object { $_.Status -eq "here only" }).Count

if (-not $Quiet) {
    Write-Host ""
    Write-Host ("SUMMARY  identical={0}  changed={1}  peer-only={2}  here-only={3}" -f $identical, $changed, $peerOnly, $hereOnly) -ForegroundColor Cyan
}

$md.Add("")
$md.Add("## Summary")
$md.Add("")
$md.Add("- identical: $identical")
$md.Add("- changed: $changed")
$md.Add("- peer only: $peerOnly")
$md.Add("- here only: $hereOnly")
$md.Add("")
$md.Add("This is a read-only report. Porting is a manual decision: copy a file, adapt its")
$md.Add("data source and paths, then run this project's compile gate and report build.")
$md.Add("")

if ($Report) {
    $reportPath = $Report
    if (-not [System.IO.Path]::IsPathRooted($reportPath)) { $reportPath = Join-Path $ProjectRoot $reportPath }
    $dir = Split-Path -Parent $reportPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $md -join "`r`n" | Set-Content -LiteralPath $reportPath -Encoding UTF8
    if (-not $Quiet) { Write-Host ""
        Write-Host "Report written: $reportPath" -ForegroundColor Cyan }
}

exit 0