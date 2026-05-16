Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path -Path $repoRoot -ChildPath 'Expecto-Comparo.ps1'
$docxfixRoot = if ($env:DOCXFIX_ROOT) { $env:DOCXFIX_ROOT } else { 'C:\Users\David\Code\docxfix' }

$env:EXPECTO_COMPARO_TEST_MODE = '1'
. $scriptPath
Remove-Item Env:\EXPECTO_COMPARO_TEST_MODE

function Assert-Equal {
    param(
        [Parameter(Mandatory)]$Expected,
        [Parameter(Mandatory)]$Actual,
        [Parameter(Mandatory)][string]$Message
    )

    if ($Expected -ne $Actual) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

function Assert-True {
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function New-TestRoot {
    $parent = Join-Path -Path $repoRoot -ChildPath '.test-tmp'
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        [void](New-Item -Path $parent -ItemType Directory)
    }

    $root = Join-Path -Path $parent -ChildPath ("expecto-comparo-tests-{0}" -f [guid]::NewGuid())
    [void](New-Item -Path $root -ItemType Directory)
    return $root
}

function Remove-TestRoot {
    param([Parameter(Mandatory)][string]$Path)

    $resolved = [System.IO.Path]::GetFullPath($Path)
    $allowedRoot = [System.IO.Path]::GetFullPath((Join-Path -Path $repoRoot -ChildPath '.test-tmp'))
    if (-not $resolved.StartsWith($allowedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove test path outside ${allowedRoot}: $resolved"
    }

    if (Test-Path -LiteralPath $resolved) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

function Initialize-FixtureGenerator {
    param([Parameter(Mandatory)][string]$Root)

    if (-not (Test-Path -LiteralPath $docxfixRoot -PathType Container)) {
        throw "docxfix root not found. Set DOCXFIX_ROOT or place it at $docxfixRoot."
    }

    $python = Join-Path -Path $docxfixRoot -ChildPath '.venv\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        $python = 'python'
    }

    $helperPath = Join-Path -Path $Root -ChildPath 'new_docx_fixture.py'
    @'
import sys
from pathlib import Path

docxfix_root = Path(sys.argv[1])
output_path = Path(sys.argv[2])
title = sys.argv[3]
body = sys.argv[4]

sys.path.insert(0, str(docxfix_root / "src"))

from docxfix.generator import DocumentGenerator
from docxfix.spec import DocumentSpec

spec = DocumentSpec(title=title, author="Expecto Comparo Tests", seed=123)
spec.add_paragraph(body)
DocumentGenerator(spec).generate(output_path)
'@ | Set-Content -LiteralPath $helperPath -Encoding UTF8

    return [pscustomobject]@{
        Python = $python
        HelperPath = $helperPath
    }
}

function New-DocxFixture {
    param(
        [Parameter(Mandatory)]$FixtureGenerator,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Body
    )

    & $FixtureGenerator.Python $FixtureGenerator.HelperPath $docxfixRoot $Path $Title $Body
    if ($LASTEXITCODE -ne 0) {
        throw "docxfix fixture generation failed for $Path"
    }

    Assert-True -Condition (Test-Path -LiteralPath $Path -PathType Leaf) -Message "Fixture was not created: $Path"
}

function Reset-UiState {
    param(
        [Parameter(Mandatory)][string]$PreviousFolder,
        [Parameter(Mandatory)][string]$CurrentFolder
    )

    $script:State.PreviousFolder = $PreviousFolder
    $script:State.CurrentFolder = $CurrentFolder
    $script:State.OutputFolder = ''
    $script:State.PreviousFiles = @()
    $script:State.CurrentFiles = @()
    $script:State.PreviousUnmatchedFiles = @()
    $script:State.CurrentUnmatchedFiles = @()
    $script:State.SuggestedPairs.Clear()

    $gridPairs.Rows.Clear()
    $lstPreviousUnmatched.Items.Clear()
    $lstCurrentUnmatched.Items.Clear()
    $lblCounts.Text = ''
}

function Assert-ScanState {
    param(
        [Parameter(Mandatory)][int]$PreviousCount,
        [Parameter(Mandatory)][int]$CurrentCount,
        [Parameter(Mandatory)][int]$PairCount,
        [Parameter(Mandatory)][int]$PreviousUnmatchedCount,
        [Parameter(Mandatory)][int]$CurrentUnmatchedCount
    )

    Assert-True -Condition ($script:State.PreviousFiles -is [array]) -Message 'PreviousFiles must always be array-shaped.'
    Assert-True -Condition ($script:State.CurrentFiles -is [array]) -Message 'CurrentFiles must always be array-shaped.'
    Assert-Equal -Expected $PreviousCount -Actual $script:State.PreviousFiles.Count -Message 'Unexpected previous file count.'
    Assert-Equal -Expected $CurrentCount -Actual $script:State.CurrentFiles.Count -Message 'Unexpected current file count.'
    Assert-Equal -Expected $PairCount -Actual $script:State.SuggestedPairs.Count -Message 'Unexpected suggested pair count.'
    Assert-Equal -Expected $PairCount -Actual $gridPairs.Rows.Count -Message 'Unexpected grid row count.'
    Assert-Equal -Expected $PreviousUnmatchedCount -Actual $lstPreviousUnmatched.Items.Count -Message 'Unexpected previous unmatched count.'
    Assert-Equal -Expected $CurrentUnmatchedCount -Actual $lstCurrentUnmatched.Items.Count -Message 'Unexpected current unmatched count.'
    Assert-Equal -Expected "Previous: $PreviousCount   Current: $CurrentCount   Suggested pairs: $PairCount" -Actual $lblCounts.Text -Message 'Unexpected count label.'
}

$testRoot = New-TestRoot

try {
    $fixtureGenerator = Initialize-FixtureGenerator -Root $testRoot

    $emptyPrevious = Join-Path -Path $testRoot -ChildPath 'empty-previous'
    $emptyCurrent = Join-Path -Path $testRoot -ChildPath 'empty-current'
    [void](New-Item -Path $emptyPrevious -ItemType Directory)
    [void](New-Item -Path $emptyCurrent -ItemType Directory)
    Reset-UiState -PreviousFolder $emptyPrevious -CurrentFolder $emptyCurrent
    Scan-Folders
    Assert-ScanState -PreviousCount 0 -CurrentCount 0 -PairCount 0 -PreviousUnmatchedCount 0 -CurrentUnmatchedCount 0

    $singlePrevious = Join-Path -Path $testRoot -ChildPath 'single-previous'
    $singleCurrent = Join-Path -Path $testRoot -ChildPath 'single-current'
    [void](New-Item -Path $singlePrevious -ItemType Directory)
    [void](New-Item -Path $singleCurrent -ItemType Directory)
    New-DocxFixture -FixtureGenerator $fixtureGenerator -Path (Join-Path -Path $singlePrevious -ChildPath 'Agreement.docx') -Title 'Previous Agreement' -Body 'Previous body'
    New-DocxFixture -FixtureGenerator $fixtureGenerator -Path (Join-Path -Path $singleCurrent -ChildPath 'Agreement.docx') -Title 'Current Agreement' -Body 'Current body'
    Reset-UiState -PreviousFolder $singlePrevious -CurrentFolder $singleCurrent
    Scan-Folders
    Assert-ScanState -PreviousCount 1 -CurrentCount 1 -PairCount 1 -PreviousUnmatchedCount 0 -CurrentUnmatchedCount 0
    Assert-Equal -Expected 'Exact filename' -Actual $script:State.SuggestedPairs[0].MatchStatus -Message 'Single-file scan should exact-match identical filenames.'

    $multiPrevious = Join-Path -Path $testRoot -ChildPath 'multi-previous'
    $multiCurrent = Join-Path -Path $testRoot -ChildPath 'multi-current'
    [void](New-Item -Path $multiPrevious -ItemType Directory)
    [void](New-Item -Path $multiCurrent -ItemType Directory)
    foreach ($name in @('Alpha.docx', 'Beta.docx')) {
        New-DocxFixture -FixtureGenerator $fixtureGenerator -Path (Join-Path -Path $multiPrevious -ChildPath $name) -Title "Previous $name" -Body "Previous $name"
        New-DocxFixture -FixtureGenerator $fixtureGenerator -Path (Join-Path -Path $multiCurrent -ChildPath $name) -Title "Current $name" -Body "Current $name"
    }
    Reset-UiState -PreviousFolder $multiPrevious -CurrentFolder $multiCurrent
    Scan-Folders
    Assert-ScanState -PreviousCount 2 -CurrentCount 2 -PairCount 2 -PreviousUnmatchedCount 0 -CurrentUnmatchedCount 0

    $unmatchedPrevious = Join-Path -Path $testRoot -ChildPath 'unmatched-previous'
    $unmatchedCurrent = Join-Path -Path $testRoot -ChildPath 'unmatched-current'
    [void](New-Item -Path $unmatchedPrevious -ItemType Directory)
    [void](New-Item -Path $unmatchedCurrent -ItemType Directory)
    [void](New-Item -Path (Join-Path -Path $unmatchedPrevious -ChildPath 'PrevOnly.docx') -ItemType File)
    [void](New-Item -Path (Join-Path -Path $unmatchedCurrent -ChildPath 'CurrOnly.docx') -ItemType File)
    Reset-UiState -PreviousFolder $unmatchedPrevious -CurrentFolder $unmatchedCurrent
    Scan-Folders
    Assert-ScanState -PreviousCount 1 -CurrentCount 1 -PairCount 0 -PreviousUnmatchedCount 1 -CurrentUnmatchedCount 1
    Assert-Equal -Expected 'PrevOnly.docx' -Actual ([string]$lstPreviousUnmatched.Items[0]) -Message 'Previous unmatched list should display only the filename.'
    Assert-Equal -Expected 'CurrOnly.docx' -Actual ([string]$lstCurrentUnmatched.Items[0]) -Message 'Current unmatched list should display only the filename.'
    Assert-Equal -Expected 'PrevOnly.docx' -Actual $script:State.PreviousUnmatchedFiles[0].Name -Message 'Previous unmatched state should retain the file info object.'
    Assert-Equal -Expected 'CurrOnly.docx' -Actual $script:State.CurrentUnmatchedFiles[0].Name -Message 'Current unmatched state should retain the file info object.'

    $rescanPrevious = Join-Path -Path $testRoot -ChildPath 'rescan-previous'
    $rescanCurrent = Join-Path -Path $testRoot -ChildPath 'rescan-current'
    [void](New-Item -Path $rescanPrevious -ItemType Directory)
    [void](New-Item -Path $rescanCurrent -ItemType Directory)
    New-DocxFixture -FixtureGenerator $fixtureGenerator -Path (Join-Path -Path $rescanPrevious -ChildPath 'Solo.docx') -Title 'Previous Solo' -Body 'Previous solo'
    New-DocxFixture -FixtureGenerator $fixtureGenerator -Path (Join-Path -Path $rescanCurrent -ChildPath 'Solo.docx') -Title 'Current Solo' -Body 'Current solo'
    Reset-UiState -PreviousFolder $rescanPrevious -CurrentFolder $rescanCurrent
    Scan-Folders
    Assert-ScanState -PreviousCount 1 -CurrentCount 1 -PairCount 1 -PreviousUnmatchedCount 0 -CurrentUnmatchedCount 0
    New-DocxFixture -FixtureGenerator $fixtureGenerator -Path (Join-Path -Path $rescanPrevious -ChildPath 'Extra.docx') -Title 'Previous Extra' -Body 'Previous extra'
    New-DocxFixture -FixtureGenerator $fixtureGenerator -Path (Join-Path -Path $rescanCurrent -ChildPath 'Extra.docx') -Title 'Current Extra' -Body 'Current extra'
    Scan-Folders
    Assert-ScanState -PreviousCount 2 -CurrentCount 2 -PairCount 2 -PreviousUnmatchedCount 0 -CurrentUnmatchedCount 0

    Assert-Equal -Expected 3 -Actual (Get-LevenshteinDistance -A 'kitten' -B 'sitting') -Message 'Levenshtein distance regression failed.'

    Write-Host 'All Expecto Comparo tests passed.'
}
finally {
    Remove-TestRoot -Path $testRoot
}
