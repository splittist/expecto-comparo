Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:MaxSuggestedOutputNameLength = 180 # conservative name cap to keep full output path under common Windows MAX_PATH deployments
$script:FuzzyMatchThreshold = 0.70
$script:HigherConfidenceFuzzyThreshold = 0.85
$script:WdCompareDestinationNew = 2 # Word constant: wdCompareDestinationNew
$script:WdGranularityWordLevel = 1 # Word constant: wdGranularityWordLevel
$script:WdFormatXMLDocument = 12 # Word constant: wdFormatXMLDocument (.docx)
$script:WdAlertsNone = 0 # Word constant: wdAlertsNone
$script:WdDoNotSaveChanges = 0 # Word constant: wdDoNotSaveChanges

$script:State = [ordered]@{
    PreviousFolder = ''
    CurrentFolder = ''
    OutputFolder = ''
    PreviousFiles = @()
    CurrentFiles = @()
    PreviousUnmatchedFiles = @()
    CurrentUnmatchedFiles = @()
    SuggestedPairs = [System.Collections.ArrayList]::new()
    LogLines = [System.Collections.Generic.List[string]]::new()
    LogPath = ''
}

function Get-DocxFiles {
    param([Parameter(Mandatory)][string]$Folder)

    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $Folder -Filter '*.docx' -File -ErrorAction SilentlyContinue | Sort-Object Name)
}

function Get-FileListDisplayName {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)

    return $File.Name
}

function Get-NormalizedBaseName {
    param([Parameter(Mandatory)][string]$Name)

    $base = [System.IO.Path]::GetFileNameWithoutExtension($Name).ToLowerInvariant()
    $base = $base -replace '(?i)\b(v|ver|version)[\s._-]*\d+\b', ' '
    $base = $base -replace '(?i)\b(rev|revision)[\s._-]*\d+\b', ' '
    $base = $base -replace '(?i)\b(19|20)\d{2}[-_ ]?[01]\d[-_ ]?[0-3]\d\b', ' '
    $base = $base -replace '\b\d{8}\b', ' '
    $base = $base -replace '(?i)\b(final|draft|clean|redline|copy)\b', ' '
    $base = $base -replace '[_\-]+', ' '
    $base = $base -replace '\s+', ' '
    return $base.Trim()
}

function Get-LevenshteinDistance {
    param(
        [Parameter(Mandatory)][string]$A,
        [Parameter(Mandatory)][string]$B
    )

    if ($A.Length -eq 0) { return $B.Length }
    if ($B.Length -eq 0) { return $A.Length }

    $width = $B.Length + 1
    $height = $A.Length + 1
    $d = [int[,]]::new($height, $width)

    for ($i = 0; $i -lt $height; $i++) { $d[$i, 0] = $i }
    for ($j = 0; $j -lt $width; $j++) { $d[0, $j] = $j }

    for ($i = 1; $i -lt $height; $i++) {
        for ($j = 1; $j -lt $width; $j++) {
            # Strings are pre-normalized to lowercase before distance scoring.
            $cost = if ($A.Chars($i - 1) -ceq $B.Chars($j - 1)) { 0 } else { 1 }
            $deletion = $d[($i - 1), $j] + 1
            $insertion = $d[$i, ($j - 1)] + 1
            $substitution = $d[($i - 1), ($j - 1)] + $cost
            $d[$i, $j] = [Math]::Min([Math]::Min($deletion, $insertion), $substitution)
        }
    }

    return $d[($height - 1), ($width - 1)]
}

function Get-SimilarityScore {
    param(
        [Parameter(Mandatory)][string]$A,
        [Parameter(Mandatory)][string]$B
    )

    if ([string]::IsNullOrWhiteSpace($A) -or [string]::IsNullOrWhiteSpace($B)) { return 0.0 }

    $distance = Get-LevenshteinDistance -A $A -B $B
    $maxLen = [Math]::Max($A.Length, $B.Length)
    if ($maxLen -eq 0) { return 1.0 }

    return [Math]::Max(0.0, 1.0 - ($distance / $maxLen))
}

function New-SafeOutputFileName {
    param(
        [Parameter(Mandatory)][string]$CurrentName,
        [Parameter(Mandatory)][string]$PreviousName
    )

    $currentBase = [System.IO.Path]::GetFileNameWithoutExtension($CurrentName)
    $previousBase = [System.IO.Path]::GetFileNameWithoutExtension($PreviousName)

    $name = "$currentBase - compared against $previousBase.docx"
    if ($name.Length -gt $script:MaxSuggestedOutputNameLength) {
        $name = "$currentBase - comparison.docx"
    }

    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $name = $name.Replace($invalidChar, '_')
    }

    return $name
}

function Get-UniqueOutputPath {
    param(
        [Parameter(Mandatory)][string]$OutputFolder,
        [Parameter(Mandatory)][string]$DesiredFileName
    )

    $base = [System.IO.Path]::GetFileNameWithoutExtension($DesiredFileName)
    $ext = [System.IO.Path]::GetExtension($DesiredFileName)
    $candidate = Join-Path -Path $OutputFolder -ChildPath $DesiredFileName
    $counter = 1

    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path -Path $OutputFolder -ChildPath ("{0} ({1}){2}" -f $base, $counter, $ext)
        $counter++
    }

    return $candidate
}

function Write-RunLog {
    param(
        [Parameter(Mandatory)][string]$Level,
        [Parameter(Mandatory)][string]$Message,
        [switch]$ToUi
    )

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.ToUpperInvariant(), $Message
    $script:State.LogLines.Add($line) | Out-Null

    if ($ToUi) {
        $txtStatus.AppendText($line + [Environment]::NewLine)
        $txtStatus.SelectionStart = $txtStatus.Text.Length
        $txtStatus.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Invoke-UiAction {
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [Parameter(Mandatory)][string]$ActionName
    )

    try {
        & $Action
    }
    catch {
        $message = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show(
            "$ActionName failed:`n`n$message",
            'Expecto Comparo',
            'OK',
            'Error'
        ) | Out-Null
    }
}

function Refresh-UnmatchedLists {
    $pairedPrevious = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $pairedCurrent = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($item in $script:State.SuggestedPairs) {
        if ($item.PreviousPath) { [void]$pairedPrevious.Add($item.PreviousPath) }
        if ($item.CurrentPath) { [void]$pairedCurrent.Add($item.CurrentPath) }
    }

    $script:State.PreviousUnmatchedFiles = @($script:State.PreviousFiles | Where-Object { -not $pairedPrevious.Contains($_.FullName) })
    $lstPreviousUnmatched.Items.Clear()
    foreach ($file in $script:State.PreviousUnmatchedFiles) {
        [void]$lstPreviousUnmatched.Items.Add((Get-FileListDisplayName -File $file))
    }

    $script:State.CurrentUnmatchedFiles = @($script:State.CurrentFiles | Where-Object { -not $pairedCurrent.Contains($_.FullName) })
    $lstCurrentUnmatched.Items.Clear()
    foreach ($file in $script:State.CurrentUnmatchedFiles) {
        [void]$lstCurrentUnmatched.Items.Add((Get-FileListDisplayName -File $file))
    }
}

function Refresh-PairGrid {
    $gridPairs.Rows.Clear()

    foreach ($pair in $script:State.SuggestedPairs) {
        $rowIndex = $gridPairs.Rows.Add(
            [bool]$pair.Include,
            $pair.PreviousName,
            $pair.CurrentName,
            $pair.MatchStatus,
            $pair.Confidence,
            $pair.OutputName
        )

        $gridPairs.Rows[$rowIndex].Tag = $pair
    }

    Refresh-UnmatchedLists
}

function Suggest-Pairs {
    $script:State.SuggestedPairs.Clear()
    $usedPrevious = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $previousByName = @{}
    $previousByNormalized = @{}

    foreach ($previous in $script:State.PreviousFiles) {
        $previousByName[$previous.Name.ToLowerInvariant()] = $previous

        $normalized = Get-NormalizedBaseName -Name $previous.Name
        if (-not $previousByNormalized.ContainsKey($normalized)) {
            $previousByNormalized[$normalized] = [System.Collections.Generic.List[object]]::new()
        }
        $previousByNormalized[$normalized].Add($previous)
    }

    foreach ($current in $script:State.CurrentFiles) {
        $currentNorm = Get-NormalizedBaseName -Name $current.Name
        $match = $null
        $status = 'Unmatched'
        $confidence = 'None'

        $exactKey = $current.Name.ToLowerInvariant()
        if ($previousByName.ContainsKey($exactKey) -and -not $usedPrevious.Contains($previousByName[$exactKey].FullName)) {
            $match = $previousByName[$exactKey]
            $status = 'Exact filename'
            $confidence = 'High'
        }
        elseif ($previousByNormalized.ContainsKey($currentNorm)) {
            $candidates = @($previousByNormalized[$currentNorm] | Where-Object { -not $usedPrevious.Contains($_.FullName) })
            if ($candidates.Count -eq 1) {
                $match = $candidates[0]
                $status = 'Normalized match'
                $confidence = 'Medium'
            }
        }

        if (-not $match) {
            $best = $null
            $bestScore = 0.0
            foreach ($previous in $script:State.PreviousFiles) {
                if ($usedPrevious.Contains($previous.FullName)) { continue }
                $score = Get-SimilarityScore -A $currentNorm -B (Get-NormalizedBaseName -Name $previous.Name)
                if ($score -gt $bestScore) {
                    $bestScore = $score
                    $best = $previous
                }
            }

            if ($best -and $bestScore -ge $script:FuzzyMatchThreshold) {
                $match = $best
                $status = 'Fuzzy suggestion'
                $confidence = if ($bestScore -ge $script:HigherConfidenceFuzzyThreshold) { 'Medium' } else { 'Low' }
            }
        }

        if ($match) {
            [void]$usedPrevious.Add($match.FullName)

            [void]$script:State.SuggestedPairs.Add([pscustomobject]@{
                Include = $true
                PreviousPath = $match.FullName
                PreviousName = $match.Name
                CurrentPath = $current.FullName
                CurrentName = $current.Name
                MatchStatus = $status
                Confidence = $confidence
                OutputName = New-SafeOutputFileName -CurrentName $current.Name -PreviousName $match.Name
            })
        }
    }

    Refresh-PairGrid
}

function Sync-PairsFromGrid {
    foreach ($row in $gridPairs.Rows) {
        if ($row.IsNewRow) { continue }

        $pair = $row.Tag
        if (-not $pair) { continue }

        $pair.Include = [bool]$row.Cells[0].Value
    }
}

function Open-FolderPicker {
    param([string]$InitialPath)

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.ShowNewFolderButton = $true

    if ($InitialPath -and (Test-Path -LiteralPath $InitialPath -PathType Container)) {
        $dialog.SelectedPath = $InitialPath
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }

    return $null
}

function Scan-Folders {
    if (-not $script:State.PreviousFolder -or -not $script:State.CurrentFolder) {
        [System.Windows.Forms.MessageBox]::Show('Select both Previous and Current folders before scanning.', 'Expecto Comparo', 'OK', 'Warning') | Out-Null
        return
    }

    $script:State.PreviousFiles = @(Get-DocxFiles -Folder $script:State.PreviousFolder)
    $script:State.CurrentFiles = @(Get-DocxFiles -Folder $script:State.CurrentFolder)

    Suggest-Pairs
    $lblCounts.Text = "Previous: $($script:State.PreviousFiles.Count)   Current: $($script:State.CurrentFiles.Count)   Suggested pairs: $($script:State.SuggestedPairs.Count)"
}

function Test-LikelyCloudPath {
    param([Parameter(Mandatory)][string]$Path)

    return ($Path -match '(?i)onedrive|sharepoint')
}

function Save-LogFile {
    if (-not $script:State.LogPath) { return }
    [System.IO.File]::WriteAllLines($script:State.LogPath, $script:State.LogLines)
}

function Start-ComparisonRun {
    Sync-PairsFromGrid

    if (-not $script:State.OutputFolder) {
        [System.Windows.Forms.MessageBox]::Show('Select an output folder first.', 'Expecto Comparo', 'OK', 'Warning') | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $script:State.OutputFolder -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show('Output folder is not accessible.', 'Expecto Comparo', 'OK', 'Error') | Out-Null
        return
    }

    $selectedPairs = @($script:State.SuggestedPairs | Where-Object { $_.Include })
    if ($selectedPairs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No included pairs to process.', 'Expecto Comparo', 'OK', 'Information') | Out-Null
        return
    }

    $txtStatus.Clear()
    $script:State.LogLines.Clear()
    $script:State.LogPath = Join-Path -Path $script:State.OutputFolder -ChildPath ("expecto-comparo-log-{0}.txt" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

    Write-RunLog -Level 'info' -Message "Previous folder: $($script:State.PreviousFolder)" -ToUi
    Write-RunLog -Level 'info' -Message "Current folder: $($script:State.CurrentFolder)" -ToUi
    Write-RunLog -Level 'info' -Message "Output folder: $($script:State.OutputFolder)" -ToUi

    $progressBar.Minimum = 0
    $progressBar.Maximum = $selectedPairs.Count
    $progressBar.Value = 0

    $word = $null
    $createdWord = $false
    $successCount = 0
    $failedCount = 0
    $skippedCount = 0

    try {
        try {
            $word = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            Write-RunLog -Level 'info' -Message 'Using existing Microsoft Word instance.' -ToUi
        }
        catch {
            Write-RunLog -Level 'info' -Message "Could not attach to existing Word instance: $($_.Exception.Message). Starting a new instance." -ToUi
            $word = New-Object -ComObject Word.Application
            $createdWord = $true
            Write-RunLog -Level 'info' -Message 'Started new Microsoft Word instance.' -ToUi
        }

        $word.Visible = [bool]$chkShowWord.Checked
        $word.DisplayAlerts = $script:WdAlertsNone

        foreach ($pair in $selectedPairs) {
            $progressBar.Value++

            $previousPath = $pair.PreviousPath
            $currentPath = $pair.CurrentPath
            $desiredOutput = New-SafeOutputFileName -CurrentName $pair.CurrentName -PreviousName $pair.PreviousName
            $outputPath = Get-UniqueOutputPath -OutputFolder $script:State.OutputFolder -DesiredFileName $desiredOutput

            Write-RunLog -Level 'info' -Message "Comparing: '$($pair.CurrentName)' against '$($pair.PreviousName)'" -ToUi
            Write-RunLog -Level 'info' -Message "Planned output: $outputPath" -ToUi

            if (-not (Test-Path -LiteralPath $previousPath -PathType Leaf)) {
                $skippedCount++
                Write-RunLog -Level 'warn' -Message "Previous file missing: $previousPath" -ToUi
                continue
            }

            if (-not (Test-Path -LiteralPath $currentPath -PathType Leaf)) {
                $skippedCount++
                Write-RunLog -Level 'warn' -Message "Current file missing: $currentPath" -ToUi
                continue
            }

            if ((Test-LikelyCloudPath -Path $previousPath) -or (Test-LikelyCloudPath -Path $currentPath)) {
                Write-RunLog -Level 'info' -Message 'OneDrive/SharePoint path detected. Ensure files are available locally if open fails.' -ToUi
            }

            $previousDoc = $null
            $currentDoc = $null
            $comparisonDoc = $null

            try {
                $readOnly = $true
                $previousDoc = $word.Documents.Open($previousPath, [ref]$false, [ref]$readOnly)
                $currentDoc = $word.Documents.Open($currentPath, [ref]$false, [ref]$readOnly)

                $revisedAuthor = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { 'ExpectoComparo' } else { $env:USERNAME }

                $comparisonDoc = $word.CompareDocuments(
                    $previousDoc,
                    $currentDoc,
                    $script:WdCompareDestinationNew, # wdCompareDestinationNew: create a new comparison document
                    $script:WdGranularityWordLevel, # compare at word level
                    $true, # compare formatting
                    $true, # compare case changes
                    $true, # compare whitespace
                    $true, # compare tables
                    $true, # compare headers
                    $true, # compare footnotes
                    $true, # compare textboxes
                    $true, # compare fields
                    $true, # compare comments
                    $true, # compare moves
                    $revisedAuthor, # revised author label
                    $true # ignore all comparison warnings
                )

                $comparisonDoc.SaveAs([ref]$outputPath, $script:WdFormatXMLDocument)
                $successCount++
                Write-RunLog -Level 'info' -Message "Success: $outputPath" -ToUi
            }
            catch {
                $failedCount++
                Write-RunLog -Level 'error' -Message "Failed comparison for '$($pair.CurrentName)' and '$($pair.PreviousName)': $($_.Exception.Message)" -ToUi
            }
            finally {
                if ($comparisonDoc) { $comparisonDoc.Close([ref]$script:WdDoNotSaveChanges) }
                if ($currentDoc) { $currentDoc.Close([ref]$script:WdDoNotSaveChanges) }
                if ($previousDoc) { $previousDoc.Close([ref]$script:WdDoNotSaveChanges) }
            }
        }
    }
    finally {
        if ($word -and $createdWord) {
            try {
                $word.Quit()
            }
            catch {
                Write-RunLog -Level 'warn' -Message "Could not quit Word cleanly: $($_.Exception.Message)" -ToUi
            }
        }

        Save-LogFile
    }

    Write-RunLog -Level 'info' -Message "Completed. Success: $successCount, Skipped: $skippedCount, Failed: $failedCount" -ToUi
    Write-RunLog -Level 'info' -Message "Log file: $($script:State.LogPath)" -ToUi
    $summaryIcon = if ($failedCount -gt 0) { 'Warning' } else { 'Information' }
    [System.Windows.Forms.MessageBox]::Show(
        "Completed: $successCount`nSkipped: $skippedCount`nFailed: $failedCount`nOutput folder: $($script:State.OutputFolder)`nLog file: $($script:State.LogPath)",
        'Expecto Comparo',
        'OK',
        $summaryIcon
    ) | Out-Null
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Expecto Comparo'
$form.Width = 1240
$form.Height = 720
$form.StartPosition = 'CenterScreen'

$btnPrevious = New-Object System.Windows.Forms.Button
$btnPrevious.Text = 'Select Previous Folder'
$btnPrevious.Width = 180
$btnPrevious.Location = New-Object System.Drawing.Point(10, 10)

$txtPrevious = New-Object System.Windows.Forms.TextBox
$txtPrevious.Width = 1000
$txtPrevious.Location = New-Object System.Drawing.Point(200, 12)
$txtPrevious.ReadOnly = $true

$btnCurrent = New-Object System.Windows.Forms.Button
$btnCurrent.Text = 'Select Current Folder'
$btnCurrent.Width = 180
$btnCurrent.Location = New-Object System.Drawing.Point(10, 45)

$txtCurrent = New-Object System.Windows.Forms.TextBox
$txtCurrent.Width = 1000
$txtCurrent.Location = New-Object System.Drawing.Point(200, 47)
$txtCurrent.ReadOnly = $true

$btnOutput = New-Object System.Windows.Forms.Button
$btnOutput.Text = 'Select Output Folder'
$btnOutput.Width = 180
$btnOutput.Location = New-Object System.Drawing.Point(10, 80)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Width = 1000
$txtOutput.Location = New-Object System.Drawing.Point(200, 82)
$txtOutput.ReadOnly = $true

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = 'Refresh / Rescan'
$btnScan.Width = 140
$btnScan.Location = New-Object System.Drawing.Point(10, 115)

$chkShowWord = New-Object System.Windows.Forms.CheckBox
$chkShowWord.Text = 'Show Word during comparison'
$chkShowWord.Width = 250
$chkShowWord.Location = New-Object System.Drawing.Point(170, 118)

$lblCounts = New-Object System.Windows.Forms.Label
$lblCounts.AutoSize = $true
$lblCounts.Location = New-Object System.Drawing.Point(450, 120)
$lblCounts.Text = 'Previous: 0   Current: 0   Suggested pairs: 0'

$gridPairs = New-Object System.Windows.Forms.DataGridView
$gridPairs.Location = New-Object System.Drawing.Point(10, 150)
$gridPairs.Width = 1200
$gridPairs.Height = 180
$gridPairs.AllowUserToAddRows = $false
$gridPairs.AllowUserToDeleteRows = $false
$gridPairs.SelectionMode = 'FullRowSelect'
$gridPairs.MultiSelect = $false
$gridPairs.AutoSizeColumnsMode = 'Fill'
$gridPairs.ScrollBars = 'Vertical'

[void]$gridPairs.Columns.Add((New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{ Name = 'Include'; HeaderText = 'Include'; FillWeight = 40 }))
[void]$gridPairs.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Previous'; HeaderText = 'Previous document'; FillWeight = 210; ReadOnly = $true }))
[void]$gridPairs.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Current'; HeaderText = 'Current document'; FillWeight = 210; ReadOnly = $true }))
[void]$gridPairs.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'MatchStatus'; HeaderText = 'Match status'; FillWeight = 110; ReadOnly = $true }))
[void]$gridPairs.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'Confidence'; HeaderText = 'Confidence'; FillWeight = 90; ReadOnly = $true }))
[void]$gridPairs.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{ Name = 'OutputName'; HeaderText = 'Output filename'; FillWeight = 240; ReadOnly = $true }))

$lblPrevUnmatched = New-Object System.Windows.Forms.Label
$lblPrevUnmatched.Text = 'Unmatched Previous Files'
$lblPrevUnmatched.Location = New-Object System.Drawing.Point(10, 340)
$lblPrevUnmatched.AutoSize = $true

$lstPreviousUnmatched = New-Object System.Windows.Forms.ListBox
$lstPreviousUnmatched.Location = New-Object System.Drawing.Point(10, 360)
$lstPreviousUnmatched.Width = 500
$lstPreviousUnmatched.Height = 130

$lblCurrUnmatched = New-Object System.Windows.Forms.Label
$lblCurrUnmatched.Text = 'Unmatched Current Files'
$lblCurrUnmatched.Location = New-Object System.Drawing.Point(710, 340)
$lblCurrUnmatched.AutoSize = $true

$lstCurrentUnmatched = New-Object System.Windows.Forms.ListBox
$lstCurrentUnmatched.Location = New-Object System.Drawing.Point(710, 360)
$lstCurrentUnmatched.Width = 500
$lstCurrentUnmatched.Height = 130

$btnPair = New-Object System.Windows.Forms.Button
$btnPair.Text = 'Pair Selected ->'
$btnPair.Width = 160
$btnPair.Location = New-Object System.Drawing.Point(530, 380)

$btnUnpair = New-Object System.Windows.Forms.Button
$btnUnpair.Text = '<- Unpair Row'
$btnUnpair.Width = 160
$btnUnpair.Location = New-Object System.Drawing.Point(530, 420)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 500)
$progressBar.Width = 860
$progressBar.Height = 24

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Start Comparison'
$btnStart.Width = 180
$btnStart.Height = 30
$btnStart.Location = New-Object System.Drawing.Point(890, 497)

$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Location = New-Object System.Drawing.Point(10, 535)
$txtStatus.Width = 1200
$txtStatus.Height = 100
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = 'Vertical'
$txtStatus.ReadOnly = $true

$form.Controls.AddRange(@(
    $btnPrevious, $txtPrevious,
    $btnCurrent, $txtCurrent,
    $btnOutput, $txtOutput,
    $btnScan, $chkShowWord, $lblCounts,
    $gridPairs,
    $lblPrevUnmatched, $lstPreviousUnmatched,
    $lblCurrUnmatched, $lstCurrentUnmatched,
    $btnPair, $btnUnpair,
    $progressBar, $btnStart,
    $txtStatus
))

$btnPrevious.Add_Click({
    Invoke-UiAction -ActionName 'Select Previous Folder' -Action {
        $selected = Open-FolderPicker -InitialPath $script:State.PreviousFolder
        if ($selected) {
            $script:State.PreviousFolder = $selected
            $txtPrevious.Text = $selected
        }
    }
})

$btnCurrent.Add_Click({
    Invoke-UiAction -ActionName 'Select Current Folder' -Action {
        $selected = Open-FolderPicker -InitialPath $script:State.CurrentFolder
        if ($selected) {
            $script:State.CurrentFolder = $selected
            $txtCurrent.Text = $selected
        }
    }
})

$btnOutput.Add_Click({
    Invoke-UiAction -ActionName 'Select Output Folder' -Action {
        $selected = Open-FolderPicker -InitialPath $script:State.OutputFolder
        if ($selected) {
            $script:State.OutputFolder = $selected
            $txtOutput.Text = $selected
        }
    }
})

$btnScan.Add_Click({
    Invoke-UiAction -ActionName 'Refresh / Rescan' -Action {
        Scan-Folders
    }
})

$btnPair.Add_Click({
    Invoke-UiAction -ActionName 'Pair Selected' -Action {
        if (-not $lstPreviousUnmatched.SelectedItem -or -not $lstCurrentUnmatched.SelectedItem) {
            [System.Windows.Forms.MessageBox]::Show('Select one unmatched previous file and one unmatched current file.', 'Expecto Comparo', 'OK', 'Information') | Out-Null
            return
        }

        $previousFile = $script:State.PreviousUnmatchedFiles[$lstPreviousUnmatched.SelectedIndex]
        $currentFile = $script:State.CurrentUnmatchedFiles[$lstCurrentUnmatched.SelectedIndex]
        $previousPath = $previousFile.FullName
        $currentPath = $currentFile.FullName
        $previousName = $previousFile.Name
        $currentName = $currentFile.Name

        [void]$script:State.SuggestedPairs.Add([pscustomobject]@{
            Include = $true
            PreviousPath = $previousPath
            PreviousName = $previousName
            CurrentPath = $currentPath
            CurrentName = $currentName
            MatchStatus = 'Manual pair'
            Confidence = 'User confirmed'
            OutputName = New-SafeOutputFileName -CurrentName $currentName -PreviousName $previousName
        })

        Refresh-PairGrid
    }
})

$btnUnpair.Add_Click({
    Invoke-UiAction -ActionName 'Unpair Row' -Action {
        if ($gridPairs.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Select a pairing row to unpair.', 'Expecto Comparo', 'OK', 'Information') | Out-Null
            return
        }

        $pair = $gridPairs.SelectedRows[0].Tag
        if (-not $pair) { return }

        [void]$script:State.SuggestedPairs.Remove($pair)
        Refresh-PairGrid
    }
})

$btnStart.Add_Click({
    Invoke-UiAction -ActionName 'Start Comparison' -Action {
        Start-ComparisonRun
    }
})

if (-not $env:EXPECTO_COMPARO_TEST_MODE) {
    [void]$form.ShowDialog()
}
