#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# WinForms requires STA. Relaunch in STA if needed.
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    $scriptPath = $PSCommandPath
    if (-not $scriptPath) {
        throw "Unable to relaunch in STA because script path is not available."
    }

    Start-Process -FilePath "powershell.exe" -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-STA",
        "-File", ('"{0}"' -f $scriptPath)
    ) | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:LogDirectory = Join-Path $env:LOCALAPPDATA "QuickPDFA\logs"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not (Test-Path -LiteralPath $script:LogDirectory)) {
        [void](New-Item -Path $script:LogDirectory -ItemType Directory -Force)
    }

    $logFile = Join-Path $script:LogDirectory ("QuickPDFA_{0}.log" -f (Get-Date -Format "yyyyMMdd"))
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8
}

function Prompt-InstallGhostscript {
    param(
        [string]$Reason = "Ghostscript is not installed or not available in PATH."
    )

    $downloadUrl = "https://ghostscript.com/releases/gsdnld.html"
    $promptText = "$Reason`n`nWould you like to open the Ghostscript download page now?"
    $answer = [System.Windows.Forms.MessageBox]::Show(
        $promptText,
        "QuickPDFA - Ghostscript Required",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
        Start-Process -FilePath $downloadUrl | Out-Null
    }
}

function Find-Ghostscript {
    if ($env:GHOSTSCRIPT_EXE -and (Test-Path -LiteralPath $env:GHOSTSCRIPT_EXE)) {
        return $env:GHOSTSCRIPT_EXE
    }

    $commandCandidates = @("gswin64c.exe", "gswin32c.exe")
    foreach ($candidate in $commandCandidates) {
        $found = Get-Command -Name $candidate -ErrorAction SilentlyContinue
        if ($found -and (Test-Path -LiteralPath $found.Source)) {
            return $found.Source
        }
    }

    $pathPatterns = @(
        (Join-Path $env:ProgramFiles "gs\*\bin\gswin64c.exe"),
        (Join-Path $env:ProgramFiles "gs\*\bin\gswin32c.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "gs\*\bin\gswin64c.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "gs\*\bin\gswin32c.exe")
    )

    $matches = foreach ($pattern in $pathPatterns) {
        Get-ChildItem -Path $pattern -File -ErrorAction SilentlyContinue
    }

    $latest = $matches |
        Sort-Object -Property LastWriteTime -Descending |
        Select-Object -First 1

    if ($latest) {
        return $latest.FullName
    }

    return $null
}

function Convert-ToPdfA {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPdf,

        [Parameter(Mandatory = $true)]
        [string]$OutputPdf,

        [Parameter(Mandatory = $true)]
        [string]$GhostscriptExe
    )

    if (-not (Test-Path -LiteralPath $InputPdf)) {
        throw "Input PDF was not found: $InputPdf"
    }

    if ([IO.Path]::GetExtension($InputPdf).ToLowerInvariant() -ne ".pdf") {
        throw "Input file is not a PDF: $InputPdf"
    }

    if ([IO.Path]::GetExtension($OutputPdf).ToLowerInvariant() -ne ".pdf") {
        throw "Output file must have a .pdf extension: $OutputPdf"
    }

    if ([IO.Path]::GetFullPath($InputPdf) -eq [IO.Path]::GetFullPath($OutputPdf)) {
        throw "Input and output paths cannot be the same file."
    }

    Write-Log ("Starting conversion. Input='{0}', Output='{1}', Ghostscript='{2}'" -f $InputPdf, $OutputPdf, $GhostscriptExe)

        $gsRoot = Split-Path (Split-Path $GhostscriptExe -Parent) -Parent
        $pdfaDefTemplate = Join-Path $gsRoot "lib\PDFA_def.ps"
        $iccProfilePath = Join-Path $gsRoot "iccprofiles\srgb.icc"

        if (-not (Test-Path -LiteralPath $pdfaDefTemplate)) {
            throw "Ghostscript PDFA definition file not found: $pdfaDefTemplate"
        }

        if (-not (Test-Path -LiteralPath $iccProfilePath)) {
                throw "Ghostscript ICC profile not found: $iccProfilePath"
        }

        $iccProfileForPs = $iccProfilePath -replace "\\", "/"
        $tempPdfaDef = Join-Path $env:TEMP ("QuickPDFA_PDFA_def_{0}.ps" -f [guid]::NewGuid().ToString("N"))
        $pdfaDefContent = Get-Content -LiteralPath $pdfaDefTemplate -Raw
        $pdfaDefContent = [Regex]::Replace(
            $pdfaDefContent,
            "/ICCProfile\s+\([^\)]*\)\s*%\s*Customise",
            "/ICCProfile ($iccProfileForPs) % Customise"
        )

        Set-Content -LiteralPath $tempPdfaDef -Value $pdfaDefContent -Encoding ASCII
        Write-Log ("Using PDFA template '{0}'" -f $pdfaDefTemplate)
        Write-Log ("Using ICC profile '{0}'" -f $iccProfilePath)
        Write-Log ("Using temp PDFA definition '{0}'" -f $tempPdfaDef)

    $args = @(
        "-dPDFA=2",
        "-dBATCH",
        "-dNOPAUSE",
        "-dNOOUTERSAVE",
        "-sDEVICE=pdfwrite",
        "-sColorConversionStrategy=RGB",
        "-sProcessColorModel=DeviceRGB",
        "-dAutoRotatePages=/None",
        "-dEmbedAllFonts=true",
        "-dSubsetFonts=true",
        "-dCompressFonts=true",
        "-dDetectDuplicateImages=true",
        "-dDownsampleColorImages=false",
        "-dDownsampleGrayImages=false",
        "-dDownsampleMonoImages=false",
        "-dPDFACompatibilityPolicy=1",
        "--permit-file-read=$iccProfilePath",
        "-sOutputFile=$outputPdf",
        $tempPdfaDef,
        $InputPdf
    )

    Write-Log ("Ghostscript args: {0}" -f ($args -join " "))

    $outputLines = @()
    $exitCode = -1
    $previousErrorAction = $ErrorActionPreference
    try {
        # Ghostscript writes useful diagnostics to stderr even on successful runs.
        # Keep ErrorActionPreference non-terminating for the native call so output is always captured.
        $ErrorActionPreference = "Continue"
        $outputLines = & $GhostscriptExe @args 2>&1
        $exitCode = $LASTEXITCODE

        $combinedText = ($outputLines | Out-String).Trim()
        if ($combinedText) {
            Write-Log ("Ghostscript output:`n{0}" -f $combinedText)
        }
        else {
            Write-Log "Ghostscript produced no stdout/stderr output."
        }

        Write-Log ("Ghostscript exit code: {0}" -f $exitCode)
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
        if (Test-Path -LiteralPath $tempPdfaDef) {
            Remove-Item -LiteralPath $tempPdfaDef -Force -ErrorAction SilentlyContinue
        }
    }

    if ($exitCode -ne 0) {
        $details = ($outputLines | Out-String).Trim()
        if (-not $details) {
            $details = "No additional error text was returned by Ghostscript."
        }
        Write-Log ("Ghostscript failed with exit code {0}." -f $exitCode)
        throw "Ghostscript failed with exit code $exitCode. $details"
    }

    if (-not (Test-Path -LiteralPath $outputPdf)) {
        Write-Log "Ghostscript returned success but output file was not created."
        throw "Ghostscript completed but no output file was created."
    }

    $outputInfo = Get-Item -LiteralPath $outputPdf
    if ($outputInfo.Length -le 0) {
        Write-Log "Ghostscript returned success but output file is 0 bytes."
        throw "Ghostscript completed but output file is 0 bytes. Check the log for details in $script:LogDirectory"
    }

    Write-Log ("Conversion succeeded. Output size={0} bytes" -f $outputInfo.Length)

    return $outputPdf
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "QuickPDFA"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(640, 320)
$form.MinimumSize = New-Object System.Drawing.Size(640, 320)
$form.MaximizeBox = $false

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Drop a PDF, then click Convert"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(16, 14)
$form.Controls.Add($titleLabel)

$dropPanel = New-Object System.Windows.Forms.Panel
$dropPanel.Location = New-Object System.Drawing.Point(16, 52)
$dropPanel.Size = New-Object System.Drawing.Size(592, 110)
$dropPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dropPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
$dropPanel.AllowDrop = $true
$form.Controls.Add($dropPanel)

$dropLabel = New-Object System.Windows.Forms.Label
$dropLabel.Text = "Drag and drop one PDF file here"
$dropLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$dropLabel.AutoSize = $true
$dropLabel.Location = New-Object System.Drawing.Point(170, 42)
$dropPanel.Controls.Add($dropLabel)

$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Location = New-Object System.Drawing.Point(16, 178)
$inputBox.Size = New-Object System.Drawing.Size(592, 24)
$inputBox.ReadOnly = $true
$form.Controls.Add($inputBox)

$convertButton = New-Object System.Windows.Forms.Button
$convertButton.Text = "Convert to PDF/A"
$convertButton.Location = New-Object System.Drawing.Point(16, 214)
$convertButton.Size = New-Object System.Drawing.Size(180, 34)
$convertButton.Enabled = $false
$form.Controls.Add($convertButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Checking Ghostscript..."
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(214, 223)
$form.Controls.Add($statusLabel)

$selectedPdf = $null
$ghostscriptExe = Find-Ghostscript

if ($ghostscriptExe) {
    $statusLabel.Text = "Ghostscript found: $ghostscriptExe"
    Write-Log ("Ghostscript detected at '{0}'" -f $ghostscriptExe)
}
else {
    $statusLabel.Text = "Ghostscript not found. Install it, then relaunch this app."
    Write-Log "Ghostscript not found during startup."
    Prompt-InstallGhostscript
}

$setSelectedFile = {
    param([string]$path)

    if (-not $path) {
        return
    }

    if (([IO.Path]::GetExtension($path)).ToLowerInvariant() -ne ".pdf") {
        [System.Windows.Forms.MessageBox]::Show(
            "Please drop a .pdf file.",
            "QuickPDFA",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $script:selectedPdf = $path
    $inputBox.Text = $path
    $convertButton.Enabled = [bool]($script:selectedPdf -and $script:ghostscriptExe)
    $statusLabel.Text = "Ready to convert."
}

$openFilePicker = {
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Title = "Open a file..."
    $openDialog.Filter = "PDF files (*.pdf)|*.pdf"
    $openDialog.DefaultExt = "pdf"
    $openDialog.CheckFileExists = $true
    $openDialog.Multiselect = $false

    if ($openDialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Log ("File selected from Open dialog: {0}" -f $openDialog.FileName)
        & $setSelectedFile $openDialog.FileName
    }
    else {
        Write-Log "Open file dialog canceled by user."
    }
}

$dropLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
$dropLabel.Add_Click({
    & $openFilePicker
})

$dropPanel.Add_Click({
    & $openFilePicker
})

$dropPanel.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $files = [string[]]$_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($files.Count -eq 1 -and ([IO.Path]::GetExtension($files[0]).ToLowerInvariant() -eq ".pdf")) {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy
            $dropPanel.BackColor = [System.Drawing.Color]::FromArgb(232, 242, 255)
        }
        else {
            $_.Effect = [System.Windows.Forms.DragDropEffects]::None
        }
    }
    else {
        $_.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

$dropPanel.Add_DragLeave({
    $dropPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
})

$dropPanel.Add_DragDrop({
    $dropPanel.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)
    $files = [string[]]$_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if ($files.Count -ge 1) {
        & $setSelectedFile $files[0]
    }
})

$convertButton.Add_Click({
    if (-not $selectedPdf) {
        return
    }

    if (-not $ghostscriptExe) {
        Write-Log "Convert blocked: Ghostscript not found."
        Prompt-InstallGhostscript
        [System.Windows.Forms.MessageBox]::Show(
            "Ghostscript was not found.",
            "QuickPDFA",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $defaultFolder = Split-Path -Path $selectedPdf -Parent
    $defaultName = [IO.Path]::GetFileNameWithoutExtension($selectedPdf)
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Title = "Save PDF/A As"
    $saveDialog.Filter = "PDF files (*.pdf)|*.pdf"
    $saveDialog.DefaultExt = "pdf"
    $saveDialog.AddExtension = $true
    $saveDialog.OverwritePrompt = $true
    $saveDialog.InitialDirectory = $defaultFolder
    $saveDialog.FileName = ("{0}_PDFA.pdf" -f $defaultName)

    if ($saveDialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
        $statusLabel.Text = "Save canceled."
        Write-Log "Save dialog canceled by user."
        return
    }

    $outputPdf = $saveDialog.FileName

    try {
        $convertButton.Enabled = $false
        $form.UseWaitCursor = $true
        $statusLabel.Text = "Converting..."
        $form.Refresh()

        $result = Convert-ToPdfA -InputPdf $selectedPdf -OutputPdf $outputPdf -GhostscriptExe $ghostscriptExe

        $statusLabel.Text = "Done: $result"
        [System.Windows.Forms.MessageBox]::Show(
            "PDF/A created successfully:`n$result",
            "QuickPDFA",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        $statusLabel.Text = "Conversion failed."
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "QuickPDFA",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $form.UseWaitCursor = $false
        $convertButton.Enabled = [bool]($selectedPdf -and $ghostscriptExe)
    }
})

[void]$form.ShowDialog()
