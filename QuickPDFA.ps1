#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# WPF requires STA. Relaunch in STA if needed.
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

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$script:LogDirectory = Join-Path $env:LOCALAPPDATA "QuickPDFA\logs"
$script:CurrentLanguage = "en"
$script:CurrentStatus = "checking"
$script:SelectedPdf = $null
$script:GhostscriptExe = $null
$script:SelectedPdfaProfile = "2a"
$script:EnableHighCompression = $false
$script:SelectedConversionMode = "strict"
$script:ExternalTagger = $null
$script:DocumentLanguage = "en-US"

$script:Translations = @{
    en = @{
        appTitle = "QuickPDFA"
        titleLabel = "Drop a PDF, then click Convert"
        dropLabel = "Drag and drop one PDF file here"
        convertButton = "Convert to PDF/A"
        toggleButton = "IT"
        pdfaOptionsLabel = "PDF/A profile"
        pdfa2a = "PDF/A-2a (default)"
        pdfa2b = "PDF/A-2b"
        pdfa1b = "PDF/A-1b"
        pdfa1a = "PDF/A-1a"
        documentLanguageLabel = "Document language (BCP 47)"
        compressCheckbox = "High compression (smaller size)"
        modeLabel = "Conversion mode"
        modeStrict = "Strict (fail if level is not possible)"
        modeBestEffort = "Best effort (fallback a -> b)"
        modeAdvanced = "Advanced pipeline (external OCR/tagging)"
        ocrStatusReady = "OCR plugin: Ready"
        ocrStatusMissing = "OCR plugin: Not configured"
        ocrSetupButton = "Setup OCR"
        dialogOcrSetupTitle = "QuickPDFA - OCR Setup"
        promptOcrSetupOptions = "OCR plugin options:`n`nYes = Select an existing tagging executable`nNo = Open OpenDataLoader installation guide`nCancel = Do nothing"
        promptAdvancedModeMissingOcr = "Advanced mode needs an OCR plugin.`n`nYes = Setup OCR now`nNo = Continue with Best effort mode`nCancel = Cancel conversion"
        promptFallbackAfterSetup = "OCR is still unavailable. Continue with Best effort mode?"
        statusCheckingGhostscript = "Checking Ghostscript..."
        statusGhostscriptFound = "Ghostscript found: {0}"
        statusGhostscriptMissing = "Ghostscript not found. Install it, then relaunch this app."
        statusReady = "Ready to convert."
        statusConverting = "Converting..."
        statusDone = "Done: {0}"
        statusSaveCanceled = "Save canceled."
        statusConversionFailed = "Conversion failed."
        warningDropPdf = "Please drop or select a .pdf file."
        dialogOpenTitle = "Open a file..."
        dialogSaveTitle = "Save PDF/A As"
        dialogSuccessBody = "PDF/A created successfully:`n{0}"
        dialogGhostscriptMissing = "Ghostscript was not found."
        dialogGhostscriptRequiredTitle = "QuickPDFA - Ghostscript Required"
        promptOpenGhostscriptPage = "{0}`n`nWould you like to open the Ghostscript download page now?"
    }
    it = @{
        appTitle = "QuickPDFA"
        titleLabel = "Trascina un PDF, poi fai clic su Converti"
        dropLabel = "Trascina qui un file PDF"
        convertButton = "Converti in PDF/A"
        toggleButton = "IT / EN"
        pdfaOptionsLabel = "Profilo PDF/A"
        pdfa2a = "PDF/A-2a (predefinito)"
        pdfa2b = "PDF/A-2b"
        pdfa1b = "PDF/A-1b"
        pdfa1a = "PDF/A-1a"
        documentLanguageLabel = "Lingua del documento (BCP 47)"
        compressCheckbox = "Alta compressione (dimensioni ridotte)"
        modeLabel = "Modalita conversione"
        modeStrict = "Rigida (errore se il livello non e possibile)"
        modeBestEffort = "Best effort (fallback da a a b)"
        modeAdvanced = "Pipeline avanzata (OCR/tag esterno)"
        ocrStatusReady = "Plugin OCR: pronto"
        ocrStatusMissing = "Plugin OCR: non configurato"
        ocrSetupButton = "Configura OCR"
        dialogOcrSetupTitle = "QuickPDFA - Configurazione OCR"
        promptOcrSetupOptions = "Opzioni plugin OCR:`n`nSi = Seleziona un eseguibile di tagging esistente`nNo = Apri la guida di installazione di OpenDataLoader`nAnnulla = Non fare nulla"
        promptAdvancedModeMissingOcr = "La modalita avanzata richiede un plugin OCR.`n`nSi = Configura OCR ora`nNo = Continua con modalita Best effort`nAnnulla = Annulla conversione"
        promptFallbackAfterSetup = "OCR ancora non disponibile. Vuoi continuare con modalita Best effort?"
        statusCheckingGhostscript = "Controllo di Ghostscript in corso..."
        statusGhostscriptFound = "Ghostscript trovato: {0}"
        statusGhostscriptMissing = "Ghostscript non trovato. Installalo e riavvia l'app."
        statusReady = "Pronto per la conversione."
        statusConverting = "Conversione in corso..."
        statusDone = "Completato: {0}"
        statusSaveCanceled = "Salvataggio annullato."
        statusConversionFailed = "Conversione non riuscita."
        warningDropPdf = "Seleziona o trascina un file .pdf."
        dialogOpenTitle = "Apri file..."
        dialogSaveTitle = "Salva PDF/A con nome"
        dialogSuccessBody = "PDF/A creato con successo:`n{0}"
        dialogGhostscriptMissing = "Ghostscript non e stato trovato."
        dialogGhostscriptRequiredTitle = "QuickPDFA - Ghostscript Richiesto"
        promptOpenGhostscriptPage = "{0}`n`nVuoi aprire ora la pagina di download di Ghostscript?"
    }
}

function Get-Text {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Key,

        [object[]]$Args
    )

    $table = $script:Translations[$script:CurrentLanguage]
    if (-not $table) {
        $table = $script:Translations["en"]
    }

    $text = $table[$Key]
    if (-not $text) {
        $text = $script:Translations["en"][$Key]
    }

    if ($Args -and $Args.Count -gt 0) {
        return [string]::Format($text, $Args)
    }

    return $text
}

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

function Normalize-DocumentLanguageTag {
    param(
        [string]$LanguageTag
    )

    $normalized = ($LanguageTag -as [string]).Trim()
    if (-not $normalized) {
        return "en-US"
    }

    $normalized = $normalized -replace '_', '-'
    if ($normalized -notmatch '^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$') {
        throw "Document language must be a BCP 47 tag such as en-US or it-IT."
    }

    return $normalized
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

function Find-ExternalTagger {
    $runCommand = {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Command,

            [string[]]$Args = @()
        )

        $output = @()
        $exitCode = -1
        $previousErrorAction = $ErrorActionPreference
        try {
            $ErrorActionPreference = "Continue"
            $output = & $Command @Args 2>&1
            $exitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previousErrorAction
        }

        return @{
            ExitCode = $exitCode
            Text = ($output | Out-String).Trim()
        }
    }

    $probeCommand = {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Command,

            [string[]]$BaseArgs = @()
        )

        $versionProbeArgs = @()
        if ($BaseArgs) {
            $versionProbeArgs += $BaseArgs
        }
        $versionProbeArgs += "--version"

        $versionProbe = & $runCommand -Command $Command -Args $versionProbeArgs
        if ($versionProbe.ExitCode -eq 0) {
            return $true
        }

        $helpProbeArgs = @()
        if ($BaseArgs) {
            $helpProbeArgs += $BaseArgs
        }
        $helpProbeArgs += "--help"

        $helpProbe = & $runCommand -Command $Command -Args $helpProbeArgs
        return ($helpProbe.ExitCode -eq 0)
    }

    if ($env:QUICKPDFA_TAGGER_EXE -and (Test-Path -LiteralPath $env:QUICKPDFA_TAGGER_EXE)) {
        $probe = & $runCommand -Command $env:QUICKPDFA_TAGGER_EXE -Args @("--version")
        $engine = if ($probe.Text -match "opendataloader") { "opendataloader" } elseif ($probe.Text -match "ocrmypdf") { "ocrmypdf" } else { "custom" }

        return @{
            DisplayName = $env:QUICKPDFA_TAGGER_EXE
            Command = $env:QUICKPDFA_TAGGER_EXE
            BaseArgs = @()
            Engine = $engine
            Source = "env"
        }
    }

    $opendataloaderCandidates = @("opendataloader-pdf")
    foreach ($candidate in $opendataloaderCandidates) {
        $found = Get-Command -Name $candidate -ErrorAction SilentlyContinue
        if ($found -and $found.Source) {
            if (& $probeCommand -Command $found.Source) {
                return @{
                    DisplayName = $found.Source
                    Command = $found.Source
                    BaseArgs = @()
                    Engine = "opendataloader"
                    Source = "path"
                }
            }
        }
    }

    $scriptPathPatterns = @(
        (Join-Path $env:APPDATA "Python\Python*\Scripts\opendataloader-pdf*.exe"),
        (Join-Path $env:APPDATA "Python\Python*\Scripts\opendataloader-pdf*.cmd"),
        (Join-Path $env:LOCALAPPDATA "Programs\Python\Python*\Scripts\opendataloader-pdf*.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\Python\Python*\Scripts\opendataloader-pdf*.cmd")
    )

    $scriptCandidates = foreach ($pattern in $scriptPathPatterns) {
        Get-ChildItem -Path $pattern -File -ErrorAction SilentlyContinue
    }

    $bestScriptCandidate = $scriptCandidates |
        Sort-Object -Property LastWriteTime -Descending |
        Select-Object -First 1

    if ($bestScriptCandidate -and (& $probeCommand -Command $bestScriptCandidate.FullName)) {
        return @{
            DisplayName = $bestScriptCandidate.FullName
            Command = $bestScriptCandidate.FullName
            BaseArgs = @()
            Engine = "opendataloader"
            Source = "python-scripts"
        }
    }

    $pythonLaunchers = @("py", "python", "python3")
    foreach ($launcher in $pythonLaunchers) {
        $foundLauncher = Get-Command -Name $launcher -ErrorAction SilentlyContinue
        if (-not $foundLauncher) {
            continue
        }

        if (& $probeCommand -Command $foundLauncher.Source -BaseArgs @("-m", "opendataloader_pdf")) {
            return @{
                DisplayName = ("{0} -m opendataloader_pdf" -f $foundLauncher.Source)
                Command = $foundLauncher.Source
                BaseArgs = @("-m", "opendataloader_pdf")
                Engine = "opendataloader"
                Source = "python-module"
            }
        }
    }

    $candidates = @("ocrmypdf.exe", "ocrmypdf")
    foreach ($candidate in $candidates) {
        $found = Get-Command -Name $candidate -ErrorAction SilentlyContinue
        if ($found -and $found.Source) {
            return @{
                DisplayName = $found.Source
                Command = $found.Source
                BaseArgs = @()
                Engine = "ocrmypdf"
                Source = "path"
            }
        }
    }

    foreach ($launcher in $pythonLaunchers) {
        $foundLauncher = Get-Command -Name $launcher -ErrorAction SilentlyContinue
        if (-not $foundLauncher) {
            continue
        }

        $ocrProbe = & $runCommand -Command $foundLauncher.Source -Args @("-m", "ocrmypdf", "--version")
        if ($ocrProbe.ExitCode -eq 0) {
            return @{
                DisplayName = ("{0} -m ocrmypdf" -f $foundLauncher.Source)
                Command = $foundLauncher.Source
                BaseArgs = @("-m", "ocrmypdf")
                Engine = "ocrmypdf"
                Source = "python-module"
            }
        }
    }

    return $null
}

function Prompt-InstallGhostscript {
    param(
        [string]$Reason = "Ghostscript is not installed or not available in PATH."
    )

    $downloadUrl = "https://ghostscript.com/releases/gsdnld.html"
    $promptText = Get-Text "promptOpenGhostscriptPage" @($Reason)
    $answer = [System.Windows.MessageBox]::Show(
        $promptText,
        (Get-Text "dialogGhostscriptRequiredTitle"),
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )

    if ($answer -eq [System.Windows.MessageBoxResult]::Yes) {
        Start-Process -FilePath $downloadUrl | Out-Null
    }
}

function Get-PdfaConformanceFromFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PdfPath
    )

    if (-not (Test-Path -LiteralPath $PdfPath)) {
        return $null
    }

    # PDF metadata is text embedded in a binary file. Latin-1 preserves byte values 1:1.
    $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
    $content = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)

    # Match both XMP styles:
    # 1) element text: <pdfaid:part>2</pdfaid:part>
    # 2) rdf attributes: pdfaid:part="2" pdfaid:conformance="B"
    $partMatch = [Regex]::Match($content, '<pdfaid:part>\s*([12])\s*</pdfaid:part>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $partMatch.Success) {
        $partMatch = [Regex]::Match($content, 'pdfaid:part\s*=\s*["'']\s*([12])\s*["'']', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }

    $confMatch = [Regex]::Match($content, '<pdfaid:conformance>\s*([ABU])\s*</pdfaid:conformance>', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $confMatch.Success) {
        $confMatch = [Regex]::Match($content, 'pdfaid:conformance\s*=\s*["'']\s*([ABU])\s*["'']', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }

    if (-not $partMatch.Success -or -not $confMatch.Success) {
        return $null
    }

    return @{
        Part = $partMatch.Groups[1].Value
        Conformance = $confMatch.Groups[1].Value.ToUpperInvariant()
    }
}

function Test-PdfHasStructTree {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PdfPath
    )

    if (-not (Test-Path -LiteralPath $PdfPath)) {
        return $false
    }

    $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
    $content = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)

    return [Regex]::IsMatch($content, "/StructTreeRoot", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Invoke-AdvancedTaggingPipeline {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPdf,

        [Parameter(Mandatory = $true)]
        [hashtable]$Tagger
    )

    $tempTaggedPdf = Join-Path $env:TEMP ("QuickPDFA_tagged_{0}.pdf" -f [guid]::NewGuid().ToString("N"))
    $tempOutputDir = Join-Path $env:TEMP ("QuickPDFA_odl_output_{0}" -f [guid]::NewGuid().ToString("N"))
    $outputLines = @()
    $exitCode = -1
    $previousErrorAction = $ErrorActionPreference

    try {
        $taggerArgs = @()
        if ($Tagger.Engine -eq "opendataloader") {
            [void](New-Item -Path $tempOutputDir -ItemType Directory -Force)
            $taggerArgs = @(
                "--format", "tagged-pdf",
                "--output-dir", $tempOutputDir,
                $InputPdf
            )
        }
        elseif ($Tagger.Engine -eq "ocrmypdf") {
            # OCRmyPDF fallback path.
            $taggerArgs = @(
                "--force-ocr",
                "--skip-big", "200",
                "--output-type", "pdf",
                "--jobs", "1",
                $InputPdf,
                $tempTaggedPdf
            )
        }
        else {
            throw "Unsupported external tagger engine '$($Tagger.Engine)'. Configure OpenDataLoader PDF (recommended) or OCRmyPDF."
        }

        $fullTaggerArgs = @()
        if ($Tagger.BaseArgs) {
            $fullTaggerArgs += $Tagger.BaseArgs
        }
        $fullTaggerArgs += $taggerArgs

        Write-Log ("Running advanced pipeline ({0}): {1} {2}" -f $Tagger.Engine, $Tagger.DisplayName, ($taggerArgs -join " "))
        $ErrorActionPreference = "Continue"
        $outputLines = & $Tagger.Command @fullTaggerArgs 2>&1
        $exitCode = $LASTEXITCODE

        $combinedText = ($outputLines | Out-String).Trim()
        if ($combinedText) {
            Write-Log ("Advanced pipeline output:`n{0}" -f $combinedText)
        }
        Write-Log ("Advanced pipeline exit code: {0}" -f $exitCode)

        if ($exitCode -eq 0 -and $Tagger.Engine -eq "opendataloader") {
            $inputFileName = [IO.Path]::GetFileName($InputPdf)
            $expectedPath = Join-Path $tempOutputDir $inputFileName
            $taggedCandidate = $null

            if (Test-Path -LiteralPath $expectedPath) {
                $taggedCandidate = Get-Item -LiteralPath $expectedPath
            }
            else {
                $taggedCandidate = Get-ChildItem -Path $tempOutputDir -Filter *.pdf -File -Recurse -ErrorAction SilentlyContinue |
                    Sort-Object -Property LastWriteTime -Descending |
                    Select-Object -First 1
            }

            if (-not $taggedCandidate) {
                throw "OpenDataLoader completed but no tagged PDF output was found in '$tempOutputDir'."
            }

            Copy-Item -LiteralPath $taggedCandidate.FullName -Destination $tempTaggedPdf -Force
        }
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
        if (Test-Path -LiteralPath $tempOutputDir) {
            Remove-Item -LiteralPath $tempOutputDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if ($exitCode -ne 0) {
        $details = ($outputLines | Out-String).Trim()
        $detailLines = @()
        if ($details) {
            $detailLines = $details -split "`r?`n"
        }

        $detailPreview = if ($detailLines.Count -gt 0) {
            ($detailLines | Select-Object -First 8) -join " `n"
        }
        else {
            "No additional backend output."
        }

        if ($Tagger.Engine -eq "opendataloader" -and ($details -match "UnsupportedClassVersionError" -or $details -match "class file version 55\.0" -or $details -match "version of the Java Runtime")) {
            throw "OpenDataLoader requires Java 11 or newer, but your system is running an older Java runtime (likely Java 8). Install Java 11+ and restart QuickPDFA. Backend details: $detailPreview"
        }

        throw "Advanced pipeline failed using backend '$($Tagger.DisplayName)' (exit code $exitCode). Details: $detailPreview"
    }

    if (-not (Test-Path -LiteralPath $tempTaggedPdf)) {
        if (Test-Path -LiteralPath $tempTaggedPdf) {
            Remove-Item -LiteralPath $tempTaggedPdf -Force -ErrorAction SilentlyContinue
        }
        throw "Advanced pipeline finished without producing a tagged PDF output. Configure a working external tagger tool (recommended: OpenDataLoader PDF) or change mode."
    }

    return $tempTaggedPdf
}

function Convert-ToPdfA {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPdf,

        [Parameter(Mandatory = $true)]
        [string]$OutputPdf,

        [Parameter(Mandatory = $true)]
        [string]$GhostscriptExe,

        [Parameter(Mandatory = $true)]
        [ValidateSet("1a", "1b", "2a", "2b")]
        [string]$PdfaProfile,

        [Parameter(Mandatory = $true)]
        [bool]$EnableCompression,

        [Parameter(Mandatory = $true)]
        [ValidateSet("strict", "best-effort", "advanced")]
        [string]$ConversionMode,

        [Parameter(Mandatory = $true)]
        [string]$DocumentLanguage
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

    $documentLanguageTag = Normalize-DocumentLanguageTag -LanguageTag $DocumentLanguage

    Write-Log ("Starting conversion. Input='{0}', Output='{1}', Ghostscript='{2}', Profile='{3}', Compression='{4}', Mode='{5}', Language='{6}'" -f $InputPdf, $OutputPdf, $GhostscriptExe, $PdfaProfile, $EnableCompression, $ConversionMode, $documentLanguageTag)

    $effectivePdfaProfile = $PdfaProfile
    $effectiveInputPdf = $InputPdf
    $tempAdvancedPdf = $null
    $savedAdvancedDebugPdf = $null

    $pdfaLevel = if ($effectivePdfaProfile -eq "1a") { 1 } else { 2 }
    $taggedProfile = $effectivePdfaProfile -in @("1a", "2a")

    if ($taggedProfile -and -not (Test-PdfHasStructTree -PdfPath $effectiveInputPdf)) {
        if ($ConversionMode -eq "strict") {
            throw "Requested PDF/A-$effectivePdfaProfile requires a tagged source PDF, but the input appears untagged (no StructTreeRoot)."
        }

        if ($ConversionMode -eq "best-effort") {
            $effectivePdfaProfile = if ($effectivePdfaProfile -eq "1a") { "1b" } else { "2b" }
            Write-Log ("Best-effort mode fallback applied. Effective profile='{0}'" -f $effectivePdfaProfile)
        }

        if ($ConversionMode -eq "advanced") {
            if (-not $script:ExternalTagger) {
                throw "Advanced mode selected, but no external OCR/tagging tool was found. Set QUICKPDFA_TAGGER_EXE or install OpenDataLoader PDF."
            }

            $tempAdvancedPdf = Invoke-AdvancedTaggingPipeline -InputPdf $effectiveInputPdf -Tagger $script:ExternalTagger
            $effectiveInputPdf = $tempAdvancedPdf

            if (-not (Test-PdfHasStructTree -PdfPath $effectiveInputPdf)) {
                Write-Log ("Advanced pipeline output is still untagged (no StructTreeRoot). Backend='{0}'." -f $script:ExternalTagger.DisplayName)
                $savedAdvancedDebugPdf = Join-Path $script:LogDirectory ("QuickPDFA_advanced_debug_{0}.pdf" -f [guid]::NewGuid().ToString("N"))
                Copy-Item -LiteralPath $effectiveInputPdf -Destination $savedAdvancedDebugPdf -Force
                Write-Log ("Saved advanced pipeline debug PDF: {0}" -f $savedAdvancedDebugPdf)
                throw ("Advanced mode requires tagged PDF output (StructTreeRoot), but backend '{0}' produced an untagged PDF. Advanced debug PDF: {1}" -f $script:ExternalTagger.DisplayName, $savedAdvancedDebugPdf)
            }
        }
    }

    $pdfaLevel = if ($effectivePdfaProfile -eq "1a" -or $effectivePdfaProfile -eq "1b") { 1 } else { 2 }
    $taggedProfile = $effectivePdfaProfile -in @("1a", "2a")

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
    $pdfaDefContent += "`r`n[{Catalog} <</Lang ($documentLanguageTag)>> /PUT pdfmark`r`n"

    Set-Content -LiteralPath $tempPdfaDef -Value $pdfaDefContent -Encoding ASCII
    Write-Log ("Using PDFA template '{0}'" -f $pdfaDefTemplate)
    Write-Log ("Using ICC profile '{0}'" -f $iccProfilePath)
    Write-Log ("Using document language '{0}'" -f $documentLanguageTag)
    Write-Log ("Using temp PDFA definition '{0}'" -f $tempPdfaDef)

    $args = @(
        "-dPDFA=$pdfaLevel",
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
        $effectiveInputPdf
    )

    if ($taggedProfile) {
        $args += @(
            "-dPreserveMarkedContent=true",
            "-dUseTaggedPDF=true"
        )
    }

    if ($EnableCompression) {
        $args += @(
            "-dPDFSETTINGS=/screen",
            "-dDownsampleColorImages=true",
            "-dDownsampleGrayImages=true",
            "-dDownsampleMonoImages=true",
            "-dColorImageDownsampleType=/Bicubic",
            "-dGrayImageDownsampleType=/Bicubic",
            "-dMonoImageDownsampleType=/Subsample",
            "-dColorImageResolution=110",
            "-dGrayImageResolution=110",
            "-dMonoImageResolution=300",
            "-dCompressPages=true"
        )
    }

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
        if ($tempAdvancedPdf -and (Test-Path -LiteralPath $tempAdvancedPdf)) {
            Remove-Item -LiteralPath $tempAdvancedPdf -Force -ErrorAction SilentlyContinue
        }
    }

    if ($exitCode -ne 0) {
        $details = ($outputLines | Out-String).Trim()
        if (-not $details) {
            $details = "No additional error text was returned by Ghostscript."
        }
        Write-Log ("Ghostscript failed with exit code {0}." -f $exitCode)
        $debugHint = if ($savedAdvancedDebugPdf) { " Advanced debug PDF: $savedAdvancedDebugPdf" } else { "" }
        throw "Ghostscript failed with exit code $exitCode. $details$debugHint"
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

    $expectedPart = if ($effectivePdfaProfile.StartsWith("1")) { "1" } else { "2" }
    $expectedConformance = if ($effectivePdfaProfile.EndsWith("a")) { "A" } else { "B" }
    $actualConformance = Get-PdfaConformanceFromFile -PdfPath $outputPdf

    if (-not $actualConformance) {
        Write-Log "Unable to determine output PDF/A conformance from metadata."
        $debugHint = if ($savedAdvancedDebugPdf) { " Advanced debug PDF: $savedAdvancedDebugPdf" } else { "" }
        throw "Conversion completed, but PDF/A conformance metadata could not be read.$debugHint"
    }

    if ($actualConformance.Part -ne $expectedPart -or $actualConformance.Conformance -ne $expectedConformance) {
        Write-Log (
            "Conformance mismatch. Requested=PDF/A-{0}{1}, Produced=PDF/A-{2}{3}" -f
            $expectedPart,
            $expectedConformance.ToLowerInvariant(),
            $actualConformance.Part,
            $actualConformance.Conformance.ToLowerInvariant()
        )

        if ($ConversionMode -eq "advanced" -and $expectedConformance -eq "A" -and $actualConformance.Part -eq $expectedPart -and $actualConformance.Conformance -eq "B") {
            Write-Log "Advanced mode returned a valid PDF/A-b file after an a-level request. Keeping the produced output because Ghostscript did not preserve the tagged structure required for level a."
            return $outputPdf
        }

        $mismatchMessage = [string]::Format(
            "Requested PDF/A-{0}{1}, but Ghostscript produced PDF/A-{2}{3}. This usually means the source PDF lacks required tagging/structure for level 'a' or the selected mode cannot enforce it.",
            $expectedPart,
            $expectedConformance.ToLowerInvariant(),
            $actualConformance.Part,
            $actualConformance.Conformance.ToLowerInvariant()
        )
        if ($savedAdvancedDebugPdf) {
            $mismatchMessage = "{0} Advanced debug PDF: {1}" -f $mismatchMessage, $savedAdvancedDebugPdf
        }
        throw $mismatchMessage
    }

    Write-Log ("Conversion succeeded. Output size={0} bytes" -f $outputInfo.Length)

    return $outputPdf
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Width="760" Height="560" MinWidth="760" MinHeight="560"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        Background="#F4F7FB"
        FontFamily="Segoe UI"
        Title="QuickPDFA">
    <Window.Resources>
        <Style x:Key="RoundedPrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="#1E6FD9" />
            <Setter Property="Foreground" Value="#FFFFFF" />
            <Setter Property="BorderBrush" Value="#1E6FD9" />
            <Setter Property="BorderThickness" Value="1" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="12"
                                SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              RecognizesAccessKey="True" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.93" />
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.86" />
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="ButtonBorder" Property="Opacity" Value="0.55" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ModeChipStyle" TargetType="RadioButton">
            <Setter Property="Foreground" Value="#1B3555" />
            <Setter Property="Margin" Value="0,0,10,0" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="ChipBorder"
                                Background="#F3F7FD"
                                BorderBrush="#C7D5E8"
                                BorderThickness="1"
                                CornerRadius="14"
                                Padding="11,6"
                                SnapsToDevicePixels="True">
                            <ContentPresenter VerticalAlignment="Center" HorizontalAlignment="Center" />
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ChipBorder" Property="Background" Value="#E9F1FD" />
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="ChipBorder" Property="Background" Value="#1E6FD9" />
                                <Setter Property="Foreground" Value="#FFFFFF" />
                                <Setter TargetName="ChipBorder" Property="BorderBrush" Value="#1E6FD9" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="16" />
            <RowDefinition Height="120" />
            <RowDefinition Height="12" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="12" />
            <RowDefinition Height="34" />
            <RowDefinition Height="12" />
            <RowDefinition Height="42" />
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>

        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="200" />
                <ColumnDefinition Width="44" />
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="TitleText"
                       Grid.Column="0"
                       FontSize="24"
                       FontWeight="SemiBold"
                       Foreground="#0B1F3A"
                       VerticalAlignment="Center" />
            <Button x:Name="LanguageToggleButton"
                    Grid.Column="1"
                    Width="188"
                    Height="34"
                    HorizontalAlignment="Right"
                    FontSize="14"
                    Background="#FFFFFF"
                    BorderBrush="#D4DCE8"
                    Foreground="#0B1F3A"
                    Cursor="Hand">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                    <Border Width="24" Height="16" BorderBrush="#AEBFD6" BorderThickness="1" Margin="0,0,8,0">
                        <Canvas x:Name="LanguageFlagCanvas" Width="24" Height="16" />
                    </Border>
                    <TextBlock x:Name="LanguageToggleText"
                               VerticalAlignment="Center"
                               FontSize="13"
                               FontWeight="SemiBold"
                               Foreground="#0B1F3A" />
                </StackPanel>
            </Button>

            <Button x:Name="InfoButton"
                    Grid.Column="2"
                    Width="34"
                    Height="34"
                    HorizontalAlignment="Right"
                    VerticalAlignment="Center"
                    FontSize="16"
                    FontWeight="Bold"
                    Content="i"
                    Background="#FFFFFF"
                    BorderBrush="#D4DCE8"
                    Foreground="#0B1F3A"
                    Cursor="Hand" />
        </Grid>

        <Border x:Name="DropZone"
                Grid.Row="2"
                CornerRadius="14"
                BorderBrush="#BFCDE2"
                BorderThickness="1.6"
                Background="#EAF2FF"
                AllowDrop="True"
                Cursor="Hand">
            <Grid>
                <TextBlock x:Name="DropText"
                           HorizontalAlignment="Center"
                           VerticalAlignment="Center"
                           FontSize="17"
                           FontWeight="SemiBold"
                           Foreground="#1E3A5F" />
            </Grid>
        </Border>

        <Border Grid.Row="4"
                CornerRadius="10"
                BorderBrush="#C7D5E8"
                BorderThickness="1"
                Background="#FFFFFF"
                Padding="12,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="12" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="10" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="8" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="10" />
                    <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>

                <TextBlock x:Name="PdfaOptionsLabel"
                           FontSize="13"
                           FontWeight="SemiBold"
                           Foreground="#1B3555" />

                <Grid Grid.Row="2">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto" />
                        <ColumnDefinition Width="12" />
                        <ColumnDefinition Width="150" />
                    </Grid.ColumnDefinitions>

                    <TextBlock x:Name="DocumentLanguageLabel"
                               Grid.Column="0"
                               VerticalAlignment="Center"
                               FontSize="13"
                               FontWeight="SemiBold"
                               Foreground="#1B3555" />

                    <TextBox x:Name="DocumentLanguageTextBox"
                             Grid.Column="2"
                             Height="28"
                             VerticalContentAlignment="Center"
                             Padding="8,0"
                             FontSize="13"
                             Text="en-US" />
                </Grid>

                <StackPanel Grid.Row="4" Orientation="Horizontal">
                    <RadioButton x:Name="Pdfa2aOption"
                                 Margin="0,0,18,0"
                                 VerticalAlignment="Center"
                                 IsChecked="True"
                                 Foreground="#1B3555" />
                    <RadioButton x:Name="Pdfa2bOption"
                                 Margin="0,0,18,0"
                                 VerticalAlignment="Center"
                                 Foreground="#1B3555" />
                    <RadioButton x:Name="Pdfa1bOption"
                                 Margin="0,0,18,0"
                                 VerticalAlignment="Center"
                                 Foreground="#1B3555" />
                    <RadioButton x:Name="Pdfa1aOption"
                                 Margin="0,0,18,0"
                                 VerticalAlignment="Center"
                                 Foreground="#1B3555" />
                    <CheckBox x:Name="HighCompressionOption"
                              VerticalAlignment="Center"
                              Foreground="#1B3555" />
                </StackPanel>

                <Grid Grid.Row="8">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="8" />
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="10" />
                        <RowDefinition Height="Auto" />
                    </Grid.RowDefinitions>

                    <TextBlock x:Name="ModeLabel"
                               Grid.Row="0"
                               VerticalAlignment="Center"
                               FontSize="13"
                               FontWeight="SemiBold"
                               Foreground="#1B3555" />

                    <WrapPanel Grid.Row="2">
                        <RadioButton x:Name="ModeStrictOption"
                                     GroupName="ConversionMode"
                                     IsChecked="True"
                                     Style="{StaticResource ModeChipStyle}" />
                        <RadioButton x:Name="ModeBestEffortOption"
                                     GroupName="ConversionMode"
                                     Style="{StaticResource ModeChipStyle}" />
                        <RadioButton x:Name="ModeAdvancedOption"
                                     GroupName="ConversionMode"
                                     Style="{StaticResource ModeChipStyle}" />
                    </WrapPanel>

                    <Grid Grid.Row="4">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="Auto" />
                        </Grid.ColumnDefinitions>

                        <TextBlock x:Name="OcrStatusText"
                                   Grid.Column="0"
                                   VerticalAlignment="Center"
                                   FontSize="12"
                                   Foreground="#3E5F83"
                                   TextTrimming="CharacterEllipsis" />

                        <Button x:Name="SetupOcrButton"
                                Grid.Column="1"
                                Margin="10,0,0,0"
                                MinWidth="100"
                                Height="28"
                                FontSize="12"
                                Background="#FFFFFF"
                                BorderBrush="#C7D5E8"
                                Foreground="#1B3555"
                                Cursor="Hand" />
                    </Grid>
                </Grid>
            </Grid>
        </Border>

        <TextBox x:Name="InputBox"
                 Grid.Row="6"
                 Height="34"
                 IsReadOnly="True"
                 FontSize="13"
                 VerticalContentAlignment="Center"
                 Padding="10,0"
                 BorderBrush="#C4D1E3"
                 Background="#FFFFFF" />

        <Grid Grid.Row="8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="230" />
                <ColumnDefinition Width="14" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>

            <Button x:Name="ConvertButton"
                    Grid.Column="0"
                    Height="42"
                    FontSize="15"
                    FontWeight="SemiBold"
                    Style="{StaticResource RoundedPrimaryButton}"
                    Cursor="Hand"
                    Background="#1E6FD9"
                    Foreground="#FFFFFF"
                    BorderBrush="#1E6FD9"
                    IsEnabled="False" />

            <TextBlock x:Name="StatusText"
                       Grid.Column="2"
                       VerticalAlignment="Center"
                       FontSize="13"
                       TextWrapping="Wrap"
                       Foreground="#27486E" />
        </Grid>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$titleText = $window.FindName("TitleText")
$languageToggleButton = $window.FindName("LanguageToggleButton")
$languageToggleText = $window.FindName("LanguageToggleText")
$languageFlagCanvas = $window.FindName("LanguageFlagCanvas")
$infoButton = $window.FindName("InfoButton")
$dropZone = $window.FindName("DropZone")
$dropText = $window.FindName("DropText")
$pdfaOptionsLabel = $window.FindName("PdfaOptionsLabel")
$documentLanguageLabel = $window.FindName("DocumentLanguageLabel")
$documentLanguageTextBox = $window.FindName("DocumentLanguageTextBox")
$pdfa2aOption = $window.FindName("Pdfa2aOption")
$pdfa2bOption = $window.FindName("Pdfa2bOption")
$pdfa1bOption = $window.FindName("Pdfa1bOption")
$pdfa1aOption = $window.FindName("Pdfa1aOption")
$highCompressionOption = $window.FindName("HighCompressionOption")
$modeLabel = $window.FindName("ModeLabel")
$modeStrictOption = $window.FindName("ModeStrictOption")
$modeBestEffortOption = $window.FindName("ModeBestEffortOption")
$modeAdvancedOption = $window.FindName("ModeAdvancedOption")
$ocrStatusText = $window.FindName("OcrStatusText")
$setupOcrButton = $window.FindName("SetupOcrButton")
$inputBox = $window.FindName("InputBox")
$convertButton = $window.FindName("ConvertButton")
$statusText = $window.FindName("StatusText")

function Show-InfoDialog {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Window]$Owner
    )

    [xml]$infoXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="430"
        Height="220"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner"
        Background="#F7FAFE"
        Title="QuickPDFA Info"
        ShowInTaskbar="False">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="16" />
            <RowDefinition Height="44" />
        </Grid.RowDefinitions>

        <TextBlock x:Name="InfoText"
                   FontSize="16"
                   FontWeight="SemiBold"
                   Foreground="#0B1F3A"
                   TextWrapping="Wrap"
                   VerticalAlignment="Center"
                   Text="Happily vibe-coded by Matteo Tognolo" />

        <Button x:Name="GitHubButton"
                Grid.Row="2"
                Width="170"
                Height="40"
                HorizontalAlignment="Left"
                FontSize="14"
                FontWeight="SemiBold"
                Content="Open on GitHub"
                Background="#1E6FD9"
                Foreground="#FFFFFF"
                BorderBrush="#1E6FD9"
                Cursor="Hand" />
    </Grid>
</Window>
"@

    $dialogReader = New-Object System.Xml.XmlNodeReader $infoXaml
    $dialogWindow = [Windows.Markup.XamlReader]::Load($dialogReader)
    $dialogWindow.Owner = $Owner

    $gitHubButton = $dialogWindow.FindName("GitHubButton")
    $gitHubButton.Add_Click({
        Start-Process -FilePath "https://github.com/tognols/quickpdfa.git" | Out-Null
    })

    [void]$dialogWindow.ShowDialog()
}

function Set-LanguageFlag {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("it", "en")]
        [string]$FlagCode
    )

    $languageFlagCanvas.Children.Clear()

    if ($FlagCode -eq "it") {
        $left = New-Object System.Windows.Shapes.Rectangle
        $left.Width = 8
        $left.Height = 16
        $left.Fill = [System.Windows.Media.Brushes]::ForestGreen
        [System.Windows.Controls.Canvas]::SetLeft($left, 0)

        $middle = New-Object System.Windows.Shapes.Rectangle
        $middle.Width = 8
        $middle.Height = 16
        $middle.Fill = [System.Windows.Media.Brushes]::White
        [System.Windows.Controls.Canvas]::SetLeft($middle, 8)

        $right = New-Object System.Windows.Shapes.Rectangle
        $right.Width = 8
        $right.Height = 16
        $right.Fill = [System.Windows.Media.Brushes]::Crimson
        [System.Windows.Controls.Canvas]::SetLeft($right, 16)

        [void]$languageFlagCanvas.Children.Add($left)
        [void]$languageFlagCanvas.Children.Add($middle)
        [void]$languageFlagCanvas.Children.Add($right)
        return
    }

    # England flag: white field with centered red cross.
    $bg = New-Object System.Windows.Shapes.Rectangle
    $bg.Width = 24
    $bg.Height = 16
    $bg.Fill = [System.Windows.Media.Brushes]::White

    $vCross = New-Object System.Windows.Shapes.Rectangle
    $vCross.Width = 4
    $vCross.Height = 16
    $vCross.Fill = [System.Windows.Media.Brushes]::Crimson
    [System.Windows.Controls.Canvas]::SetLeft($vCross, 10)

    $hCross = New-Object System.Windows.Shapes.Rectangle
    $hCross.Width = 24
    $hCross.Height = 4
    $hCross.Fill = [System.Windows.Media.Brushes]::Crimson
    [System.Windows.Controls.Canvas]::SetTop($hCross, 6)

    [void]$languageFlagCanvas.Children.Add($bg)
    [void]$languageFlagCanvas.Children.Add($vCross)
    [void]$languageFlagCanvas.Children.Add($hCross)
}

function Update-OcrUiState {
    if (-not $ocrStatusText -or -not $setupOcrButton) {
        return
    }

    if ($script:ExternalTagger) {
        $ocrStatusText.Text = Get-Text "ocrStatusReady"
        $ocrStatusText.ToolTip = $script:ExternalTagger.DisplayName
    }
    else {
        $ocrStatusText.Text = Get-Text "ocrStatusMissing"
        $ocrStatusText.ToolTip = $null
    }
}

function Prompt-SetupExternalTagger {
    $answer = [System.Windows.MessageBox]::Show(
        (Get-Text "promptOcrSetupOptions"),
        (Get-Text "dialogOcrSetupTitle"),
        [System.Windows.MessageBoxButton]::YesNoCancel,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($answer -eq [System.Windows.MessageBoxResult]::Yes) {
        $openDialog = New-Object Microsoft.Win32.OpenFileDialog
        $openDialog.Title = Get-Text "dialogOpenTitle"
        $openDialog.Filter = "Executable files (*.exe)|*.exe|All files (*.*)|*.*"
        $openDialog.CheckFileExists = $true
        $openDialog.Multiselect = $false

        if ($openDialog.ShowDialog() -eq $true) {
            [Environment]::SetEnvironmentVariable("QUICKPDFA_TAGGER_EXE", $openDialog.FileName, "User")
            Write-Log ("External tagger configured by user: {0}" -f $openDialog.FileName)
            return
        }

        Write-Log "OCR executable selection canceled by user."
        return
    }

    if ($answer -eq [System.Windows.MessageBoxResult]::No) {
        Start-Process -FilePath "https://opendataloader.org/docs/quick-start-python" | Out-Null
        Write-Log "Opened OpenDataLoader installation guide."
    }
}

function Refresh-ExternalTagger {
    $script:ExternalTagger = Find-ExternalTagger
    if ($script:ExternalTagger) {
        Write-Log ("External tagger detected via {0}: '{1}'" -f $script:ExternalTagger.Source, $script:ExternalTagger.DisplayName)
    }
    else {
        Write-Log "No external OCR/tagging tool detected (advanced mode may fail)."
    }

    Update-OcrUiState
}

function Update-UiLanguage {
    $window.Title = Get-Text "appTitle"
    $titleText.Text = Get-Text "titleLabel"
    $dropText.Text = Get-Text "dropLabel"
    $pdfaOptionsLabel.Text = Get-Text "pdfaOptionsLabel"
    $documentLanguageLabel.Text = Get-Text "documentLanguageLabel"
    $pdfa2aOption.Content = Get-Text "pdfa2a"
    $pdfa2bOption.Content = Get-Text "pdfa2b"
    $pdfa1bOption.Content = Get-Text "pdfa1b"
    $pdfa1aOption.Content = Get-Text "pdfa1a"
    $highCompressionOption.Content = Get-Text "compressCheckbox"
    $modeLabel.Text = Get-Text "modeLabel"
    $modeStrictOption.Content = Get-Text "modeStrict"
    $modeBestEffortOption.Content = Get-Text "modeBestEffort"
    $modeAdvancedOption.Content = Get-Text "modeAdvanced"
    $setupOcrButton.Content = Get-Text "ocrSetupButton"
    $convertButton.Content = Get-Text "convertButton"

    if ($script:CurrentLanguage -eq "it") {
        Set-LanguageFlag -FlagCode "en"
        $languageToggleText.Text = "EN"
    }
    else {
        Set-LanguageFlag -FlagCode "it"
        $languageToggleText.Text = Get-Text "toggleButton"
    }

    switch ($script:CurrentStatus) {
        "checking" { $statusText.Text = Get-Text "statusCheckingGhostscript" }
        "ghostscriptMissing" { $statusText.Text = Get-Text "statusGhostscriptMissing" }
        "ready" { $statusText.Text = Get-Text "statusReady" }
        "converting" { $statusText.Text = Get-Text "statusConverting" }
        "saveCanceled" { $statusText.Text = Get-Text "statusSaveCanceled" }
        "conversionFailed" { $statusText.Text = Get-Text "statusConversionFailed" }
        default {
            if ($script:GhostscriptExe) {
                $statusText.Text = Get-Text "statusGhostscriptFound" @($script:GhostscriptExe)
            }
            else {
                $statusText.Text = Get-Text "statusGhostscriptMissing"
            }
        }
    }

    Update-OcrUiState
}

$setSelectedFile = {
    param([string]$path)

    if (-not $path) {
        return
    }

    if (([IO.Path]::GetExtension($path)).ToLowerInvariant() -ne ".pdf") {
        [System.Windows.MessageBox]::Show(
            (Get-Text "warningDropPdf"),
            (Get-Text "appTitle"),
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        return
    }

    $script:SelectedPdf = $path
    $inputBox.Text = $path
    $convertButton.IsEnabled = [bool]($script:SelectedPdf -and $script:GhostscriptExe)
    $script:CurrentStatus = "ready"
    $statusText.Text = Get-Text "statusReady"
}

$openFilePicker = {
    $openDialog = New-Object Microsoft.Win32.OpenFileDialog
    $openDialog.Title = Get-Text "dialogOpenTitle"
    $openDialog.Filter = "PDF files (*.pdf)|*.pdf"
    $openDialog.DefaultExt = "pdf"
    $openDialog.CheckFileExists = $true
    $openDialog.Multiselect = $false

    if ($openDialog.ShowDialog() -eq $true) {
        Write-Log ("File selected from Open dialog: {0}" -f $openDialog.FileName)
        & $setSelectedFile $openDialog.FileName
    }
    else {
        Write-Log "Open file dialog canceled by user."
    }
}

$languageToggleButton.Add_Click({
    if ($script:CurrentLanguage -eq "it") {
        $script:CurrentLanguage = "en"
    }
    else {
        $script:CurrentLanguage = "it"
    }

    Update-UiLanguage
})

$pdfa2aOption.Add_Checked({ $script:SelectedPdfaProfile = "2a" })
$pdfa2bOption.Add_Checked({ $script:SelectedPdfaProfile = "2b" })
$pdfa1bOption.Add_Checked({ $script:SelectedPdfaProfile = "1b" })
$pdfa1aOption.Add_Checked({ $script:SelectedPdfaProfile = "1a" })
$highCompressionOption.Add_Checked({ $script:EnableHighCompression = $true })
$highCompressionOption.Add_Unchecked({ $script:EnableHighCompression = $false })
$modeStrictOption.Add_Checked({ $script:SelectedConversionMode = "strict" })
$modeBestEffortOption.Add_Checked({ $script:SelectedConversionMode = "best-effort" })
$modeAdvancedOption.Add_Checked({ $script:SelectedConversionMode = "advanced" })

$setupOcrButton.Add_Click({
    Prompt-SetupExternalTagger
    Refresh-ExternalTagger
})

$infoButton.Add_Click({
    Show-InfoDialog -Owner $window
})

$dropZone.Add_MouseLeftButtonUp({
    & $openFilePicker
})

$dropText.Add_MouseLeftButtonUp({
    & $openFilePicker
})

$dropZone.Add_DragEnter({
    if ($_.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $files = [string[]]$_.Data.GetData([System.Windows.DataFormats]::FileDrop)
        if ($files.Count -eq 1 -and ([IO.Path]::GetExtension($files[0]).ToLowerInvariant() -eq ".pdf")) {
            $_.Effects = [System.Windows.DragDropEffects]::Copy
            $dropZone.Background = [System.Windows.Media.Brushes]::LightCyan
        }
        else {
            $_.Effects = [System.Windows.DragDropEffects]::None
        }
    }
    else {
        $_.Effects = [System.Windows.DragDropEffects]::None
    }
    $_.Handled = $true
})

$dropZone.Add_DragLeave({
    $dropZone.Background = [System.Windows.Media.Brushes]::AliceBlue
})

$dropZone.Add_Drop({
    $dropZone.Background = [System.Windows.Media.Brushes]::AliceBlue
    $files = [string[]]$_.Data.GetData([System.Windows.DataFormats]::FileDrop)
    if ($files.Count -ge 1) {
        & $setSelectedFile $files[0]
    }
    $_.Handled = $true
})

$convertButton.Add_Click({
    if (-not $script:SelectedPdf) {
        return
    }

    if (-not $script:GhostscriptExe) {
        Write-Log "Convert blocked: Ghostscript not found."
        Prompt-InstallGhostscript
        [System.Windows.MessageBox]::Show(
            (Get-Text "dialogGhostscriptMissing"),
            (Get-Text "appTitle"),
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        return
    }

    # Detect newly installed external taggers without requiring app restart.
    Refresh-ExternalTagger

    if ($script:SelectedConversionMode -eq "advanced" -and -not $script:ExternalTagger) {
        $advancedAnswer = [System.Windows.MessageBox]::Show(
            (Get-Text "promptAdvancedModeMissingOcr"),
            (Get-Text "dialogOcrSetupTitle"),
            [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Warning
        )

        if ($advancedAnswer -eq [System.Windows.MessageBoxResult]::Yes) {
            Prompt-SetupExternalTagger
            Refresh-ExternalTagger

            if (-not $script:ExternalTagger) {
                $fallbackAfterSetup = [System.Windows.MessageBox]::Show(
                    (Get-Text "promptFallbackAfterSetup"),
                    (Get-Text "appTitle"),
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Question
                )

                if ($fallbackAfterSetup -ne [System.Windows.MessageBoxResult]::Yes) {
                    return
                }

                $modeBestEffortOption.IsChecked = $true
            }
        }
        elseif ($advancedAnswer -eq [System.Windows.MessageBoxResult]::No) {
            $modeBestEffortOption.IsChecked = $true
        }
        else {
            return
        }
    }

    $defaultFolder = Split-Path -Path $script:SelectedPdf -Parent
    $defaultName = [IO.Path]::GetFileNameWithoutExtension($script:SelectedPdf)
    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Title = Get-Text "dialogSaveTitle"
    $saveDialog.Filter = "PDF files (*.pdf)|*.pdf"
    $saveDialog.DefaultExt = "pdf"
    $saveDialog.AddExtension = $true
    $saveDialog.OverwritePrompt = $true
    $saveDialog.InitialDirectory = $defaultFolder
    $saveDialog.FileName = ("{0}_PDFA.pdf" -f $defaultName)

    if ($saveDialog.ShowDialog() -ne $true) {
        $script:CurrentStatus = "saveCanceled"
        $statusText.Text = Get-Text "statusSaveCanceled"
        Write-Log "Save dialog canceled by user."
        return
    }

    $outputPdf = $saveDialog.FileName

    try {
        $convertButton.IsEnabled = $false
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        $script:CurrentStatus = "converting"
        $statusText.Text = Get-Text "statusConverting"

        $result = Convert-ToPdfA `
            -InputPdf $script:SelectedPdf `
            -OutputPdf $outputPdf `
            -GhostscriptExe $script:GhostscriptExe `
            -PdfaProfile $script:SelectedPdfaProfile `
            -EnableCompression $script:EnableHighCompression `
            -ConversionMode $script:SelectedConversionMode `
            -DocumentLanguage $documentLanguageTextBox.Text

        $script:CurrentStatus = "done"
        $statusText.Text = Get-Text "statusDone" @($result)
        [System.Windows.MessageBox]::Show(
            (Get-Text "dialogSuccessBody" @($result)),
            (Get-Text "appTitle"),
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
    }
    catch {
        $script:CurrentStatus = "conversionFailed"
        $statusText.Text = Get-Text "statusConversionFailed"
        [System.Windows.MessageBox]::Show(
            $_.Exception.Message,
            (Get-Text "appTitle"),
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    }
    finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        $convertButton.IsEnabled = [bool]($script:SelectedPdf -and $script:GhostscriptExe)
    }
})

Update-UiLanguage
$script:GhostscriptExe = Find-Ghostscript
Refresh-ExternalTagger

if ($script:GhostscriptExe) {
    $script:CurrentStatus = "ghostscriptFound"
    $statusText.Text = Get-Text "statusGhostscriptFound" @($script:GhostscriptExe)
    Write-Log ("Ghostscript detected at '{0}'" -f $script:GhostscriptExe)
}
else {
    $script:CurrentStatus = "ghostscriptMissing"
    $statusText.Text = Get-Text "statusGhostscriptMissing"
    Write-Log "Ghostscript not found during startup."
    Prompt-InstallGhostscript
}

[void]$window.ShowDialog()
