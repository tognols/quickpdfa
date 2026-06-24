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

$script:Translations = @{
    en = @{
        appTitle = "QuickPDFA"
        titleLabel = "Drop a PDF, then click Convert"
        dropLabel = "Drag and drop one PDF file here"
        convertButton = "Convert to PDF/A"
        toggleButton = "IT"
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

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="720" Height="390" MinWidth="720" MinHeight="390"
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
    </Window.Resources>
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="18" />
            <RowDefinition Height="120" />
            <RowDefinition Height="16" />
            <RowDefinition Height="34" />
            <RowDefinition Height="20" />
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

        <TextBox x:Name="InputBox"
                 Grid.Row="4"
                 Height="34"
                 IsReadOnly="True"
                 FontSize="13"
                 VerticalContentAlignment="Center"
                 Padding="10,0"
                 BorderBrush="#C4D1E3"
                 Background="#FFFFFF" />

        <Grid Grid.Row="6">
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

function Update-UiLanguage {
    $window.Title = Get-Text "appTitle"
    $titleText.Text = Get-Text "titleLabel"
    $dropText.Text = Get-Text "dropLabel"
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

        $result = Convert-ToPdfA -InputPdf $script:SelectedPdf -OutputPdf $outputPdf -GhostscriptExe $script:GhostscriptExe

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
