Param(
    [string]$GamePath = "",
    [switch]$Quiet = $false,
    [switch]$StartGame = $false,
    [switch]$OverrideUnknown = $false
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.IO.Compression.FileSystem
$null = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")


$FoldersToSkip = @("Among Us_Data", "BepInEx")
$FoldersToSkipBepInEx = @("config")
$FilesToSkip = @("Among Us.exe", "baselib.dll", "GameAssembly.dll", "UnityCrashHandler32.exe", "UnityPlayer.dll", "version.txt")

function ReadPathManual {
    $Result = [Microsoft.VisualBasic.Interaction]::InputBox("Couldn't autodetect the Among Us folder. $([System.Environment]::NewLine)Please enter it manually.", "Enter Gamepath")
    if ($Result -eq "") {
        return $Result
    }

    while (!(ValidateGamePath -GamePath $Result)) {
        $Result = [Microsoft.VisualBasic.Interaction]::InputBox("The given folder doesn't contain a working Among Us installation. Please try again.", "Wrong Input")
        if ($Result -eq "") {
            return $Result
        }
    }

    return $Result
}

function ValidateGamePath {
    param (
        $GamePath
    )
    return (Test-Path $GamePath) -And (Test-Path "$GamePath\Among Us.exe")
}

function GetGamePath {
    try {
        $SteamPath = (Get-ItemProperty -path 'HKLM:\SOFTWARE\WOW6432Node\Valve\Steam').InstallPath
    }
    catch [System.Management.Automation.ItemNotFoundException] {}

    try {
        if ($null -eq $SteamPath) {
            $SteamPath = (Get-ItemProperty -path 'HKLM:\SOFTWARE\Valve\Steam').InstallPath
        }
    }
    catch [System.Management.Automation.ItemNotFoundException] {}

    $GamePath = "$SteamPath\steamapps\common\Among Us"

    if ($null -eq $SteamPath -Or !(Test-Path $GamePath)) {
        if ($Quiet) {
            return ""
        }
        $GamePath = ReadPathManual
    }

    return $GamePath
}

function CleanUp {
    param (
        $GamePath,
        $IgnoreConfig = $False
    )

    $Content = Get-Item -Path $GamePath

    foreach ($it in $Content.GetDirectories()) {
        if ($FoldersToSkip.contains($it.Name)) { continue }
        $it.Delete($True)
    }

    foreach ($it in $Content.GetFiles()) {
        if ($FilesToSkip.contains($it.Name)) { continue }
        $it.Delete()
    }

    $BepInExPath = "$GamePath\BepInEx"

    if (Test-Path $BepInExPath) {

        if ($IgnoreConfig) {
            Remove-Item -Path $BepInExPath -Recurse
            return
        }

        $BepInExContent = Get-Item -Path $BepInExPath

        foreach ($it in $BepInExContent.GetDirectories()) {
            if ($FoldersToSkipBepInEx.contains($it.Name)) { continue }
            $it.Delete($True)
        }

        foreach ($it in $BepInExContent.GetFiles()) {
            $it.Delete()
        }
    }
}

function GetLatesVersion {
    $LatestRelease = Invoke-WebRequest https://github.com/Eisbison/TheOtherRoles/releases/latest -Headers @{"Accept" = "application/json" }
    return ($LatestRelease.Content | ConvertFrom-Json).tag_name
}

function StartGame {
    Start-Process "steam://rungameid/945360"
}

function DownloadAndExtractRelease {
    param (
        $GamePath,
        $Version
    )

    $url = "https://github.com/Eisbison/TheOtherRoles/releases/download/$Version/TheOtherRoles.zip"
    $download_path = "$GamePath\TheOtherRoles_$Version.zip"

    Invoke-WebRequest -Uri $url -OutFile $download_path

    Expand-Archive -Path $download_path -DestinationPath $GamePath -force

    $Version > "$GamePath\version.txt"

    Remove-Item -Path $download_path

    if (!($Quiet)) {
        $Result = [System.Windows.MessageBox]::Show("The Other Roles $LatestVersion is successfully installed. $([System.Environment]::NewLine)Start the game now?", "Successfully installed", 4)
        if ($Result -eq "Yes") { StartGame }
    }
    elseif ($StartGame) { StartGame }
}

if ($GamePath -eq "" -Or !(ValidateGamePath -GamePath $GamePath)) {
    $GamePath = GetGamePath
}

if ($GamePath -eq "") { exit 1 }

$LatestVersion = GetLatesVersion

if (!(Test-Path "$GamePath\version.txt")) {

    if (Test-Path "$GamePath\BepInEx") {

        if ($Quiet) {
            if (!($OverrideUnknown)) { exit 1 }
        }
        else {
            $Result = [System.Windows.MessageBox]::Show("An unknown installation was found. $([System.Environment]::NewLine)Clean and process?", "Unknown installation found", 4)

            if ($Result -eq "No") { exit 1 }
        }

        CleanUp -GamePath $GamePath -IgnoreConfig $true
    }
    else {
        CleanUp -GamePath $GamePath
    }

    DownloadAndExtractRelease -GamePath $GamePath -Version $LatestVersion
    exit 0
}

$InstalledVersion = Get-Content -Path "$GamePath\version.txt" -TotalCount 1

if ($LatestVersion -eq $InstalledVersion) {
    if (!($Quiet)) {
        $Result = [System.Windows.MessageBox]::Show("The latest Version ($LatestVersion) is already installed. $([System.Environment]::NewLine)Start the game now?", "Already up to date", 4)
        if ($Result -eq "Yes") { StartGame }
    }
    elseif ($StartGame) { StartGame }
    exit 0
}

CleanUp -GamePath $GamePath

DownloadAndExtractRelease -GamePath $GamePath -Version $LatestVersion
exit 0
