# Builds a folder you can carry to another machine and install Outlook Refined from, with the
# settings and the in-Outlook choices this machine is using.
#
#   powershell -ExecutionPolicy Bypass -File make-handoff.ps1
#
# No administrator rights needed; it only reads. The folder lands on the Desktop unless you pass
# -Out <folder>, and holds four files: the installer, the mod source, the settings, and a note
# with the one command to run on the other machine.
#
# The settings file carries your account addresses and labels, because that is what makes the
# other machine look like this one. Keep the folder somewhere private.

[CmdletBinding()]
param(
    [string]$Out = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'outlook-refined-handoff'),
    [string]$Windhawk = 'C:\Program Files\Windhawk'
)

$ErrorActionPreference = 'Stop'
$ModId = 'outlook-refined'
$Full = "local@$ModId"

function Fail($message) { Write-Host "error: $message" -ForegroundColor Red; exit 1 }
function Step($message) { Write-Host "  $message" }

$installer = Join-Path $PSScriptRoot 'install-mod.ps1'
$source = Join-Path (Split-Path -Parent $PSScriptRoot) 'mods\outlook-refined.wh.cpp'
if (-not (Test-Path -LiteralPath $installer)) { Fail "install-mod.ps1 is not beside this script in $PSScriptRoot" }
if (-not (Test-Path -LiteralPath $source)) { Fail "no built mod at $source. Run: node build.mjs" }

$iniPath = Join-Path $Windhawk 'windhawk.ini'
$registryKey = 'HKLM\SOFTWARE\Windhawk'
if (Test-Path -LiteralPath $iniPath) {
    $iniBytes = [IO.File]::ReadAllBytes($iniPath)
    if ($iniBytes.Length -ge 2 -and $iniBytes[0] -eq 0xFF -and $iniBytes[1] -eq 0xFE) {
        $ini = [Text.Encoding]::Unicode.GetString($iniBytes)
    } else {
        $ini = [Text.Encoding]::UTF8.GetString($iniBytes)
    }
    $m = [regex]::Match($ini, '(?m)^RegistryKey=(.+?)\s*$')
    if ($m.Success) { $registryKey = $m.Groups[1].Value.Trim() }
}
$regSub = $registryKey -replace '^HKLM\\', ''

$text = [IO.File]::ReadAllText($source, (New-Object Text.UTF8Encoding $false))
$version = ([regex]::Match($text, '(?m)^\s*//\s*@version\s+(\S+)\s*$')).Groups[1].Value
if (-not $version) { Fail "no @version in $source" }

# What this machine has in the registry now. Both halves are read as they are: the settings are
# what the Windhawk settings page wrote, and LocalStorage is what you chose inside Outlook (folds,
# pane widths, favourites order).
$hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64')
$settings = @()
$key = $hklm.OpenSubKey("$regSub\Engine\Mods\$Full\Settings")
if ($key) {
    foreach ($name in $key.GetValueNames()) {
        $kind = 'String'
        if ($key.GetValueKind($name) -eq [Microsoft.Win32.RegistryValueKind]::DWord) { $kind = 'DWord' }
        $settings += [ordered]@{ name = $name; kind = $kind; value = $key.GetValue($name) }
    }
    $key.Close()
}
$storage = @()
$writable = $hklm.OpenSubKey("$regSub\Engine\ModsWritable\$Full\LocalStorage")
if ($writable) {
    foreach ($name in $writable.GetValueNames()) {
        $storage += [ordered]@{ name = $name; value = [string]$writable.GetValue($name) }
    }
    $writable.Close()
}
if ($settings.Count -eq 0) { Write-Host 'note: this machine has no saved settings for the mod, so only the mod itself travels.' }

$dir = [IO.Directory]::CreateDirectory($Out).FullName
Copy-Item -LiteralPath $installer -Destination (Join-Path $dir 'install-mod.ps1') -Force
Copy-Item -LiteralPath $source -Destination (Join-Path $dir 'outlook-refined.wh.cpp') -Force
Step "install-mod.ps1 and outlook-refined.wh.cpp ($ModId $version)"

$payload = [ordered]@{
    mod      = $Full
    version  = $version
    exported = (Get-Date).ToString('s')
    from     = $env:COMPUTERNAME
    settings = $settings
    storage  = $storage
}
$json = $payload | ConvertTo-Json -Depth 6
[IO.File]::WriteAllText((Join-Path $dir 'settings.json'), $json, (New-Object Text.UTF8Encoding $false))
Step "settings.json ($($settings.Count) settings, $($storage.Count) saved choices)"

$note = @"
Outlook Refined $version, from $env:COMPUTERNAME on $(Get-Date -Format 'yyyy-MM-dd').

To install it on this machine:

1. Install Windhawk from https://windhawk.net if it is not already there.
2. Open PowerShell as administrator.
3. Run, from this folder:

     powershell -ExecutionPolicy Bypass -File install-mod.ps1 -Source .\outlook-refined.wh.cpp -Settings .\settings.json

4. Close every Outlook window and start Outlook again.

The mod then appears in Windhawk under Installed mods, with the same settings as the machine it
came from. To check the compile without installing anything, add -DryRun and drop the admin
PowerShell; that writes nothing.

settings.json holds your account addresses and labels, so keep this folder private.
"@
[IO.File]::WriteAllText((Join-Path $dir 'README.txt'), $note, (New-Object Text.UTF8Encoding $false))
Step 'README.txt'

Write-Host ''
Write-Host "Ready: $dir"
Write-Host 'Carry that folder to the other machine, then follow README.txt in it.'
