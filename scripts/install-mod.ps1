# Installs Outlook Refined into Windhawk without touching the mod editor.
#
# Run it from an elevated PowerShell on the machine that should get the mod:
#
#   powershell -ExecutionPolicy Bypass -File install-mod.ps1
#
# With no arguments it takes the mod source from this repository when the script is sitting in it,
# and otherwise downloads the published copy from GitHub. Pass -Settings <json> to restore the
# settings and the in-Outlook choices that make-handoff.ps1 exported from another machine, and
# -DryRun to compile and report without writing anything (that one needs no administrator rights).
#
# It leaves Windhawk the way its own Compile Mod button does: the DLL under a fresh name in
# Engine\Mods, the source in ModsSource, and the registry entry under
# <RegistryKey>\Engine\Mods\local@outlook-refined with LibraryFileName written last, because the
# engine watches that value and loads the new file when it changes. Older builds of this mod are
# deleted, except one a running Outlook still holds, which is reported instead.
#
# Everything but the compiler flags is read from the target machine. The engine version, the data
# folder and the registry root come from windhawk.ini; the include list, the architecture and the
# link options come from the mod's own metadata block. The compile runs elevated with the rest,
# which is acceptable only because the source is checked first (@id, @include, @architecture, no
# byte-order mark). Keep that order if you edit this script.

[CmdletBinding()]
param(
    [string]$Source = '',
    [string]$Settings = '',
    [string]$Windhawk = 'C:\Program Files\Windhawk',
    [string]$Ref = 'main',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
$ModId = 'outlook-refined'
$Full = "local@$ModId"
$RawUrl = "https://raw.githubusercontent.com/ae-reid-havens/outlook-refined/$Ref/mods/outlook-refined.wh.cpp"

function Fail($message) { Write-Host "error: $message" -ForegroundColor Red; exit 1 }
function Step($message) { Write-Host "  $message" }

function IniField($text, $key) {
    $m = [regex]::Match($text, "(?m)^$key=(.+?)\s*$")
    if (-not $m.Success) { Fail "windhawk.ini names no $key" }
    return $m.Groups[1].Value.Trim()
}

function MetaField($meta, $name) {
    return ([regex]::Match($meta, "(?m)^\s*//\s*@$name\s+(.+?)\s*$")).Groups[1].Value
}

function MetaList($meta, $name) {
    $values = @([regex]::Matches($meta, "(?m)^\s*//\s*@$name\s+(\S+)\s*$") | ForEach-Object { $_.Groups[1].Value })
    return ($values -join '|')
}

Write-Host "Outlook Refined into Windhawk"

# 1. Administrator rights, unless this is a dry run.
if (-not $DryRun) {
    $me = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $me.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Fail 'this needs an elevated PowerShell. Open one as administrator and run it again, or add -DryRun to compile only.'
    }
}

# 2. Windhawk itself, and where it keeps things on this machine.
if (-not (Test-Path -LiteralPath (Join-Path $Windhawk 'windhawk.exe'))) {
    Fail "Windhawk is not installed at $Windhawk. Get it from https://windhawk.net, install it, then run this again (or pass -Windhawk <path>)."
}
$iniBytes = [IO.File]::ReadAllBytes((Join-Path $Windhawk 'windhawk.ini'))
if ($iniBytes.Length -ge 2 -and $iniBytes[0] -eq 0xFF -and $iniBytes[1] -eq 0xFE) {
    $ini = [Text.Encoding]::Unicode.GetString($iniBytes)
} else {
    $ini = [Text.Encoding]::UTF8.GetString($iniBytes)
}
$enginePath = IniField $ini 'EnginePath'
$appData = [Environment]::ExpandEnvironmentVariables((IniField $ini 'AppDataPath'))
$registryKey = IniField $ini 'RegistryKey'
if ($registryKey -notmatch '^HKLM\\') { Fail "windhawk.ini puts its registry key at $registryKey, which this script does not handle" }
$regSub = $registryKey -replace '^HKLM\\', ''
Step "Windhawk $enginePath, data in $appData"

# 3. The mod source: an explicit path, this repository, or the published copy.
if (-not $Source) {
    $inRepo = Join-Path (Split-Path -Parent $PSScriptRoot) 'mods\outlook-refined.wh.cpp'
    if (Test-Path -LiteralPath $inRepo) { $Source = $inRepo }
    else {
        $beside = Join-Path $PSScriptRoot 'outlook-refined.wh.cpp'
        if (Test-Path -LiteralPath $beside) { $Source = $beside }
    }
}
$work = Join-Path $env:LOCALAPPDATA 'outlook-refined\windhawk-build'
Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
[void][IO.Directory]::CreateDirectory($work)
if ($Source) {
    if (-not (Test-Path -LiteralPath $Source)) { Fail "no mod source at $Source" }
    $srcBytes = [IO.File]::ReadAllBytes($Source)
    Step "source: $Source"
} else {
    $download = Join-Path $work 'downloaded.wh.cpp'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $RawUrl -OutFile $download -UseBasicParsing
    $srcBytes = [IO.File]::ReadAllBytes($download)
    Step "source: $RawUrl"
}

# 4. Check it before the compiler sees it. A byte-order mark alone stops Windhawk parsing it.
if ($srcBytes.Length -ge 3 -and $srcBytes[0] -eq 0xEF -and $srcBytes[1] -eq 0xBB -and $srcBytes[2] -eq 0xBF) {
    Fail 'the source starts with a UTF-8 byte-order mark, which hides the metadata block from Windhawk'
}
$text = (New-Object Text.UTF8Encoding $false).GetString($srcBytes)
$block = [regex]::Match($text, '(?s)// ==WindhawkMod==(.*?)// ==/WindhawkMod==')
if (-not $block.Success) { Fail 'the source has no // ==WindhawkMod== block' }
$meta = $block.Groups[1].Value
$id = MetaField $meta 'id'
if ($id -cne $ModId) { Fail "this source is mod id '$id', not $ModId" }
$version = MetaField $meta 'version'
if ($version -cnotmatch '^[0-9A-Za-z.\-]{1,32}$') { Fail "version '$version' is not usable" }
$include = MetaList $meta 'include'
if (-not $include) { Fail 'the source names no @include, so it would hook every process' }
if ($include -notmatch '(?i)(^|\|)olk\.exe(\||$)') { Fail "the source includes $include, and this installer is for olk.exe only" }
$exclude = MetaList $meta 'exclude'
$architecture = MetaField $meta 'architecture'
$compilerOptions = @((MetaField $meta 'compilerOptions') -split '\s+' | Where-Object { $_ })
$archList = @($architecture -split '[\s|]+' | Where-Object { $_ })
$targets = @()
if ($archList.Count -eq 0 -or $archList -contains 'x86') { $targets += , @('32', 'i686-w64-mingw32') }
if ($archList.Count -eq 0 -or $archList -contains 'x86-64') { $targets += , @('64', 'x86_64-w64-mingw32') }
if ($targets.Count -eq 0) { Fail "the source asks for architecture '$architecture', which this installer does not know" }
Step "$ModId $version, include $include, architecture $(if ($architecture) { $architecture } else { 'any' })"

# 4b. The settings defaults, from the mod's own metadata block. Windhawk's API returns zero for an
#     integer setting and an empty string for a string setting the registry does not hold (see
#     Wh_GetIntSetting in windhawk_api.h), so a mod installed with no settings at all comes up with
#     every switch off, no pane width and no headings. The settings page writes these when you open
#     it; this writes them now so a fresh install behaves the way the mod documents.
#     Only the plain values are taken. A name with nothing after the colon introduces a list of
#     records, and a [..] default is a list as well; the mod reads a row that is not there as the
#     end of the list, which is exactly what an empty list should look like.
$defaults = [ordered]@{}
$settingsBlock = [regex]::Match($text, '(?s)// ==WindhawkModSettings==\s*/\*(.*?)\*/')
if ($settingsBlock.Success) {
    foreach ($line in ($settingsBlock.Groups[1].Value -split "`n")) {
        $entry = [regex]::Match($line.TrimEnd("`r"), '^- ([A-Za-z_][A-Za-z0-9_]{0,63}):\s*(.*?)\s*$')
        if (-not $entry.Success) { continue }
        $name = $entry.Groups[1].Value
        $value = $entry.Groups[2].Value
        if (-not $value -or $value.StartsWith('[')) { continue }
        if ($value -ceq 'true' -or $value -ceq 'false') { $defaults[$name] = @('DWord', [int]($value -ceq 'true')) }
        elseif ($value -match '^\d{1,9}$') { $defaults[$name] = @('DWord', [int]$value) }
        else { $defaults[$name] = @('String', ($value -replace '^"(.*)"$', '$1')) }
    }
}
Step "$($defaults.Count) settings defaults read from the source"

# 4c. The settings file, read before anything is built so that -DryRun checks it too.
$saved = $null
if ($Settings) {
    if (-not (Test-Path -LiteralPath $Settings)) { Fail "no settings file at $Settings" }
    $saved = Get-Content -LiteralPath $Settings -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($saved.mod -ne $Full) { Fail "that settings file is for $($saved.mod), not $Full" }
    if (-not $saved.settings) { Fail "$Settings carries no settings, so there is nothing to restore" }
    Step "settings file: $(@($saved.settings).Count) settings and $(@($saved.storage).Count) saved choices, exported from $($saved.from)"
}

# 5. Compile, with Windhawk's own clang and the flags its Compile Mod button uses.
$clang = Join-Path $Windhawk 'Compiler\bin\clang++.exe'
if (-not (Test-Path -LiteralPath $clang)) { Fail "Windhawk's compiler is missing from $clang" }
$srcFile = Join-Path $work 'mod.wh.cpp'
[IO.File]::WriteAllBytes($srcFile, $srcBytes)
Push-Location (Join-Path $Windhawk 'Compiler')
try {
    foreach ($target in $targets) {
        $sub = $target[0]
        $triple = $target[1]
        $lib = Join-Path $Windhawk "$enginePath\$sub\windhawk.lib"
        if (-not (Test-Path -LiteralPath $lib)) { Fail "no engine library at $lib" }
        $dllOut = Join-Path $work "$sub\mod.dll"
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $dllOut))
        $clangArgs = @(
            '-std=c++23', '-O2', '-shared', '-DUNICODE', '-D_UNICODE',
            '-DWINVER=0x0A00', '-D_WIN32_WINNT=0x0A00', '-D_WIN32_IE=0x0A00', '-DNTDDI_VERSION=0x0A000008',
            '-D__USE_MINGW_ANSI_STDIO=0', '-DWH_MOD',
            "-DWH_MOD_ID=L`"$Full`"", "-DWH_MOD_VERSION=L`"$version`"",
            $lib, '-x', 'c++', $srcFile, '-include', 'windhawk_api.h',
            '-target', $triple, '-Wl,--export-all-symbols'
        ) + $compilerOptions + @('-o', $dllOut)
        # Stop on error has to come off for this one call. Anything clang writes to stderr, a
        # warning included, reaches PowerShell as a NativeCommandError, which would abort a build
        # that was going to succeed and would hide a real compile error behind one truncated line.
        $stopping = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $log = & $clang @clangArgs 2>&1 | Out-String
        $compiled = $LASTEXITCODE
        $ErrorActionPreference = $stopping
        if ($compiled -ne 0) { Fail "clang++ ($triple) failed:`n$log" }
        if ($log.Trim()) { Write-Host $log.Trim() }
        Step "built $sub\mod.dll ($([int]((Get-Item -LiteralPath $dllOut).Length / 1KB)) KB)"
    }
} finally { Pop-Location }

if ($DryRun) {
    Write-Host ''
    Write-Host "Dry run. Compiled into $work and wrote nothing."
    Write-Host "Would set Include=$include Exclude=$exclude Architecture=$architecture Version=$version under HKLM\$regSub\Engine\Mods\$Full"
    Write-Host "Would fill in any missing settings from: $($defaults.Keys -join ', ')"
    exit 0
}

# 6. Copy the finished files in under a name no earlier build used.
$mods = Join-Path $appData 'Engine\Mods'
do {
    $dllName = "${Full}_${version}_$(Get-Random -Minimum 100000 -Maximum 1000000).dll"
    $taken = $false
    foreach ($target in $targets) { if (Test-Path -LiteralPath (Join-Path $mods "$($target[0])\$dllName")) { $taken = $true } }
} while ($taken)
foreach ($target in $targets) {
    $sub = $target[0]
    $dir = [IO.Directory]::CreateDirectory((Join-Path $mods $sub)).FullName
    Copy-Item -LiteralPath (Join-Path $work "$sub\mod.dll") -Destination (Join-Path $dir $dllName) -Force
    Step "installed $sub\$dllName"
}
$modsSource = [IO.Directory]::CreateDirectory((Join-Path $appData 'ModsSource')).FullName
[IO.File]::WriteAllBytes((Join-Path $modsSource "$Full.wh.cpp"), $srcBytes)

# 7. The registry entry. LibraryFileName goes last, after the settings are in place, so the engine
#    never sees a new library alongside settings that have not landed yet.
$hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry64')
$key = $hklm.CreateSubKey("$regSub\Engine\Mods\$Full")
$fresh = $null -eq $key.GetValue('LibraryFileName')
if ($null -eq $key.GetValue('Disabled')) {
    $key.SetValue('Disabled', 0, 'DWord')
    $key.SetValue('LoggingEnabled', 0, 'DWord')
}
$key.SetValue('Include', $include, 'String')
$key.SetValue('Exclude', $exclude, 'String')
$key.SetValue('Architecture', $architecture, 'String')
$key.SetValue('Version', $version, 'String')

# A settings file replaces the settings outright, and the defaults then fill in anything it does
# not mention, such as a setting added since it was exported. Without a settings file nothing
# already in the registry is touched, so running this over an existing install keeps its settings.
#
# The replacement has to clear the key first. List settings live as indexed value names
# (accountLabels[0].account and so on) and AddList in the mod skips an index that is missing rather
# than stopping there, so restoring two rows over a machine that had four would leave rows 2 and 3
# of the old set live.
$restored = 0
$storeCount = 0
if ($saved) { $key.DeleteSubKeyTree('Settings', $false) }
$settingsKey = $key.CreateSubKey('Settings')
if ($saved) {
    foreach ($entry in $saved.settings) {
        if ($entry.kind -eq 'DWord') { $settingsKey.SetValue($entry.name, [int]$entry.value, 'DWord') }
        else { $settingsKey.SetValue($entry.name, [string]$entry.value, 'String') }
        $restored++
    }
    if ($saved.storage) {
        $storeKey = $hklm.CreateSubKey("$regSub\Engine\ModsWritable\$Full\LocalStorage")
        foreach ($entry in $saved.storage) { $storeKey.SetValue($entry.name, [string]$entry.value, 'String'); $storeCount++ }
        $storeKey.Close()
    }
    Step "restored $restored settings and $storeCount saved choices from $(Split-Path -Leaf $Settings)"
}
$added = @()
foreach ($name in $defaults.Keys) {
    if ($null -eq $settingsKey.GetValue($name)) {
        $settingsKey.SetValue($name, $defaults[$name][1], $defaults[$name][0])
        $added += $name
    }
}
$settingsKey.Close()
if ($added.Count) { Step "filled in $($added.Count) settings at their defaults: $($added -join ', ')" }
if ($restored -or $added.Count) {
    $key.SetValue('SettingsChangeTime', [int]([DateTimeOffset]::UtcNow.ToUnixTimeSeconds() -band 0x7fffffff), 'DWord')
}

$key.SetValue('LibraryFileName', $dllName, 'String')
$disabled = 1 -eq $key.GetValue('Disabled')
$key.Close()

# 8. Earlier builds of this mod, except one a running Outlook still has open.
foreach ($target in $targets) {
    $dir = Join-Path $mods $target[0]
    foreach ($file in [IO.Directory]::GetFiles($dir, "${Full}_*.dll")) {
        $name = [IO.Path]::GetFileName($file)
        if ($name -ne $dllName -and $name -match '_[0-9]+\.dll$') {
            try { [IO.File]::Delete($file) } catch { Step "kept $name, a running Outlook still has it open" }
        }
    }
}

Write-Host ''
Write-Host "$ModId $version is installed$(if ($fresh) { '' } else { ' (updated in place)' })."
if ($disabled) {
    Write-Host 'This machine has the mod turned off, and that was left alone. Turn it on in Windhawk to use it.' -ForegroundColor Yellow
}
Write-Host 'Close every Outlook window and start Outlook again to load it.'
Write-Host 'It shows up in Windhawk under Installed mods, where you can change its settings or turn it off.'
