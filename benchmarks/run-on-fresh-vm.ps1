<#
.SYNOPSIS
    Runs the SheetStream benchmark suite on a clean Windows machine and collects everything needed to reproduce
    and publish the numbers.

.DESCRIPTION
    Needs only Windows and the .NET 10 SDK (plus Git if you want the script to download the code for you).
    Double-click run-on-fresh-vm.cmd, or run this script from PowerShell. With no options it does this:

      1. Shows the hardware (CPU, cores, memory), the operating system, the .NET SDK and the package versions.
      2. Builds the benchmark project in Release.
      3. Runs the correctness check (--verify). If any library fails it, the script stops: no timing is
         worth publishing for a library that did not write the data.
      4. Runs the benchmark scenarios, one after the other.
      5. Runs the peak-memory measurement (--peak) for the same scenarios.
      6. Copies the BenchmarkDotNet reports and a hardware.txt into a new, time-stamped folder
         (benchmarks\results\<date>-<time>-<machine>) so a run can be kept and compared with another.

    Use -DryRun first to see exactly what would happen without running anything.

    A full run of the default scenarios takes roughly an hour. Do not use the machine for anything else while it
    runs: other programs change the timings.

.PARAMETER Scenarios
    Which scenarios to run, separated by commas. Default: SpreadCheetahOwn, StringHeavy, StyledDates, Long, Wide.
    Known names: SpreadCheetahOwn, StringHeavy, StyledDates, Long, Wide, Million, VeryWide, Read, or All.
    Million, VeryWide and Read are not in the default set: the first two need 16 GB of memory and take a long time.

.PARAMETER DryRun
    Only print what would be done. Nothing is built, run, copied or created (the hardware is still read and shown).

.PARAMETER SkipPeak
    Skip the peak-memory measurement (--peak).

.PARAMETER RepoUrl
    Only needed when this script is NOT inside a checkout of the SheetStream.NET repository (for example when you
    copied just this file to a new machine). The repository is cloned into a folder next to the script. Needs Git.

.PARAMETER NoPause
    Do not wait for a key press at the end (for unattended runs). By default the window stays open so the result
    can be read.

.EXAMPLE
    .\run-on-fresh-vm.ps1 -DryRun

.EXAMPLE
    .\run-on-fresh-vm.ps1 -Scenarios SpreadCheetahOwn,Long
#>
[CmdletBinding()]
param(
    [string[]]$Scenarios = @('SpreadCheetahOwn', 'StringHeavy', 'StyledDates', 'Long', 'Wide'),
    [switch]$DryRun,
    [switch]$SkipPeak,
    [string]$RepoUrl = '',
    [switch]$NoPause
)

# Written for Windows PowerShell 5.1 as well as PowerShell 7, because a clean Windows machine only has 5.1.
# Native programs (dotnet) write progress to stderr; "Stop" would turn that into errors, so errors are checked
# by hand through the exit codes instead.
$ErrorActionPreference = 'Continue'

# The scenario names this script knows, the BenchmarkDotNet class that runs each one, and what to tell the user.
$KnownScenarios = [ordered]@{
    'SpreadCheetahOwn' = 'SpreadCheetahOwnBenchmarks'
    'StringHeavy'      = 'StringHeavyBenchmarks'
    'StyledDates'      = 'StyledDatesBenchmarks'
    'Long'             = 'LongBenchmarks'
    'Wide'             = 'WideBenchmarks'
    'Million'          = 'MillionBenchmarks'
    'VeryWide'         = 'VeryWideBenchmarks'
    'Read'             = 'ReadBenchmarks'
}

function Say([string]$Text, [string]$Color = 'Gray') { Write-Host $Text -ForegroundColor $Color }
function Heading([string]$Text) { Write-Host ''; Write-Host "== $Text" -ForegroundColor Cyan }

function Wait-ForKey {
    if ($NoPause) { return }
    Write-Host ''
    Write-Host 'Press Enter to close this window.' -ForegroundColor Yellow
    [void](Read-Host)
}

# Stops the script with a message a person can act on, after keeping the window open long enough to read it.
function Stop-WithMessage([string]$What, [string]$Next) {
    Write-Host ''
    Write-Host "STOPPED: $What" -ForegroundColor Red
    Write-Host "What to do: $Next" -ForegroundColor Yellow
    if ($script:Transcribing) { try { Stop-Transcript | Out-Null } catch { } }
    Wait-ForKey
    exit 1
}

# Runs a dotnet command, shows its output and keeps a copy in a log file. In a dry run it only prints the command.
function Invoke-Dotnet([string]$Label, [string[]]$Arguments, [string]$LogFile) {
    $shown = 'dotnet ' + ($Arguments -join ' ')
    if ($DryRun) { Say "  Would run: $shown" 'DarkYellow'; return 0 }
    Say "  Running: $shown" 'Gray'
    # Strings only (Windows PowerShell 5.1 wraps stderr lines in error records), shown and saved at once.
    # Out-Host keeps the output out of the return value, which must be just the exit code.
    & dotnet @Arguments 2>&1 | ForEach-Object { "$_" } | Tee-Object -FilePath $LogFile | Out-Host
    return $LASTEXITCODE
}

Say 'SheetStream benchmarks on a clean machine' 'White'
if ($DryRun) { Say 'DRY RUN: nothing will be built, run, copied or created. Remove -DryRun to do it for real.' 'Yellow' }

# ---------------------------------------------------------------- which scenarios
$requested = @()
foreach ($item in $Scenarios) { foreach ($part in ($item -split '[,;\s]+')) { if ($part) { $requested += $part } } }
if ($requested.Count -eq 0) { $requested = @('SpreadCheetahOwn', 'StringHeavy', 'StyledDates', 'Long', 'Wide') }
if ($requested -contains 'All') { $requested = @($KnownScenarios.Keys) }
$selected = @()
foreach ($name in $requested) {
    $match = $KnownScenarios.Keys | Where-Object { $_ -ieq $name } | Select-Object -First 1
    if (-not $match) {
        Stop-WithMessage "'$name' is not a scenario this script knows." ("Use one or more of: " + ($KnownScenarios.Keys -join ', ') + ', or All. Example: -Scenarios SpreadCheetahOwn,Long')
    }
    if ($selected -notcontains $match) { $selected += $match }
}

# ---------------------------------------------------------------- find (or fetch) the code
Heading 'Where the code is'
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
$projectDir = Join-Path $scriptDir 'SheetStream.Benchmarks'
if (-not (Test-Path (Join-Path $projectDir 'SheetStream.Benchmarks.csproj'))) {
    if (-not $RepoUrl) {
        Stop-WithMessage 'This script is not next to the SheetStream.Benchmarks folder, so there is no code to run.' `
            'Run it from the benchmarks folder of the SheetStream.NET repository, or pass -RepoUrl <address of the repository> to download it.'
    }
    $cloneDir = Join-Path $scriptDir 'SheetStream.NET'
    if ($DryRun) {
        Say "  Would download $RepoUrl into $cloneDir (needs Git)." 'DarkYellow'
        $projectDir = Join-Path $cloneDir 'benchmarks\SheetStream.Benchmarks'
    }
    else {
        if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
            Stop-WithMessage 'Git is not installed, so the code cannot be downloaded.' 'Install Git for Windows, or copy the whole SheetStream.NET folder to this machine and run the script from its benchmarks folder.'
        }
        if (-not (Test-Path $cloneDir)) {
            Say "  Downloading $RepoUrl ..."
            & git clone $RepoUrl $cloneDir 2>&1 | ForEach-Object { Say "  $_" }
            if ($LASTEXITCODE -ne 0) { Stop-WithMessage 'The download failed.' 'Check the address and the internet connection, then run the script again.' }
        }
        else { Say "  Using the existing download in $cloneDir" }
        $projectDir = Join-Path $cloneDir 'benchmarks\SheetStream.Benchmarks'
    }
    $benchRoot = Split-Path $projectDir -Parent
}
else {
    $benchRoot = $scriptDir
    Say "  Using the code in $benchRoot"
}
$repoRoot = Split-Path $benchRoot -Parent

# BenchmarkDotNet refuses to run when two projects with the same name are visible below the repository.
$worktrees = Join-Path $repoRoot '.claude\worktrees'
if ((Test-Path $worktrees) -and (Get-ChildItem $worktrees -Directory -ErrorAction SilentlyContinue)) {
    Stop-WithMessage "There are extra working copies in $worktrees. BenchmarkDotNet stops with a 'duplicate project' error when it finds them." `
        "Remove the folders inside $worktrees (they are temporary copies made by tools) and run the script again."
}

# ---------------------------------------------------------------- the machine
Heading 'This machine'
$cpu = $null; $os = $null; $mem = $null
try { $cpu = @(Get-CimInstance Win32_Processor)[0]; $os = Get-CimInstance Win32_OperatingSystem; $mem = Get-CimInstance Win32_ComputerSystem } catch { }
$cpuName = 'unknown'; $physical = 0; $logical = 0; $ramGb = 'unknown'; $osName = [Environment]::OSVersion.VersionString
if ($cpu) {
    $cpuName = ($cpu.Name -replace '\s+', ' ').Trim()
    $physical = (Get-CimInstance Win32_Processor | Measure-Object NumberOfCores -Sum).Sum
    $logical = (Get-CimInstance Win32_Processor | Measure-Object NumberOfLogicalProcessors -Sum).Sum
}
if ($mem) { $ramGb = [math]::Round($mem.TotalPhysicalMemory / 1GB, 1) }
if ($os) { $osName = "$($os.Caption) (build $($os.BuildNumber))" }
$vmHint = ''
if ($mem -and $mem.Model) { $vmHint = $mem.Manufacturer + ' ' + $mem.Model }
$powerPlan = ''
try { $powerPlan = ((powercfg /getactivescheme) -join ' ') -replace '^.*\((.*)\).*$', '$1' } catch { }

$sdk = ''
try { $sdk = (& dotnet --version 2>$null) } catch { }
if (-not $sdk) {
    Stop-WithMessage 'The .NET SDK was not found.' 'Install the .NET 10 SDK from https://dotnet.microsoft.com/download, open a new window and run this script again.'
}
$sdkMajor = 0
[void][int]::TryParse(($sdk -split '\.')[0], [ref]$sdkMajor)
if ($sdkMajor -lt 10) {
    Stop-WithMessage "The .NET SDK on this machine is version $sdk, but the benchmarks need version 10 or newer." 'Install the .NET 10 SDK from https://dotnet.microsoft.com/download and run this script again.'
}

$commit = ''
if (Get-Command git -ErrorAction SilentlyContinue) { try { $commit = (& git -C $repoRoot rev-parse HEAD 2>$null) } catch { } }

# The package versions come straight from the project file, so the list can never drift from what is built.
$packages = @()
try {
    [xml]$csproj = Get-Content (Join-Path $projectDir 'SheetStream.Benchmarks.csproj') -Raw
    foreach ($ref in $csproj.Project.ItemGroup.PackageReference) { if ($ref.Include) { $packages += ('{0} {1}' -f $ref.Include, $ref.Version) } }
} catch { }

$facts = @(
    "CPU:             $cpuName",
    "Cores:           $physical physical, $logical logical",
    "Memory:          $ramGb GB",
    "Operating system: $osName",
    "Machine type:    $vmHint",
    "Power plan:      $powerPlan",
    ".NET SDK:        $sdk",
    "Code version:    $(if ($commit) { $commit } else { 'not a Git checkout' })",
    "Scenarios:       $($selected -join ', ')",
    "Date:            $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')",
    '',
    'Packages under test (from SheetStream.Benchmarks.csproj):'
) + ($packages | ForEach-Object { "  $_" })
$facts | ForEach-Object { Say "  $_" }

if ($logical -gt 0 -and $logical -lt 4) { Say '  Note: fewer than 4 logical cores. Results will be valid for this machine only.' 'Yellow' }
if ($ramGb -is [double] -and $ramGb -lt 16 -and (($selected -contains 'Million') -or ($selected -contains 'VeryWide'))) {
    Say '  Note: Million and VeryWide need about 16 GB of memory; with less, the in-memory libraries may fail.' 'Yellow'
}
if ($powerPlan -and $powerPlan -notmatch 'High|Ultimate|Alto') { Say "  Tip: the power plan is '$powerPlan'. A high-performance plan gives steadier numbers." 'Yellow' }

# ---------------------------------------------------------------- results folder
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$resultDir = Join-Path $benchRoot ("results\$stamp-$($env:COMPUTERNAME)")
Heading 'Where the results will go'
Say "  $resultDir"
if (-not $DryRun) {
    New-Item -ItemType Directory -Path $resultDir -Force | Out-Null
    $facts | Set-Content -Path (Join-Path $resultDir 'hardware.txt') -Encoding UTF8
    try { Start-Transcript -Path (Join-Path $resultDir 'run-log.txt') | Out-Null; $script:Transcribing = $true } catch { }
}

# ---------------------------------------------------------------- build
Heading 'Step 1 of 4: build'
Push-Location $projectDir
try {
    $code = Invoke-Dotnet 'build' @('build', '-c', 'Release') (Join-Path $resultDir 'build.txt')
    if ($code -ne 0) { Stop-WithMessage 'The build failed.' "Read the messages above (also saved in $resultDir\build.txt). The most common cause is a missing .NET 10 SDK or no internet access to download the packages." }

    # ---------------------------------------------------------------- correctness guard
    Heading 'Step 2 of 4: correctness check (every library must write exactly the data it was given)'
    $code = Invoke-Dotnet 'verify' @('run', '-c', 'Release', '--no-build', '--', '--verify') (Join-Path $resultDir 'verify.txt')
    if ($code -ne 0) {
        Stop-WithMessage 'At least one library did not write the data it was given, so no timing was taken.' "Open $resultDir\verify.txt and look for the lines that say FAILED. A failing library must be fixed, or left out of the comparison, before any number is published."
    }
    if (-not $DryRun) { Say '  All libraries passed. Timing can start.' 'Green' }

    # ---------------------------------------------------------------- benchmarks
    Heading 'Step 3 of 4: timing (BenchmarkDotNet; about 10 to 15 minutes per 100,000-row scenario)'
    $artifacts = Join-Path $projectDir 'BenchmarkDotNet.Artifacts'
    if ((Test-Path $artifacts) -and -not $DryRun) {
        # Keep any older reports out of the way, so everything copied afterwards belongs to this run.
        $old = "$artifacts.before-$stamp"
        Rename-Item $artifacts $old
        Say "  Older reports were moved to $old" 'Gray'
    }
    elseif ($DryRun -and (Test-Path $artifacts)) { Say '  Would move the older BenchmarkDotNet.Artifacts folder out of the way.' 'DarkYellow' }
    foreach ($name in $selected) {
        Say "  Scenario $name" 'White'
        # '*.<Class>.*' selects exactly one class. Do not shorten it: the matching ignores letter case, so a
        # filter like '*Read*' would also pick up every benchmark whose name contains "spread".
        $filter = '*.' + $KnownScenarios[$name] + '.*'
        $code = Invoke-Dotnet "bench-$name" @('run', '-c', 'Release', '--no-build', '--', '--filter', $filter) (Join-Path $resultDir "benchmark-$name.txt")
        if ($code -ne 0) { Say "  The $name benchmark ended with an error (exit code $code). The other scenarios will still run; see benchmark-$name.txt." 'Red'; $script:AnyFailed = $true }
    }

    # ---------------------------------------------------------------- peak memory
    if ($SkipPeak) {
        Heading 'Step 4 of 4: peak memory (skipped, -SkipPeak)'
    }
    else {
        Heading 'Step 4 of 4: peak memory (every library in a fresh process, three times each)'
        $peakArgs = @('run', '-c', 'Release', '--no-build', '--', '--peak') + $selected
        $code = Invoke-Dotnet 'peak' $peakArgs (Join-Path $resultDir 'peak.txt')
        if ($code -ne 0) { Say "  The peak-memory run ended with an error (exit code $code); see peak.txt." 'Red'; $script:AnyFailed = $true }
    }
}
finally {
    Pop-Location
}

# ---------------------------------------------------------------- collect
Heading 'Collecting the results'
$artifacts = Join-Path $projectDir 'BenchmarkDotNet.Artifacts'
if ($DryRun) {
    Say "  Would copy $artifacts into $resultDir\BenchmarkDotNet.Artifacts" 'DarkYellow'
}
else {
    if (Test-Path $artifacts) {
        Copy-Item $artifacts (Join-Path $resultDir 'BenchmarkDotNet.Artifacts') -Recurse -Force
        Say '  Copied the BenchmarkDotNet reports.'
    }
    else { Say '  No BenchmarkDotNet reports were found (no scenario produced one).' 'Yellow'; $script:AnyFailed = $true }
}

Heading 'Done'
if ($DryRun) {
    Say '  That was a dry run. Nothing was changed. Run again without -DryRun to measure.' 'Yellow'
}
else {
    Say "  Everything is in: $resultDir" 'Green'
    Say '    hardware.txt   the machine, SDK and package versions' 'Gray'
    Say '    verify.txt     what every library wrote in the correctness check' 'Gray'
    Say '    benchmark-*.txt, BenchmarkDotNet.Artifacts\results\*-report-github.md   the timings' 'Gray'
    Say '    peak.txt       peak memory, cold write time and output size' 'Gray'
    if ($script:AnyFailed) { Say '  Some steps ended with an error. Read the messages above before using these numbers.' 'Red' }
    if ($script:Transcribing) { try { Stop-Transcript | Out-Null } catch { } }
}
Wait-ForKey
if ($script:AnyFailed) { exit 2 }

