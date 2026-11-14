<#
.SYNOPSIS
    Opens workbooks in the installed Excel (COM automation) and reports any that Excel had to repair.

.DESCRIPTION
    The Open XML validator proves a file is schema-correct, not that Excel accepts it. Excel is stricter in
    places and looser in others, and what users see when it disagrees is the "We found a problem with some
    content" box. This script is the second gate: every workbook must open with nothing repaired.

    Excel reports a repair in two ways and the script checks both. With alerts off it repairs silently and
    marks the window caption "[Repaired]"; and it writes a recovery log (error*.xml) to the temp folder,
    which the script looks for by time. A file Excel cannot open at all fails with its message.

    For each workbook that opens cleanly the script prints its sheets and how many tables, charts, pivot
    tables and comments Excel found, so a part that was silently dropped shows up as a missing count.

    Exit code 0 = every workbook opened clean, 1 = at least one was repaired or refused. Excel is always
    closed, even on failure. Needs Excel on the machine, so this is a local check, not part of CI.

.PARAMETER Path
    One or more workbooks, or folders whose .xlsx files are all checked.

.EXAMPLE
    dotnet run --project samples/SheetStream.Sample -- showcase
    .\tests\excel-check.ps1 -Path .\Showcase.xlsx

.EXAMPLE
    .\tests\excel-check.ps1 -Path .\artifacts\acceptance
#>
param(
    [Parameter(Mandatory = $true)][string[]]$Path
)

$ErrorActionPreference = 'Stop'

$files = @()
foreach ($item in $Path) {
    $resolved = Resolve-Path -LiteralPath $item
    foreach ($entry in $resolved) {
        if (Test-Path -LiteralPath $entry.Path -PathType Container) {
            $files += Get-ChildItem -LiteralPath $entry.Path -Filter *.xlsx -File | Where-Object { $_.Name -notlike '~$*' } | ForEach-Object { $_.FullName }
        }
        else { $files += $entry.Path }
    }
}
if ($files.Count -eq 0) { Write-Output "No workbooks to check."; exit 1 }

$excel = $null
$failures = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false   # No dialogs: a damaged file is then repaired silently or refused.
    $excel.AskToUpdateLinks = $false
    Write-Output "Excel $($excel.Version), $($files.Count) workbook(s)"

    foreach ($file in $files) {
        $name = Split-Path $file -Leaf
        $workbook = $null
        $before = Get-Date
        try {
            $workbook = $excel.Workbooks.Open($file, 0, $true)   # UpdateLinks=0, ReadOnly=true
            $caption = ''
            try { $caption = [string]$excel.ActiveWindow.Caption } catch { }
            # Excel writes the recovery log beside its other temp files, named error<digits>_<nn>.xml.
            $logs = @(Get-ChildItem -Path $env:TEMP -Filter 'error*.xml' -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $before.AddSeconds(-1) })
            if ($caption -match 'Repaired' -or $logs.Count -gt 0) {
                $failures++
                Write-Output ("REPAIRED  {0}" -f $name)
                foreach ($log in $logs) {
                    $text = Get-Content -LiteralPath $log.FullName -Raw
                    foreach ($m in [regex]::Matches($text, '<(?:removedPart|removedRecord|repairedRecord|removedFeature|repairedPart)[^>]*>([^<]*)<')) { Write-Output ("            " + $m.Groups[1].Value) }
                }
                continue
            }

            $sheets = @(); $tables = 0; $charts = 0; $pivots = 0; $comments = 0
            foreach ($ws in $workbook.Worksheets) {
                $sheets += $ws.Name
                $tables += $ws.ListObjects.Count
                $charts += $ws.ChartObjects().Count
                $pivots += $ws.PivotTables().Count
                $comments += $ws.Comments.Count
            }
            Write-Output ("OK        {0}  sheets={1} tables={2} charts={3} pivots={4} comments={5}  [{6}]" -f $name, $sheets.Count, $tables, $charts, $pivots, $comments, ($sheets -join ', '))
        }
        catch {
            $failures++
            Write-Output ("REFUSED   {0}: {1}" -f $name, $_.Exception.Message)
        }
        finally {
            if ($null -ne $workbook) {
                try { $workbook.Close($false) } catch { }
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
            }
        }
    }
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($failures -gt 0) { Write-Output "RESULT: FAILED - $failures of $($files.Count) workbook(s) were repaired or refused by Excel"; exit 1 }
Write-Output "RESULT: OK - Excel opened all $($files.Count) workbook(s) without repairing anything"
exit 0
