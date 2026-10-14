<#
.SYNOPSIS
    Opens an .xlsx in the installed Excel (COM automation), refreshes every pivot table and prints its cells.

.DESCRIPTION
    The Open XML validator proves a file is schema-correct, not that Excel accepts it. This script is the
    second gate for pivot tables: Excel must open the workbook without repairing it (a repaired workbook
    silently drops the pivot parts, so the pivot count is compared with -ExpectedPivots), every pivot must
    refresh from its source (RefreshTable throws when Excel rejects the cache), and the refreshed values are
    printed as a table so the sums can be checked by eye or by a caller.

    Exit code 0 = everything refreshed, 1 = a check failed. Excel is always closed, even on failure.

.PARAMETER Path
    The workbook to check.

.PARAMETER ExpectedPivots
    Total number of pivot tables the workbook should contain across all sheets. -1 skips the count check.

.EXAMPLE
    .\excel-check-pivot.ps1 -Path .\Showcase.xlsx -ExpectedPivots 1
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [int]$ExpectedPivots = -1
)

$ErrorActionPreference = 'Stop'
$fullPath = (Resolve-Path -LiteralPath $Path).Path
$excel = $null
$workbook = $null
$failed = $false

function Write-RangeTable {
    param($Range)
    # Value2 returns a 2-D array for several cells, a scalar for one cell.
    $values = $Range.Value2
    $rows = $Range.Rows.Count
    $cols = $Range.Columns.Count
    if ($rows -eq 1 -and $cols -eq 1) { Write-Output ("    " + $values); return }
    for ($r = 1; $r -le $rows; $r++) {
        $cells = @()
        for ($c = 1; $c -le $cols; $c++) {
            $v = $values[$r, $c]
            if ($null -eq $v) { $v = '' }
            elseif ($v -is [double]) { $v = $v.ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture) }
            $cells += ('{0,-18}' -f $v)
        }
        Write-Output ("    " + ($cells -join ' | '))
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false   # No repair / compatibility dialogs; a damaged file then fails to open or loses parts.
    $excel.AskToUpdateLinks = $false

    Write-Output "Opening $fullPath in Excel $($excel.Version)"
    $workbook = $excel.Workbooks.Open($fullPath, 0, $true)   # UpdateLinks=0, ReadOnly=true
    $readOnly = $workbook.ReadOnly
    $sheetCount = $workbook.Worksheets.Count
    Write-Output "Opened. ReadOnly=$readOnly Sheets=$sheetCount"

    $total = 0
    foreach ($ws in $workbook.Worksheets) {
        $pivots = $ws.PivotTables()
        if ($pivots.Count -eq 0) { continue }
        Write-Output ("Sheet '{0}': {1} pivot table(s)" -f $ws.Name, $pivots.Count)
        for ($i = 1; $i -le $pivots.Count; $i++) {
            $pt = $pivots.Item($i)
            $total++
            Write-Output ("  Pivot '{0}'  source={1}  cache records before refresh={2}" -f $pt.Name, $pt.SourceData, $pt.PivotCache().RecordCount)
            try {
                $ok = $pt.RefreshTable()
                Write-Output ("  RefreshTable returned {0}; records after refresh={1}; range={2}" -f $ok, $pt.PivotCache().RecordCount, $pt.TableRange2.Address(0, 0))
                Write-RangeTable -Range $pt.TableRange2
            }
            catch {
                $failed = $true
                Write-Output ("  FAILED: Excel refused to refresh pivot '{0}': {1}" -f $pt.Name, $_.Exception.Message)
            }
        }
    }

    if ($ExpectedPivots -ge 0 -and $total -ne $ExpectedPivots) {
        $failed = $true
        Write-Output ("FAILED: expected {0} pivot table(s) but Excel sees {1}. Excel probably repaired the file and dropped the pivot parts." -f $ExpectedPivots, $total)
    }
    elseif ($total -eq 0) {
        $failed = $true
        Write-Output "FAILED: the workbook contains no pivot tables."
    }
}
catch {
    $failed = $true
    Write-Output ("FAILED: {0}" -f $_.Exception.Message)
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    if ($null -ne $workbook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
    if ($null -ne $excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($failed) { Write-Output "RESULT: FAILED"; exit 1 }
Write-Output "RESULT: OK - every pivot table refreshed in Excel"
exit 0
