<#
.SYNOPSIS
    Builds the Excel-authored test corpus: workbooks saved by the installed Excel, each with a text file that
    records what Excel itself says every cell holds.

.DESCRIPTION
    The library's own tests mostly read files the library wrote, which proves the two halves agree with each
    other, not that either agrees with Excel. The files this script makes are written by Excel through COM
    automation, so they carry everything Excel puts in a workbook (themes, shared formulas, calc chains,
    extension lists, its own number formats), and the expectations beside them come from Excel as well:
    the typed value of every cell, the date each date cell displays, merges, hyperlinks, sheet order and
    visibility.

    The corpus is checked in, so the tests need no Excel. Run this script only to regenerate it or to add a
    workbook: it needs Excel on the machine. Document properties are overwritten so that no user name from
    the machine that built the corpus ends up in the files.

    Output: tests/corpus/excel/<name>.xlsx and <name>.expected.tsv. The .tsv is tab-separated:
        sheet <tab> kind <tab> address <tab> value
    kind is: meta (sheet state), n (number), s (text), b (boolean), e (error text), d (date cell, the value is
    the date and time Excel displays), t (elapsed-time cell, the value is in days), merge, link.

.EXAMPLE
    .\tests\corpus\build-excel-corpus.ps1
#>
param(
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'excel')
)

$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputFolder | Out-Null
$inv = [System.Globalization.CultureInfo]::InvariantCulture

# Sets a cell's value through reflection. PowerShell's own COM binder remembers the type of the first value
# it passes to a property and then refuses a value of another type, which a script that writes numbers, text
# and booleans cannot live with.
function Set-Value($range, $value) {
    [void]$range.GetType().InvokeMember('Value2', [System.Reflection.BindingFlags]::SetProperty, $null, $range, @(, $value))
}

function Escape([string]$text) { $text.Replace('\', '\\').Replace("`t", '\t').Replace("`r", '\r').Replace("`n", '\n') }

function IsElapsed([string]$format) { $format -match '\[(h+|m+|s+)\]' }

function ColumnName([int]$column) {
    $name = ''
    while ($column -gt 0) { $rem = ($column - 1) % 26; $name = [char](65 + $rem) + $name; $column = [int](($column - 1 - $rem) / 26) }
    $name
}

# One line for a cell, from what Excel says it holds.
function Describe([string]$sheet, [string]$address, $value, $value2, [string]$format, $shown, $cell) {
    if ($null -eq $value) { return $null }
    if ($value -is [string]) { return "$sheet`ts`t$address`t$(Escape $value)" }
    if ($value -is [bool]) { return "$sheet`tb`t$address`t$(if ($value) { 1 } else { 0 })" }
    if ($value -is [int]) { return "$sheet`te`t$address`t$($cell.Text)" }          # An error comes back as its numeric code.
    if (IsElapsed $format) { return "$sheet`tt`t$address`t$(([double]$value2).ToString('R', $inv))" }
    if ($value -is [datetime]) { return "$sheet`td`t$address`t$shown" }             # The date Excel displays, in the workbook's own date system.
    return "$sheet`tn`t$address`t$(([double]$value).ToString('R', $inv))"
}

# Writes what Excel reports for every cell that holds something, on every sheet.
function Write-Expected($workbook, [string]$path) {
    $lines = New-Object System.Collections.Generic.List[string]
    $dateText = '"yyyy-mm-dd hh:mm:ss.000"'
    foreach ($ws in $workbook.Worksheets) {
        $name = $ws.Name
        $state = switch ($ws.Visible) { -1 { 'visible' } 0 { 'hidden' } 2 { 'veryHidden' } default { 'visible' } }
        $lines.Add("$name`tmeta`t`t$state")
        if ($ws.Application.WorksheetFunction.CountA($ws.Cells) -eq 0) { continue }
        $used = $ws.UsedRange
        $rows = $used.Rows.Count; $cols = $used.Columns.Count
        $firstRow = $used.Row; $firstCol = $used.Column

        if ([double]$rows * $cols -le 250000) {
            # A compact sheet: fetch everything in three calls and consult single cells only where needed.
            $values = $used.Value(); $values2 = $used.Value2
            $shown = $ws.Evaluate("TEXT($($used.Address($false, $false)),$dateText)")
            $single = ($rows -eq 1 -and $cols -eq 1)
            $columnFormats = @{}
            for ($c = 1; $c -le $cols; $c++) { $uniform = $used.Columns.Item($c).NumberFormat; if ($uniform -is [string]) { $columnFormats[$c] = $uniform } }
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    $value = if ($single) { $values } else { $values[$r, $c] }
                    if ($null -eq $value) { continue }
                    $cell = $null
                    $format = $columnFormats[$c]
                    if ($null -eq $format -or $value -is [int]) { $cell = $used.Cells.Item($r, $c); $format = [string]$cell.NumberFormat }
                    $address = (ColumnName ($firstCol + $c - 1)) + ($firstRow + $r - 1)
                    $line = Describe $name $address $value $(if ($single) { $values2 } else { $values2[$r, $c] }) $format $(if ($single) { $shown } else { $shown[$r, $c] }) $cell
                    if ($null -ne $line) { $lines.Add($line) }
                }
            }
            # Merged areas, each once. MergeCells is False for the whole range when there are none.
            if ($used.MergeCells -isnot [bool] -or $used.MergeCells) {
                $seen = @{}
                foreach ($cell in $used.Cells) {
                    if ($cell.MergeCells -eq $true) {
                        $area = $cell.MergeArea.Address($false, $false)
                        if (-not $seen.ContainsKey($area)) { $seen[$area] = $true; $lines.Add("$name`tmerge`t$area`t") }
                    }
                }
            }
        }
        else {
            # A sparse sheet whose corners are far apart: visit only the cells that hold a constant or a formula.
            foreach ($kind in 2, -4123) {      # xlCellTypeConstants, xlCellTypeFormulas
                $found = $null
                try { $found = $ws.Cells.SpecialCells($kind) } catch { }
                if ($null -eq $found) { continue }
                foreach ($cell in $found.Cells) {
                    $address = $cell.Address($false, $false)
                    $line = Describe $name $address $cell.Value() $cell.Value2 ([string]$cell.NumberFormat) $ws.Evaluate("TEXT($address,$dateText)") $cell
                    if ($null -ne $line) { $lines.Add($line) }
                }
            }
        }
        foreach ($link in $ws.Hyperlinks) {
            $lines.Add("$name`tlink`t$($link.Range.Address($false, $false))`t$(Escape ([string]$link.Address))|$(Escape ([string]$link.SubAddress))|$(Escape ([string]$link.ScreenTip))")
        }
    }
    [System.IO.File]::WriteAllLines($path, $lines, (New-Object System.Text.UTF8Encoding($false)))
}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.UserName = 'SheetStream.NET test corpus'   # Becomes "last modified by" in the saved files.
$built = @()

# Replaces what Excel stamps into a workbook from the machine that saved it: the names in the document
# properties (author, last modified by, company) and the folder path in the workbook part. None of that has a
# place in a public repository.
function Remove-PersonalProperties([string]$file) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($file, 'Update')
    try {
        foreach ($part in 'docProps/core.xml', 'docProps/app.xml', 'xl/workbook.xml') {
            $entry = $zip.GetEntry($part)
            if ($null -eq $entry) { continue }
            $reader = New-Object System.IO.StreamReader($entry.Open())
            $xml = $reader.ReadToEnd(); $reader.Dispose()
            $xml = [regex]::Replace($xml, '(<dc:creator>)[^<]*(</dc:creator>)', '${1}SheetStream.NET test corpus${2}')
            $xml = [regex]::Replace($xml, '(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)', '${1}SheetStream.NET test corpus${2}')
            $xml = [regex]::Replace($xml, '(<Company>)[^<]*(</Company>)', '${1}${2}')
            $xml = [regex]::Replace($xml, '(<Manager>)[^<]*(</Manager>)', '${1}${2}')
            # Excel records the folder the workbook was saved in.
            $xml = [regex]::Replace($xml, '(<x15ac:absPath url=")[^"]*(")', '${1}C:\corpus\${2}')
            $entry.Delete()
            $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($part).Open(), (New-Object System.Text.UTF8Encoding($false)))
            $writer.Write($xml); $writer.Dispose()
        }
    }
    finally { $zip.Dispose() }
}

function Save-Corpus($workbook, [string]$name) {
    $file = Join-Path $OutputFolder "$name.xlsx"
    if (Test-Path $file) { [System.IO.File]::Delete($file) }
    $excel.Calculate()
    $workbook.SaveAs($file, 51)   # 51 = xlOpenXMLWorkbook
    Write-Expected $workbook (Join-Path $OutputFolder "$name.expected.tsv")
    $workbook.Close($false)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    Remove-PersonalProperties $file
    $script:built += $name
    Write-Output ("  {0,-22} {1,8:N0} bytes" -f $name, (Get-Item $file).Length)
}

try {
    Write-Output "Excel $($excel.Version) -> $OutputFolder"

    # ---------------------------------------------------------------- values of every type
    $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Types'
    $numbers = 0, 1, -1, 42, 0.1, -0.30000000000000004, 1e-7, 123456789.123456789, 1e15, 9.99999999999999e307, 2.675, 1.005, (1 / 3)
    for ($i = 0; $i -lt $numbers.Count; $i++) { Set-Value ($ws.Cells.Item($i + 1, 1)) ([double]$numbers[$i])}
    $texts = 'plain', ' leading and trailing ', "two`nlines", 'a & b < c > d "q" ''s''', 'ñandú café', '日本語', 'emoji 😀', '_x0041_ stays literal', '=not a formula', '00123', ('long ' * 2000)
    for ($i = 0; $i -lt $texts.Count; $i++) { $cell = $ws.Cells.Item($i + 1, 2); $cell.NumberFormat = '@'; Set-Value ($cell) ($texts[$i])}
    Set-Value ($ws.Cells.Item(1, 3)) ($true); Set-Value ($ws.Cells.Item(2, 3)) ($false)
    $errors = '=NA()', '=1/0', '="a"+1', '=SQRT(-1)', '=missingname', '=A1 B2', '=INDEX(A1:A2,5)'
    for ($i = 0; $i -lt $errors.Count; $i++) { $ws.Cells.Item($i + 1, 4).Formula = $errors[$i] }
    Set-Value ($ws.Cells.Item(30, 6)) ('far corner')      # A gap of empty rows and columns.
    Save-Corpus $wb 'types'

    # ---------------------------------------------------------------- number formats, dates and times
    $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Formats'
    $rows = @(
        @(46000.75, 'yyyy-mm-dd hh:mm:ss'), @(46000.75, 'm/d/yyyy'), @(46000.75, 'dddd, mmmm d, yyyy'), @(46000.75, 'hh:mm'),
        @(46000.75, 'h:mm:ss AM/PM'), @(0.5, 'hh:mm:ss'), @(1, 'yyyy-mm-dd'), @(59, 'yyyy-mm-dd'), @(61, 'yyyy-mm-dd'), @(2958465, 'yyyy-mm-dd'),
        @(46000.999999, 'yyyy-mm-dd hh:mm:ss'), @(1.25, '[h]:mm:ss'), @(0.0625, '[h]:mm'), @(100.5, '[m]:ss'), @(36.5, '[s]'),
        @(1234.5678, '#,##0.00'), @(0.256, '0.0%'), @(1234.5, '0.00E+00'), @(2.5, '# ?/?'), @(1234.5, '$#,##0.00'),
        @(12.5, '0.0 "days"'), @(7, '0 "months"'), @(3, '"d"0'), @(1500, '[Red]0;[Blue]-0'), @(42, 'General'),
        @(5, '0 \d'), @(46000, 'yyyy'), @(46000, 'mmm-yy'), @(46000.5, '[$-409]d-mmm-yyyy h:mm'), @(0.75, 'mm:ss.0')
    )
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $cell = $ws.Cells.Item($i + 1, 1)
        $cell.NumberFormat = $rows[$i][1]
        Set-Value ($cell) ([double]$rows[$i][0])
        $ws.Cells.Item($i + 1, 2).NumberFormat = '@'
        Set-Value ($ws.Cells.Item($i + 1, 2)) ([string]$rows[$i][1])
    }
    # Codes that Excel stores as built-in format ids (14 to 22, 45 to 47, 9, 11) instead of as custom formats.
    $builtIn = 'm/d/yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yyyy h:mm', 'mm:ss', '[h]:mm:ss', 'mm:ss.0', '0%', '0.00E+00'
    for ($i = 0; $i -lt $builtIn.Count; $i++) {
        $cell = $ws.Cells.Item($i + 1, 4)
        $cell.NumberFormat = [string]$builtIn[$i]
        Set-Value ($cell) ([double](45500.25 + $i))
    }
    Save-Corpus $wb 'formats'

    # ---------------------------------------------------------------- the 1904 date system
    $wb = $excel.Workbooks.Add(); $wb.Date1904 = $true; $ws = $wb.Worksheets.Item(1); $ws.Name = 'Mac dates'
    $serials = 0, 1, 365.5, 44538.75, 0.25
    for ($i = 0; $i -lt $serials.Count; $i++) { $cell = $ws.Cells.Item($i + 1, 1); $cell.NumberFormat = 'yyyy-mm-dd hh:mm:ss'; Set-Value ($cell) ([double]$serials[$i])}
    $ws.Cells.Item(1, 2).NumberFormat = '[h]:mm'; Set-Value ($ws.Cells.Item(1, 2)) (2.5)
    Set-Value ($ws.Cells.Item(2, 2)) (1234.5)
    Save-Corpus $wb 'dates-1904'

    # ---------------------------------------------------------------- formulas
    $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Calc'
    for ($r = 1; $r -le 12; $r++) { Set-Value ($ws.Cells.Item($r, 1)) ($r * 1.5); Set-Value ($ws.Cells.Item($r, 2)) ("item $r")}
    $ws.Range('C1:C12').Formula = '=A1*2'                       # Filled down: Excel stores it as a shared formula.
    $ws.Range('D1').Formula = '=SUM(A1:A12)'
    $ws.Range('D2').Formula = '=AVERAGE(A1:A12)'
    $ws.Range('D3').Formula = '=B1&" and "&B2'
    $ws.Range('D4').Formula = '=A1>1'
    $ws.Range('D5').Formula = '=IF(A1>100,"big","small")'
    $ws.Range('D6').Formula = '=ROUND(A3/7,3)'
    $ws.Range('D7').Formula = '=DATE(2026,11,13)'; $ws.Range('D7').NumberFormat = 'yyyy-mm-dd'
    $ws.Range('D8').Formula = '=VLOOKUP("item 4",B1:C12,2,FALSE)'
    $ws.Range('E1:E3').FormulaArray = '=A1:A3*10'               # A legacy array formula.
    $other = $wb.Worksheets.Add([Type]::Missing, $ws); $other.Name = 'Other sheet'
    $other.Range('A1').Formula = "=Calc!D1"
    $other.Range('A2').Formula = "=SUM(Calc!A1:A3)"
    Save-Corpus $wb 'formulas'

    # ---------------------------------------------------------------- text: shared strings and rich text
    $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Text'
    for ($r = 1; $r -le 40; $r++) { Set-Value ($ws.Cells.Item($r, 1)) (@('repeat', 'again', 'repeat', 'unique ' + $r)[$r % 4])}
    Set-Value ($ws.Range('B1')) ('Bold start, plain end')
    $ws.Range('B1').Characters(1, 10).Font.Bold = $true
    Set-Value ($ws.Range('B2')) ('red green blue')
    $ws.Range('B2').Characters(1, 3).Font.Color = 255
    $ws.Range('B2').Characters(5, 5).Font.Color = 65280
    $ws.Range('B2').Characters(11, 4).Font.Italic = $true
    Set-Value ($ws.Range('B3')) ("wrapped`nrich`ntext")
    $ws.Range('B3').Characters(9, 4).Font.Underline = 2
    Set-Value ($ws.Range('B4')) ('  spaces kept  ')
    Set-Value ($ws.Range('B5')) ('''')                                # An apostrophe-prefixed empty text.
    Save-Corpus $wb 'text'

    # ---------------------------------------------------------------- layout, sheets and annotations
    $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Layout'
    Set-Value ($ws.Range('A1')) ('Title across three columns'); $ws.Range('A1:C1').Merge()
    Set-Value ($ws.Range('A3')) ('Tall'); $ws.Range('A3:A5').Merge()
    Set-Value ($ws.Range('E2')) ('Block'); $ws.Range('E2:G4').Merge()
    for ($r = 7; $r -le 12; $r++) { Set-Value ($ws.Cells.Item($r, 1)) ($r); Set-Value ($ws.Cells.Item($r, 2)) ("row $r")}
    $ws.Rows.Item(9).Hidden = $true
    $ws.Columns.Item(4).Hidden = $true
    $ws.Columns.Item(2).ColumnWidth = 30
    $ws.Rows.Item(8).RowHeight = 40
    $ws.Range('A10:A12').Rows.Group() | Out-Null
    [void]$ws.Hyperlinks.Add($ws.Range('B14'), 'https://example.com/path?x=1&y=2', '', 'Open the site', 'External link')
    [void]$ws.Hyperlinks.Add($ws.Range('B15'), '', "'It''s a sheet'!A1", 'Jump inside', 'Internal link')
    [void]$ws.Hyperlinks.Add($ws.Range('B16'), 'mailto:someone@example.com?subject=Hello', '', '', 'Mail link')
    [void]$ws.Range('C7').AddComment('A note on a cell')
    $excel.ActiveWindow.FreezePanes = $false
    $second = $wb.Worksheets.Add([Type]::Missing, $ws); $second.Name = "It's a sheet"; Set-Value ($second.Range('A1')) ('quoted name')
    $third = $wb.Worksheets.Add([Type]::Missing, $second); $third.Name = 'Hidden'; Set-Value ($third.Range('A1')) ('hidden sheet'); $third.Visible = 0
    $fourth = $wb.Worksheets.Add([Type]::Missing, $third); $fourth.Name = 'Very hidden'; Set-Value ($fourth.Range('A1')) ('very hidden sheet'); $fourth.Visible = 2
    $fifth = $wb.Worksheets.Add([Type]::Missing, $fourth); $fifth.Name = 'Ñandú 日本'; Set-Value ($fifth.Range('B2')) ('unicode sheet name')
    [void]$wb.Names.Add('TitleCell', '=Layout!$A$1')
    Save-Corpus $wb 'layout'

    # ---------------------------------------------------------------- tables, validation, conditional formats, a pivot and a chart
    $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Data'
    $headers = 'Region', 'Product', 'Units', 'Price', 'When'
    for ($c = 0; $c -lt $headers.Count; $c++) { Set-Value ($ws.Cells.Item(1, $c + 1)) ($headers[$c])}
    $regions = 'North', 'South', 'East', 'West'
    for ($r = 2; $r -le 41; $r++) {
        Set-Value ($ws.Cells.Item($r, 1)) ($regions[$r % 4])
        Set-Value ($ws.Cells.Item($r, 2)) ("Product $($r % 7)")
        Set-Value ($ws.Cells.Item($r, 3)) ($r * 3)
        Set-Value ($ws.Cells.Item($r, 4)) (9.99 + $r)
        Set-Value ($ws.Cells.Item($r, 5)) (46000 + $r); $ws.Cells.Item($r, 5).NumberFormat = 'yyyy-mm-dd'
    }
    $table = $ws.ListObjects.Add(1, $ws.Range('A1:E41'), $null, 1)   # 1 = xlSrcRange, 1 = xlYes (has headers)
    $table.Name = 'Sales'
    $table.ShowTotals = $true
    $table.ListColumns.Item('Units').TotalsCalculation = 1            # xlTotalsCalculationSum
    [void]$ws.Range('G2:G10').Validation.Add(3, 1, 1, 'North,South,East,West')   # xlValidateList
    [void]$ws.Range('C2:C41').FormatConditions.AddDatabar()
    $rule = $ws.Range('D2:D41').FormatConditions.Add(1, 5, '=30')    # xlCellValue, xlGreater
    $rule.Interior.Color = 65535
    $pivotSheet = $wb.Worksheets.Add([Type]::Missing, $ws); $pivotSheet.Name = 'Pivot'
    $cache = $wb.PivotCaches().Create(1, 'Sales')                      # xlDatabase
    $pivot = $cache.CreatePivotTable($pivotSheet.Range('A3'), 'ByRegion')
    $pivot.PivotFields('Region').Orientation = 1                       # xlRowField
    [void]$pivot.AddDataField($pivot.PivotFields('Units'), 'Sum of Units', -4157)   # xlSum
    $chartSheet = $wb.Worksheets.Add([Type]::Missing, $pivotSheet); $chartSheet.Name = 'Chart'
    Set-Value ($chartSheet.Range('A1')) ('chart below')
    $chart = $chartSheet.ChartObjects().Add(20, 40, 360, 220).Chart
    $chart.SetSourceData($ws.Range('B1:C8'))
    $chart.ChartType = 51                                              # xlColumnClustered
    Save-Corpus $wb 'features'

    # ---------------------------------------------------------------- the edges of the grid, and a sheet with some size to it
    $wb = $excel.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Edges'
    Set-Value ($ws.Range('A1')) ('first')
    Set-Value ($ws.Range('XFD1')) ('last column')
    Set-Value ($ws.Range('A1048576')) ('last row')
    Set-Value ($ws.Range('XFD1048576')) (16384)
    Set-Value ($ws.Range('ZZ500')) (3.14)
    $big = $wb.Worksheets.Add([Type]::Missing, $ws); $big.Name = 'Rows'
    # Filled by formulas and then turned into plain values inside Excel, which is far quicker than sending
    # 18,000 cells across one at a time. Every tenth note is an empty text, the kind a formula leaves behind.
    $big.Range('A1:A3000').Formula = '=ROW()-1'
    $big.Range('B1:B3000').Formula = '="name "&MOD(ROW()-1,50)'
    $big.Range('C1:C3000').Formula = '=(ROW()-1)*0.01'
    $big.Range('D1:D3000').Formula = '=MOD(ROW()-1,3)=0'
    $big.Range('E1:E3000').Formula = '=45000+MOD(ROW()-1,400)'
    $big.Range('F1:F3000').Formula = '=IF(MOD(ROW()-1,10)=0,"","note "&(ROW()-1))'
    $excel.Calculate()
    [void]$big.Range('A1:F3000').Copy()
    [void]$big.Range('A1:F3000').PasteSpecial(-4163)   # xlPasteValues
    $excel.CutCopyMode = 0
    $big.Range('E1:E3000').NumberFormat = 'yyyy-mm-dd'
    Save-Corpus $wb 'edges-and-rows'
}
finally {
    foreach ($open in @($excel.Workbooks)) { try { $open.Close($false) } catch { } }
    try { $excel.Quit() } catch { }
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Built $($built.Count) workbook(s): $($built -join ', ')"
