param(
  [Parameter(Mandatory=$true)]
  [string]$WorkbookPath,
  [string]$PivotSheetName = '统计透视',
  [string]$PivotTitle = '推荐人客户数统计（按资金账号非重复计数）',
  [string]$TableName = '订单明细'
)

$ErrorActionPreference = 'Stop'

function Get-HeaderColumnIndex {
  param(
    [Parameter(Mandatory=$true)]$Worksheet,
    [Parameter(Mandatory=$true)][string]$HeaderText
  )

  $xlToLeft = -4159
  $lastCol = $Worksheet.Cells.Item(1, $Worksheet.Columns.Count).End($xlToLeft).Column
  for ($c = 1; $c -le $lastCol; $c++) {
    $value = $Worksheet.Cells.Item(1, $c).Value2
    if ($null -ne $value -and $value.ToString().Trim() -eq $HeaderText) {
      return $c
    }
  }
  return $null
}

function Release-ComObject {
  param([Parameter(ValueFromPipeline=$true)]$ComObject)
  process {
    if ($null -ne $ComObject) {
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
  }
}

$excel = $null
$workbook = $null
$sourceSheet = $null
$pivotSheet = $null
$pivotCache = $null
$pivotTable = $null
$dataField = $null
$table = $null

try {
  if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
  }

  $fullPath = (Resolve-Path -LiteralPath $WorkbookPath).Path

  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  $workbook = $excel.Workbooks.Open($fullPath)

  foreach ($ws in @($workbook.Worksheets)) {
    if ($ws.Name -eq $PivotSheetName) {
      $ws.Delete()
      break
    }
  }

  foreach ($ws in @($workbook.Worksheets)) {
    if ($ws.Name -eq $PivotSheetName) { continue }
    $acctCol = Get-HeaderColumnIndex -Worksheet $ws -HeaderText '资金账号'
    $refCol = Get-HeaderColumnIndex -Worksheet $ws -HeaderText '推荐人姓名'
    if ($acctCol -and $refCol) {
      $sourceSheet = $ws
      break
    }
  }

  if ($null -eq $sourceSheet) {
    throw 'No source sheet with headers 资金账号 and 推荐人姓名 was found.'
  }

  $helperCol = Get-HeaderColumnIndex -Worksheet $sourceSheet -HeaderText '客户去重标记'
  if ($helperCol) {
    $sourceSheet.Columns.Item($helperCol).Delete()
  }

  $acctCol = Get-HeaderColumnIndex -Worksheet $sourceSheet -HeaderText '资金账号'
  $refCol = Get-HeaderColumnIndex -Worksheet $sourceSheet -HeaderText '推荐人姓名'
  if (-not $acctCol -or -not $refCol) {
    throw 'Headers 资金账号 and 推荐人姓名 are required.'
  }

  $xlUp = -4162
  $xlToLeft = -4159
  $lastRow = $sourceSheet.Cells.Item($sourceSheet.Rows.Count, $acctCol).End($xlUp).Row
  if ($lastRow -lt 2) {
    throw 'Data row count is less than 1. Cannot create pivot table.'
  }

  $lastCol = $sourceSheet.Cells.Item(1, $sourceSheet.Columns.Count).End($xlToLeft).Column
  if ($lastCol -lt [Math]::Max($acctCol, $refCol)) {
    $lastCol = [Math]::Max($acctCol, $refCol)
  }

  $sourceRange = $sourceSheet.Range($sourceSheet.Cells.Item(1,1), $sourceSheet.Cells.Item($lastRow, $lastCol))

  if ($sourceSheet.ListObjects.Count -gt 0) {
    $table = $sourceSheet.ListObjects.Item(1)
    $table.Resize($sourceRange)
  } else {
    $xlSrcRange = 1
    $xlYes = 1
    $table = $sourceSheet.ListObjects.Add($xlSrcRange, $sourceRange, $null, $xlYes)
  }

  try {
    $table.Name = $TableName
  } catch {
    # Keep existing table name when rename fails.
  }

  $pivotSheet = $workbook.Worksheets.Add()
  $pivotSheet.Name = $PivotSheetName
  $pivotSheet.Range('A1').Value2 = $PivotTitle

  $xlDatabase = 1
  $xlPivotTableVersion15 = 5
  $pivotCache = $workbook.PivotCaches().Create($xlDatabase, $table.Range, $xlPivotTableVersion15)
  $pivotTable = $pivotCache.CreatePivotTable($pivotSheet.Range('A3'), '推荐人客户统计')

  $xlRowField = 1
  $pivotRowField = $pivotTable.PivotFields('推荐人姓名')
  $pivotRowField.Orientation = $xlRowField
  $pivotRowField.Position = 1

  $dataField = $pivotTable.AddDataField($pivotTable.PivotFields('资金账号'), '客户数(非重复资金账号)')
  $xlDistinctCount = 11
  $dataField.Function = $xlDistinctCount

  $pivotTable.PivotCache().RefreshOnFileOpen = $true
  [void]$pivotSheet.Columns('A:B').AutoFit()

  $workbook.Save()
  $workbook.Close($true)

  Write-Output "Done: rebuilt pivot table in $fullPath"
  Write-Output "Sheet: $PivotSheetName"
  Write-Output 'Value field: Distinct Count of 资金账号'
}
finally {
  if ($null -ne $workbook) {
    try { $workbook.Close($false) } catch {}
  }
  if ($null -ne $excel) {
    try { $excel.Quit() } catch {}
  }

  $dataField | Release-ComObject
  $pivotTable | Release-ComObject
  $pivotCache | Release-ComObject
  $pivotSheet | Release-ComObject
  $table | Release-ComObject
  $sourceSheet | Release-ComObject
  $workbook | Release-ComObject
  $excel | Release-ComObject

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
