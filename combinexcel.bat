@echo off
setlocal enabledelayedexpansion

set "folder_path=C:\Users\DENIS\Documents\_AGED CASE FILE DATA\DISPATCHED FILES - 2024"
set "output_file=C:\Users\DENIS\Documents\_AGED CASE FILE DATA\DISPATCHED FILES - 2024\combined_file.xlsx"

set "powershell_script=[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.Excel') | Out-Null; $excel = New-Object -ComObject Excel.Application; $workbook = $excel.Workbooks.Add(); $worksheet = $workbook.Worksheets.Item(1); $row = 1; Get-ChildItem '%folder_path%' -Filter *.xlsx | ForEach-Object { $currentFile = $_.FullName; $currentWorkbook = $excel.Workbooks.Open($currentFile); $currentSheet = $currentWorkbook.Sheets.Item(1); $rowCount = $currentSheet.UsedRange.Rows.Count; $columnCount = $currentSheet.UsedRange.Columns.Count; $currentSheet.UsedRange.Copy($worksheet.Cells($row, 1)); $row += $rowCount + 2; $currentWorkbook.Close(); }; $workbook.SaveAs('%output_file%'); $workbook.Close(); $excel.Quit();"

echo Running PowerShell script...
powershell -Command "%powershell_script%"

echo Excel files combined successfully.
