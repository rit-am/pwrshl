#Declare the file path and sheet name
$file = "C:\Users\"+ $env:UserName + "\Downloads\data\xls\spreadsheet.xlsx" 
#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file) 
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false 
for ($iterateSheet=1; $iterateSheet -le $workbook.sheets.count; $iterateSheet++){
    $sheet=$workbook.Sheets.Item($iterateSheet)
    $varSheetName=$sheet.Name
    Write-Host ("Sheet Name: "+$varSheetName)
    $rowMax = ($sheet.UsedRange.Rows).count 
    Write-Host ("Total Rows: "+$rowMax)
    $objRange = $sheet.UsedRange 
    $lastrow = $objRange.SpecialCells(11).row ;$lastcol = $objRange.SpecialCells(11).column
    write-host "Lastrow:", $lastrow, " Last Column:" $lastcol
    for ($row=1; $row -le $lastrow; $row++)
        { 
        $data = " "
        for ($col=1; $col -le $lastcol-1; $col++) 
            { 
            $cellvalue = $sheet.Cells.Item($row,$col).text ;
            $data = $data + " " + $cellvalue 
            }
        Write-Host ("Data: "+$data)
        }
    }
#close excel file
$objExcel.quit() ;$objExcel.quit() 


