$ascii=$NULL;For($a=33;$a–le126;$a++){$ascii+=,[char][byte]$a}
$file = "C:\Users\"+ $env:UserName + "\Downloads\data\xls\pwrshlt.xlsx"
if(![System.IO.File]::Exists($file)){Write-Host("# file with path $path doesn't exist")}else{GET-SpreadSheet}
Function GET-rndmstrng() {Param([int]$length=10,[string[]]$sourcedata);
    For ($loop=1;$loop–le$length;$loop++){$TempPassword+=($sourcedata | GET-RANDOM)};return $TempPassword}
Function GET-SpreadSheet() {$objExcel=New-Object -ComObject Excel.Application; 
    $workbook=$objExcel.Workbooks.Open($file);$objExcel.Visible=$false;$objExcel.DisplayAlerts=$false 
    for ($iterateSheet=1;$iterateSheet-le$workbook.sheets.count;$iterateSheet++){
        $sheet=$workbook.Sheets.Item($iterateSheet);$varSheetName=$sheet.Name;Write-Host("Sheet Name: "+$varSheetName)
        $rowMax=($sheet.UsedRange.Rows).count;Write-Host("Total Rows: "+$rowMax)
        $objRange=$sheet.UsedRange;$lastrow=$objRange.SpecialCells(11).row ;$lastcol=$objRange.SpecialCells(11).column
        for ($row=1;$row-le$lastrow;$row++){$data = " "
            for ($col=1;$col-le$lastcol;$col++) { 
                $cellvalue=$sheet.Cells.Item($row,$col).text;$data=$data+" " +$cellvalue 
                $rndmstrng=GET-rndmstrng –length $cellvalue.length –sourcedata $ascii;$sheet.Cells.Item($row,$col)=$rndmstrng}
            Write-Host("Data: "+$data)}}
    $workbook.save();$objExcel.quit();
}