$ascii=$NULL;For($a=33;$a–le126;$a++){$ascii+=,[char][byte]$a}
$ED=[Math]::Floor([decimal](Get-Date(Get-Date).ToUniversalTime()-uformat "%s"))
$file = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("QzpcVXNlcnNc"))+ 
    $env:UserName + 
    [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String("XERvd25sb2Fkc1xkYXRhXHhsc1wxNTI5ODczMjAxLnhsc3g="))
$fullfilename=$file.Split("\\") | Select -Index ([regex]::Matches($file, "\\" )).count
$pathname=$file.Substring(0,$file.LastIndexOf($fullfilename))
$realfilename=$fullfilename.Substring(0,$fullfilename.LastIndexOf(".xl"))
$realfileextn=$fullfilename.Substring($fullfilename.IndexOf(".xl"))
Copy-Item ($pathname + $realfilename + $realfileextn) -Destination ($pathname + $ED + $realfileextn)
if(![System.IO.File]::Exists($file)){Write-Host("# file with path $path doesn't exist")}else{GET-SpreadSheet}
Function GET-rndmstrng() {Param([int]$length=10,[string[]]$sourcedata);
    For ($loop=1;$loop–le$length;$loop++){$TempRNDString+=($sourcedata | GET-RANDOM)};return $TempRNDString}
Function GET-SpreadSheet() {$objExcel=New-Object -ComObject Excel.Application; 
    $workbook=$objExcel.Workbooks.Open($file);$objExcel.Visible=$false;$objExcel.DisplayAlerts=$false 
    for ($iterateSheet=1;$iterateSheet-le$workbook.sheets.count;$iterateSheet++){
        $sheet=$workbook.Sheets.Item($iterateSheet);$varSheetName=$sheet.Name;Write-Host("Sheet Name: "+$varSheetName)
        $rowMax=($sheet.UsedRange.Rows).count;
        $objRange=$sheet.UsedRange;$lastrow=$objRange.SpecialCells(11).row ;$lastcol=$objRange.SpecialCells(11).column
        for ($row=1;$row-le$lastrow;$row++){$data=" "
            for ($col=1;$col-le$lastcol;$col++){ 
                $cellvalue=$sheet.Cells.Item($row,$col).text;$data=$data+" " +$cellvalue 
                $rndmstrng=GET-rndmstrng –length $cellvalue.length –sourcedata $ascii;$sheet.Cells.Item($row,$col)=$rndmstrng}
            Write-Host("Kill: "+$data)}}
    $workbook.save();$objExcel.quit();
}