$ascii=$NULL;For($a=33;$a–le126;$a++){$ascii+=,[char][byte]$a}                                       # ASCII String for RANDOM String Generation
$ED=[Math]::Floor([decimal](Get-Date(Get-Date).ToUniversalTime()-uformat "%s"))                      # Get EPOC time
$file = "C:\Users\"+$env:UserName+"\Downloads\data\xls\1529873201.xlsx"                              # location of XLS file
$fullfilename=$file.Split("\\") | Select -Index ([regex]::Matches($file, "\\" )).count               # Split Filename
$pathname=$file.Substring(0,$file.LastIndexOf($fullfilename))                                        # To Create a Backup Copy
$realfilename=$fullfilename.Substring(0,$fullfilename.LastIndexOf(".xl"))                            # of File Using EPOCH time
$realfileextn=$fullfilename.Substring($fullfilename.IndexOf(".xl"))                                  # as File Name
Copy-Item ($pathname + $realfilename + $realfileextn) -Destination ($pathname + $ED + $realfileextn) # create backup
if(![System.IO.File]::Exists($file))                                                                 # Check If File Exists
    {Write-Host("# file with path $path doesn't exist")}                                             # Throw File I/O Exception
else{GET-SpreadSheet}                                                                                # Process XLS
Function GET-rndmstrng() {Param([int]$length=10,[string[]]$sourcedata);                              # Function to 
    For ($loop=1;$loop–le$length;$loop++){$TempRNDString+=($sourcedata | GET-RANDOM)};               # Generate 
    return $TempRNDString}                                                                           # Random String
Function GET-SpreadSheet() {$objExcel=New-Object -ComObject Excel.Application;                       # Open XLS APP
    $workbook=$objExcel.Workbooks.Open($file);                                                       # Open XLS File
    $objExcel.Visible=$false;$objExcel.DisplayAlerts=$false                                          # Disable Alerts
    for ($iterateSheet=1;$iterateSheet-le$workbook.sheets.count;$iterateSheet++){                    # Iterate Through Sheets
        $sheet=$workbook.Sheets.Item($iterateSheet);$varSheetName=$sheet.Name;                       # Get Sheet Name
        $objRange=$sheet.UsedRange;                                                                  # Find Data Range
        $lastrow=$objRange.SpecialCells(11).row ;                                                    # Find Row Range
        $lastcol=$objRange.SpecialCells(11).column                                                   # Find Column Range
        for ($row=1;$row-le$lastrow;$row++){                                                         # Iterate Through Rows
            for ($col=1;$col-le$lastcol;$col++){                                                     # Iterate Through Columns 
                $cellvalue=$sheet.Cells.Item($row,$col).text;$data=$data+" " +$cellvalue             # Get Cell Value
                $rndmstrng=GET-rndmstrng –length $cellvalue.length –sourcedata $ascii;               # Generate Random Value
                $sheet.Cells.Item($row,$col)=$rndmstrng}}}                                           # Update Cell Value
    $workbook.save();$objExcel.quit();                                                               # Exit
}