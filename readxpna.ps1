cls
Write-Host "Start Execution    @ " + (Get-Date(Get-Date).ToUniversalTime()-uformat "%s")
$file = "C:\Users\"+$env:UserName+"\Downloads\data\txt\xpnet.log"   
GET-Process-Xpnet_Audit $file
Write-Host "Complete Execution @ " + (Get-Date(Get-Date).ToUniversalTime()-uformat "%s")





Function GET-Process-Xpnet_Audit() {

    Param(
        [parameter(Mandatory=$true)]
        [String]
        $StringAuditFileName
        )

    $StringArrayText=Get-Content -Path $StringAuditFileName;Write-host("x");$StringArrayText.GetType() #|Format-Table -AutoSize
    $IntegerTextFileLength=$StringArrayText.Count
    #$IntegerTextFileLength=4
    $IntegerTextFileLengthStart=0
    $RecordHeaderFound=0;$RecordFooterFound=0;
    $IntegerRecordCounter=0

    For ($loop=$IntegerTextFileLengthStart;$loop–le$IntegerTextFileLength;$loop++){
        if($RecordHeaderFound-eq 0){
            if($StringArrayText[$loop].Length -eq 79){
                $StringRecordHeaderCheck2=$StringArrayText[$loop].Substring(46,4);$StringRecordHeaderCheck1=$StringArrayText[$loop].Substring(0,1)
                    if(($StringRecordHeaderCheck1-eq"R")-and($StringRecordHeaderCheck2-eq"Len ")){$IntegerRecordHeaderFoundAt=$loop;$RecordHeaderFound=1;}
                }
        }
        if(($RecordHeaderFound-eq $true)-and($RecordFooterFound-eq $false)){
            if($StringArrayText[$loop].Length -gt 5)
                {if((Select-String -InputObject $StringArrayText[$loop] -Pattern "RRN" -Quiet -SimpleMatch)-and
                    (Select-String -InputObject $StringArrayText[$loop] -Pattern "37" -Quiet -SimpleMatch)-and 
                    (Select-String -InputObject $StringArrayText[$loop] -Pattern ":" -Quiet -SimpleMatch)-and
                    (($StringArrayText[$loop].IndexOf("37"))-le($StringArrayText[$loop].IndexOf("RRN")))){$IntegerRRNFoundAt=$loop;}
                }
            }
        if(($RecordFooterFound-eq 0)-and($RecordHeaderFound-eq 1)){
            if($StringArrayText[$loop].Length -eq 42){
                $StringRecordFooterCheck1=$StringArrayText[$loop].Substring(20,22);$StringRecordFooterCheck2=$StringArrayText[$loop].Substring(0,1)
                if(($StringRecordFooterCheck1-eq"--- End of Message ---")-and($StringRecordFooterCheck2-eq"(")) 
                    {$IntegerRecordFooterFoundAt=$loop;$RecordFooterFound=1;$IntegerRecordCounter=$IntegerRecordCounter+1;
                        Write-Host("Record       : #" + $IntegerRecordCounter + " |");
                        $StringMsgTypeFoundAt=$IntegerRecordHeaderFoundAt + 2;
                        if($StringArrayText[$StringMsgTypeFoundAt].substring(0,11) -eq"B24 ISO8583"){
                            Write-Host("Message Type : #" + $StringMsgTypeFoundAt  + " |Data : " + $StringArrayText[$StringMsgTypeFoundAt]);
                            Write-Host("Header Found : #" + $IntegerRecordHeaderFoundAt + " |Data : " + $StringArrayText[$IntegerRecordHeaderFoundAt]);
                            Write-Host("RRN          : #" + $IntegerRRNFoundAt  + " |Data : " + $StringArrayText[$IntegerRRNFoundAt])
                            Write-Host("Footer Found : #" + $IntegerRecordFooterFoundAt  + " |Data : " + $StringArrayText[$IntegerRecordFooterFoundAt]);
                            
  
                            }
                        else {
                            Write-Host("Currently Processing : #" + $StringMsgTypeFoundAt  + " |Data : " + $StringArrayText[$StringMsgTypeFoundAt]);
                            Write-Host("Message Type : #" + $StringMsgTypeFoundAt  + " |Data : " + $StringArrayText[$StringMsgTypeFoundAt]);
                            }
                        Write-Host(" ");
                        $RecordHeaderFound=$RecordFooterFound=$StringMsgTypeFoundAt=
                            $IntegerRRNFoundAt=$IntegerRecordHeaderFoundAt=$IntegerRecordFooterFoundAt=0;
                    }
                }
            }
        }
    }