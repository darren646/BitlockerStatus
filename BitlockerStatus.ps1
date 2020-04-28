Import-Module ActiveDirectory
$listbitlockerstatus=@("hostname,IsEnableBitlocker")
$ListOU=@("OU=Romania-PC-Give-Admin,OU=WorkStations,OU=DomainComputers,DC=abc,DC=def,DC=com",
"OU=Romania-PC-Remove-Admin,OU=WorkStations,OU=DomainComputers,DC=abc,DC=def,DC=com")

$ListOU | ForEach-Object{
    $OU=$_.ToString()

    Get-ADComputer -Filter 'ObjectClass -eq "computer"' -SearchBase $OU | foreach-object {
        $Computer = $_.name
        #Check if the Computer Object exists
        $Computer_Object = Get-ADComputer -Filter {cn -eq $Computer} -Property msTPM-OwnerInformation, msTPM-TpmInformationForComputer
        if($Computer_Object -eq $null){
        Write-Host "Error..."
        }
        #Check if the computer object has had a BitLocker Recovery Password
        $Bitlocker_Object = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $Computer_Object.DistinguishedName -Properties 'msFVE-RecoveryPassword' | Select-Object -Last 1
        if($Bitlocker_Object.'msFVE-RecoveryPassword'){
        $BitLocker_Key = $BitLocker_Object.'msFVE-RecoveryPassword' -imatch 1
        }
        else
        {
        $BitLocker_Key = "none"
        }

        #Display Output
        $strToReport = $Computer +','+ $BitLocker_Key
        Write-Host $strToReport 

        $listbitlockerstatus+=$strToReport

    }

}




Write-Host "Total computers is "
Write-Host $listbitlockerstatus.Count
$now=Get-Date
$now.ToString("yyyyMMdd")
$outfile="C:\workstsstatus\Bitlocker gv status GSC Romania "+$now.ToString("yyyyMMdd")+".csv"
$listbitlockerstatus |Out-File -Encoding utf8 $outfile #-append

