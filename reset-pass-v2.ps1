#!!!!!!!!!!!!!!!!!!!!сброс!!!!!!!!!!!!!! 
#Сфио или логинТ
#todo: убрать на тонких требование смены
$string="¬веди фио пользовател€ или часть (например, мор ю). ћожно и логин или его начальную часть (например, mor).  ому —Ѕ–јјјј—џ¬ј≈ћ?"
$fio = Read-Host $string 
if($fio.Contains(' ')){
#fio 
$fio_name=$fio.replace(" ","* ")+"*"
get-aduser -Filter {name -like $fio_name}|%{
    #Write-Host $_.SamAccountName $_.Name
    $Action= Read-Host ("—брасываем "+$_.SamAccountName +" " +$_.Name+" (y/N)?")
    if ($Action="y"){
        Set-ADAccountPassword -Identity $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText УP@ssw0rdФ -Force -Verbose)
        Unlock-ADAccount -Identity $_
        Set-ADUser -Identity $_ -ChangePasswordAtLogon $true
        }
    }
}
Else{
#username пока не отлаживал
$fio_username=$fio+"*"
get-aduser -Filter {SamAccountName -like $fio_username}|%{
    #Write-Host $_.SamAccountName $_.Name
    $Action= Read-Host ("—брасываем "+$_.SamAccountName +" " +$_.Name+" (y/N)?")
    if ($Action="y"){
        Set-ADAccountPassword -Identity $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText УP@ssw0rdФ -Force -Verbose)
        Unlock-ADAccount -Identity $_
        Set-ADUser -Identity $_ -ChangePasswordAtLogon $true
        }
    }
}
