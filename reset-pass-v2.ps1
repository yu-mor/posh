#!!!!!!!!!!!!!!!!!!!!сброс!!!!!!!!!!!!!! 
#‘фио или логин’

$string="Введи фио пользователя или часть (например, мор ю). Можно и логин или его начальную часть (например, mor). Кому СБРААААСЫВАЕМ?"
$fio = Read-Host $string 
if($fio.Contains(' ')){
#fio 
$fio_name=$fio.replace(" ","* ")+"*"
get-aduser -Filter {name -like $fio_name}|%{
    #Write-Host $_.SamAccountName $_.Name
    $Action= Read-Host ("Сбрасываем "+$_.SamAccountName +" " +$_.Name+" (y/N)?")
    if ($Action="y"){
        #не успеваевает до ChangePasswordAtLogon $true
        #Set-ADUser -Identity $_ -PasswordNeverExpires $false
        Set-ADAccountPassword -Identity $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText “123qweASD” -Force -Verbose)
        Unlock-ADAccount -Identity $_
        Set-ADUser -Identity $_ -ChangePasswordAtLogon $true
        }
    }
}
Else{
#username
$fio_username=$fio+"*"
get-aduser -Filter {SamAccountName -like $fio_username}|%{
    #Write-Host $_.SamAccountName $_.Name
    $Action= Read-Host ("Сбрасываем "+$_.SamAccountName +" " +$_.Name+" (y/N)?")
    if ($Action="y"){
        #не успеваевает до ChangePasswordAtLogon $true
        #Set-ADUser -Identity $_ -PasswordNeverExpires $false
        Set-ADAccountPassword -Identity $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText “123qweASD” -Force -Verbose)
        Unlock-ADAccount -Identity $_
        Set-ADUser -Identity $_ -ChangePasswordAtLogon $true
        }
    }
}
