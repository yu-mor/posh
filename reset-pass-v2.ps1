#!!!!!!!!!!!!!!!!!!!!�����!!!!!!!!!!!!!! 
#���� ��� �����
#todo: ������ �� ������ ���������� �����
$string="����� ��� ������������ ��� ����� (��������, ��� �). ����� � ����� ��� ��� ��������� ����� (��������, mor). ���� �������������?"
$fio = Read-Host $string 
if($fio.Contains(' ')){
#fio 
$fio_name=$fio.replace(" ","* ")+"*"
get-aduser -Filter {name -like $fio_name}|%{
    #Write-Host $_.SamAccountName $_.Name
    $Action= Read-Host ("���������� "+$_.SamAccountName +" " +$_.Name+" (y/N)?")
    if ($Action="y"){
        Set-ADAccountPassword -Identity $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText �P@ssw0rd� -Force -Verbose)
        Unlock-ADAccount -Identity $_
        Set-ADUser -Identity $_ -ChangePasswordAtLogon $true
        }
    }
}
Else{
#username ���� �� ���������
$fio_username=$fio+"*"
get-aduser -Filter {SamAccountName -like $fio_username}|%{
    #Write-Host $_.SamAccountName $_.Name
    $Action= Read-Host ("���������� "+$_.SamAccountName +" " +$_.Name+" (y/N)?")
    if ($Action="y"){
        Set-ADAccountPassword -Identity $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText �P@ssw0rd� -Force -Verbose)
        Unlock-ADAccount -Identity $_
        Set-ADUser -Identity $_ -ChangePasswordAtLogon $true
        }
    }
}
