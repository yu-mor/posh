# фио или логин
$string="Кого ищем? Введи фио пользователя или часть (например, мор ю). Можно и логин или его начальную часть (например, mor)"
$fio = Read-Host $string

#todo проверку на кириллицу
#фио
if($fio.Contains(' ')){
#возможно, споткнётся на двух пробелах подряд
$fio_name=$fio.replace(" ","* ")+"*"
get-aduser -Filter {name -like $fio_name} -Properties PasswordLastSet, LockedOut, LastBadPasswordAttempt, LastLogonDate, PasswordExpired, Description, whenCreated -Server agat-sr-dc02 |%{$_ ; Unlock-ADAccount $_}
}
else{
#username пока тупо звездочка в конце
$fio_username=$fio+"*" 
get-aduser -Filter {SamAccountName -like $fio_username} -Properties PasswordLastSet, LockedOut, LastBadPasswordAttempt, LastLogonDate, PasswordExpired, Description, whenCreated -Server agat-sr-dc02 |%{$_ ; Unlock-ADAccount $_}
}

