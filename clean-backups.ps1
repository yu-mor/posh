#make test fileset 
#1..400|%{$n=$_;New-Item -Path d:\test\ -Name $n".txt" -ItemType File;(get-item ("d:\test\"+$n+".txt")).CreationTime=(Get-Date).AddDays(0-$n);(get-item ("d:\test\"+$n+".txt")).LastWriteTime=(Get-Date).AddDays(0-$n)}
#14 �������, ����� 10 ����������, ����� 6 ������ �����������. ��� �������� ���.
#����� � ����� � � ��������� � ������, ������������ � "_", �� �������

$path="d:\test\"
get-childitem -path $path -Directory -Exclude _* | %{ 

    #������������ �� 14 ���� � ������, ���� �� 1 ����� � ���� �� �����������,  �������
    get-childitem -path $_.FullName | Where-Object { ($_.CreationTime -lt (Get-Date).AddDays(-15))-and ($_.CreationTime.DayOfWeek.value__ -ne 0) }|Remove-Item

    #�����������: �� 3 ������ � ������, ���� �� ����� ������ �� �����,  �������  - ��� ������ ��������� (����������)
    get-childitem -path $_.FullName | ? { ($_.CreationTime -lt (Get-Date).AddMonths(-3)) -and ($_.CreationTime.Day -ne 1) }| % {
        $month=$_.CreationTime.Month
        $year=$_.CreationTime.Year
        #Write-Host $year $month
	    Get-ChildItem -path $_.DirectoryName | ? {($_.CreationTime.Month -eq $month) -and ($_.CreationTime.Year -eq $year)}|Sort-Object -Property CreationTime|Select-Object -skip 1|Remove-Item
    }

    #�� ���� ������� - ��� ����� � ������ ��������
    get-childitem -path $_.FullName | ? { ($_.CreationTime -lt (Get-Date).AddYears(-1))}|Remove-Item


}