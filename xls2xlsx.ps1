#Remove-Module Resize-Image
#Import-Module "D:\POSH\Resize-Image.psm1"

$excel = New-Object -ComObject Excel.Application
$excel.visible = $true
$dir="P:\Заказы\"#со слэшем
#не трогаем свежие файлы            -
Get-ChildItem -PATH $dir -include "*.xls" -Recurse | Where-Object { ($_.CreationTime -lt (Get-Date).AddDays(-1)) -and ($_.Length -gt 2560000) } |ForEach-Object{#-filter "*.xls" пропускает xlsx тоже
    (Get-Date).ToString() +"; try "+ $_.FullName| Out-File -FilePath ("P:\ОИТ\МордасовЮрий\1\del.log") -Append 
    $excel.displayalerts = $false
    #Trap [System.Runtime.InteropServices.COMException]
    #{
    
    #Continue
    #}
    Try{
    $excel.Workbooks.Open($_.FullName,0)
    }
    Catch [System.Runtime.InteropServices.COMException]{
    (Get-Date).ToString() +"; bad "| Out-File -FilePath ("D:\logs\del.log") -Append
    return
    }
    $excel.displayalerts = $true
    $excel.ActiveWorkbook.SaveAs(($_.DirectoryName+"\"+$_.BaseName+".xlsx"),51)
    $excel.ActiveWorkbook.Close($false)
    if(Test-Path($_.DirectoryName+"\"+$_.BaseName+".xlsx")){#если получился xlsx
        #распаковываем
        Rename-Item ($_.DirectoryName+"\"+$_.BaseName+".xlsx") ($_.BaseName+".zip") -Force
        expand-archive ($_.DirectoryName+"\"+$_.BaseName+".zip") -DestinationPath ($dir+"xls2xlsx") -force
        #&'c:\Program Files\7-Zip\7z.exe' 'x', ($_.DirectoryName+"\"+$_.BaseName+".zip"),' -o',($dir+"xls2xlsx")
        ##сжимаем картинки png
        Get-ChildItem ($dir+"xls2xlsx\xl\media\*.png")| foreach{
            #echo ($_.DirectoryName+"new\"+$_.Name)
            #Resize-Image -InputFile $old -Width 1024 -OutputFile $old
            $wia = New-Object -ComObject wia.imagefile
            $wia.LoadFile($_.FullName)
            Write-Host $wia.Width, $wia.Height
            $wip = New-Object -ComObject wia.imageprocess
            $scale = $wip.FilterInfos.Item("Scale").FilterId                    
            $wip.Filters.Add($scale)
            $wip.Filters[1].Properties("MaximumWidth") = 1024
            $wip.Filters[1].Properties("MaximumHeight") = 1024
            #aspect ratio should be set as false if you want the pics in exact size 
            $wip.Filters[1].Properties("PreserveAspectRatio") = $true 
            $wip.Apply($wia) 
            $newimg = $wip.Apply($wia)
            Remove-Item($_.FullName)
            $newimg.SaveFile($_.FullName)    
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wia)
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wip)
        }
    #todo:
    #конвертим emf в png и сжимаем
    #правим \xl\drawings\_rels\*.xml emf->png
    #правим \xl\drawings\*.xml что именно?)))

    #compress-archive -path ($dir+"xls2xlsx\xl") -destinationpath ($_.DirectoryName+"\"+$_.BaseName+".zip") -update #зараза, не заменяет файлы, а добавляет
    &'c:\Program Files\7-Zip\7z.exe' 'u',($_.DirectoryName+"\"+$_.BaseName+".zip"), ($dir+"xls2xlsx\xl"),' *.png -r'
    Remove-Item ($dir+"xls2xlsx") -Recurse
    Rename-Item ($_.DirectoryName+"\"+$_.BaseName+".zip") ($_.BaseName+".xlsx") -Force
    #аттрибуты
    (Get-Item ($_.DirectoryName+"\"+$_.BaseName+".xlsx")).CreationTime=$_.CreationTime
    (Get-Item ($_.DirectoryName+"\"+$_.BaseName+".xlsx")).LastAccessTime=$_.LastAccessTime
    (Get-Item ($_.DirectoryName+"\"+$_.BaseName+".xlsx")).LastWriteTime=$_.LastWriteTime
    Remove-Item $_.FullName #удаляем исходник
    (Get-Date).ToString() +"; DEL "+ $_.FullName| Out-File -FilePath ("P:\ОИТ\МордасовЮрий\1\del.log") -Append 
    #(Get-Date).ToString() +"; "+ $_.FullName| Out-File -FilePath ($dir+"del.log") -Append 
    }
}
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)

