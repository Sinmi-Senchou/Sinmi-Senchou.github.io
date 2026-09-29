@echo off
chcp 65001 >nul
echo 正在根据属性索引提取 MP3 信息...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$shell=New-Object -ComObject Shell.Application; ^
$folder=$shell.Namespace((Get-Location).Path); ^
$items=@($folder.Items() ^| Where-Object { $_.Path -like '*.mp3' }); ^
if ($items.Count -eq 0) { Write-Host '当前文件夹没有 MP3 文件。'; exit }; ^
$titleIdx=21; $lenIdx=27; $artistIdx=20; ^
$results=foreach($item in $items){ ^
    $name=$item.Name; ^
    $title=$folder.GetDetailsOf($item,$titleIdx); ^
    $duration=$folder.GetDetailsOf($item,$lenIdx); ^
    $artist=$folder.GetDetailsOf($item,$artistIdx); ^
    [PSCustomObject]@{'文件名'=$name;'时长'=$duration;'标题'=$title;'参与创作的艺术家'=$artist} ^
}; ^
$csvContent = $results ^| ConvertTo-Csv -NoTypeInformation; ^
$utf8 = New-Object System.Text.UTF8Encoding $true; ^
[System.IO.File]::WriteAllLines('.\mp3_info.csv', $csvContent, $utf8); ^
Write-Host '提取完成，结果保存在 mp3_info.csv'"
pause