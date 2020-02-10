<# 
    Inbox (24,460) - sdmoore68@gmail.com - Gmail
https://mail.google.com/mail/u/0/#inbox

myhealth.bankofamerica.com
https://myhealth.bankofamerica.com/login.aspx

windows - How to enable access to a win10 shared folder from android VLC? - Super User
https://superuser.com/questions/1344535/how-to-enable-access-to-a-win10-shared-folder-from-android-vlc

TweetDeck
https://tweetdeck.twitter.com/

Torrents - TD
https://www.torrentday.com/t?q=who+was&qf=&96=on&25=on&11=on&5=on&3=on&21=on&22=on&13=on&44=on&48=on&1=on&24=on&32=on&31=on&33=on&46=on&82=on&14=on&26=on&7=on&34=on&2=on&29=on&42=on&20=on&30=on&95=on&47=on&43=on&45=on&28=on&12=on#torrents

get-childitem extended properties - Google Search
https://www.google.com/search?safe=active&biw=1536&bih=758&sxsrf=ACYBGNRTNtDxS4Difw2T8ramFagU2coKyg%3A1578757682139&ei=Mu4ZXoebCInX5gKVsZuwAg&q=get-childitem+extended+properties&oq=get-childitem+extended+proper&gs_l=psy-ab.3.0.0i22i30.4693.8735..9948...0.4..0.117.1401.12j3......0....1..gws-wiz.......0i71j0j0i13i30j0i13i5i30.VgXGAjBkmEA

Filtering files by their metadata (extended properties) | Loose Scripts Sink Ships
https://martin77s.wordpress.com/2015/02/22/filtering-files-by-their-metadata-extended-properties/

Test-Path
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.management/test-path?view=powershell-7

[SOLVED] Measure-Object Doesn't Like Me - PowerShell - Spiceworks
https://community.spiceworks.com/topic/1125978-measure-object-doesn-t-like-me

Convert a string to datetime in PowerShell - Stack Overflow
https://stackoverflow.com/questions/38717490/convert-a-string-to-datetime-in-powershell

about_Hash_Tables - PowerShell | Microsoft Docs
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_hash_tables?view=powershell-7

Copy URLs Of All Open Tabs In All Browsers [Chrome, Firefox, Opera, IE]
https://www.itechtics.com/copy-urls-of-all-open-tabs-in-all-browsers-chrome-firefox-opera-ie/

TabCopy - Chrome Web Store
https://chrome.google.com/webstore/detail/tabcopy/micdllihgoppmejpecmkilggmaagfdmb
#>
function Get-FileMetaData{
    [cmdletbinding()]
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [Alias('Fullname')]
        [string[]]$path
    )
    Begin{
        $oshell = New-Object -ComObject shell.Application
        #$objArray = @{}
    }
    Process{
        $path | ForEach-Object{
            # "Processing <$_>"
            if (Test-Path -Path $_ -PathType Leaf){
                $FileItem = Get-Item -Path $_
                $ofolder = $oshell.namespace($FileItem.DirectoryName)
                $ofile = $ofolder.parsename($FileItem.Name)
                
                $props = @{}

                0..287 | ForEach-Object{
                    $ExtPropName = $ofolder.GetDetailsof($ofolder.items, $_)
                    $ExtValName = $ofolder.GetDetailsof($ofile, $_)
                    if(-not $props.ContainsKey($ExtPropName) -and ($ExtValName -ne '')){
                        $props.Add($ExtPropName,$ExtValName)
                    }
                }
                New-Object psobject -Property $props | Tee-Object -Variable obj
            }
        }
    }
    End{
        $oshell = $null
    }
}
#Example
$FilePath = "D:\SharedDoc\Belinda\CHICOM\Day03" #""E:\Lock\F"
$filter = "*.mov"
Get-ChildItem -Path $FilePath -Recurse -Filter $filter | Select-Object name, CreationTime, Directory | Sort-Object -Property CreationTime -OutVariable fileOrder | Out-Null
Write-Host "test `n`n"
$fileOrder | Select-Object -First 1 -OutVariable OldestFile | Out-Null
$fileOrder | Select-Object -Last 1 -OutVariable NewestFile | Out-Null
$OldestDate = $OldestFile.CreationTime.Date
$NewestDate = $NewestFile.CreationTime.Date
$varDay = $OldestDate
$FileTimeTotal = $null
while ($varDay -lt $NewestDate.AddDays(1)) {
    # $lastconstraint = $NewestDate
    $DayTimeTotal = $null
    Write-Host "Files for " $varDay
    Get-ChildItem -Path $FilePath -Filter $filter -Recurse | Where-Object {$_.CreationTime -ge $varDay -and $_.CreationTime -lt $varDay.AddDays(1)} | Sort-Object -Property CreationTime | Get-FileMetaData -OutVariable FilesByDate | Format-Table Name, 'Date Created', length
    foreach($test in $FilesByDate){
        $DayTimeTotal += [timespan]$test.Length
    }
    Write-Host "Total time for " $varday.Date" = " $DayTimeTotal "`n"
    $varDay = $varDay.AddDays(1)
    $FileTimeTotal += $DayTimeTotal
}
Write-Host "Total time of Files: " $FileTimeTotal
break 
#Example #2
Get-Childitem -Path $FilePath -recurse -Filter $filter | Get-FileMetaData -OutVariable testvar | Select-Object name, length
Write-Host `n 
foreach($test in $testvar){
        $DayTimeTotal += [timespan]$test.length
    }
Write-Host $DayTimeTotal