Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null 
$outlook = new-object -comobject outlook.application 
$namespace = $outlook.GetNameSpace("MAPI") 
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type] 

$sendFolder = $namespace.getDefaultFolder($olFolders::olFolderSentMail)

#prepare the date variable
$today=Get-Date

$allSentMails = $sendFolder.items

#$thisMonthEmails=$allSentMails |?{$_.senton.month -eq $today.Month -and $_.senton.year -eq $today.Year}
#$lastMonthEmails=$allSentMails |?{$_.senton.month -eq $today.AddMonths(-1).Month -and $_.senton.year -eq $today.AddMonths(-1).Year}
$last7daysEmails=$allSentMails | ?{$_.senton -lt $today -and $_.senton -gt $today.AddDays(-7)}

#total email amount sent this month
function get-allMailCount($mails){
    $mails.Length
}

function get-AllMailLength($mails){
    $mails | % {$totalLength+=[math]::Round(($_.size/1MB),2)}
}

function Get-TableYearMonthCount($mails){
    $resultList=@()
    foreach($m in $mails){
        $sentTime=$m.senton
        $object = New-Object –TypeName PSObject -Property @{
                        year=$sentTime.year
                        month=$sentTime.month
                     }
        $resultList+=$object
    }

    $resultList=$resultList |Group-Object year,month | select Count,Name
    $final=@()
    $final | Add-Member "Year" ""
    $final | Add-Member "Month" ""
    $final | Add-Member "Count" ""
    foreach($r in $resultList){
         if($r.name -eq ""){
            $object = New-Object –TypeName PSObject -Property @{
                    Year=""
                    Month=""
                    Count=$r.Count
             }
         }else{
            $object = New-Object –TypeName PSObject -Property @{
                Year=$r.name.split(",")[0]
                Month=[int]$r.name.split(",")[1]
                Count=$r.Count
            }
         
         $final+=$object
        }
    }
    return $final | sort year,month
}

# get  "receiver | count" table, sort by count
function Get-Table-ReceiverCountSize($mails){
    $resultList=@()
    foreach ($m in $mails){
        foreach ($r in $m.Recipients){
            if ($resultList.receiver -contains $r.name){
                foreach($item in $resultList){
                    if($item.receiver -eq $r.name){
                        $item.count++
                        $item.size+=[math]::Round(($m.size/1MB),2)
                     } 
                }
            }else{
                $object = New-Object –TypeName PSObject -Property @{
                    receiver=$r.name
                    count=1
                    size=[math]::Round(($m.size/1MB),2)
                 }
                $resultList+=$object
            }
        }
    }
    return $resultList
}
