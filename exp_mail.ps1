#Import libery Exchange Management PowerShell
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
#Set log file path
$Logfile = "logfile.log"
$LogPath = "Path:\to\logfile\"
#Write log function
function WriteLog($LogString) {
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    $LogMessage | Out-File -FilePath $LogPath$LogFile -Append -Encoding utf8
    }
#Get User list from OU
$DisableUsers = Get-ADUser -SearchBase ‘OU=Выгрузить,OU=Уволенные,DC=DOMAIN,DC=LOCAL’ -filter *
#Declare a variable to combine multiple export requests.
$BatchName = 'ExportMailRequest'
#Create export paths
$CYear = (Get-Date).year 
$CurrentYear = "$CYear" #Create export paths
$MainDir = "\\Paths\to\folder\EMAIL_ARCHIVES\"
$ExportPath = $MainDir + $CurrentYear + "\"
#Check if the current year folder already exists, if not, then create it.
if ((Test-Path $ExportPath -PathType Container) -eq $false){
    New-Item -Path $MainDir -Name $CurrentYear -ItemType "directory"
}
#Upload mail in a cycle to a .pst file from the main and archive mailboxes
foreach($User in $DisableUsers){
    $PrimaryPath = $ExportPath + $User.SamAccountName + "_" + $User.Surname + "_" + $User.GivenName + "_mail.pst"
    $ArhivePath = $ExportPath + $User.SamAccountName + "_" + $User.Surname + "_" + $User.GivenName + "_archive.pst"
#Using the BatchName parameter we combine requests to be able to track the status of the entire upload at once.
    New-MailboxExportRequest -Mailbox $User.SamAccountName -BatchName $BatchName -FilePath $PrimaryPath
    New-MailboxExportRequest -Mailbox $User.SamAccountName -BatchName $BatchName -FilePath $ArhivePath -IsArchive
}
#Check whether the upload was successful, if not, wait
$i=1;
while ((Get-MailboxExportRequest -BatchName $BatchName | Where {($_.Status -eq “Queued”) -or ($_.Status -eq “InProgress”)})) {
    sleep 60
    Write-Host "Скрипт работает $i минут. Ожидаем завершения.."
    $i=$i+1
}
#After the export is complete, delete all requests
Get-MailboxExportRequest -Status Completed | Remove-MailboxExportRequest -Confirm:$false
#Cleaning up mailing lists. We get the list.
$DistribList = Get-DistributionGroup
#We remove users from lists in a loop
foreach($List in $DistribList){
    foreach($User in $DisableUsers){        
        Remove-DistributionGroupMember -Identity $List -Member $User -Confirm:$false -ErrorAction Ignore  
    }
} 
#Disable mailboxes to clear the address book
$msg = ""
foreach($User in $DisableUsers){
    Disable-Mailbox -Identity $User.SamAccountName -Archive -Confirm:$false
    Disable-Mailbox -Identity $User.SamAccountName -Confirm:$false
    $dat = get-date -Format "dd.MM.yyyy"
    $date = "$dat"
#Write information about the upload in the Description field
    $descr = "PST-OK - " + $date 
    Set-ADUser $User -Description $descr
#Write a message to the log
    $msg = "Выгружен архив почты пользователя " + $User.SamAccountName
    WriteLog $msg
}
#Updating the Global Address List so that users can see the changes
Get-GlobalAddressList | Update-GlobalAddressList
Get-OfflineAddressBook | Update-OfflineAddressBook
Get-AddressList | Update-AddressList
