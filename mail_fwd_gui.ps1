Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
$Logfile = "logfile.log"
$LogPath = "Path:\to\log\dir\"
#Write log function
function WriteLog($LogString) {
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    $LogMessage | Out-File -FilePath $LogPath+$LogFile -Append -Encoding utf8
    }
#Exit on error function
function ExitErr {
    $MSG = "Пользователь не выбран." 
    [Microsoft.VisualBasic.Interaction]::MsgBox("$MSG", "OKOnly,SystemModal,Information", "Title")
    exit
    }
#get ok(6) or cancel
function getYon($formTitle, $textTitle){
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $formTitle
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = 'CenterScreen'
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(95,70)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'Включить'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(170,70)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = 'Отключить'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = $textTitle
    $form.Controls.Add($label)
    $form.Topmost = $true
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { $resss = 6 }
    if ($result -eq [System.Windows.Forms.DialogResult]::CANCEL) { $resss = 0 }
    return $resss
    }
#Import dc session
$Session = New-PSSession -ComputerName dc_controller_name.domain.local
Import-Module -PSsession $Session -Name ActiveDirectory
$UserName = "DOMAIN\User"
$PlainPassword = "Password"
$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ` -ArgumentList $UserName, $SecurePassword
#Import exchange session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange_server_name.domain.local/PowerShell/ -Authentication Kerberos -Credential $Credentials
Import-PSSession $Session
#Select organization from OU list
$OU = Get-ADOrganizationalUnit -SearchBase "OU=Accounts, DC=DOMAIN, DC=LOCAL" -Filter * | Select Name, DistinguishedName | Out-GridView -Title "Select organization?" -OutputMode Single
$act = getYon "Выберите действие?" "Включить или отключить переадресацию?"
if ($act -eq 6 ) {
    $Susr1 = Get-ADUser -SearchBase $OU.DistinguishedName -filter * | Select Name, SamAccountName , UserPrincipalName | Out-GridView -Title "От которого будут переадресовыватся сообщения?" -OutputMode Single
    $Susr2 = Get-ADUser -SearchBase $OU.DistinguishedName -filter * | Select Name, SamAccountName , UserPrincipalName | Out-GridView -Title "На которого будут переадресовыватся сообщения?" -OutputMode Single
    if ( $Susr1 -eq $null ) { ExitErr }
    if ( $Susr2 -eq $null ) { ExitErr }
    Set-Mailbox -Identity $Susr1.SamAccountName -DeliverToMailboxAndForward $true -ForwardingAddress $Susr2.SamAccountName
    
    $MSG = "Включена переадресация почты с пользователя " + $Susr1.Name + " на пользователя " + $Susr2.Name
    WriteLog $MSG
    [Microsoft.VisualBasic.Interaction]::MsgBox("$MSG", "OKOnly,SystemModal,Information", "Title")
}
if ($act -eq 0 ) {
    $Susr = Get-ADUser -SearchBase $OU.DistinguishedName -filter * | Select Name, SamAccountName , UserPrincipalName | Out-GridView -Title "Выберите пользователя у которого отключить переадресацию?" -OutputMode Single
    if ( $Susr -eq $null ) { ExitErr }
    Set-Mailbox -Identity $Susr.SamAccountName -DeliverToMailboxAndForward $false -ForwardingAddress $null
    $MSG = "Отключена переадресация почты у пользователя " + $Susr.Name
    WriteLog $MSG
    [Microsoft.VisualBasic.Interaction]::MsgBox("$MSG", "OKOnly,SystemModal,Information", "Title")
}
