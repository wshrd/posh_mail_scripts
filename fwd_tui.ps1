Import-Module Microsoft.PowerShell.ConsoleGuiTools
#Set log file path
$Logfile = "logfile.log"
$LogPath = "/Path/to/log/dir/"
#Write log function
function WriteLog($LogString) {
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    $LogMessage | Out-File -FilePath $LogPath$LogFile -Append -Encoding utf8
    }
#Exit on error function
function ExitErr {
    $MSG = "Пользователь не выбран." 
    [Microsoft.VisualBasic.Interaction]::MsgBox("$MSG", "OKOnly,SystemModal,Information", "Title")
    exit
    }
$Session = New-PSSession -ComputerName dc_controller_name.domain.local
Import-Module -PSsession $Session -Name ActiveDirectory
$UserName = "DOMAIN\USER"
$PlainPassword = "Password"
$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ` -ArgumentList $UserName, $SecurePassword
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange_server_name.domain.local/PowerShell/ -Authentication Kerberos -Credential $Credentials
Import-PSSession $Session
#Select OU
$OU = Get-ADOrganizationalUnit -SearchBase "OU=Accounts, DC=DOMAIN, DC=LOCAL" -Filter * | Select Name, DistinguishedName | Out-ConsoleGridView -Title "Select organization?" -OutputMode Single

$module = (Get-Module Microsoft.PowerShell.ConsoleGuiTools -List).ModuleBase
Add-Type -Path (Join-path $module Terminal.Gui.dll)
[Terminal.Gui.Application]::Init()
$act = [Terminal.Gui.MessageBox]::Query("Mail Forwarding", "Enable or Disable mail forwarding?", @("Disable", "Enable"))
[Terminal.Gui.Application]::shutdown()

if ($act -eq "1" ) {
    $Susr1 = Get-ADUser -SearchBase $OU.DistinguishedName -filter * | Select Name, SamAccountName , UserPrincipalName | Out-ConsoleGridView -Title "Select user to enable mail forwarding?" -OutputMode Single
    $Susr2 = Get-ADUser -SearchBase $OU.DistinguishedName -filter * | Select Name, SamAccountName , UserPrincipalName | Out-ConsoleGridView -Title "Select user to whom to forward mail?" -OutputMode Single
    if ( $Susr1 -eq $null ) { ExitErr }
    if ( $Susr2 -eq $null ) { ExitErr }
    Set-Mailbox -Identity $Susr1.SamAccountName -DeliverToMailboxAndForward $true -ForwardingAddress $Susr2.UserPrincipalName
    $msg =  "Включена переадресация почты с пользователя " + $Susr1.Name + " на пользователя " + $Susr2.Name
    [Terminal.Gui.Application]::Init()
    WriteLog $msg
    $ms = [Terminal.Gui.MessageBox]::Query("Mail Forwarding", $msg, @("OK"))
    [Terminal.Gui.Application]::shutdown()
}

if ($act -eq "0" ) {
    $Susr = Get-ADUser -SearchBase $OU.DistinguishedName -filter * | Select Name, SamAccountName , UserPrincipalName | Out-ConsoleGridView -Title "Select user to disable mail forwarding?" -OutputMode Single
    if ( $Susr -eq $null ) { ExitErr }
    Set-Mailbox -Identity $Susr.SamAccountName -DeliverToMailboxAndForward $false -ForwardingAddress $null
    $msg = "Отключена переадресация почты у пользователя " + $Susr.Name
    [Terminal.Gui.Application]::Init()
    WriteLog $msg
    $ms = [Terminal.Gui.MessageBox]::Query("Mail Forwarding", $msg, @("OK"))
    [Terminal.Gui.Application]::shutdown()
}
