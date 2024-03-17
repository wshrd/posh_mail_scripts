$Session = New-PSSession -ComputerName dc_controller_name.domain.local
Import-Module -PSsession $Session -Name ActiveDirectory 3>$null
$UserName = "DOMAIN\USER"
$PlainPassword = "Password"
$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ` -ArgumentList $UserName, $SecurePassword
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange_server_name.domain.local/PowerShell/ -Authentication Kerberos -Credential $Credentials
Import-PSSession $Session

Write-host "Enable(1) or Disable(0) mail forwarding?" 
$act = Read-Host

if ($act -eq "1" ) {
    Write-Host "Enter username(SamAccountName) to enable mail forwarding"
    $Susr1 = Read-Host
    Write-Host "Enter the username(SamAccountName) to whom to forward mail"
    $Susr2 = Read-Host
    Set-Mailbox -Identity $Susr1 -DeliverToMailboxAndForward $true -ForwardingAddress $Susr2
    $msg =  "Mail forwarding enabled from " + $Susr1 + " to " + $Susr2
    Write-host $msg
}

if ($act -eq "0" ) {
    Write-Host "Enter username(SamAccountName) to disable mail forwarding"
    $Susr = Read-Host
    Set-Mailbox -Identity $Susr -DeliverToMailboxAndForward $false -ForwardingAddress $null
    $msg = "Mail forwarding disabled for " + $Susr
    Write-host $msg 
}
