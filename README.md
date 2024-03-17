## Power Shell scripts for manage mailboxes on Exchange server 
###  Script fwd_shell.ps1
***
Enable or disable mail forwarding console version.
### Script mail_fwd_gui.ps1
***
Enable or disable mail forwarding GUI version. Power Shell GridView, System.Windows.Forms, System.Drawing is used.
### Script fwd_tui.ps1
***
Enable or disable mail forwarding in the TUI version. ConsoleGridView is used. Working on Power Shell in Linux terminal. Required Microsoft.PowerShell.ConsoleGuiTools to work. 
```powershell
Install-Module Microsoft.PowerShell.ConsoleGuiTools
```
### Script exp_mail.ps1
***
Mailbox export script.
Executed according to the scheduler on the Exchange server.
