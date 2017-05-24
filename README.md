## run-remote-desktop-from-google-spreadsheet.
This is used to maintain active RDP sessions to a list of Windows servers, as noted in a Google spreadsheet.
It makes use of a 3rd party program (https://github.com/prasmussen/gdrive) to download a copy of the sheet.
It then imports the sheet, and looks for any entry for which you are responsible, and that has the appropriate status.
For each valid entry, it then checks to see if there is an active RDP session, and if not, it will start one.

This is useful if you are having to make changes to a long list of servers, especially changes which require reboots, to make sure
that you do not get lost in your list.

#### Requirements:
* The gdrive software
* A properly formatted google spreadsheet

#### Instructions:
1. Download the gdrive software (https://github.com/prasmussen/gdrive) and rename the executable to *gdrive.exe*
1. Run the gdrive.exe and go through the authorization process to connect it to your google account
1. Prepare a Google Spreadsheet
	1. Create a new sheet, with at least three columns.  One for the server names, one for the user responsible, and one for the status
	1. The first row should be column headers. Example: Host, Task Admin, Host Status
	1. Fill out the rest of the rows with the hosts, and the responsible user
	1. Make note of the sheet name, and ensure that it is unique.
1. Edit the powershell script, to ensure variables fit your situation
	1. *$SheetName* should be the unique sheet name you found above
	1. *$AdminNameHeader* should be the column header for the responsible user
	1. *$HostNameHeader* should be the column header for the server names
	1. *$StatusHeader* should be the column header for the status of each server
	1. *$StatusToUse* is the value the script will look for in the status column.  If it matches this string, it will maintain an RDP session.
	1. *$ShowDebugMessages* to yes/no, if you want to see more detailed script output if something doesn't work right.
1. Run the script!
1. As you finish with a server or are ready to start a new one, change the status in the google sheet.
	
##### Example Sheet Layout
Host | Task Admin | Host Status
---- | ---------- | -----------
server1 | bpeters |
server2 | bpeters | patching
server3 | someuser | done
server4 | someuser | patching
server5 | bpeters | done

Based on the above sheet, if I (bpeters) run the script, it will try to maintain an RDP session to *server2* since it's assigned to me, and has the status *patching*

## check-bitlocker-key-backup-status
This script will poll an Active Directory group for computers that have the bitlocker policies applied.
It will then connect to Active Directory, and check to see if the bitlocker keys are backed up against the computer object.
Next, it will connect to your MBAM database server, and see if keys are properly backed up there.

#### Requirements:
* Bitlocker policies applied to a specific group within Active Directory
* A functional Bitlocker MBAM setup (https://technet.microsoft.com/en-us/windows/hh826072.aspx)
* Network access to the MBAM database

#### Instructions:

1. Edit the powershell script, to ensure variables fit your situation
	1. $ADComputerGroup = The name of the group for which the policies are applied
	1. $SQLServer = Your MBAM SQL server
	1. $SQLDBName = The MBAM database
	1. $path = Your Active Directory DN
1. Run the script!

