# check-bitlocker-key-backup-status
This script will poll an Active Directory group for computers that have the bitlocker policies applied.
It will then connect to Active Directory, and check to see if the bitlocker keys are backed up against the computer object.
Next, it will connect to your MBAM database server, and see if keys are properly backed up there.
#run-remote-desktop-from-google-spreadsheet.
This is used to maintain active RDP sessions to a list of Windows servers, as noted in a Google spreadsheet.
It makes use of a 3rd party program (https://github.com/prasmussen/gdrive) to download a copy of the sheet.
It then imports the sheet, and looks for any entry for which you are responsible, and that has the appropriate status.
For each valid entry, it then checks to see if there is an active RDP session, and if not, it will start one.

This is useful if you are having to make changes to a long list of servers, especially changes which require reboots, to make sure
that you do not get lost in your list.
