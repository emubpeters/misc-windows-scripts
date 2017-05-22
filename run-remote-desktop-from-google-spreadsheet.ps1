########################
# Written by: Ben Peters
# 5/19/2017
#
# This tool will read a google spreadsheet, for any lines for which you are marked as a task admin.
# It will check the status line, and if it's set to your desired status, it will try to open an RDP session if you
# do not already have it open.  Cool.
#
# Instructions:
# 1 - Copy https://github.com/prasmussen/gdrive to somewhere on your machine (it assumes somewhere on C drive).
# 2 - Rename the executable you downloaded to just 'gdrive.exe'
# 3 - Open up a command prompt, and run the gdrive program.  It should walk you through OAUTH authorization.
# 4 - Edit this script, and change the relevant paths / file names in the "Variable Configuration" section.
#
#########################

###############
# Variable Configuration
###############
# How long should I wait between checks (In seconds)?
$Delay = 30

# Who are you? I.E. who should it look for to know the ask is assigned to you?
$user = 'bpeters'

# What's the name of the Google sheet you are working with?
$SheetName = '<name of sheet>'

# What is the header for the column containing the name of the person responsible?
$AdminNameHeader = 'Task Admin'

# What is the header for the column containing the hostnames?
$HostNameHeader = 'Host'

# What is the header for the column containing the status, and what status should I look for?
$StatusHeader = 'Host Status'
$StatusToUse = 'patching'

# Where is the gdrive program located?
$InputFilePath = "C:\Users\<username>\Downloads\"

# Debugging?
$ShowDebugMessages = "yes"

###################
# Script stuff
###################

# Get the google drive file name, then download the most recent copy
$Command = $InputFilePath + "gdrive.exe list --no-header --query `"name contains '$SheetName'`""
$GCommand = Invoke-Expression $Command
$GoogleFile = $GCommand -split " "
$GoogleFile = $GoogleFile -split " "
$GoogleFile = $GoogleFile[0]
$InputFile = $InputFilePath + $SheetName 

# Loop this baby until we're done!
while ($true) {

    clear
    
    if ($ShowDebugMessages -eq "yes") {
        write-host "----------start debug----------"
        write-host "Path to gdrive.exe executable: " $InputFilePath
        write-host "Name of Google worksheet: " $SheetName
        write-host "What we are searching on in sheet: " $AdminNameHeader "=" $user ", " $StatusHeader "=" $StatusToUse
        write-host "Delay between runs (seconds): " $Delay
        write-host "----------end debug----------"
        write-host ""
    }

    # Download and Import the working file to see status
    write-host $(Get-Date -Format g) "Downloading current spreadsheet..."
    write-host "-----------------------------------"
    $Command = $InputFilePath + "gdrive.exe export $GoogleFile --force"
    Invoke-Expression $Command
    $ServersBeingPatched = Import-Csv $InputFile -EA Ignore | Where-Object {$_.$StatusHeader -eq $StatusToUse -and $_.$AdminNameHeader -eq $user}

    # Get all instances of the remote desktop app that are running
    $RunningRemoteDesktop = Get-Process mstsc -EA SilentlyContinue | where {$_.mainWindowTitle}

    # Find out which of the ones listed as patching doesn't have an active RDP Session
    foreach ($server in $ServersBeingPatched) {
        $connected = "no"
        foreach ($session in $RunningRemoteDesktop) {
            if ($session.MainWindowTitle.ToLower() -match $server.Host.ToLower()) {
                $connected = "yes"
            }
        }

        # If we're not connected to an active session, try to do so, if it's actually online
        if ($connected -eq "no") {
            if (test-connection -Count 1 $server.Host) {
                write-host "    " $server.Host " is online!  Attempting RDP session..."
                mstsc /v: $server.Host
            } else {
                write-host "    "  $server.Host " is offline... will try reconnecting at next pass."
            }
        } else {
            write-host "    "  $server.Host " is currently connected."
        }

    }

    write-host "---------------------------------------------"
    write-host "$(Get-Date -Format g) Sleeping" $Delay "seconds..."
    Start-Sleep $Delay

}