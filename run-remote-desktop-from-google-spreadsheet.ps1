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

# Who are you? I.E. who should it look for to know the ask is assigned to you? (Defaults to current user)
$user = $env:UserName

# What's the name of the Google sheet you are working with?
$SheetName = '<sheet name>'

# What is the header for the column containing the name of the person responsible?
$AdminNameHeader = 'Task Admin'

# What is the header for the column containing the hostnames?
$HostNameHeader = 'Host'

# What is the header for the column containing the status, and what status should I look for?
$StatusHeader = 'Host Status'
$StatusToUse = 'patching'

# Debugging?
$ShowDebugMessages = "yes"

###################
# Error Check
###################

# Try and find the gdrive software.  Search first for gdrive.exe
clear
write-host "Searching for gdrive.exe...."
write-host ""
$GDriveSearch = Get-ChildItem -Path C:\ -Filter gdrive.exe -Recurse -ErrorAction SilentlyContinue -Force
if ($GDriveSearch) {
    $GDrivePath = $GDriveSearch.Directory.ToString() + "\"
    write-host "Found gdrive.exe at" $GDrivePath
    write-host ""
} else {
    write-host "Error: Cannot find gdrive.exe.  Please download that first.  See the comments"
    write-host "       at the top of this script for the github path."
    break
}


###################
# Script Body
###################

# Make sure we're executing this script from the same location as the gdrive software, otherwise download will fail.
$ExectPath = (Get-Location).Path + "\"

# Get the google drive file name, then download the most recent copy
$Command = $GDrivePath + "gdrive.exe list --no-header --query `"name contains '$SheetName'`""
$GCommand = Invoke-Expression $Command

# Make sure the output has one and only one line
$NumberOfResults = $GCommand | Measure-Object -Line
if ($NumberOfResults.Lines -ne 1) {
    write-host "Error: Invalid number of matching results found:" $NumberOfResults.Lines
    write-host "       Possible Matches:" $NumberOfResults.Lines
    break
}

# Otherwise, we're good - get the ID of the file to download, and set the input path once it's downloaded
$GoogleFile = $GCommand -split " "
$GoogleFile = $GoogleFile -split " "
$GoogleFile = $GoogleFile[0]

# Set our input file path.  Gdrive.exe will download it to wherever this script is run from. 
$ExectPath = (Get-Location).Path + "\"
$InputFile = $ExectPath + $SheetName

# Loop this baby until we're done!
while ($true) {
    
    if ($ShowDebugMessages -eq "yes") {
        write-host "----------start debug----------"
        write-host "   Path to gdrive.exe executable: " $GDrivePath
        write-host "   Name of Google worksheet: " $SheetName
        write-host "   What we are searching on in sheet: " $AdminNameHeader "=" $user ", " $StatusHeader "=" $StatusToUse
        write-host "   ID of the Google Sheet: " $GoogleFile
        write-host "   Delay between runs (seconds): " $Delay
        write-host "   Number of results in Google matching search: " $NumberOfResults.Lines
        write-host "----------end debug----------"
        write-host ""
        write-host ""
        write-host ""
        write-host ""
    }

    write-host $InputFile

    # Download and Import the working file to see status
    write-host $(Get-Date -Format g) "Downloading current spreadsheet..."
    write-host "-----------------------------------"
    $Command = $GDrivePath + "gdrive.exe export $GoogleFile --force"
    Invoke-Expression $Command
    write-host ""
    $ServersBeingPatched = Import-Csv $InputFile -EA Ignore | Where-Object {$_.$StatusHeader -eq $StatusToUse -and $_.$AdminNameHeader -eq $user}

    # Get all instances of the remote desktop app that are running
    $RunningRemoteDesktop = Get-Process mstsc -EA SilentlyContinue | where {$_.mainWindowTitle}

    # Find out which of the ones listed as patching doesn't have an active RDP Session
    $Count = 0
    foreach ($server in $ServersBeingPatched) {
        $Count++
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

    # No systems match query, mention that.
    if ($Count -eq 0) {
        write-host "    No lines match your query at this time."
    }

    # Write delay info
    write-host ""
    write-host "---------------------------------------------"
    write-host "$(Get-Date -Format g) Sleeping" $Delay "seconds..."
    Start-Sleep $Delay

}