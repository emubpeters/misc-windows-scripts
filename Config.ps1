##########################################
# Written by: Ben Peters
# 5/22/2017
#
# This file should contain most of the configuration data for the scripts in this repo.
# That way, you need only edit this file, and all the scripts should work.
#
##########################################

###############
# General Variables
###############

# Who are you?
$user = 'bpeters'

# Debugging?
$ShowDebugMessages = "yes"


###############
# Variables for RDP script
###############

# How long should I wait between checks (In seconds)?
$Delay = 30

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