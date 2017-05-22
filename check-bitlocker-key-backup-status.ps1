###########################
# Author: Ben Peters (bpeters@emich.edu)
#
# This script will double check a group in active directory for all computer accounts in there, and will report any 
# computers that do not have keys stored in Active Directory or the MBAM database.
###########################

# Basic Configs
$ADComputerGroup = 'some_group_name'
$SQLServer = "mbamdb.server.example.com"
$SQLDBName = "MBAM Recovery and Hardware"
$path = "DC=something,DC=somethingelse,DC=com"

Import-Module ActiveDirectory

$TestOutput = "==================================`n"
$TestOutput = $TestOutput + "Individual Test Output`n"
$TestOutput = $TestOutput + "==================================`n`n"
$ComputersWithNoADKey = ""
$ComputersWithNoMBAMKey = ""
$NumberOfComputers = 0
$NumberMissingMBAMKeys = 0
$NumberMissingADKeys = 0

# get all the members of the bitlocker fixed drives group
$Members = Get-ADGroupMember -identity "$ADComputerGroup" | where {$_.objectclass-eq “computer”} | Select-Object -Property name


# Go through each computer in the group
foreach ($Name in $Members) {
    
    # Increment counter
    $NumberOfComputers++

    # Get details about this computer
    $Data = Get-ADComputer $Name.name -Properties distinguishedName,name,lastLogon,location,operatingSystem,operatingSystemServicePack,operatingSystemVersion
    
    # Format some output
    $TestOutput = $TestOutput + $Name.name + "`n"
    $TestOutput = $TestOutput +  "-------------------------" + "`n"
    $TestOutput = $TestOutput + "     " + $Data.operatingSystem + "/" + $Data.operatingSystemServicePack + "/" + $Data.operatingSystemVersion + "`n"
    
    #####################
    # See if it has keys stored in AD or not.
    #####################
    $path = $Data.distinguishedName
    $children = Get-Childitem -Path AD:\$path | where {$_.objectclass-eq “msFVE-RecoveryInformation”}

    # If keys are found, great.  If not, mention that.
    if (!$children) {
        $TestOutput = $TestOutput +   "     ***No AD key found!***" + "`n"
        $ComputersWithNoADKey = $ComputersWithNoADKey + $Name.name+ " : " + [datetime]$Data.lastLogon + " : " + $Data.location + "`n"
        $NumberMissingADKeys++
    } else {
        $TestOutput = $TestOutput +   "     AD Key OK." + "`n"
    }

    ####################
    # See if it has keys stored in the MBAM database
    ####################
    $SQLQuery = "SELECT LastUpdateTime FROM [MBAM Recovery and Hardware].[RecoveryAndHardwareCore].[Machines] WHERE Name like '$Name'"
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True" 
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $SqlQuery
    $SqlCmd.Connection = $SqlConnection 
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd 
    $DataSet = New-Object System.Data.DataSet
    $SqlAdapter.Fill($DataSet) 
    $SqlConnection.Close()
 
    if (!$DataSet.Tables[0]) {
        $TestOutput = $TestOutput +   "     ***No MBAM key found!***" + "`n"
        $ComputersWithNoMBAMKey = $ComputersWithNoMBAMKey + $Name.name + " (" + $path + ")`n"
        $NumberMissingMBAMKeys++
    } else {
        $TestOutput = $TestOutput +   "     MBAM key OK." + "`n"
    }

    $TestOutput = $TestOutput +   "`n"
    clear

}

###########################
# Verify keys are in MBAM if they are in AD
###########################
$keys = Get-ADObject -LDAPFilter "(objectClass=msFVE-RecoveryInformation)" -SearchBase $path -SearchScope Subtree
$ADToMBMMissingNum = 0
$ADtoMBAMCheckOutput = ""
$ADtoMBAMCheckOutput = $ADtoMBAMCheckOutput + "`n`n"
$ADtoMBAMCheckOutput = $ADtoMBAMCheckOutput + "==================================`n"
$ADtoMBAMCheckOutput = $ADtoMBAMCheckOutput + "Keys In AD, But Missing From MBAM: " + $ADToMBMMissingNum + "`n"
$ADtoMBAMCheckOutput = $ADtoMBAMCheckOutput + "==================================`n"
$ADtoMBAMCheckOutput = $ADtoMBAMCheckOutput + "`n`n"

foreach ($item in $keys) {
    $data = $item.ToString()
    $data = $data.Substring(29,36)
    $data = $data.ToLower()


    ####################
    # See if it has keys stored in the MBAM database
    ####################
    $SQLServer = "mbamdb.ad.emich.edu"
    $SQLDBName = "MBAM Recovery and Hardware"
    $SQLQuery = "SELECT * FROM [MBAM Recovery and Hardware].[RecoveryAndHardwareCore].[Keys] WHERE RecoveryKeyId = '`$data'"
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True" 
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $SqlQuery
    $SqlCmd.Connection = $SqlConnection 
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd 
    $DataSet = New-Object System.Data.DataSet
    $SqlAdapter.Fill($DataSet) 
    $SqlConnection.Close()
 
    if (!$DataSet.Tables[0]) {
       $ADtoMBAMCheckOutput = $ADtoMBAMCheckOutput + $data + " not found in MBAM, but is in AD!`n"
       $ADToMBMMissingNum++
    } 

    clear
}

# Compile output from AD test
$ADOutput = ""
$ADOutput = $ADOutput + "`n`n"
$ADOutput = $ADOutput + "==================================`n"
$ADOutput = $ADOutput + "Computers Missing Keys In AD: " + $NumberMissingADKeys + "`n"
$ADOutput = $ADOutput + "==================================`n"
$ADOutput = $ADOutput + $ComputersWithNoADKey
$ADOutput = $ADOutput + "`n`n"

# Compile output from MBAM test
$MBAMOutput = ""
$MBAMOutput = $MBAMOutput + "==================================`n"
$MBAMOutput = $MBAMOutput + "Computers Missing Keys In MBAM: " + $NumberMissingMBAMKeys + "`n"
$MBAMOutput = $MBAMOutput + "==================================`n"
$MBAMOutput = $MBAMOutput + $ComputersWithNoMBAMKey
$MBAMOutput = $MBAMOutput + "`n`n"

# Compile final output
$FinalOut = $ADOutput + $MBAMOutput + $ADtoMBAMCheckOutput + $TestOutput

# Post final output
write-host $FinalOut