###############################################################################################################
# DistroMaker v4
# Updates Exchange distribution list membership based on AD atributes.  
# Preferable to Dynamic Distribution Groups in two cases:
#   - where list mebmership is dependant on fields other than Container, State, Company, or Department fields 
# 	- where users want to view list membership in Outlook or other AD-aware applications.  
#
# Written by Rob Pennoyer, Silverline TG.  
# Free to distribute and modify.  
#
###############################################################################################################
#
# See Usage section below for help.
#
# Tip: try running new filters on your own before actually implementing the commands below.  
#
###############################################################################################################

# Connect to Exchange server and import Exchange PS commands
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://nyf-exchcas01.nyfoundling.org/powershell -Authentication Kerberos
Import-PSSession $session

# Query AD once to gather all users so we don't search all of AD for each list
#
# This uses Get-ADUSer (the AD command) instead of Get-User (the Exchange command) because Get-User does not
# return all intersting fields, like EmployeeType or Enabled.
#
# Filter here should be generic; you want to get the entire population of active users with mailboxes.  
#
# N.B. msExchRecipientTypeDetails "1" is UserMailbox.  This helps filter out shared mailboxes, resource 
# mailboxes, etc., as well as AD users that don't have mailboxes.
#
# (mxExchHideFromAddressLists -ne "True") doesn't evaluate properly in -filter, so we do it afterward in where.

$AllUsers = Get-ADUser -Properties * -filter {
    (msExchRecipientTypeDetails -eq "1") 
    -AND (enabled -eq "true") 
    -AND ((EmployeeType -eq "A") -OR (EmployeeType -eq "L") -OR (EmployeeType -eq "C") -OR (EmployeeType -eq "S"))
    } | where {$_.msExchHideFromAddressLists -ne "true"} 


###############################################################################################################
# Replace-DistributionGroupMembers
###############################################################################################################
#
# Replace-DistributionGroupMembers -ListName "[ListName]" -UserList ($AllUsers | where {[filter]})
#
###############################################################################################################

function Replace-DistributionGroupMembers {
	Param ([string]$ListName,[array]$UserList)
	    
	# Obtain the lists' current members and remove them from the list.
	$CurrentMembers = Get-DistributionGroupMember $ListName -ResultSize Unlimited
    ForEach ($Member in $CurrentMembers) { 
		Remove-DistributionGroupMember -Identity $ListName -Member $Member.name -Confirm:$false
		}
	
	# put the filtered ones in
	ForEach ($User in $UserList) { 
		Add-DistributionGroupMember -Identity $ListName -Member $User.name 
		}

}


###############################################################################################################
#
# Usage
#
###############################################################################################################
#
# Example 1 -- all users with department "Human Resources":
# Replace-DistributionGroupMembers -ListName "All HR" -UserList ($AllUsers | 
#	where {$_.Department -eq "Human Resources"})
#
# Example 2 -- all users with Title that includes "case" and location is "33NB":
# Replace-DistributionGroupMembers -ListName "All HR" -UserList ($AllUsers | 
#	where {$_.Title -like "*case*" -and $_.Office -eq "33NB"})
#
# Common filter variables for the where filter:
# Department, Title, Office, City, StateOrProvince, PostalCode
# Common filter comparison operators:  -eq -ne -like -notlike -
#
###############################################################################################################


Replace-DistributionGroupMembers -ListName "_Test1" -UserList ($AllUsers | 
	where {$_.Department -like "Testdept1"})

#Replace-DistributionGroupMembers -ListName "IT Department" -UserList ($AllUsers | 
#	where {$_.Department -like "Information Tech*"})







###############################################################################################################
#
# Cleanup
#
###############################################################################################################

Remove-PSSession $session
	
