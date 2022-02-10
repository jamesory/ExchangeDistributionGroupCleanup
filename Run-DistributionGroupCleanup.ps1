<#  
.SYNOPSIS  
	Distribution Group Cleanup Script - for Exchange 203 and 2016
 
.DESCRIPTION  
	This script will cleanup all scripts that have not been used within a certain timeframe.
 
.NOTES  
    Version					: 1.3 - Removed unneeded functions. This will send an email with all inactive groups to IT Email Administrators and sends the email with one group per line.
    Date Created			: 12/29/2019
    Change Log			 	: 1.1 - Change 'GetTransportServer' to 'Get-TransportService' and added a variable list for $To, $From and $SMTPServer
        				 	: 1.0 - Script first set up
    Wish list				: HTML Report to IT?
    Rights Required			: Local admin on server
    Sched Task Req'd		: No
    Exchange Version		: 2013 / 2016
    Author					: Damian Scoles
    Dedicated Blog			: http://justaucguy.wordpress.com/
    Disclaimer				: You are on your own.  This was not written by, support by, or endorsed by Microsoft.
    Code stolen from		: None 

.EXAMPLE
        .\Run-DistributionGroupCleanup.ps1
		To be run once per month as a recurring task
 
.INPUTS
		None. You cannot pipe objects to this script.
#>

# Global variable section

# Testing Dates - use this for finding older results
# $onemonth = ((get-date).addmonths(-13))
# $current = ((get-date).addmonths(-12))

# Production Dates
$current = get-date
$onemonth = ((get-date).addmonths(-3))

# Arrays
$activegroups2 = @()
$activegroups = @()
$inactivegroups = @() 
$allgroups = @()
$smtp = @()


# Other Variables - below are samples, make sure to change for your environment
$From = "exchange@test.com"
$SMTPServer = "mail.test.com"
$To = "it-alert-mail@ttest"
$AdminAddress = "exchange@test.com"

# Load AD Module for PowerShell
import-module activedirectory
# Load Exchange Powershell Module
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

	
# Get a list of the active groups
$servers = get-transportservice
foreach ($name in $servers) {
     $activegroups2 += (Get-MessageTrackingLog -Server $name.name -EventId Expand -ResultSize Unlimited -start $onemonth -end $current | Sort-Object RelatedRecipientAddress | Group-Object RelatedRecipientAddress | Sort-Object Name | select-object name)
}

$activegroups2 = $activegroups2 | sort-object name | group-object name
foreach ($line in $activegroups2) {
	$activegroups += $line.name
}

# Get a list of all groups
$allgroups2 = get-distributiongroup -resultsize unlimited | Select-Object -Property @{Label="Name";Expression={$_.PrimarySmtpAddress}}
foreach ($line in $allgroups2) {
	$allgroups += $line.name
}

# Find inactive groups by comparing active groups to all groups
$InactiveGroups2 = Compare-Object $activegroups $allgroups
foreach ($line in $inactivegroups2) {
	$smtp2=$line.inputobject
	$address=$smtp2.local+"@"+$smtp2.domain
	$inactivegroups += $address
}

# Set custom attribute 10 for active groups to 0
foreach ($line in $ActiveGroups){
 	set-distributiongroup -identity $line -CustomAttribute10 0 -warningaction silentlycontinue
}

# Set custom attribute 10 for inactive groups - increase by 1
# Hide or disable group
foreach ($line in $InactiveGroups){
		[string]$email = $line
		[int]$number = (get-distributiongroup -identity $email).CustomAttribute10
		$number += 1
		set-distributiongroup -identity $email -CustomAttribute10 $number 
          
	
		}
		if ($number += 1) {
        $OFS = "`r`n"
        [String]$InactiveGroups
		Send-MailMessage -From "$From" -To "$to" -Subject  "Distribution Group - IT Alert - Distribution Groups Not emailed in the past 90 days" -Body "$InactiveGroups" -SmtpServer "$SMTPServer"
                      }
	
	