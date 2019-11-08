# Charles


# The output of this report is e-mailed
#
# 
#
# On first run a file called 'exceptionList.txt' in the same directory as itself if the file doesn't exist. This means the script will need
# read-write access to the location on the first run and read-only access from then on. Once created you can add e-mail addresses into the
# text file (one per line) that you want excluded from the report.
#
# 
#
# Requirements: The account that runs this script needs at least the "View-Only Organization Management" and the "Records Management" role in Exchange
#
# This script was created by Michael Fields
# Updated: 11/7/2019
 

)
 
# ------------------------- (START) Customize this section per your deployment -------------------------

 
# Location of script, no trailing slash
$scriptLocation = "C:\Users\Michael.Fields\Documents"
 
# E-mail stuff
# Multiple e-mail addresses should be in this format "<email1@somewhere.com>, <email2@somewhere.com>"

clear

Start-Sleep -s 10

Read-Host "Beep Boop, Greetings, I am the report bot, Charles. Please answer the following questions."
Pause -s 2
Read-Host "Which report would you like to run?"

function Show-Menu
{
    param (
        [string]$Title = 'Charles'
 )
 Clear-Host
 Write-Host "========= $Title ==============="

 Write-Host "1: Rejected Email Report-90 days."
 Write-Host "Q: Press 'q' to quit."

 Pause -s 2

 $selection = Read-Host "Please make a selection"

 switch ($selection)
 {
     '1' {
         'You chose option = '1'
     'Q' {
         'You chose option = 'Q'
     }


      pause
 }


if ($selection -eq '1' {

    $to = Read-Host "Where should I send the email?"
    Read-Host "Would you like to be cc'd?"
    $selection = Read-Host "Please make a selection."
    Write-Host "1: yes"
    Write-Host "2: no"

    switch ($selection)
 {
     'yes' {
         'You chose option = 'yes'
     'no' {
         'You chose option = 'no'
     }
 

      pause
 }

 if ($selection -eq "yes") {
    
    Read-Host "Okay, I got this boss!(Charles says enthusistically)"
    $to
    $cc = Read-Host "What is your email?"
    $from = "john.vandesteeg@usda.gov"
    $subject = "FW: MAIL Processing Rejection"
    $smtpServer = "smtp.usda.gov"
}

else ($selection -eq "no") {
    Read-Host "No Problem. Generating report"
    $to
    $from = "john.vandesteeg@usda.gov"
    $subject = "FW: MAIL Processing Rejection"
    $smtpServer = "smtp.usda.gov"
    

    
}
 
# Exchange server you want to run this PowerShell against
$exchangePowerShellServer = "exchange01.usda.gov"

# ------------------------- (END) Customize this section per your deployment -------------------------


 

}

else {$selection -eq 'q'}

Read-Host = "Beep Boop. You make Charles mad. Why you waste Charles time? Charles go away now!"

Pause -s 8

clear
 
}

   
 
# Debugging - This is useful if you're executiong the script under your account and need to test with different credentials
#$UserCredential = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://brockton-03.it.int.viu.ca/PowerShell/ -Authentication Kerberos -Credential $UserCredential
 
# Production - This assumes you're running this as a scheduled task under a user account with the proper credentials so we don't prompt for credentials
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangePowerShellServer/PowerShell/ -Authentication Kerberos
 
# Assuming we've gotten here in the script start a Exchange session
Import-PSSession $Session -AllowClobber
 
# Check if the exception list exists, create it if it doesn't and then read the exception list into the script skipping blank lines and comments
if (!(Test-Path "$($scriptLocation)\exceptionList.txt")) {
    New-Item -path $($scriptLocation) -name exceptionList.txt -type "file" -value "# One entry per line, E-mail addresses are NOT case sensitive" | Out-Null
}
$exceptionList = Get-Content "$($scriptLocation)\exceptionList.txt" | Where-Object { $_ -notmatch "#" -or $_ -notmatch "" }
 
# Get a list of transport servers in case there is more than one
$transportServers = Get-TransportService
 
# Initialize the variable so we can append to it. Initial stats array is cast differently because we can't use "Group-Object" on an ArrayList and we can't use .Remove on a Array
$mailStats = @()
[System.Collections.ArrayList]$mailStatsToDelete = @()
[System.Collections.ArrayList]$mailStatsArrayList = @()
 
# Search the message tracking log on each transport server
foreach ($server in $transportServers) {

#Set the timeframe for the search
 $startTime = (get-date -Hour 00 -Minute 00 -Second 00).AddDays(-90)
 $endTime = (get-date -Hour 23 -Minute 59 -Second 59).AddDays(-1)
 $subject = "90 day rejected email report"
 
# Search the message tracking log within a time frame on each transport server only looking for the 'RECEIVEINTERNAL' EventID
    $mailStats += Get-MessageTrackingLog -Server $server.name -Start $startTime -End $endTime -EventID 'RECEIVEINTERNAL' 
 
}
 

 
# Convert the Array to an ArrayList so we can use the .Remove method and so we can sort things later
foreach ($stat in $mailStats) {
 
    $line = "" | Select-Object Email,Total
    $line.Email = $stat.Name
    $line.Total = $stat.Group.Recipients.Count
    $mailStatsArrayList += $line
 
}
 
# If there are entries in the exception list grab all of the matches and store them in a seperate array
if ($exceptionList.Count -ne 0) {
    # Go through the list of exceptions and find any matches and store them
    foreach ($exception in $exceptionList) {
 
        foreach ($stat in $mailStatsArrayList) {
 
            if ($stat.Email -like $exception) {
 
                $mailStatsToDelete += $stat
 
            }
 
        }
 
    }
 
    # Remove the matched exceptions from the final results
    foreach ($stat in $mailStatsToDelete) {
 
        $mailStatsArrayList.Remove($stat)
 
    }
}
 
# Check and see if there are any results to report
if ($mailStatsArrayList.Count -ne 0) {
 
    # Sort and format the output into a table for the e-mail
    $results = $mailStatsArrayList | Sort-Object -Property Total -Descending | Format-Table -Property @{Expression={$_.Total};Label="Count";Width=15; Alignment="left"},@{Expression={$_.Email};Label="Sender"; Width=250; Alignment="left"} |Out-String -width 300
 
    # Send the e-mail
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($from, $to, $cc, $subject, $results)
 
}

Read-Host "Beep Boop. The script is now complete. Your email has been sent. If you have CC'd yourself, please check your Inbox. Have a super-duper day!"

Pause -s 10

# Clean-up our session
Remove-PSSession $Session
