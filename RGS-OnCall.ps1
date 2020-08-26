<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2014 v4.1.67
	 Created on:   	30/10/2014 8:13 p.m.
	 Created by:   	Paul.Bloem
	 Blog: 	UCSorted.com
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>
# Script Config
Import-Module Lync
#The following Items need to be manually edited
#
$Team = "IT Support" 						#The Team the script is used for. This is txt only and only for labelling purposes
$queue = "IT Support On Call Queue" 				#Lync Queue Name. MUST be a valid RGS Queue. Can be queried with Get-CsRgsQueue |fl name
$sipdomain = "@lyncsorted.co.nz" 				#Lync Sip Domain. Can be queried with Get-CsSipdomain
$lyncpool = "service:ApplicationServer:frontendpool.lyncsorted.local" 	#Lync front end pool identity. Can be queried with Get-CsPool
$logfile = "\\<network share>\On Call\Lync\IT Support-log.txt"	#Specify a log file for this script

# Menu
#The Menu items below are displayed as TXT for selection purpose only. Once selected the next section (Selection Logic) is engaged
cls

Write-Host "* Change" $team "On Call Team Member *" -ForegroundColor Cyan
Write-Host " "
Write-Host "Change the" $team "On Call Team member to:"
Write-Host " "
Write-Host "1. Engineer 1" -foregroundcolor Yellow
Write-Host "2. Engineer 2" -foregroundcolor Yellow
Write-Host "3. Engineer 3" -foregroundcolor Yellow
Write-Host "4. Engineer 4" -foregroundcolor Yellow
Write-Host "5. Engineer 5" -ForegroundColor Yellow
Write-Host "6. Engineer 6" -ForegroundColor Yellow
Write-Host "7. Other" -foregroundcolor Yellow
Write-Host " "
$a = Read-Host "Select 1-7: "
Write-Host " "

# Selection Logic to set sip URI based on Menu Selection
# This needs to be manually edited to match the above Menu items. Also the Engineers DID needs to be edited as fwddest - in E.164 format

switch ($a)
{
	1 {
		$fwdname = "Engineer 1"
		$fwddest = "+6495501111"
	  }
	2 {
		$fwdname = "Engineer 2"
		$fwddest = "+6495502222"
	  }
	3 {
		$fwdname = "Engineer 3"
		$fwddest = "+6495503333"
	  }
	4 {
		$fwdname = "Engineer 4"
		$fwddest = "+6495504444"
	  }
	5 {
		$fwdname = "Engineer 5"
		$fwddest = "+6495505555"
	  }
	6 {
		$fwdname = "Engineer 6"
		$fwddest = "+6495506666"
	  }
	7 {
		# Other
		Write-Host "Enter phone number in E.164 format e.g.: " -NoNewLine
		Write-Host "+6495501234" -ForegroundColor Magenta
		$fwddest = Read-Host "Phone Number"
	}
}

# Changing E.164 phone numbers to sip format
$fwddest = "sip:" + $fwddest
$fwddest = $fwddest + $sipdomain

# Changing on call number for the RGS Queue
$b = Get-CsRgsQueue -Identity $lyncpool -Name $queue
$b.OverflowAction = New-CsRgsCallAction -Action TransferToPstn -Uri $fwddest
Set-CsRgsQueue -Instance $b

# Write to log file
$date = Get-Date
Write-output "$date, $team, $fwdname, $fwddest" | Out-File -append $logfile

Write-Host "On call person changed to: " -NoNewline
Write-Host $fwdname "-" $fwddest -ForegroundColor Green
Write-Host " "
