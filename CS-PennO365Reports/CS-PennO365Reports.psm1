# To use this variable in functions in this module, call it this way:
# $script:CSCenters
$CSCenters = [ordered]@{
	AC 		= @{ [string]'Name' = 'Annenberg Center for Performing Arts'; [string]'Code' = '19'; [string]'Short' = 'AC'; [string]'Email' = "o365-ac-lsps@isc.upenn.edu"}
	ASC 	= @{ [string]'Name' = 'Annenberg School for Communications'; [string]'Code' = '36'; [string]'Short' = 'ASC'; [string]'Email' = "o365-asc-lsps@isc.upenn.edu"}
	BSD 	= @{ [string]'Name' = 'Business Services'; [string]'Code' = '93'; [string]'Short' = 'BSD'; [string]'Email' = "o365-bsd-lsps@isc.upenn.edu"}
	BUS 	= @{ [string]'Name' = 'Morris Arboretum'; [string]'Code' = '60'; [string]'Short' = 'BUS'; [string]'Email' = "o365-bus-lsps@isc.upenn.edu"}
	CHAS 	= @{ [string]'Name' = 'College Houses and Academic Services'; [string]'Code' = '86'; [string]'Short' = 'CHAS'; [string]'Email' = "o365-chas-lsps@isc.upenn.edu"}
	DAR 	= @{ [string]'Name' = 'Development and Alumni Relations'; [string]'Code' = '90'; [string]'Short' = 'DAR'; [string]'Email' = "o365-dar-lsps@isc.upenn.edu"}
	DES 	= @{ [string]'Name' = 'Weitzman School of Design'; [string]'Code' = '33'; [string]'Short' = 'DES'; [string]'Email' = "o365-des-lsps@isc.upenn.edu"}
	DOF 	= @{ [string]'Name' = 'Division of Finance'; [string]'Code' = '87'; [string]'Short' = 'DOF'; [string]'Email' = "o365-dof-lsps@isc.upenn.edu"}
	DRIA 	= @{ [string]'Name' = 'Division of Recreation & Intercollegiate Athletics'; [string]'Code' = '24'; [string]'Short' = 'DRIA'; [string]'Email' = "o365-dria-lsps@isc.upenn.edu"}
	EVP 	= @{ [string]'Name' = 'Executive Vice President'; [string]'Code' = '98'; [string]'Short' = 'EVP'; [string]'Email' = "o365-evp-lsps@isc.upenn.edu"}
	FRES 	= @{ [string]'Name' = 'Facilities and Real Estate Services'; [string]'Code' = '96'; [string]'Short' = 'FRES'; [string]'Email' = "o365-fres-lsps@isc.upenn.edu"}
	GSE 	= @{ [string]'Name' = 'Graduate School of Education'; [string]'Code' = '32'; [string]'Short' = 'GSE'; [string]'Email' = "o365-gse-lsps@isc.upenn.edu"}
	HR 		= @{ [string]'Name' = 'Human Resources'; [string]'Code' = '92'; [string]'Short' = 'HR'; [string]'Email' = "o365-hr-lsps@isc.upenn.edu"}
	ISC 	= @{ [string]'Name' = 'Information Systems and Computing'; [string]'Code' = '91'; [string]'Short' = 'ISC'; [string]'Email' = "o365-isc-lsps@isc.upenn.edu"}
	ICA 	= @{ [string]'Name' = 'Institute of Contemporary Art'; [string]'Code' = '61'; [string]'Short' = 'ICA'; [string]'Email' = "o365-ica-lsps@isc.upenn.edu"}
	LAW 	= @{ [string]'Name' = 'Law School'; [string]'Code' = '56'; [string]'Short' = 'LAW'; [string]'Email' = "o365-law-lsps@isc.upenn.edu"}
	LDC 	= @{ [string]'Name' = 'Linguistic Data Consortium'; [string]'Code' = '02'; [string]'Short' = 'LDC'; [string]'Email' = "o365-ldc-lsps@isc.upenn.edu"}
	LIB 	= @{ [string]'Name' = 'University Library'; [string]'Code' = '50'; [string]'Short' = 'LIB'; [string]'Email' = "o365-lib-lsps@isc.upenn.edu"}
	MUS 	= @{ [string]'Name' = 'University Museum'; [string]'Code' = '26'; [string]'Short' = 'MUS'; [string]'Email' = "o365-mus-lsps@isc.upenn.edu"}
	NUR 	= @{ [string]'Name' = 'School of Nursing'; [string]'Code' = '06'; [string]'Short' = 'NUR'; [string]'Email' = "o365-nur-lsps@isc.upenn.edu"}
	OACP 	= @{ [string]'Name' = 'Audit Compliance and Privacy'; [string]'Code' = '78'; [string]'Short' = 'OACP'; [string]'Email' = "o365-oacp-lsps@isc.upenn.edu"}
	PDM 	= @{ [string]'Name' = 'School of Dental Medicine'; [string]'Code' = '51'; [string]'Short' = 'PDM'; [string]'Email' = "o365-pdm-lsps@isc.upenn.edu"}
	PGL 	= @{ [string]'Name' = 'Penn Global'; [string]'Code' = '62'; [string]'Short' = 'PGL'; [string]'Email' = "o365-pgl-lsps@isc.upenn.edu"}
	PSOM 	= @{ [string]'Name' = 'Perelman School of Medicine'; [string]'Code' = '40'; [string]'Short' = 'PSOM'; [string]'Email' = "o365-psom-lsps@isc.upenn.edu"}
	PRES 	= @{ [string]'Name' = 'President''s Center'; [string]'Code' = '81'; [string]'Short' = 'PRES'; [string]'Email' = "o365-pres-lsps@isc.upenn.edu"}
	PROV 	= @{ [string]'Name' = 'Provost''s Center'; [string]'Code' = '83'; [string]'Short' = 'PROV'; [string]'Email' = "o365-prov-lsps@isc.upenn.edu"}
	RHS 	= @{ [string]'Name' = 'Residential and Hospitality Services'; [string]'Code' = '95'; [string]'Short' = 'RHS'; [string]'Email' = "o365-rhs-lsps@isc.upenn.edu"}
	SAS 	= @{ [string]'Name' = 'School of Arts and Sciences'; [string]'Code' = '2'; [string]'Short' = 'SAS'; [string]'Email' = "o365-sas-lsps@isc.upenn.edu"}
	SEAS 	= @{ [string]'Name' = 'School of Engineering and Applied Science'; [string]'Code' = '13'; [string]'Short' = 'SEAS'; [string]'Email' = "o365-seas-lsps@isc.upenn.edu"}
	SP2 	= @{ [string]'Name' = 'School of Social Policy and Practice'; [string]'Code' = '35'; [string]'Short' = 'SP2'; [string]'Email' = "o365-sp2-lsps@isc.upenn.edu"}
	VET 	= @{ [string]'Name' = 'School of Veterinary Medicine'; [string]'Code' = '58'; [string]'Short' = 'VET'; [string]'Email' = "o365-vet-lsps@isc.upenn.edu"}
	VPUL 	= @{ [string]'Name' = 'Student Services'; [string]'Code' = '85'; [string]'Short' = 'VPUL'; [string]'Email' = "o365-vpul-lsps@isc.upenn.edu"}
	WEL 	= @{ [string]'Name' = 'Health and Wellness'; [string]'Code' = '89'; [string]'Short' = 'WEL'; [string]'Email' = "o365-wel-lsps@isc.upenn.edu"}
	WHA 	= @{ [string]'Name' = 'Wharton School'; [string]'Code' = '7'; [string]'Short' = 'WHA'; [string]'Email' = "o365-wha-lsps@isc.upenn.edu"}
	XPN 	= @{ [string]'Name' = 'XPN'; [string]'Code' = '81'; [string]'Short' = 'XPN'; [string]'Email' = "o365-xpn-lsps@isc.upenn.edu"}
}


<#
.SYNOPSIS
Gets mailbox statistics, and puts them into a report.

.DESCRIPTION
This function can be called in two ways.  You can either generate a report
for a specific center, or generate a series of reports, one for every center at Penn.

.PARAMETER Center
You can generate a report for a specific center using this parameter.  If you do not
specify a center, it will default to creating reports for each center.

.EXAMPLE
Get-CSBulkMailboxStatisticsReport

.EXAMPLE
Get-CSBulkMailboxStatisticsReport -Center "ISC"
#>
function Get-CS-BulkMailboxStatisticsReport {
	[cmdletbinding()]
	Param(
		[string]$Center = "ALL"
	)

	Write-Debug "Connecting to Exchange Online"
	Connect-ExchangeOnline -ShowBanner:$false

	$ReportDate = Get-Date -Format "yyyy-MM-dd"
	Write-Debug "ReportDate = $($ReportDate)"

	# Check that the center exists
	if ($Center -eq "ALL") {
		Write-Debug "Called for All Centers"
		foreach ($Center in $Script:CSCenters.Keys) {
			Write-Host "Creating mailbox statistics report for $($Center) ($($($Script:CSCenters[$Center].Name)))" -ForegroundColor Green
			$CenterReport = Get-CS-MailboxStatisticsReportForCenter -Center $($Script:CSCenters[$Center].Code)
			$CenterReportFilename = "Mailbox Statistics Report - $($($Script:CSCenters[$Center].Name)) - $($ReportDate).csv"
			Write-Host "Writing Report to ""$($CenterReportFilename)"""
			$CenterReport | Export-Csv -Path (Join-Path -Path (Get-Location).Path -ChildPath $CenterReportFilename)
		}
	}
	else {
		Write-Debug "Called for $($Center)"
		if ($Center -in $script:CSCenters.Keys) {
			Write-Debug "Calling Get-CS-MailboxStatisticsReportForCenter for report using code $(($Script:CSCenters[$Center].Code))"
			Write-Host "Creating mailbox statistics report for $($Center) ($($Script:CSCenters[$Center].Name))" -ForegroundColor Green
			$CenterReport = Get-CS-MailboxStatisticsReportForCenter -Center $($Script:CSCenters[$Center].Code)
			Write-Debug "Report returned with $($CenterReport.Count) entries"
			$CenterReportFilename = "Mailbox Statistics Report - $($($Script:CSCenters[$Center].Name)) - $($ReportDate).csv"
			Write-Debug "Report Filename: $($CenterReportFilename)"
			Write-Host "Writing Report to ""$($CenterReportFilename)"""
			$CenterReport | Export-Csv -Path (Join-Path -Path (Get-Location) -ChildPath $CenterReportFilename)
		}
		else {
			Write-Host "Invalid Center Code Specified.  Please use one of the following:" -ForegroundColor Red
			$Centers = $script:CSCenters.Keys
			Write-Host $Centers
		}
	}
	Disconnect-ExchangeOnline -Confirm:$false
}


<#
.SYNOPSIS
Internal helper function to create mailbox statistics report a report for a 
specific center

.DESCRIPTION
This internal function creates a report for a specific center, and returns 
the report as an array of PSObjects to the calling function.

This function should never be called directly.

.PARAMETER Center
The center to do the report for.

.OUTPUTS
An Array object containing PSObjects
#>
function Get-CS-MailboxStatisticsReportForCenter {
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$true)]
		[string]$CenterCode
	)

	Write-Debug "Get-CS-MailboxStatisticsReportForCenter called for center $($CenterCode)"

	# Get the mailboxes for the center
	Write-Debug "Retrieving mailboxes"
	#Write-Debug "Mailboxes Command:`n`$Mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties CustomAttribute4,CustomAttribute7,ArchiveStatus -Filter { ((CustomAttribute4 -eq ""FACULTY"") -or (CustomAttribute4 -eq ""STAFF"")) -and (CustomAttribute7 -like ""$CenterCode;*"") }"
	$CustomAttribute7SearchString = $CenterCode+ ";*"
	$Mailboxes = Get-EXOMailbox `
		-ResultSize Unlimited `
		-Properties CustomAttribute4,CustomAttribute7,ArchiveStatus `
		-Filter "((CustomAttribute4 -eq ""FACULTY"") -or (CustomAttribute4 -eq ""STAFF"")) -and (CustomAttribute7 -like ""$($CustomAttribute7SearchString)"")"
	
	$TotalMailboxes = ($Mailboxes | Measure-Object).Count
	Write-Debug "Mailboxes retrieved: $($TotalMailboxes)"

	$Report = @()

	$MailboxesProcessed = 0
	Write-Debug "Iterating through mailboxes to get stats"
	foreach ($Mailbox in $Mailboxes) {
		$PercentMailboxesProcessed = $MailboxesProcessed / $TotalMailboxes

		Write-Progress `
			-Activity 			"Getting Mailbox Statistics for Center ""$($Center)""" `
			-PercentComplete 	($PercentMailboxesProcessed * 100) `
			-Status 			"$($MailboxesProcessed) of $($TotalMailboxes) Processed" 


		$MailboxStatistics = Get-EXOMailboxStatistics -Identity $Mailbox.UserPrincipalName

		if ($Mailbox.ArchiveStatus -eq "Active") {
			$ArchiveStatistics 				= Get-EXOMailboxStatistics -Identity $Mailbox.Identity -Archive
			$ArchiveDisplayname 			= $ArchiveStatistics.DisplayName
			$ArchiveDeletedItemCount 		= $ArchiveStatistics.DeletedItemCount
			$ArchiveItemCount 				= $ArchiveStatistics.ItemCount
			$ArchiveTotalDeletedItemSize 	= $ArchiveStatistics.TotalDeletedItemSize
			$ArchiveTotalItemSize 			= $ArchiveStatistics.TotalItemSize
			# Clean up
			$ArchiveStatistics 				= $null
		}
		else {
			$ArchiveDisplayname 			= "N/A"
			$ArchiveDeletedItemCount 		= "N/A"
			$ArchiveItemCount 				= "N/A"
			$ArchiveTotalDeletedItemSize 	= "N/A"
			$ArchiveTotalItemSize 			= "N/A"
		}
		$ReportRowHash = [ordered]@{
			#"Username" = $Mailbox.Identity
			"UserPrincipalName" = $Mailbox.UserPrincipalName
			"Mailbox DisplayName" = $Mailbox.DisplayName
			"Mailbox Total Items Count" = $MailboxStatistics.ItemCount
			"Mailbox Total Items Size" = $MailboxStatistics.TotalItemSize
			"Mailbox Deleted Items Count" = $MailboxStatistics.DeletedItemCount
			"Mailbox Deleted Items Size" = $MailboxStatistics.TotalDeletedItemSize
			"Archive DisplayName" = $ArchiveDisplayname
			"Archive Total Items Count" = $ArchiveItemCount
			"Archive Total Items Size" = $ArchiveTotalItemSize
			"Archive Deleted Items Count" = $ArchiveDeletedItemCount
			"Archive Deleted Items Size" = $ArchiveTotalDeletedItemSize
		}
		$ReportRow = New-Object psobject -Property $ReportRowHash
		$Report += $ReportRow
		$MailboxesProcessed += 1
	}
	Write-Debug "Returning report"
	return $Report
}