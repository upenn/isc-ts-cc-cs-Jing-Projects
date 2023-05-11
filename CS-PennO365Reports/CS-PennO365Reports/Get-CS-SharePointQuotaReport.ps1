function Get-CS-SharePointQuotaReport {
	[cmdletbinding()]
	Param(
		[Parameter(Mandatory=$true)]
		[ValidateSet("D7","D30","D90","D180")]
		[string]$ReportPeriod,
		[ValidateRange(70, 100)]
		[double]$QuotaThreshold = 90
	)


	#region Import Modules
	Import-Module Microsoft.Graph.Sites
	#endregion Import Modules


	#region Variables
	[double]$QuotaThresholdPercent = $QuotaThreshold / 100
	Write-Debug "Quota Threshold Percent: $($QuotaThresholdPercent)"
	#$SharePointServiceListsAndLibrariesLimit = 2000
	#$SharePointServiceUsersLimit = 1000000
	$SharePointServiceMaxFilesLimit = 300000
	#endregion Variables


	#region Connect to Microsoft Graph
	Write-Host "Connecting to Microsoft Graph" -ForegroundColor Green
	$CertificateThumbprint = "517D1E540CBEF6BCD060E28E9C4CF60D133CD68F"
	$AppId = "34d2b8ed-32ec-48c1-9df0-7e35e9481cfc"
	$TenantId = "6c4d949d-b91c-4c45-9aae-66d76443110d"
	Connect-MgGraph -AppId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
	#endregion Connect to Microsoft Graph


	#region Get SharePoint Site Usage Detail report
	$Uri = "https://graph.microsoft.com/beta/reports/getSharePointSiteUsageDetail(period='$($ReportPeriod)')"
	$ReportTmpFilename = "$(Get-Date -Format "yyyyMMddHHmmss")-SharePointSiteUsageDetail.csv"
	$ReportTmpFilePath = Join-Path -Path (Get-Location) -ChildPath $ReportTmpFilename
	# Note that this *must* output to a csv file.  There appears to be no way around this.
	Invoke-MgGraphRequest -Uri $Uri -OutputFilePath $ReportTmpFilePath
	$SharePointUsageReport = Import-Csv -Path $ReportTmpFilePath
	#endregion Get SharePoint Site Usage Detail report


	#region Site usage
	Write-Host "Getting list of sites at or above $($QuotaThresholdPercent * 100)% of quota."
	$Report = @()
	foreach ($UsageReportRow in $SharePointUsageReport) {
		Write-Debug "Analyzing $($UsageReportRow.'Site URL')"
		[double]$SiteUsagePercent = [double]$UsageReportRow.'Storage Used (Byte)' / [double]$UsageReportRow.'Storage Allocated (Byte)'
		Write-Debug "Site quota percent $($SiteUsagePercent) vs threshold of $($QuotaThresholdPercent)"
		if ($SiteUsagePercent -ge $QuotaThresholdPercent) {
			$Report += $UsageReportRow
		}
	}
	$SharePoingQuotaPercentReportFilename = "SharePoint Site Quota Report - $(Get-Date -Format "yyyy-MM-dd").csv"
	$SharePoingQuotaPercentReportFilepath = Join-Path -Path (Get-Location) -ChildPath $SharePoingQuotaPercentReportFilename
	Write-Host "Writing report to $($SharePoingQuotaPercentReportFilepath)"
	$Report | Export-Csv -Path $SharePoingQuotaPercentReportFilepath
	#endregion Site usage


	#region File Count
	Write-Host "Getting list of sites at or above $($SharePointServiceMaxFilesLimit) files"
	$Report = @()
	foreach ($UsageReportRow in $SharePointUsageReport) {
		[int]$FileCount = $UsageReportRow.'File Count'
		if ($FileCount -ge $SharePointServiceMaxFilesLimit) {
			$Report += $UsageReportRow
		}
	}
	$SharePoingFileLimitReportFilename = "SharePoint Site File Limit Report - $(Get-Date -Format "yyyy-MM-dd").csv"
	$SharePoingFileLimitReportFilepath = Join-Path -Path (Get-Location) -ChildPath $SharePoingFileLimitReportFilename
	Write-Host "Writing report to $($SharePoingQuotaPercentReportFilepath)"
	$Report | Export-Csv -Path $SharePoingFileLimitReportFilepath
	#endregion File Count

	Remove-Item -Path $ReportTmpFilePath
}