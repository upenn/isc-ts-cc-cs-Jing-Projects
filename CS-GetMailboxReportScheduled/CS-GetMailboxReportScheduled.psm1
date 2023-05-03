<#
Get-CS-MailboxAIOScheduled
 * it's for task scheduler to schedule weekly and daily mailbox report
 * this script has to run in PowerShell, otherwise certificate authentication won't work.
 * Parameter - Freq mandantory and positional.
 * 2 ps1 files called by task scheduler to run both reports
 * report will be on C:\CS\PennO365Data\MailboxReport\
#>
$data = Import-Csv -Path "C:\CS\PennO365Data\Center.csv"
$CSCenters = @{}
foreach ($row in $data) {
    $hash = [ordered]@{
        'Name' = $row.Name
        'Center_Code' = $row.Center_Code # was code
        'ARS_MU' = $row.ARS_MU  # was short
        'Email' = $row.Email
    }
    $CSCenters.Add($row.ARS_MU, $hash)
}
<#
    .Synopsis
       This cmdlet is all in one version for mailbox report. 
    .DESCRIPTION
       this cmdlet returns a report of all mailboxes in all organization with daily report or weekly.
       daily report returns all new mailboxes created in previous day, if it runs on Monay, it returns all mailboxes created since Last Friday till Sunday.
       weekly report returns all mailboxes in the tenant.
    .EXAMPLE
        Get-CS-MailboxAIOScheduled Daily
    .EXAMPLE
       Get-CS-MailboxAIOScheduled Weekly
    .INPUTS
       parameters
    .OUTPUTS
       CSV file in user's desktop
#>
function Get-CS-MailboxAIOScheduled{
    Param(
        [Parameter(Mandatory=$false)]
        [string]$Freq,
        [string]$Center = 'all',
        [datetime]$StartDate,
        [datetime]$EndDate        
        )

        $CenterCode = $CSCenters[$center].Center_Code
        $CenterName = $CSCenters[$center].Name
                   
        $DateStamp = $(Get-Date -f yyyyMMdd_HHmmss).tostring()
        $MyLocation = [Environment]::GetFolderPath("Desktop")
#        $ReportPath = "$($MyLocation)\MailboxReport_"
        $ReportPath = "C:\CS\PennO365Data\MailboxReport\MailboxReport_"
       
    $ReadOnlyClientId  = "34d2b8ed-32ec-48c1-9df0-7e35e9481cfc"
    $TenantId = "6c4d949d-b91c-4c45-9aae-66d76443110d"
    $certSubject = "CN=aad-aar-ro.penno365.isc.upenn.edu"
    $Organization = "penno365.onmicrosoft.com"    
    $Cert = Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object {$_.Subject -eq $certSubject}
    $CertThumbprint = $cert.Thumbprint
    
    Connect-ExchangeOnline -ShowBanner:$false `
                           -AppId 					$ReadOnlyClientId `
                           -Organization 			$Organization `
                           -CertificateThumbprint 	$certThumbprint `
                           -InformationAction SilentlyContinue | Out-Null
   # Write-Output "connected to exchangeonline!"

    Connect-MgGraph -ClientId $ReadOnlyClientId `
                    -TenantId $TenantId `
                    -CertificateThumbprint 	$certThumbprint `
                    -ErrorAction SilentlyContinue `
                    -Errorvariable ConnectionError `
                    -InformationAction SilentlyContinue | Out-Null

   # Write-Output "connected to MgGraph!"
    Select-MgProfile -Name "beta" 
   # Write-Output "connected to MgProfile!"
   # Write-Progress "Getting mailboxes. Please wait..." 
    $ReadOnlyClientId  = $null
    $TenantId = $null
    $certSubject = $null
    $Organization = $null   
    $Cert = $null
    $CertThumbprint = $null

    if ($Freq -eq "Daily"){
        $StartDate = (Get-ReportDate).StartDate
        #Write-Output "$StartDate"
        $EndDate = (Get-ReportDate).EndDate
        #Write-Output "$EndDate"
 
        $Mailboxes = Get-EXOMailbox -ResultSize unlimited -Filter "(WhenMailboxCreated -ge '$($Startdate.ToUniversalTime())') -and (WhenMailboxCreated -lt '$($Enddate.ToUniversalTime())')" `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress 
                   
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        Get-Output
    }
    elseif($Freq -eq "Weekly"){
        $StartDate = (Get-Date).AddYears(-99)
        Write-Output "Weekly: $StartDate"
        $EndDate = Get-Date
        Write-Output "Weekly: $EndDate"

        $Mailboxes = Get-EXOMailbox -ResultSize unlimited -Filter "WhenMailboxCreated -ge '$Startdate' -and WhenMailboxCreated -le '$Enddate'" `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ForwardingSmtpAddress
                
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        Get-Output
    }

   Disconnect-ExchangeOnline -Confirm:$false *> $null
   Disconnect-MgGraph | Out-Null
}
<#
    internal function, hash table to create report.
#>
function Get-CS-AllMailboxHash{
    Param(
        [Parameter(Mandatory=$false)]
        $Mailboxes
    )
        $total = $Mailboxes.count
        if ($total -eq 0){
            Write-Error "There is no mailbox!" -ErrorAction Stop
        }
        $i = 0
        $MailboxData = @()
        $Mailboxes | ForEach-Object {
          #  $i++
          #  $percent = ($i / $total * 100)
          #   Write-Progress -Activity "Processing mailboxes" -PercentComplete $percent -Status "$i of $total processed"
           # Write-Progress -Activity "`n     Exported user count:$MailboxCount "`n"Currently Processing:$UPN"
           #  $MailboxCount = $Mailboxes.count

           $MailboxHash = [ordered]@{
            PennID                = (Get-MgUser -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue).EmployeeId
            #PennID                = ''
            Name                  = ($_.UserPrincipalName -split '@')[0]
            DisplayName           = $_.DisplayName
            UserPrincipalName     = $_.UserPrincipalName
            ManagingCenter        = ($_.CustomAttribute7 -split ';')[0]
            ManagingCenterName    = $_.CustomAttribute5
            PCOMCenter            = ($_.CustomAttribute7 -split ';')[0]
            BudgetCode            = $_.CustomAttribute14
            PrimarySMTP           = $_.PrimarySmtpAddress
            AccountType           = $_.RecipientTypeDetails
            Alias                 = $_.Alias
            EmailAddresses        = $_.EmailAddresses -join ';'
            ForwardingSmtpAddress = $_.ForwardingSmtpAddress
            MailboxEnabled        = $_.isMailboxEnabled
            WhenMailboxCreated    = $_.WhenMailboxCreated
            ArchiveStatus         = $_.ArchiveStatus
            License               = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue).SkuPartNumber -join ";"
            #License               = ''
       }
       <#
       try {
        $MgUser = Get-MgUser -UserId $_.UserPrincipalName 
        #$MgUser = Get-MgUser -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue
        $MailboxHash.PennID = $MgUser.EmployeeId
        #$MailboxHash.License = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue).SkuPartNumber -join ";"
        $MailboxHash.License = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName).SkuPartNumber -join ";"
        $ErrorUser = $MgUser.UserPrincipalName
    }
    catch {
        
        Write-Warning "Failed to retrieve user's PennID and License information for $ErrorUser" -ErrorAction:SilentlyContinue
    } #>
            $MailboxRow = New-Object psobject -Property $MailboxHash
            $MailboxData += $MailboxRow
        }
       return $MailboxData | Sort-Object WhenMailboxCreated -Descending
}
function Get-Output {
    
    $ReportPath = $ReportPath + $($Freq)+"_"
    $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
    Write-Output "Please find $($Center)_Mailboxes report in $($ReportPath+$DateStamp).csv"   
}

function Get-ReportDate {

    # Get the current date and time
    $now = Get-Date

    # If it's Monday, get the date range for Friday to Sunday
    if ($now.DayOfWeek -eq "Monday") {
        $lastFriday = (Get-Date).AddDays(-3).Date
        $startDate = Get-Date "$($lastFriday.ToString("MM/dd/yyyy")) 00:00:00"
        $yesterday = (Get-Date).AddDays(-1).Date
        $endDate = Get-Date "$($yesterday.ToString("MM/dd/yyyy")) 23:59:59"
    }
    # For all other days, get the date range for the previous day
    else {
        $yesterday = (Get-Date).AddDays(-1).Date
        $startDate = Get-Date "$($yesterday.ToString("MM/dd/yyyy")) 00:00:00"
        $endDate = Get-Date "$($yesterday.ToString("MM/dd/yyyy")) 23:59:59"
    }

    # Return the start and end dates as an array or custom object
    return @{
        StartDate = $startDate
        EndDate = $endDate
    }
}

