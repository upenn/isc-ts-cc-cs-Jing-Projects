$data = Import-Csv -Path "C:\CS\PennO365Data\Center.csv"
$CSCenters = @{}
foreach ($row in $data) {
    $hash = [ordered]@{
        'Name' = $row.Name
        'Center_Code' = $row.Center_Code 
        'ARS_MU' = $row.ARS_MU  
        'Email' = $row.Email
    }
    $CSCenters.Add($row.ARS_MU, $hash)
}

Function Get-CS-CalendarDelegationC{

Connect-ExchangeOnline -ShowBanner:$false
$DateStamp = $(Get-Date -f yyyyMMdd_HHmmss).tostring()
$MyLocation = [Environment]::GetFolderPath("Desktop")
$ReportPath = "$($MyLocation)\CalendarDelegation_$($Center)_$DateStamp"
$ExportCSV = "$($ReportPath).csv"
# Get all user mailboxes
$mailboxes = Get-EXOMailbox -ResultSize Unlimited | Select-Object PrimarySmtpAddress, RecipientTypeDetails,CustomAttribute5 | Sort-Object PrimarySmtpAddress

# Initialize the results array
$permissionsReport = @()

# Loop through each mailbox and retrieve calendar folder permissions
foreach ($mailbox in $mailboxes) {
    $calendarFolderPath = $mailbox.PrimarySmtpAddress.ToString() + ":\Calendar"

    # Get calendar folder permissions
    # $permissions = Get-MailboxFolderPermission -Identity $calendarFolderPath | Where-Object { $_.User.DisplayName -ne ("Default" -or "Anonymous") }
    $permissions = Get-MailboxFolderPermission -Identity $calendarFolderPath | Where-Object { $_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" }


    # Add each permission to the report
    foreach ($permission in $permissions) {
        $permissionsReport += [PSCustomObject]@{
            #Mailbox            = $mailbox.DisplayName
            DisplayName         = $mailbox.DisplayName
            #Email              = $mailbox.PrimarySmtpAddress
            UserPrinciPalName   = $mailbox.PrimarySmtpAddress
            MailboxType         = $mailbox.RecipientTypeDetails
            AccessRights        = $permission.AccessRights
            #User               = $permission.User.DisplayName
            UserWithAccess      = $mailbox.UserPrinciPalName      
           # Identity            = $permission.Identity
            ManagingCenterName  = $mailbox.CustomAttribute5
            
        }
    }
}

# Export the report to a CSV file
# $permissionsReport | Export-Csv -Path "c:\temp\CalendarDelegationReport5.csv" -NoTypeInformation
$permissionsReport | Export-Csv -Path $ExportCSV -Notype -Append 
Write-Host "Calendar permissions report saved to $ExportCSV"
}