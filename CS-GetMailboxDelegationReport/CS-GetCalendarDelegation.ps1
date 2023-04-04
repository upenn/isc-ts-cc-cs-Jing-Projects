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


    $DateStamp = $(Get-Date -f yyyyMMdd_HHmmss).tostring()
    $MyLocation = [Environment]::GetFolderPath("Desktop")
    $ReportPath = "$($MyLocation)\CalendarDelegation_$($Center)_$DateStamp"
    $ExportCSV = "$($ReportPath).csv"

    Connect-ExchangeOnline -ShowBanner:$false
    Write-Progress "Getting mailboxes. Please wait..." 

    # Get all user mailboxes
  #  $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties CustomAttribute5 | Select-Object PrimarySmtpAddress, RecipientTypeDetails CustomAttribute5 | Sort-Object PrimarySmtpAddress
    $mailboxes = Get-EXOMailbox -ResultSize 10 -Properties CustomAttribute5,CustomAttribute7,UserPrincipalName,PrimarySmtpAddress | Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails, CustomAttribute5, CustomAttribute7 | Sort-Object PrimarySmtpAddress

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
                DisplayName         = $mailbox.DisplayName         
                UserPrinciPalName   = $mailbox.PrimarySmtpAddress
                MailboxType         = $mailbox.RecipientTypeDetails
                AccessRights        = "Calendar:" + $permission.AccessRights
                Trustee             = $permission.User.DisplayName

                trusteeEmailAddress = (Get-EXORecipient -Identity $permission.User.DisplayName).PrimarySmtpAddress
                PCOMCenter           = ($mailbox.CustomAttribute7 -split ';')[0]
            }
            
        }
    }
# Export the report to a CSV file
# $permissionsReport | Export-Csv -Path "c:\temp\CalendarDelegationReport5.csv" -NoTypeInformation
    $permissionsReport | Export-Csv -Path $ExportCSV -Notype -Append 
    Write-Host "Calendar permissions report saved to $ExportCSV"
}

#########for 4/4, need to re=write foreach in sepeate function, add try/catch to get trustee's email address.
#### add center name.