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

Function Get-CS-CalendarDelegationReport{
    param(
    [Parameter(Mandatory=$false)]
    [string]$Center = 'all'
    )

    $CenterCode = $CSCenters[$center].Center_Code
    $CenterName = $CSCenters[$center].Name

    $DateStamp = $(Get-Date -f yyyyMMdd_HHmmss).tostring()
    $MyLocation = [Environment]::GetFolderPath("Desktop")
    $ReportPath = "$($MyLocation)\CalendarDelegation_$($Center)_$DateStamp"
    $ExportCSV = "$($ReportPath).csv"

    if($Center -ne 'all' -and $Center -ne "None" -and $Center -ne '' -and $Center -notin $Script:CSCenters.Keys){
        Write-Error "Please put correct Center Name" -ErrorAction Stop
    }
    $MBUserCount=0

    Connect-ExchangeOnline -ShowBanner:$false
    Write-Progress "Getting mailboxes. Please wait..." 

    #Getting all User mailbox
    if ($Center -eq "all" -or $Center -eq ""){

        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties CustomAttribute5, CustomAttribute7, UserPrincipalName, PrimarySmtpAddress |
        Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails, CustomAttribute5, CustomAttribute7 |
        Sort-Object PrimarySmtpAddress

        Get-CS-CalendarDelegationPermissionHash

<#
    Get-EXOMailbox -ResultSize Unlimited `
      -Properties CustomAttribute5, CustomAttribute7 |
      ForEach-Object{
        $MBUserCount++
        Get-MBPermission 
#>
    }elseif ($Center -eq "none"){

        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties CustomAttribute5, CustomAttribute7, UserPrincipalName, PrimarySmtpAddress |
        Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails, CustomAttribute5, CustomAttribute7 |
        Where-Object {($_.Customattribute7 -eq "") -and ($_.CustomAttribute5 -eq "")} |
        Sort-Object PrimarySmtpAddress

        Get-CS-CalendarDelegationPermissionHash

<#

    Get-EXOMailbox -ResultSize Unlimited `
      -Properties CustomAttribute5, CustomAttribute7 |
      Where-Object {($_.Customattribute7 -eq "") -and ($_.CustomAttribute5 -eq "")} |
      ForEach-Object{
        $MBUserCount++
        Get-MBPermission
        #>
      }elseif($Center -in $Script:CSCenters.Keys){
    $center = $CSCenters[$center].ARS_MU

    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties CustomAttribute5, CustomAttribute7, UserPrincipalName, PrimarySmtpAddress |
    Select-Object DisplayName, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails, CustomAttribute5, CustomAttribute7 |
    where-object {($_.CustomAttribute7 -split ';')[0] -eq $CenterCode -or $_.CustomAttribute5 -eq $CenterName} |
    Sort-Object PrimarySmtpAddress

    Get-CS-CalendarDelegationPermissionHash
  }
  Disconnect-ExchangeOnline -Confirm:$false *> $null
}

function Get-TrusteeEmail($TrusteeUserPrincipalName) {
    $user = Get-User -Identity $TrusteeUserPrincipalName -ErrorAction SilentlyContinue
    if ($user) {
        return $user.WindowsEmailAddress
    } else {
        return $TrusteeUserPrincipalName
    }
}

function Get-CS-CalendarDelegationPermissionHash{
    Param{
        $mailboxes
    }
    # Initialize the results array
    $Result  = "" 
    $Results = @()

    # Loop through each mailbox and retrieve calendar folder permissions
    foreach ($mailbox in $mailboxes) {
        $calendarFolderPath = $mailbox.PrimarySmtpAddress.ToString() + ":\Calendar"
        $permissions = Get-MailboxFolderPermission -Identity $calendarFolderPath | Where-Object { $_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" } 
        $MBUserCount++
        Write-Progress -Activity "`n     Processed mailbox count: $MBUserCount "`n"  Currently Processing: $($mailbox.DisplayName)"

        # Add each permission to the report
        foreach ($permission in $permissions) {

            $Displayname = $mailbox.DisplayName
            $Trustee = $permission.User.DisplayName
            $AccessRights = "Calendar:" + $permission.AccessRights
            $UserPrincipalName = $mailbox.PrimarySmtpAddress   

            $TrusteeEmailAddresses = @()
            $TrusteeMailbox = Get-EXOMailbox -Identity $permission.user -ErrorAction:SilentlyContinue
            if ($TrusteeMailbox) {
                $TrusteeEmailAddresses += $TrusteeMailbox.PrimarySmtpAddress
            } else {
                $TrusteeEmailAddresses += $Trustee
            }
    
            # Combine all email addresses for the same access rights into one field, separated by a semicolon
            $AccessRightGroup = $Results | Where-Object { $_.UserPrinciPalName -eq $UserPrincipalName -and $_.AccessRights -eq $AccessRights }
            if ($AccessRightGroup) {
                $AccessRightGroup.TrusteeEmailAddress += ";" + $TrusteeEmailAddresses
            } else {
                $Result = [ordered]@{
                    DisplayName          = $Displayname
                    UserPrinciPalName    = $UserPrincipalName                
                    PCOMCenter           = ($mailbox.CustomAttribute7 -split ';')[0]
                    ManagingCenterName   = $mailbox.CustomAttribute5                
                    MailboxType          = $mailbox.RecipientTypeDetails
                    AccessRights         = $AccessRights 
                  #  Trustee              = $Trustee
                    TrusteeEmailAddress  = $TrusteeEmailAddresses -join ";"
                }
    
                $ResultRow = New-Object psobject -Property $Result 
                $Results += $ResultRow
            }

        }
    }

    $Results | Export-Csv -Path $ExportCSV -Notype -Append 
    Write-Host "Calendar permissions report saved to $ExportCSV"
}