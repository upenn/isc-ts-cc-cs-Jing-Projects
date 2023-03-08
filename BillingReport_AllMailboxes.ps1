Import-Module ExchangeOnlineManagement
Import-Module AzureAD # -UseWindowsPowerShell (need to be added on VB Code)

Connect-ExchangeOnline
Connect-AzureAD
Select-MgProfile -Name "beta"
Connect-Graph -Scopes User.Read.All

# Get all mailboxes in the tenant
$Mailboxes = Get-EXOMailbox -ResultSize unlimited -Properties  CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ImmutableId

$MailboxData = @()

# Iterate through each user
$Mailboxes | ForEach-Object {

    #Collect the user data
    $MailboxData += New-Object PSObject -property $([ordered]@{

         
            #PennID                = (Get-AzureADUser -ObjectId $_.UserPrincipalName).ImmutableId
            PennID                = (Get-MgUser -UserId $_.UserPrincipalName).EmployeeId
            Name                  = $_.Name
            DisplayName           = $_.DisplayName
            UserPrincipalName     = $_.UserPrincipalName
            ManagingCenter        = ($_.CustomAttribute7 -split ';')[0]
            ManagingCenterName    = $_.CustomAttribute5
            PCOMCenter            = ($_.CustomAttribute7 -split ';')[0]
            BudgetCode            = $_.CustomAttribute14
            PrimarySMTP           = $_.PrimarySmtpAddress
            AccountType           = $_.RecipientTypeDetails
            Alias                 = $_.Alias
            EmailAddresses        = $_.EmailAddresses
            ForwardingSmtpAddress = $_.ForwardingSmtpAddress
            MailboxEnabled        = $_.isMailboxEnabled
            WhenMailboxCreated    = $_.WhenMailboxCreated
            ArchiveStatus         = $_.ArchiveStatus
            License               = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName).SkuPartNumber -join ";"
          
    })
}

#Export to CSV
$MailboxData |Sort-Object WhenMailboxCreated -Descending | Export-Csv "C:\Temp\WeeklyO365BillingReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" -NoTypeInformation