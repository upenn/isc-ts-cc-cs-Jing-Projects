# Import Modules
Import-Module ExchangeOnlineManagement
#(Import-Module AzureAD) is enough if it's not on Jing's VB Code.
Import-Module AzureAD -UseWindowsPowerShell

Connect-ExchangeOnline
Connect-AzureAD
Connect-Graph -Scopes User.Read.All

# Calculate the date and time
$CurrentDate = Get-Date 
$LastWeek = $CurrentDate.AddDays(-7)
# $OneMonthAgo = (Get-Date).AddMonths(-1)

# Get a list of all mailboxes created in the past week
$Mailboxes = Get-EXOMailbox -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ImmutableId `
 | Where-Object {[System.DateTime]$_.WhenMailboxCreated -gt $LastWeek}

$MailboxData = @()

# Iterate through each user
$Mailboxes | ForEach-Object {
  
    #Collect the user data
    $MailboxData += New-Object PSObject -property $([ordered]@{

            PennID                = (Get-AzureADUser -ObjectId $_.UserPrincipalName).ImmutableId
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
$MailboxData |Sort-Object WhenMailboxCreated -Descending | Export-Csv "C:\Temp\WeeklyO365MailboxCreation_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" -NoTypeInformation