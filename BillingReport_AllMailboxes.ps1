#Import the Azure AD module

Import-Module ExchangeOnlineManagement
Import-Module AzureAD

Connect-ExchangeOnline
Connect-AzureAD
#Connect-MgGraph
Connect-Graph -Scopes User.Read.All

# Get a list of all users created in the past month
#$Users = Get-Mailbox -ResultSize unlimited | Where-Object {[System.DateTime]$_.WhenMailboxCreated -gt $LastWeek}
#$Mailboxes = Get-Mailbox -ResultSize unlimited | Where-Object {[System.DateTime]$_.WhenMailboxCreated -gt $LastWeek}
#$Mailboxes = Get-EXOMailbox -ResultSize 100 -Properties  CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ImmutableId
$Mailboxes = Get-EXOMailbox -ResultSize unlimited -Properties  CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ImmutableId

$MailboxData = @()

# Iterate through each user
$Mailboxes | ForEach-Object {
  #  Write-Host "Getting created date for" $_.UserPrincipalName
 
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

#$MailboxData | Export-Csv "C:\Temp\Office365UsersCreationHistory2.csv" -NoTypeInformation
$MailboxData |Sort WhenMailboxCreated -Descending | Export-Csv "C:\scr\Jhu-test\WeeklyO365BillingReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" -NoTypeInformation

