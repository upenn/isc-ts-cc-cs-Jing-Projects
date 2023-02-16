# Import the Azure AD module
Import-Module ExchangeOnlineManagement
 
# Connect to Azure AD
#Connect-AzureAD
# Connect to exhange online
Connect-ExchangeOnline
 
# Calculate the date and time one month ago
# $OneMonthAgo = (Get-Date).AddMonths(-1)
# $OneWeekAgo = (Get-Date).AddWeeks(-1)
 
$CurrentDate = Get-Date 
$LastWeek = $CurrentDate.AddDays(-7)
#$LastWeek = (Get-Date).AddDays(-1)
#$Last3Days = $CurrentDate.AddDays(-3)
#$Lastmonth = $CurrentDate.AddDays(-30)

# Get a list of all users created in the past month
#$Users = Get-Mailbox -ResultSize unlimited | Where-Object {[System.DateTime]$_.WhenMailboxCreated -gt $LastWeek}
$Mailboxes = Get-Mailbox -ResultSize unlimited | Where-Object {[System.DateTime]$_.WhenMailboxCreated -gt $LastWeek}

#Get-Mailbox -ResultSize unlimited -RecipientTypeDetails $mailboxTypes -Properties GrantSendOnBehalfTo, ForwardingSMTPAddress| select UserPrincipalName, DisplayName, PrimarySMTPAddress, RecipientType, RecipientTypeDetails, GrantSendOnBehalfTo, ForwardingSMTPAddress

#Get-Mailbox -ResultSize unlimited
#$Users = Get-Mailbox  | Where-Object {[System.DateTime]$_.WhenMailboxCreated -gt (get-date).AddDays(-7)}
#$UserData = @()
$MailboxData = @()

# Iterate through each user
$Mailboxes | ForEach-Object {
  #  Write-Host "Getting created date for" $_.UserPrincipalName
 
    #Collect the user data
    $MailboxData += New-Object PSObject -property $([ordered]@{

            PennID                = $_.ImmutableId
            Name                  = $_.Name
            DisplayName           = $_.DisplayName
            UserPrincipalName     = $_.UserPrincipalName
            ManagingCenter        = $_.CustomAttribute7
            ManagingCenterName    = $_.CustomAttribute5
            PCOMCenter            = $_.CustomAttribute7
            BudgetCode            = $_.CustomAttribute14
            PrimarySMTP           = $_.PrimarySmtpAddress
            AccountType           = $_.RecipientTypeDetails
            Alias                 = $_.Alias
            EmailAddresses        = $_.EmailAddresses
            ForwardingSmtpAddress = $_.ForwardingSmtpAddress
            MailboxEnabled        = $_.isMailboxEnabled
            WhenMailboxCreated    = $_.WhenMailboxCreated
            ArchiveStatus         = $_.ArchiveStatus
          #  License               = ($_.Licenses).AccountSkuId    

    })
}
 #Get-Mailbox -Identity 'jhu18@upenn.edu' | Select-Object EmployID, Name,Displayname,userPrincipalName, CustomAttribute7, CustomAttribute5,CustomAttribute14,PrimarySMTP,RecipientTypeDetails, Alias, EmailAddresses, ForwardingSmtpAddress, MailboxEnabled,WhenMailboxCreated, ArchiveStatus,AssignedPlans

#Export to CSV

#$MailboxData | Export-Csv "C:\Temp\Office365UsersCreationHistory2.csv" -NoTypeInformation
$MailboxData | Export-Csv "C:\Temp\Office365UsersCreationHistory_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" -NoTypeInformation

Write-Host $now "Office365 Users Report Generated in C:\Temp\"

###########################
 #          "PennID"                = $PennID
 #          "Name"                  = $thisUserMailboxINfo.name
 #          "DisplayName"           = $thisUserMailboxInfo.DisplayName
 #          "UserPrincipalName"     = $thisUserMailboxInfo.UserPrincipalName
 #          "ManagingCenter"        = $thisMC
 #          "ManagingCenterName"    = $thisMCName
 #          "PCOMCenter"            = $PCOMCenter
 #          "BudgetCode"            = $thisUserMailboxInfo.CustomAttribute14
 #          "PrimarySMTP"           = $thisUserMailboxInfo.PrimarySmtpAddress
 #          "AccountType"           = $thisUserMailboxInfo.RecipientTypeDetails
 #          "Alias"                 = $thisUserMailboxInfo.Alias
 #          "EmailAddresses"        = $thisUserMailboxInfo.emailaddresses -join ';'
 #          "ForwardingSmtpAddress" = $thisUserMailboxInfo.ForwardingSmtpAddress
 #          "MailboxEnabled"        = $thisUserMailboxInfo.IsMailboxEnabled
 #          "WhenMailboxCreated"    = $thisUserMailboxInfo.WhenMailboxCreated
 #          "ArchiveStatus"         = $thisUserMailboxINfo.ArchiveStatus
 #          "License"               = $thisUserLicenses