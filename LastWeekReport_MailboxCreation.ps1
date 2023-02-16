# Import the Azure AD module
Import-Module AzureAD
 
# Connect to Azure AD
Connect-AzureAD
 
# Calculate the date and time one month ago
# $OneMonthAgo = (Get-Date).AddMonths(-1)
# $OneWeekAgo = (Get-Date).AddWeeks(-1)
 
$CurrentDate = Get-Date
$LastWeek = $CurrentDate.AddDays(-7)
$Lastmonth = $CurrentDate.AddDays(-30)

# Get a list of all users created in the past month
$Users = Get-AzureADUser -All:$true | Where-Object {[System.DateTime]$_.ExtensionProperty.createdDateTime -gt $LastWeek}
 
$UserData = @()
# Iterate through each user
$Users | ForEach-Object {
  #  Write-Host "Getting created date for" $_.UserPrincipalName
 
    #Collect the user data
    $UserData += New-Object PSObject -property $([ordered]@{

            PennID                = $_.ImmutableId
          # Name                  = $_.name
            DisplayName           = $_.DisplayName
            UserPrincipalName     = $_.UserPrincipalName
          #  ManagingCenter       = $thisMC
          #  ManagingCenterName   = $thisMCName
          #  PCOMCenter           = $PCOMCenter
          #  BudgetCode            = $_.CustomAttribute14
          #  PrimarySMTP           = $_.PrimarySmtpAddress
          #  AccountType           = $_.RecipientTypeDetails
          #  Alias                 = $_.Alias
          #  EmailAddresses        = $_.emailaddresses -join ';'
          #  ForwardingSmtpAddress = $_.ForwardingSmtpAddress
          #  MailboxEnabled        = $_.IsMailboxEnabled
          #  WhenMailboxCreated    = $_.WhenMailboxCreated
          #  ArchiveStatus         = $_.ArchiveStatus
          #  License              = ($_.Licenses).AccountSkuId    
           CreatedDateTime       = $_.ExtensionProperty.createdDateTime
        
    })
}
 
#Export to CSV

$UserData | Export-Csv "C:\scr\Jhu-test\Office365UsersCreationHistory.csv" -NoTypeInformation



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