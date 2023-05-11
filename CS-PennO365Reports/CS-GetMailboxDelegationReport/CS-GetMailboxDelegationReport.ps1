<#
Get-CS-MailboxDelegationReport with paramter: 
 * This cmdlet is getting  Mailbox Delegation Report by Center or for All center.
 * report by center ARS_MU number: -center ISC
 * report by all centers: -center all
 * report will be on current user's Desktop
#>

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

<#
    .Synopsis
       This cmdlet is getting  Mailbox  Delegation Report by Center or for All center.
    .DESCRIPTION
       By default, this cmdlet returns a report of all mailboxe calendars in the organization.
       -center is optional, by default is for All.
    .EXAMPLE
        Get-CS-MailboxDelegationReport All
    .EXAMPLE
       Get-CS-MailboxDelegationReport
    .EXAMPLE
       Get-CS-MailboxDelegationReport -Center ISC
    .INPUTS
       parameters
    .OUTPUTS
       CSV file in current user's desktop.
#>
function Get-CS-MailboxDelegationReport{

  param(
    [Parameter(Mandatory=$false)]
    [string]$Center = 'all',
    [switch]$FullAccess,
    [switch]$SendAs,
    [switch]$SendOnBehalf
  )

  $CenterCode = $CSCenters[$center].Center_Code
  $CenterName = $CSCenters[$center].Name
  
  $DateStamp = $(Get-Date -f yyyyMMdd_HHmmss).tostring()
  $MyLocation = [Environment]::GetFolderPath("Desktop")
  $ReportPath = "$($MyLocation)\MailboxDelegation_$($Center)_$DateStamp"
  $ExportCSV = "$($ReportPath).csv"

  if($Center -ne 'all' -and $Center -ne "None" -and $Center -ne '' -and $Center -notin $Script:CSCenters.Keys){
    Write-Error "Please put correct Center Name" -ErrorAction Stop
  }


  Connect-ExchangeOnline -ShowBanner:$false 
  Write-Progress "Getting mailboxes. Please wait..." 

  $MBUserCount=0
  
 #Check for AccessRights filter
  if(($FullAccess.IsPresent) -or ($SendAs.IsPresent) -or ($SendOnBehalf.IsPresent)){}
  else{
    $FilterPresent='False'
  }
 
 #Getting all User mailbox
if ($Center -eq "all" -or $Center -eq ""){

  
  Get-EXOMailbox -ResultSize Unlimited `
    -Properties CustomAttribute5, CustomAttribute7 |
    ForEach-Object{
      $MBUserCount++
      Get-MBPermission
  }
}elseif ($Center -eq "none"){
  Get-EXOMailbox -ResultSize Unlimited `
    -Properties CustomAttribute5, CustomAttribute7 |
    Where-Object {($_.Customattribute7 -eq "") -and ($_.CustomAttribute5 -eq "")} |
    ForEach-Object{
      $MBUserCount++
      Get-MBPermission
    }
}elseif($Center -in $Script:CSCenters.Keys){
  $center = $CSCenters[$center].ARS_MU
  Get-EXOMailbox -ResultSize unlimited `
    -Properties CustomAttribute5, CustomAttribute7 | 
    where-object {($_.CustomAttribute7 -split ';')[0] -eq $CenterCode -or $_.CustomAttribute5 -eq $CenterName} | 
    ForEach-Object{
      $MBUserCount++
      Get-MBPermission
    }
  }
# Write-Host "Script executed successfully"
  if((Test-Path -Path $ExportCSV) -eq "True"){
    write-host "Please find report: $ExportCSV"
  }
  Else{
    Write-Host No mailbox found!
  }

  Disconnect-ExchangeOnline -Confirm:$false *> $null
}

#Getting Mailbox permission
function Get-MBPermission{
    $upn                = $_.UserPrincipalName
    $DisplayName        = $_.Displayname
    $MBType             = $_.RecipientTypeDetails
    $Print              = 0 
    Write-Progress -Activity "`n     Processed mailbox count: $MBUserCount "`n"  Currently Processing: $DisplayName"
  
    #Getting delegated Full access permission for mailbox
    if(($FilterPresent -eq 'False') -or ($FullAccess.IsPresent)){
      $FullAccessPermissions = (Get-EXOMailboxPermission -Identity $upn |
                              where{ ($_.AccessRights -contains "FullAccess") `
                              -and ($_.IsInherited -eq $false) `
                              -and -not ($_.User -match "NT AUTHORITY" ) }).User
  
      if([string]$FullAccessPermissions -ne ""){
        $Print = 1
        $UserWithAccessEmail = ""
        $AccessRights = "FullAccess"
        foreach($FullAccessPermission in $FullAccessPermissions){
          $UserWithAccessEmail = $UserWithAccessEmail+$FullAccessPermission
          if($FullAccessPermissions.indexof($FullAccessPermission) -lt (($FullAccessPermissions.count)-1)){
            $UserWithAccessEmail = $UserWithAccessEmail+";"
          }
        }
        Get-MBPermissionHash
      }
    }  
    #Getting delegated SendAs permission for mailbox
    if(($FilterPresent -eq 'False') -or ($SendAs.IsPresent)){
      $SendAsPermissions = (Get-EXORecipientPermission -Identity $upn | 
                            where{ -not (($_.Trustee -match "NT AUTHORITY") )}).Trustee
      if([string]$SendAsPermissions -ne ""){
        $Print = 1
        $UserWithAccessEmail = ""
        $AccessRights = "SendAs"
        foreach($SendAsPermission in $SendAsPermissions){
  
          $UserWithAccessEmail = $UserWithAccessEmail+$SendAsPermission
          if($SendAsPermissions.indexof($SendAsPermission) -lt (($SendAsPermissions.count)-1)){
            $UserWithAccessEmail = $UserWithAccessEmail+";"
          }
        }
        Get-MBPermissionHash
      }
    }  
    #Getting delegated SendOnBehalf permission for mailbox
    if(($FilterPresent -eq 'False') -or ($SendOnBehalf.IsPresent)){
      $SendOnBehalfPermissions = (Get-EXOMailbox -Identity $upn -Properties GrantSendOnBehalfTo).GrantSendOnBehalfTo
  
      if([string]$SendOnBehalfPermissions -ne ""){
          $Print = 1
          $UserWithAccessEmail = ""
          $AccessRights = "SendOnBehalfTo"
          foreach($SendOnBehalfPermission in $SendOnBehalfPermissions){
              $UserUPN = (Get-EXOMailbox -Identity $SendOnBehalfPermission.ToString() -ErrorAction SilentlyContinue).UserPrincipalName
              $UserWithAccessEmail = $UserWithAccessEmail + $UserUPN

              if($SendOnBehalfPermissions.IndexOf($SendOnBehalfPermission) -lt (($SendOnBehalfPermissions.Count)-1)){
                  $UserWithAccessEmail = $UserWithAccessEmail + ";"
              }
          }
          Get-MBPermissionHash
      }
    }
  }

function Get-MBPermissionHash{ 

    $Result  = "" 
    $Results = @()
  
    if($Print -eq 1){
      $Result = [ordered]@{
        UserPrinciPalName   = $upn
        DisplayName         = $_.Displayname
        PCOMCenter          = ($_.CustomAttribute7 -split ';')[0]
        ManagingCenterName  = $_.CustomAttribute5
        MailboxType         = $MBType
        #AccessRights            = "Mailbox:$AccessRights"
        AccessRights            = $AccessRights
        #UserWithAccessName     = ($UserWithAccessEmail -split '@')[0]
        UserWithAccessEmail     = $UserWithAccessEmail
      } 
      $Results = New-Object PSObject -Property $Result |
                Sort-Object UserPrinciPalName -Descending |
                Export-Csv -Path $ExportCSV -Notype -Append 
    }
  }