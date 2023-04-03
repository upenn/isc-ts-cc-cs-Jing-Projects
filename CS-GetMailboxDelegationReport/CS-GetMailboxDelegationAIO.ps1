

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

function Get-CS-MailboxDelegation{

  param(
    [Parameter(Mandatory=$false)]
    [string]$Center = 'all',
    [switch]$FullAccess,
    [switch]$SendAs,
    [switch]$SendOnBehalf,
    [switch]$CalendarPermissions
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

  Write-Host "Connecting Exchange Online PowerShell..."
  Connect-ExchangeOnline -ShowBanner:$false 

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

Write-Host "Script executed successfully"
  if((Test-Path -Path $ExportCSV) -eq "True"){
    write-host "Please find report: $ExportCSV"
  }
  Else{
    Write-Host No mailbox found!
  }
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
        $UserWithAccess = ""
        $AccessRights = "FullAccess"
        foreach($FullAccessPermission in $FullAccessPermissions){
          $UserWithAccess = $UserWithAccess+$FullAccessPermission
          if($FullAccessPermissions.indexof($FullAccessPermission) -lt (($FullAccessPermissions.count)-1)){
            $UserWithAccess = $UserWithAccess+","
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
        $UserWithAccess = ""
        $AccessRights = "SendAs"
        foreach($SendAsPermission in $SendAsPermissions){
  
          $UserWithAccess = $UserWithAccess+$SendAsPermission
          if($SendAsPermissions.indexof($SendAsPermission) -lt (($SendAsPermissions.count)-1)){
            $UserWithAccess = $UserWithAccess+","
          }        }
        Get-MBPermissionHash
     }    }
     #Getting delegated SendOnBehalf permission for mailbox
    if(($FilterPresent -eq 'False') -or ($SendOnBehalf.IsPresent)){
      $SendOnBehalfPermissions = (Get-EXOMailbox -Identity $upn -Properties GrantSendOnBehalfTo).GrantSendOnBehalfTo
  
      if([string]$SendOnBehalfPermissions -ne ""){
          $Print = 1
          $UserWithAccess = ""
          $AccessRights = "SendOnBehalfTo"
          foreach($SendOnBehalfPermission in $SendOnBehalfPermissions){
              $UserUPN = (Get-EXOMailbox -Identity $SendOnBehalfPermission.ToString()).UserPrincipalName
              $UserWithAccess = $UserWithAccess + $UserUPN
              if($SendOnBehalfPermissions.IndexOf($SendOnBehalfPermission) -lt (($SendOnBehalfPermissions.Count)-1)){
                  $UserWithAccess = $UserWithAccess + ","
              }          }
          Get-MBPermissionHash
      }  }
  #Getting Calendar permission
  if(($FilterPresent -eq 'False') -or ($CalendarPermissions.IsPresent)){
    $calendarFolderPath = $mailbox.PrimarySmtpAddress.ToString() + ":\Calendar"
    $CalendarPermissions = Get-MailboxFolderPermission -Identity $calendarFolderPath | 
                           Where-Object { $_.User.DisplayName -ne "Default" -and $_.User.DisplayName -ne "Anonymous" }
    if([string]$CalendarPermissions -ne ""){
      $Print = 1
      $UserWithAccess = ""
     # $AccessRights = "SendOnBehalfTo"
      foreach($CalendarPermission in $CalendarPermissions){
        $UserUPN = (Get-EXOMailbox -Identity $SendOnBehalfPermission.ToString()).UserPrincipalName
        $UserWithAccess = $UserWithAccess + $UserUPN
        if($SendOnBehalfPermissions.IndexOf($SendOnBehalfPermission) -lt (($SendOnBehalfPermissions.Count)-1)){
      $UserWithAccess = $UserWithAccess + ","
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
      $Result = @{
        DisplayName        = $_.Displayname
        UserPrinciPalName  = $upn
        MailboxType        = $MBType
        AccessRights       = $AccessRights
        UserWithAccess     = $userwithAccess
        ManagingCenterName = $_.CustomAttribute5
      } 
      $Results = New-Object PSObject -Property $Result 
      $Results | select-object DisplayName,UserPrinciPalName,MailboxType,AccessRights,UserWithAccess,ManagingCenterName |
                Export-Csv -Path $ExportCSV -Notype -Append 
    }
  }