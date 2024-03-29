﻿<#
Get-CS-MailboxAIO (alias: Get-MailboxAIO) with paramter: 
 * Parameter -Center is mandantory and positional.
 * 
 * report by center ARS_MU number: -center ISC
 * report by all centers: -center all
 * report without any center information: -center none.
 * report by date: -LastNumberOfDays, -StartDate/-EndDate
 * this cmdlet needs $CSCcenters, Get-CS-MailboxHash
 * report will be on current user's Desktop
 * change to -filter instead of where-object to get better performance
#>

$data = Import-Csv -Path "C:\CS\PennO365Data\Center.csv"
$CSCenters = @{}
foreach ($row in $data) {
    $hash = [ordered]@{
        'Name' = $row.Name
        'Center_Code' = $row.Center_Code # was code
        'ARS_MU' = $row.ARS_MU  # was short
        'Email' = $row.Email
    }
    $CSCenters.Add($row.ARS_MU, $hash)
}
<#
    .Synopsis
       This cmdlet is all in one version for mailbox report. 
    .DESCRIPTION
       By default, this cmdlet returns a report of all mailboxes in the organization with selected parameters. 
       -center is Mandatory and positional Parameter
       -LastNumberofDays
       -StartDate -Enddate
    .EXAMPLE
        Get-CS-MailboxAIO -center All
    .EXAMPLE
       Get-CS-MailboxAIO -center ISC 
    .EXAMPLE
       Get-CS-MailboxAIO -center NONE 
    .EXAMPLE
       Get-CS-MailboxAIO All -lastNumberOfDays 7
    .EXAMPLE
       Get-CS-MailboxAIO All -StartDate 2/1/2023 -EndDate 3/1/2023  
    .EXAMPLE
       Get-CS-MailboxAIO -center ICA -StartDate 2/1/2023 -EndDate 3/1/2023  
    .EXAMPLE
       Get-CS-MailboxAIO -center ICA -LastNumberOfDays 365
    .INPUTS
       parameters
    .OUTPUTS
       CSV file in C:\Temp folder
#>
function Get-CS-MailboxAIO{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Center,
        [int]$LastNumberOfDays,
        [datetime]$StartDate,
        [datetime]$EndDate
    )
        $CenterCode = $CSCenters[$center].Center_Code
        $CenterName = $CSCenters[$center].Name
           
        $DateStamp = $(Get-Date -f yyyyMMdd_HHmmss).tostring()
        $MyLocation = [Environment]::GetFolderPath("Desktop")
        $ReportPath = "$($MyLocation)\MailboxReport_"
       
        if($Center -ne 'all' -and $Center -ne "None" -and $Center -ne '' -and $Center -notin $Script:CSCenters.Keys){
            Write-Error "Please put correct Center Name" -ErrorAction Stop

        }elseif($StartDate -gt $null -and -not $EndDate){
            Write-Error "-EndDate is missing."  -ErrorAction Stop
          
        }elseif($EndDate -gt $null -and -not $StartDate){
            Write-Error "StartDate is missing." -ErrorAction Stop

        }elseif($Startdate -ge $EndDate -and $EndDate -ne $null){
            Write-Error "The Startdate has to be early than the EndDate." -ErrorAction Stop
        }

        #Write-Output "Connect to real world!" 
        #break
        Connect-ExchangeOnline -ShowBanner:$false 
        Connect-MgGraph -Scopes "Directory.Read.All"  -ErrorAction SilentlyContinue -Errorvariable ConnectionError |Out-Null
        Select-MgProfile -Name "beta" 
        Write-Progress "Getting mailboxes. Please wait..." 

    if($center -eq "all" -and $LastNumberOfDays -ge 1){

        $Mailboxes = Get-EXOMailbox -ResultSize unlimited -Filter "WhenMailboxCreated -ge '$((Get-Date).ToUniversalTime().AddDays(-$LastNumberOfDays))'" `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress
                    
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes                    
        Get-Output
    }elseif($center -eq "all" -and $Startdate -lt $EndDate){  
        
        $Mailboxes = Get-EXOMailbox -ResultSize unlimited -Filter "(WhenMailboxCreated -ge '$($Startdate.ToUniversalTime())') -and (WhenMailboxCreated -lt '$($Enddate.ToUniversalTime())')" `
           -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress 
                   
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        Get-Output
    }elseif(($Center -in $Script:CSCenters.Keys) -and ($LastNumberOfDays -ge 1)){
             
        $Mailboxes = Get-EXOMailbox -ResultSize unlimited -Filter "WhenMailboxCreated -ge '$((Get-Date).ToUniversalTime().AddDays(-$LastNumberOfDays))'" `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            where-object {($_.CustomAttribute7 -split ';')[0] -eq $CenterCode -or $_.CustomAttribute5 -eq $CenterName}
          #  -and $_.WhenMailboxCreated -ge (Get-Date).AddDays(-$LastNumberOfDays)}                   

          $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        Get-Output  
    }
    elseif (($Center -in $Script:CSCenters.Keys) -and ($Startdate -lt $EndDate)){

        $Mailboxes = Get-EXOMailbox -ResultSize unlimited -Filter "(WhenMailboxCreated -ge '$($Startdate.ToUniversalTime())') -and (WhenMailboxCreated -lt '$($Enddate.ToUniversalTime())')" `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            Where-Object {($_.CustomAttribute7 -split ';')[0] -eq $CenterCode -or $_.CustomAttribute5 -eq $CenterName}
          #  -and ($_.WhenMailboxCreated -ge $Startdate -and $_.whenmailboxcreated -le $Enddate)}
        
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        Get-Output
    }
    elseif ($Center -eq "all"){

        $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ForwardingSmtpAddress
                
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        Get-Output
    }
    elseif ($Center -in $Script:CSCenters.Keys){    
    
        $Mailboxes = Get-EXOMailbox -ResultSize unlimited -Filter "(CustomAttribute7 -eq '$($CenterCode)*') -or (CustomAttribute5 -eq '$CenterName')" `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress 
          #  Where-Object {(($_.CustomAttribute7 -split ';')[0] -eq $CenterCode) -or ($_.CustomAttribute5 -eq $CenterName)} 
   
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        Get-Output
    }
    elseif ($Center -eq "none"){        
           
        $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            Where-Object {($_.Customattribute7 -eq "") -and ($_.CustomAttribute5 -eq "")}
                
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        Get-Output
    }    
    Disconnect-ExchangeOnline -Confirm:$false *> $null
    #Disconnect-MgGraph
}
<#
    internal function, hash table to create report.
#>
function Get-CS-AllMailboxHash{
    Param(
        [Parameter(Mandatory=$false)]
        $Mailboxes
    )
        $total = $Mailboxes.count
        if ($total -eq 0){
            Write-Error "There is no mailbox!" -ErrorAction Stop
        }
        $i = 0
        $MailboxData = @()
        $Mailboxes | ForEach-Object {
            $i++
 #           $percent = ($i / $total * 100)
 #           Write-Progress -Activity "Processing mailboxes" -PercentComplete $percent -Status "$i of $total processed"

           $MailboxHash = [ordered]@{
            PennID                = (Get-MgUser -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue).EmployeeId
            #PennID                = ''
            Name                  = ($_.UserPrincipalName -split '@')[0]
            DisplayName           = $_.DisplayName
            UserPrincipalName     = $_.UserPrincipalName
            ManagingCenter        = ($_.CustomAttribute7 -split ';')[0]
            ManagingCenterName    = $_.CustomAttribute5
            PCOMCenter            = ($_.CustomAttribute7 -split ';')[0]
            BudgetCode            = $_.CustomAttribute14
            PrimarySMTP           = $_.PrimarySmtpAddress
            AccountType           = $_.RecipientTypeDetails
            Alias                 = $_.Alias
            EmailAddresses        = $_.EmailAddresses -join ';'
            ForwardingSmtpAddress = $_.ForwardingSmtpAddress
            MailboxEnabled        = $_.isMailboxEnabled
            WhenMailboxCreated    = $_.WhenMailboxCreated
            ArchiveStatus         = $_.ArchiveStatus
            License               = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue).SkuPartNumber -join ";"
            #License               = ''
       }
      <# try {
        #$MgUser = Get-MgUser -UserId $_.UserPrincipalName 
        $MgUser = Get-MgUser -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue
        $MailboxHash.PennID = $MgUser.EmployeeId
        $MailboxHash.License = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue).SkuPartNumber -join ";"
        #$MailboxHash.License = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName).SkuPartNumber -join ";"
        $ErrorUser = $MgUser.UserPrincipalName
    }
    catch {        
        Write-Warning "Failed to retrieve user's PennID and License information for $ErrorUser" -ErrorAction:SilentlyContinue
    } #>
            $MailboxRow = New-Object psobject -Property $MailboxHash
            $MailboxData += $MailboxRow
        }
       
       return $MailboxData | Sort-Object WhenMailboxCreated -Descending
}
function Get-Output {
    $ReportPath = $ReportPath + $($Center)+"_"
    $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
    Write-Output "Please find $($Center)_Mailboxes report in $($ReportPath+$DateStamp).csv"   
}
