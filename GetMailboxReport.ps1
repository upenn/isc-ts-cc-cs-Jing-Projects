<#
Get-CS-MailboxAIO (alias: Get-MailboxAIO) with paramter: 
 * report by date: -LastNumberOfDays, -StartDate/-EndDate
 * report by center ARS_MU number: -center ISC
 * report by all centers: -center all
 * report without any center inforamtion: -center none.
 * this cmdlet needs $CSCcenters, Get-CS-MailboxHash
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
       By default, this cmdlet returns a report of all mailboxes in the organization, or depends on parameter,
       the option of center's ARS_MU name, and/or last # of days of new mailbox creation, and/or StartDate, EndDate of new mailbox creation.
    .EXAMPLE
       Get-CS-MailboxAIO -lastNumberOfDays 7 //tested
    .EXAMPLE
       Get-CS-MailboxAIO -StartDate 2/1/2023 -EndDate 3/1/2023  //tested
    .EXAMPLE
        Get-CS-MailboxAIO -center All
    .EXAMPLE
       Get-CS-MailboxAIO -center ISC //tested 
    .EXAMPLE
       Get-CS-MailboxAIO -center NONE //tested 
    .EXAMPLE
       Get-CS-MailboxAIO -center ICA -LastNumberOfDays 365
    .INPUTS
       parameters
    .OUTPUTS
       CSV file in C:\Temp folder
#>
function Get-CS-MailboxAIO{
    [CmdletBinding()]
    [Alias('Get-MailboxAIO')]
    Param(
        [string]$Center,
        [int]$LastNumberOfDays,
        [datetime]$StartDate,
        [datetime]$EndDate
    )
        $CenterCode = $CSCenters[$center].Center_Code
        $CenterName = $CSCenters[$center].Name
   
        $DateStamp = (Get-Date -f yyyyMMdd_HHmmss).tostring()
        $ReportPath = "C:\Temp\MailboxReport"

        #Write-Host "Center-$Center,LastNumberOfDays-$LastNumberOfDays,StartDate-$StartDate, EndDate-$EndDate " 
        
        if($Startdate -gt $EndDate -and $EndDate -ne $null){
            Write-Host "The Startdate is after the EndDate." 
            break

        }elseif($LastNumberOfDays -lt 0 ){
                Write-Host "-LastNumberOfDays can't be negative." 
                break
        }elseif($StartDate -eq $null -and $EndDate -eq $null -and $LastNumberOfDays -lt 1 -and $Center -eq ''){
            Write-Host "Please add Parameters"
            break
        }elseif($StartDate -gt $null -and -not $EndDate){
            Write-Host "-EndDate is missing." 
            break
        }elseif($EndDate -gt $null -and -not $StartDate){
            Write-Host "StartDate is missing." 
            break
        }elseif($Center -ne 'all' -and $Center -ne "None" -and $Center -ne '' -and $Center -notin $Script:CSCenters.Keys){
            Write-host "Please put correct Center short name/ARS_MU name!"
            break
        }

        Connect-ExchangeOnline -ShowBanner:$false
        Connect-MgGraph | out-null
        Select-MgProfile -Name "beta"

    if(($Center -in $Script:CSCenters.Keys) -and ($LastNumberOfDays -ge 1)){
        Write-Output "Center-$Center, Center_Code-$CenterCode, CenterName-$CenterName, LastNumberofDays-$LastNumberofDays" # testing only
              
        $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            where-object {($_.CustomAttribute7 -split ';')[0] -eq $CenterCode -or $_.CustomAttribute5 -eq $CenterName `
            -and $_.WhenMailboxCreated -ge (Get-Date).AddDays(-$LastNumberOfDays)}                   
   
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

        $ReportPath = $ReportPath +"_"+ $($Center)+"_Last_"+$LastNumberOfDays+"_Days"
        $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
         Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"         
    }
    elseif (($Center -in $Script:CSCenters.Keys) -and ($Startdate -lt $EndDate)){

        Write-Output "Center-$Center, Center_Code-$CenterCode, CenterName-$CenterName,StartDate-$Startdate, EndDate-$Enddate" # testing only

        $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            Where-Object {($_.CustomAttribute7 -split ';')[0] -eq $CenterCode -or $_.CustomAttribute5 -eq $CenterName `
            -and ($_.WhenMailboxCreated -ge $Startdate -and $_.whenmailboxcreated -le $Enddate)}
        
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

        $ReportPath = $ReportPath +"_"+ $($Center)+"_From$(($StartDate.ToString("yyyyMMdd")))to$(($EndDate.ToString("yyyyMMdd")))_"
        $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
        Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"
    }
    elseif ($Center -eq "all"){

    Write-Output "CenterShortName1-$Center" # testing only

    $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
        -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ForwardingSmtpAddress
                
    $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

    $ReportPath = $ReportPath +"_"+ $($Center)+"_"
    $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
    Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"
    }
    elseif ($Center -in $Script:CSCenters.Keys){
        
        Write-Output "CenterShortName-$Center, Center_Code-$CenterCode, CenterName-$CenterName" # testing only
              
        $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            Where-Object {(($_.CustomAttribute7 -split ';')[0] -eq $CenterCode) -or ($_.CustomAttribute5 -eq $CenterName)} 
    
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

        $ReportPath = $ReportPath +"_"+ $($Center)+"_"
        $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
        Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"
    }
    elseif ($Center -eq "none"){
            
        Write-Output "CenterShortName-$Center" # testing only
                
        $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            Where-Object {($_.Customattribute7 -eq "") -and ($_.CustomAttribute5 -eq "")}
                
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

        $ReportPath = $ReportPath +"_"+ $($Center)+"_"                                      
        $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
         Write-Output "Please find $center _Mailboxes report in $($ReportPath+$DateStamp).csv"
    }
    elseif($LastNumberOfDays -ge 1){
    Write-Output "LastNumberOfDays-$LastNumberOfDays" # testing only

    $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
        -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
        Where-Object {$_.WhenMailboxCreated -ge (Get-Date).AddDays(-$LastNumberOfDays)}
            
    $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
            
    $ReportPath = $ReportPath + "_Last"+$LastNumberOfDays + "Days_"
    $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
            
    }
    elseif($Startdate -lt $EndDate){  
        Write-Output "Startdate-$Startdate, EndDate-$EndDate" # testing only

        $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            Where-Object {$_.WhenMailboxCreated -ge $Startdate -and $_.whenmailboxcreated -le $Enddate} 
            
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
        $ReportPath = $ReportPath + "_From$(($StartDate.ToString("yyyyMMdd")))to$(($EndDate.ToString("yyyyMMdd")))_"
        $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
    } 
    
}
<#
    internal function, hash table to create report.
#>
function Get-CS-AllMailboxHash{
    [CmdletBinding()]
    [Alias('Get-AllMailboxReport')]
    Param(
        [Parameter(Mandatory=$false)]
        #[string]$Center,
        $Mailboxes
    )
        $total = $Mailboxes.count
        
       # $Total = ($Mailboxes | Measure-Object).Count
        $i = 0
        $MailboxData = @()
        $Mailboxes | ForEach-Object {
            $i++
            $percent = ($i / $total * 100)
            Write-Progress -Activity "Processing mailboxes" -PercentComplete $percent -Status "$i of $total processed"

            $MailboxHash = [ordered]@{
                PennID                = (Get-MgUser -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue).EmployeeId
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
                EmailAddresses        = $_.EmailAddresses -join ';'
                ForwardingSmtpAddress = $_.ForwardingSmtpAddress
                MailboxEnabled        = $_.isMailboxEnabled
                WhenMailboxCreated    = $_.WhenMailboxCreated
                ArchiveStatus         = $_.ArchiveStatus
                License               = (Get-MgUserLicenseDetail -UserId $_.UserPrincipalName -ErrorAction:SilentlyContinue).SkuPartNumber -join ";"
            }
            $MailboxRow = New-Object psobject -Property $MailboxHash

            $MailboxData += $MailboxRow
            }
       return $MailboxData | Sort-Object WhenMailboxCreated -Descending
}
