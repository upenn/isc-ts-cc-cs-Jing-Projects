<#
Get-CS-MailboxAIO (alias: Get-MailboxAIO) with paramter: 
 * report by date: -LastNumberOfDays, -StartDate/-EndDate
 * report by center short number: -center ISC
 * report by all centers: -center all
 * report without any center inforamtion: -center none.
 * this cmdlet needs $CSCcenters, Get-CS-MailboxHash
 * 
 following cmdlets can work individully, see details in each cmdlet.
 Get-CS-NewMailbox
 Get-CS-AllMailbox
 Get-CS-allmailBoxBycenter
#>

# $script:CSCenters
$CSCenters = [ordered]@{
	AC 	= @{ 'Name' = 'Annenberg Center for Performing Arts'; 'Code' = '19'; 'Short' = 'AC'; 'Email' = "o365-ac-lsps@isc.upenn.edu"}
	ASC 	= @{ 'Name' = 'Annenberg School for Communications'; 'Code' = '36'; 'Short' = 'ASC'; 'Email' = "o365-asc-lsps@isc.upenn.edu"}
	BSD 	= @{ 'Name' = 'Business Services'; 'Code' = '93'; 'Short' = 'BSD'; 'Email' = "o365-bsd-lsps@isc.upenn.edu"}
	BUS 	= @{ 'Name' = 'Morris Arboretum'; 'Code' = '60'; 'Short' = 'BUS'; 'Email' = "o365-bus-lsps@isc.upenn.edu"}
	CHAS 	= @{ 'Name' = 'College Houses and Academic Services'; 'Code' = '86'; 'Short' = 'CHAS'; 'Email' = "o365-chas-lsps@isc.upenn.edu"}
	DAR 	= @{ 'Name' = 'Development and Alumni Relations'; 'Code' = '90'; 'Short' = 'DAR'; 'Email' = "o365-dar-lsps@isc.upenn.edu"}
	DES 	= @{ 'Name' = 'Weitzman School of Design'; 'Code' = '33'; 'Short' = 'DES'; 'Email' = "o365-des-lsps@isc.upenn.edu"}
	DOF 	= @{ 'Name' = 'Division of Finance'; 'Code' = '87'; 'Short' = 'DOF'; 'Email' = "o365-dof-lsps@isc.upenn.edu"}
	DRIA 	= @{ 'Name' = 'Division of Recreation & Intercollegiate Athletics'; 'Code' = '24'; 'Short' = 'DRIA'; 'Email' = "o365-dria-lsps@isc.upenn.edu"}
	EVP 	= @{ 'Name' = 'Executive Vice President'; 'Code' = '98'; 'Short' = 'EVP'; 'Email' = "o365-evp-lsps@isc.upenn.edu"}
	FRES 	= @{ 'Name' = 'Facilities and Real Estate Services'; 'Code' = '96'; 'Short' = 'FRES'; 'Email' = "o365-fres-lsps@isc.upenn.edu"}
	GSE 	= @{ 'Name' = 'Graduate School of Education'; 'Code' = '32'; 'Short' = 'GSE'; 'Email' = "o365-gse-lsps@isc.upenn.edu"}
	HR 	= @{ 'Name' = 'Human Resources'; 'Code' = '92'; 'Short' = 'HR'; 'Email' = "o365-hr-lsps@isc.upenn.edu"}
	ISC 	= @{ 'Name' = 'Information Systems and Computing'; 'Code' = '91'; 'Short' = 'ISC'; 'Email' = "o365-isc-lsps@isc.upenn.edu"}
	ICA 	= @{ 'Name' = 'Institute of Contemporary Art'; 'Code' = '61'; 'Short' = 'ICA'; 'Email' = "o365-ica-lsps@isc.upenn.edu"}
	LAW 	= @{ 'Name' = 'Law School'; 'Code' = '56'; 'Short' = 'LAW'; 'Email' = "o365-law-lsps@isc.upenn.edu"}
	LDC 	= @{ 'Name' = 'Linguistic Data Consortium'; 'Code' = '02'; 'Short' = 'LDC'; 'Email' = "o365-ldc-lsps@isc.upenn.edu"}
	LIB 	= @{ 'Name' = 'University Library'; 'Code' = '50'; 'Short' = 'LIB'; 'Email' = "o365-lib-lsps@isc.upenn.edu"}
	MUS 	= @{ 'Name' = 'University Museum'; 'Code' = '26'; 'Short' = 'MUS'; 'Email' = "o365-mus-lsps@isc.upenn.edu"}
	NUR 	= @{ 'Name' = 'School of Nursing'; 'Code' = '06'; 'Short' = 'NUR'; 'Email' = "o365-nur-lsps@isc.upenn.edu"}
	OACP 	= @{ 'Name' = 'Audit Compliance and Privacy'; 'Code' = '78'; 'Short' = 'OACP'; 'Email' = "o365-oacp-lsps@isc.upenn.edu"}
	PDM 	= @{ 'Name' = 'School of Dental Medicine'; 'Code' = '51'; 'Short' = 'PDM'; 'Email' = "o365-pdm-lsps@isc.upenn.edu"}
	PGL 	= @{ 'Name' = 'Penn Global'; 'Code' = '62'; 'Short' = 'PGL'; 'Email' = "o365-pgl-lsps@isc.upenn.edu"}
	PSOM 	= @{ 'Name' = 'Perelman School of Medicine'; 'Code' = '40'; 'Short' = 'PSOM'; 'Email' = "o365-psom-lsps@isc.upenn.edu"}
	PRES 	= @{ 'Name' = 'President''s Center'; 'Code' = '81'; 'Short' = 'PRES'; 'Email' = "o365-pres-lsps@isc.upenn.edu"}
	PROV 	= @{ 'Name' = 'Provost''s Center'; 'Code' = '83'; 'Short' = 'PROV'; 'Email' = "o365-prov-lsps@isc.upenn.edu"}
	RHS 	= @{ 'Name' = 'Residential and Hospitality Services'; 'Code' = '95'; 'Short' = 'RHS'; 'Email' = "o365-rhs-lsps@isc.upenn.edu"}
	SAS 	= @{ 'Name' = 'School of Arts and Sciences'; 'Code' = '2'; 'Short' = 'SAS'; 'Email' = "o365-sas-lsps@isc.upenn.edu"}
	SEAS 	= @{ 'Name' = 'School of Engineering and Applied Science'; 'Code' = '13'; 'Short' = 'SEAS'; 'Email' = "o365-seas-lsps@isc.upenn.edu"}
	SP2 	= @{ 'Name' = 'School of Social Policy and Practice'; 'Code' = '35'; 'Short' = 'SP2'; 'Email' = "o365-sp2-lsps@isc.upenn.edu"}
	VET 	= @{ 'Name' = 'School of Veterinary Medicine'; 'Code' = '58'; 'Short' = 'VET'; 'Email' = "o365-vet-lsps@isc.upenn.edu"}
	VPUL 	= @{ 'Name' = 'Student Services'; 'Code' = '85'; 'Short' = 'VPUL'; 'Email' = "o365-vpul-lsps@isc.upenn.edu"}
	WEL 	= @{ 'Name' = 'Health and Wellness'; 'Code' = '89'; 'Short' = 'WEL'; 'Email' = "o365-wel-lsps@isc.upenn.edu"}
	WHA 	= @{ 'Name' = 'Wharton School'; 'Code' = '7'; 'Short' = 'WHA'; 'Email' = "o365-wha-lsps@isc.upenn.edu"}
	XPN 	= @{ 'Name' = 'XPN'; 'Code' = '81'; 'Short' = 'XPN'; 'Email' = "o365-xpn-lsps@isc.upenn.edu"}
}

function Get-CSCenter{
    # testing CSCenter only
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Center = "ALL"
    )

    $Code = $CSCenters[$center].Code
    $CenterName = $CSCenters[$center].Name
    $CenterEmail = $CSCenters[$center].Email

    if ($Center -eq "ALL"){
        foreach ($CenterCode in $Script:CSCenters.Keys){
            Write-Output "Here is Centercode-$CenterCode, code-$($CSCenters[$CenterCode].Code), CenterName-$($CSCenters[$CenterCode].Name)"  
        }
    }elseif($Center -in $Script:CSCenters.keys){
            Write-Output "Here is CenterShortName-$Center, code-$Code, CenterName-$CenterName, CenterEmail-$CenterEmail"
        }
    }

<#
    .Synopsis
       This cmdlet is all in one version for mailbox report. 
    .DESCRIPTION
       By default, this cmdlet returns a report of all mailboxes in the organization, or depends on parameter,
       the option of center's short name, and/or last # of days of new mailbox creation, and/or StartDate, EndDate of new mailbox creation.
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
        [string]$Center="all",
        [int]$LastNumberOfDays,
        [datetime]$StartDate,
        [datetime]$EndDate
    )
        Connect-ExchangeOnline -ShowBanner:$false
        Connect-MgGraph -ErrorAction:SilentlyContinue
        Select-MgProfile -Name "beta"
   
        $CenterCode = $CSCenters[$center].Code
        $CenterName = $CSCenters[$center].Name
   
        $DateStamp = (Get-Date -f yyyyMMdd_HHmmss).tostring()
        $ReportPath = "C:\Temp\MailboxReport"

        if($Startdate -gt $EndDate){
            Write-Host "The Startdate is after the EndDate." 
            break
        }
        elseif(($Center -in $Script:CSCenters.Keys) -and ($LastNumberOfDays -ge 1)){
#####################################testing only          
            Write-Output "Here is Center-$Center, code-$CenterCode, CenterName-$CenterName, LastNumberofDays-$LastNumberofDays" # testing only
              
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
 #####################################testing only
            Write-Output "Here is Center-$Center, code-$CenterCode, CenterName-$CenterName,StartDate-$Startdate, EndDate-$Enddate" # testing only

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
 #####################################testing only
            Write-Output "Here is CenterShortName1-$Center, code-$CenterCode, CenterName-$CenterName" # testing only

            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ForwardingSmtpAddress
                
            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

            $ReportPath = $ReportPath +"_"+ $($Center)+"_"
            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
            Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"
            }
        elseif ($Center -in $Script:CSCenters.Keys){
 #####################################testing only            
            Write-Output "Here is CenterShortName-$Center, code-$CenterCode, CenterName-$CenterName" # testing only
              
            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
                Where-Object {(($_.CustomAttribute7 -split ';')[0] -eq $CenterCode) -or ($_.CustomAttribute5 -eq $CenterName)} 
    
            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

            $ReportPath = $ReportPath +"_"+ $($Center)+"_"
            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
            Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"
        }
        elseif ($Center -eq "none"){
            $ReportPath = $ReportPath +"_"+ $($Center)+"_"
            Write-Output "Here is CenterShortName-$Center" # testing only
                
            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
                Where-Object {($_.Customattribute7 -eq "") -and ($_.CustomAttribute5 -eq "")}
                
            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
                                      
            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
            Write-Output "Please find $center _Mailboxes report in $($ReportPath+$DateStamp).csv"
        }
        elseif($LastNumberOfDays -ge 1){
            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
                Where-Object {$_.WhenMailboxCreated -ge (Get-Date).AddDays(-$LastNumberOfDays)}
            
            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
            
            $ReportPath = $ReportPath + "_Last"+$LastNumberOfDays + "Days_"
            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
            
        }
        elseif($Startdate -lt $EndDate){  
            ## $startdate/$enddate
            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
                Where-Object {$_.WhenMailboxCreated -ge $Startdate -and $_.whenmailboxcreated -le $Enddate} 
            
            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
            $ReportPath = $ReportPath + "_From$(($StartDate.ToString("yyyyMMdd")))to$(($EndDate.ToString("yyyyMMdd")))_"
            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
        }else{ #no parameter, return all mailboxes
            #####################################testing only
                           Write-Output "Here is CenterShortName2-$Center, code-$CenterCode, CenterName-$CenterName" # testing only
               
                           $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                               -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ForwardingSmtpAddress
                               
                           $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

                           $ReportPath = "All_"+$ReportPath
                           $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
                           Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"
        }          
}

<#
.Synopsis
   Get new mailboxes report created in last # of days or in certain peroid with -StartDate and -EndDate
.DESCRIPTION
   Get new mailboxes report created in last # of days or in certain peroid 
   parameter will be -LastNumberOfDays, -StartDate and -EndDate
   -StartDate and -EndDate it's from midnight to midnight
.EXAMPLE
   Get-CS-NewMailbox 7 (get last 7 days new mailboxes report)   
.EXAMPLE
   Get-CS-NewMailbox -StartDate 1/1/2023 -EndDate 3/1/2023 (get new mailboxes created between two dates)
   #>
function Get-CS-NewMailbox
{
    [CmdletBinding()]
    [Alias('Get-NewMailbox')]
    Param
    ( 
        [int] $LastNumberOfDays,
        [datetime]$StartDate,
        [datetime]$EndDate
    )
        Connect-ExchangeOnline -ShowBanner:$false
        Connect-MgGraph -ErrorAction:SilentlyContinue
        Select-MgProfile -Name "beta"

        $DateStamp = (Get-Date -f MMdd_HHmmss).tostring()
        
        $ReportPath = "C:\Temp\NewMailboxReport"

        if($LastNumberOfDays -ge 1){
            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
                Where-Object {$_.WhenMailboxCreated -ge (Get-Date).AddDays(-$LastNumberOfDays)}

            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

            $ReportPath = $ReportPath + "_Last"+$LastNumberOfDays + "Days_"
            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
        }
        elseif(-not $StartDate){
            Write-Host "-LastNumberOfDays or - StartDate -EndDate missing." 
            break
        }
        elseif(-not $EndDate){
            Write-Host "-EndDate missing." 
            break
        }
        elseif($Startdate -gt $EndDate){
            Write-Host "The Startdate is after the EndDate." 
            break
        }
        else{
            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
            -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
            Where-Object {$_.WhenMailboxCreated -ge $Startdate -and $_.whenmailboxcreated -le $Enddate} 

            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
            $ReportPath = $ReportPath + "_From$(($StartDate.ToString("yyyyMMdd")))to$(($EndDate.ToString("yyyyMMdd")))_"
            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
        }          
        
    $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
    $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
    Write-Output "Please find New Mailboxes report from $Startdate to $EndDate in $($ReportPath+$DateStamp).csv"
}

<#
.Synopsis
   Get all mailboxes report in the organization. 
.DESCRIPTION
    This cmdlet returns a report of all mailboxes report in the organization. 
.EXAMPLE
   Get-CS-AllMailbox
.EXAMPLE
   Get-AllMailbox
   #>
function Get-CS-AllMailbox{
    [CmdletBinding()]
    [Alias('Get-AllMailbox')]
    Param()
     
        Connect-ExchangeOnline -ShowBanner:$false
        Connect-MgGraph -ErrorAction:SilentlyContinue
        Select-MgProfile -Name "beta"

        $DateStamp = (Get-Date -f MMdd_HHmmss).tostring()
        $ReportPath = "C:\Temp\AllMailboxReport"

        $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
        -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, `
                    isMailboxEnabled, ArchiveStatus,ImmutableId,ForwardingSmtpAddress
        $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
  
        $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
        Write-Output "Please find All Mailboxes report in $($ReportPath+$DateStamp).csv"
}

<#
.Synopsis
   get mailbox report by center short name or All or None as parameter
   "short name" returns all mailbox with ID/center name even one of them is missing.
    "All" means all mailboxes.
    "none" means no (PCOM ID and center name)
.DESCRIPTION
    this cmdlet return a report of all mailboxes in the center/all organization or without center
.EXAMPLE
    Get-CS-AllMailboxByCenter ALL
.EXAMPLE
    Get-CS-AllMailboxByCenter ISC
.EXAMPLE
    Get-CS-AllMailboxByCenter none
   #>
function Get-CS-AllMailboxByCenter{
    [CmdletBinding()]
    [Alias('Get-AllMailboxByCenter')]
    Param(
        [string]$Center = "ALL"
    )
        Connect-ExchangeOnline -ShowBanner:$false
        Connect-MgGraph -ErrorAction:SilentlyContinue
        Select-MgProfile -Name "beta"

        $CenterCode = $CSCenters[$center].Code
        $CenterName = $CSCenters[$center].Name

        $DateStamp = (Get-Date -f MMdd_HHmmss).tostring()
        $ReportPath = "C:\Temp\$($Center)_MailboxReport"

        if ($Center -eq "all"){
            
            Write-Output "Here is CenterShortName-$Center, code-$CenterCode, CenterName-$CenterName" # testing only

            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated, isMailboxEnabled, ArchiveStatus,ForwardingSmtpAddress
            
            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
            Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"
            }
        elseif ($Center -in $Script:CSCenters.Keys){
            
            Write-Output "Here is CenterShortName-$Center, code-$CenterCode, CenterName-$CenterName" # testing only
            
            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
                Where-Object {(($_.CustomAttribute7 -split ';')[0] -eq $CenterCode) -or ($_.CustomAttribute5 -eq $CenterName)} 

            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes

            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
            Write-Output "Please find $Center _Mailboxes report in $($ReportPath+$DateStamp).csv"
        }
        elseif ($Center -eq "none"){

            Write-Output "Here is CenterShortName-$Center" # testing only
            
            $Mailboxes = Get-EXOMailbox -ResultSize unlimited `
                -Properties CustomAttribute5,Customattribute7,CustomAttribute14,WhenMailboxCreated,isMailboxEnabled,ArchiveStatus,ForwardingSmtpAddress |
                Where-Object {($_.Customattribute7 -eq "") -and ($_.CustomAttribute5 -eq "")}
            
            $MailboxData = Get-CS-AllMailboxHash -Mailboxes $Mailboxes
                                  
            $MailboxData | Export-Csv "$($ReportPath+$DateStamp).csv" -NoTypeInformation
            Write-Output "Please find $center _Mailboxes report in $($ReportPath+$DateStamp).csv"
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
