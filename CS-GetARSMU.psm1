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
   Get ARSMU name by center code
.DESCRIPTION
   use center code to get center name
.EXAMPLE
   Get-CS-ARSMU -Center_Code 91
.EXAMPLE
   Get-CS-ARSMU 91
#>
function Get-CS-ARSMU{
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Center_Code
    )

$Center_Info = $CSCenters.Values | Where-Object { $_.Center_Code -eq $Center_Code }
$Center_Name = $Center_Info.Name

if ($null -ne $Center_Name){
    Write-Host "The name for center $Center_Code is $Center_Name"
}
else{
    Write-Host "$Center_Code is not a valid Center code"

}
}