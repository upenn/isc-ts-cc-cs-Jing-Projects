Get-CS-O365ServiceHealthStatus
function Get-CS-O365ServiceHealthStatus {

    $TenantID = '6c4d949d-b91c-4c45-9aae-66d76443110d'
    $ClientId = 'de762a73-e4ba-458c-b9e6-64541e608766'
    $TenantID = '6c4d949d-b91c-4c45-9aae-66d76443110d'
    $CertificateThumbprint = '10BF18C022368A7F0FF8F1070DFA876C56250DAF'

    $DateStamp = $(Get-Date -f yyyyMMdd_HHmmss).tostring()

    #install-module Microsoft.Graph.Authentication
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint -ContextScope CurrentUser -Environment Global | Out-Null

    #Import-Module Microsoft.Graph.Devices.ServiceAnnouncement -ErrorAction Stop
    #install-module Microsoft.Graph.Devices.ServiceAnnouncement

    $issues = Get-MgServiceAnnouncementIssue -Top 20 `
            -Property Title, Id, StartDateTime, LastModifiedDateTime, Classification, Origin,Status,IsResolved, ImpactDescription,Feature,service `
            -OrderBy 'StartDateTime desc'

    $issues | Select-Object @{Name="Title"; Expression={$_.Title}},
                            @{Name="Id"; Expression={$_.Id}},
                            @{Name="Estimated Start Time"; Expression={Get-Date $_.StartDateTime}},
                            @{Name="LastModifiedDateTime"; Expression={Get-Date $_.LastModifiedDateTime}},
                            @{Name="Issue Type"; Expression={$_.Classification}},
                            @{Name="Issue origin"; Expression={$_.Origin}},
                            @{Name="Status"; Expression={$_.Status}},
                            @{Name="IsResolved"; Expression={$_.IsResolved}},                  
                            @{Name="ImpactDescription"; Expression={$_.ImpactDescription}},
                            @{Name="service"; Expression={$_.service}},
                            @{Name="Feature"; Expression={$_.Feature}} |
                            Out-File -FilePath "C:\CS\PennO365Data\O365SeviceHealthStatus\PennO365_ServiceHealthIssue_$($DateStamp).txt"
}