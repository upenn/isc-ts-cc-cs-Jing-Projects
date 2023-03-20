## [1.0.0] - 2023-02-16
- Initial version. Implement Get-AllMailboxReport/Get-NewMailboxReport

## [1.1.0] - 2023-02-28
- Implement Get-CS-MailboxAIO with function Get-CS-AllMailboxHash
- with Parameter -Center (Mandatory and positional Parameter), -StartDate/-EndDate and -LastNumberofDays


## [1.2.0] - 2023-03-08
- Add comment, help 
- Test on Module
- Create Center.csv, added to PennO365Data, use import cvs
- Test Private/Public key to transfer file to mailservice server.

## [1.3.0] - 2023-03-17
- remove Alias without CS
- use write-error -ErrorAction Stop to replace break
- MgGraph works better on PowerShell instead of Windows PowerShell

## [1.3.1] - 2023-03-20
- add CHANGELOG.md

