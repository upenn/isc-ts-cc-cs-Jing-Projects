## Version 0.2
**3/2/2023 (Nate)**
- Added Changelog
- Removed installer and uninstaller scripts
- Moved module and module manifest

## Version 1.0
**3/28/2023**
- Fixed issue with the filter in Get-CS-MailboxStatisticsReportForCenter by encapsulating the filter statement in quotes, instead of in elipses
- Added a number of useful debug statements
- For reports for a single center, the default filename is used for that center to output the report
- Added -Confirm:$false flag when disconnecting from Exchange Online

**Things that need to be done**
- Currently this script has an embedded variable for the centers, and center data; this needs to use centralized variables instead.
- Pipe detection for output
	- When outputting without a connecting pipe, write to files
	- If outputting to a pipe, send the $Report variable instead
	- This can be useful if you want to do something like send emails without having an intermediary step of saving to a file, then pulling the file back in.
- Practical input options to use certificate based authentication, instead of interactive authentication.  Alternately, detect an active connection and use that if one exists.

## Version 1.0.1
**3/29/2023**
- In the Get-CS-MailboxStatisticsReportForCenter function, now using the UserPrincipalName to lookup mailbox statistics for the mailbox
	- Newer accounts have the EDOID GUID for the Identity, Id, Name and CN part of the DistinguishedName fields
	- Furthermore, the UserPrincipalName will always identify the correct account.
<<<<<<< HEAD
	
	
## Get-CS-MailboxAIO

## [1.0.0] - 2023-02-16
- Initial version. Implement Get-AllMailboxReport/Get-NewMailboxReport

## [1.1.0] - 2023-02-28
- Implement Get-CS-MailboxAIO with function Get-CS-AllMailboxHash
- with Parameter -Center (Mandatory and positional Parameter), -StartDate/-EndDate and -LastNumberofDays

## [1.2.0] - 2023-03-08
- Add comments and Help 
- Test on Module
- Create Center.csv for Centers, added to PennO365Data, use import Center.csv
- Test Private/Public key to transfer file to mailservice server.

## [1.3.0] - 2023-03-17
- remove Alias without CS

## [1.3.1] - 2023-03-23
- changed Name without objectID

## [1.3.2] -2023-04-03
- added function get-output
- added "processing" message

## [1.3.3] -2023-04-10
- added disconnect-exchangeonline to avoid duplicate records on previous scripting
=======
- When reporting the filename used to save a report for a center, the filename is now quotes encapsulated to make it easier to read.

## Version 1.1
**4/18/2023**
- Added a new function, *Get-CS-SharePointQuotaReport*, to get SharePoint usage reports
	- This function makes use of the PennO365 - Automation Runbooks Auth Read-Only App Registration in Azure AD.
	- You must [install the authentication certificate](https://confluence.apps.upenn.edu/confluence/display/TS/Import+PFX+Certificate+to+Individual+User+Account) in your local certificate store in order to use this function
>>>>>>> 3b0958928bb3cbd9e11ef086ac7e3754ce5351b8
