## Description
Allows the synchronisation of classes, teachers and students to Microsoft Teams education classes from a CSV file.

## Table of Contents
- [Overview](#overview)
- [Prerequisites](#prerequisites)
  * [Azure Active Directory App Registration](#azure-active-directory-app-registration)
  * [PowerShell](#powershell)
  * [Account Matching](#account-matching)
- [Varaibles](#variables)
- [High Level Script Process](#high-level-script-process)
- [Considerations](#considerations)
- [File Information](#file-information)

    
## Overview
 The Classes to Microsoft Teams synchronisation task uses Microsoft Graph to create Microsoft teams with the education class template for each class, adds class teachers, and supervisors as team owners, and students as team members. Students are added/removed from teams when they are added/removed from classes in the CSV.

By default, teams are created with the name format: ‘prefix +ClassName + suffix’ for example: ‘**10a-M1 (2020)**’ if the team was created in the academic year 2020 and the suffix was set to ' **(2020)**'.
The class description is set to ‘Automated team for class + the team name’ this allows HelloID to determine which teams are synchronised from this process.
The process does not create a team if a class does not have a teacher or students assigned to it in the CSV file.

## Prerequisites

* The HelloID agent is installed.
* A CSV file containing the class name, supervisor id, and student id. In the format outlined in the sampe.csv file.
* PowerShell 5.1 or higher installed on the server running the HelloID agent.
* An app registration configured in Azure Active Directory with the permissions for HelloID to perform the required actions.
* Staff and Student IDs in the CSV file must be present in the ‘**employeeId**’ attribute in AzureAD

### Azure Active Directory App Registration

HelloID uses the Microsoft Graph API to communicate with Azure, and Microsoft Teams. This requires an app registration to be in place for HelloID. More information on registering an application with Azure AD can be found here: https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app

The API permissions required are as follows:
*	Microsoft Graph
*	Directory.Read.All
*	Directory.ReadWrite.All
*	Group.Create
*	Group.Read.All
*	Group.ReadWrite.All
*	GroupMember.Read.All
*	GroupMember.ReadWrite.All
*	Team.Create
*	Team.ReadBasic.All
*	TeamMember.Read.All
*	TeamMember.ReadWrite.All
*	TeamsSettings.Read.All
*	TeamsSettings.ReadWrite.All
*	User.Read
*	User.Read.All
*	User.ReadWrite.All

The following information will be required for HelloID to connect to the app registration once configured:
*	Azure AD tenant ID
*	Azure AD tenant URL 
*	Application ID
*	Client secret

### PowerShell
The task requires PowerShell 5.1 to be installed on the HelloID agent server. PowerShell 5.1 is installed on Windows Server 2016 and 2019 by default. Please check the Microsoft link below to install PowerShell 5.1 if it is not already installed:
https://docs.microsoft.com/en-us/powershell/scripting/windows-powershell/install/windows-powershell-system-requirements?view=powershell-7

### Account Matching
Staff and student IDs must be present in the Azure AD ‘employeeId’ attribute to ensure HelloID can match staff and student records from the CSV to the associated Azure AD account. 
If you are using Azure AD Connect and have the staff and student IDs in the employeeID attribute of Active Directory, then this should already be in place.

## Variables
The following variables are required:

| Variable Name | Type | Description   |
| :------------- | ---- | :----------- |
| AADAppId | String | The Azure Active Directory application registration app ID.
| AADAppSecret | String | The Azure Active Directory application registration app secret.
| AADTenantId | String | The Azure Active Directory Tenant ID.
| AcademicYear | String | The current Academic year. For Example: ‘**2020**’ (Optional if academic year is to be used to name teams).
| AdditionalOwners | String | UserPrincipalNames of any owners you wish to be added to all teams. Separated by commas. For Example: ‘**John<span>Smith@tenant</span>.onmicrosoft.com, Jane<span>Smith@tenant</span>.onmicrosoft.com**’
| ClassNameFieldNameInSourceData | String | The name of the field containing the class name in the CSV file. For example: ‘Class’
| GroupDescriptionSearchString | String | Allows HelloID to determine which groups are part of this automation process. Recommended: '**Automated team for class***' 
| GroupPrefix | String | Allows a prefix to be added to the start of the Microsoft 365 group name. For example: ‘**Team-**‘
| GroupSuffix | String | Allows a suffix to be added to the end of the Microsoft 365 group name. For example: ‘ **(2020)**‘
| LowerProcessLimit | Integer | Stops the process from running if fewer classes than specified are returned. For example ‘**500**’. To disable the limit specify ‘**-1**’
| MemberIdFieldNameInSourceData | String | The name of the field containing the student IDs in the CSV file. For example: ‘**Student ID**’
| OwnerIdFieldNameInSourceData | String | The name of the field containing the staff IDs in the CSV file. For example: ‘**Staff ID**’
| RemoveOwners | Boolean | Stops the process removing owners from teams. If this variable is set to ‘true’ the process will add new owners to the team, but not remove existing. Allowing team owners to add extra owners to the team without them being removed by this process. If set to ‘false’ this will enforce the owners to match the source CSV file.
| SourceTeamsDataPath |  String | Specifies the location of the source CSV file to automate MS Teams creation from. This must be accessible from the server running the HelloID agent. For example: '**C:\Data\Classes.csv**'

## High level script process:

1.	Data is returned from the source CSV file and fields added to support standard field names used by the script.  A ‘Name’ field is added containing the class name, A ‘OwnerId’ field is added containing staff IDs, and a ‘MemberId’ field is added containing student IDs. Any ‘/’ are replaced with ‘-‘. This is done using the ‘Get-SourceData’ function.

2.	The script obtains an access token from the Graph API using the ‘Get-MSGraphAuthorization’ function. 

3.	Current Microsoft 365 groups are retrieved from the Graph API using the ‘Get-365Groups’ function. By default, this retrieves the top 999 groups from Graph and then filtering the description using the specified ‘groupDescriptionSearchString’ variable.

4.	Current active 365 Users are retrieved from the Graph API using the ‘Get-365Users’ function. The users’ employeeId, displayName, and userPrincipalNames are returned here.

5.	New Teams are created using the ‘educationClass’ template with a single owner, where they exist in the source CSV file, but do not currently exist in Microsoft 365.

6.	Microsoft 365 Teams are retrieved from the Graph API again and memberships for each Team are compared with the source CSV file. 

7.	Team members are added/removed as required.

8.	Team owners are added/removed as required.

## Considerations
Items to consider when using this script:

* How to supply the academic year to the script (if required). We do have a script available that automatically determines the academic year and saves it to a HelloID variable.
* Academic year changes/end of year rollup. Generally we implement the process to stop synchronising before students finish for the summer break. This means no classes change over the summer and students are still in the assigned classes until the process is configured to run again in the new year.
* The best way to archive/delete these teams in the new academic year (if applicable)

## File Information

| Filename | Description |
| :------- | :---------- |
| HelloIDVariablesExample.png | An example image showing HelloID variable configuration.
| Sample.csv | An example csv source data file.
| SyncClassesToMSTeamsFromCsv.ps1 | The PowerShell script used to create the HelloID task.
| TeamsAdminPortalExample.png | An image showing an example of Teams created by the script in the Microsoft 365 Admin Centre.

