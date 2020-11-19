#region Microsoft Graph API URIs
$baseUri = "https://graph.microsoft.com/"
$groupsUri = $baseUri + "v1.0/groups"
$teamsUri = $baseUri +  "/v1.0/teams"
$usersUri = $baseUri + "v1.0/users"
$directoryObjectsUri = $baseUri + "v1.0/directoryObjects"

#endregion Microsoft Graph API URIs

#region Functions
function Get-MSGraphAuthorization {
param(  [Parameter(Mandatory=$true)][String]$TenantID,
        [Parameter(Mandatory=$true)][String]$AppID,
        [Parameter(Mandatory=$true)][String]$AppSecret)
    # Set TLS to accept TLS, TLS 1.1 and TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

    #Get access token
    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"
 
    $body = @{
                grant_type = "client_credentials"
                client_id = "$AADAppId"
                client_secret = "$AADAppSecret"
                resource = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token
    $accessTokenExpiry =   get-date ([timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($response.expires_on))) -UFormat %s
    $authorization = @{
                        Authorization = "Bearer $accesstoken"
                        'Content-Type' = "application/json"
                        Accept = "application/json"
    }
    return $authorization, $accessTokenExpiry
}

function Get-SourceData {
    param([Parameter(Mandatory=$true)][String]$FilePath,[Parameter(Mandatory=$true)][String]$classNameFieldName,[Parameter(Mandatory=$true)][String]$ownerIdFieldName ,[Parameter(Mandatory=$true)][String]$memberIdFieldName)
           $sourceData = Import-Csv $FilePath | Where-Object {($_.$memberIdFieldName -ne '' -and $_.$memberIdFieldName -ne $null) -and ($_.$ownerIdFieldName -ne $null -and $_.$ownerIdFieldName -ne '')} | `
           Add-Member -MemberType AliasProperty -Name Name -Value $classNameFieldName -PassThru -Force | `
           Add-Member -MemberType AliasProperty -Name OwnerId -Value $ownerIdFieldName -PassThru -Force | `
           Add-Member -MemberType AliasProperty -Name MemberId -Value $memberIdFieldName -PassThru  -Force     
           $sourceData  = $sourceData | ForEach-Object {
                            $_.Name = $_.Name -replace '/', '-'   # Replace '/' with a '-'
                            $_.Name = $groupPrefix + $_.Name + $groupSuffix
                            $_                                    
                        }
           return $sourceData | Select-Object -Property Name, OwnerId, MemberId 
}

function Get-365Groups {
param(  [Parameter(Mandatory=$true)][String]$groupsUri,
        [Parameter(Mandatory=$true)][String]$searchString,
        [Parameter(Mandatory=$true)]$authorization)
        $groupsResponse = Invoke-RestMethod -Uri "$($groupsUri)?`$top=999" -Method GET -Headers $authorization -Verbose:$false
        $groups = $groupsResponse.value 
        while($groupsResponse.'@odata.nextLink') {
            $groupsResponse = Invoke-RestMethod -Uri $groupsResponse.'@odata.nextLink' -Method GET -Headers $authorization -Verbose:$false
            $groups += $groupsResponse.value
        }

        return $groups | Where-Object {$_.description -like 'Automated team for class*'} | Select-Object -Property * -ErrorAction SilentlyContinue
}

function Get-365Users {
param(  [Parameter(Mandatory=$true)]$authorization,
        [Parameter(Mandatory=$true)]$Uri)
        $usersResponse = Invoke-RestMethod -Uri "$Uri`?`$select=id,employeeid,displayName,userPrincipalName&`$filter=accountEnabled%20eq%20true" -Method GET -Headers $authorization -Verbose:$false
        $users = $usersResponse.value
        while($usersResponse.'@odata.nextLink') {
        $usersResponse = Invoke-RestMethod -Uri $usersResponse.'@odata.nextLink' -Method GET -Headers $authorization -Verbose:$false
        $users += $usersResponse.value
  }
        return $users
}
#endregion Functions

Hid-Write-Status -Message "Synchronisation started." -Event Information 
Hid-Write-Status -Message "Academic year: $AcademicYear." -Event Information 
Hid-Write-Status -Message "Source data path: $sourceTeamsDataPath." -Event Information 
Hid-Write-Summary -Message "Synchronisation started." -Event Information -Icon fa-tasks
Hid-Write-Summary -Message "Academic year: $AcademicYear ." -Event Information -Icon fa-tasks

#region Get source data
try {
        Hid-Write-Status -Message "Retrieving source data..." -Event Information
        # Get source data. If the data is not returned or the number of classes returned is less than lower process limit, then throw an error
        $sourceData = Get-SourceData -FilePath $sourceTeamsDataPath -classNameFieldName $classNameFieldNameInSourceData -ownerIdFieldName $ownerIdFieldNameInSourceData -memberIdFieldName $memberIdFieldNameInSourceData
        if(-not($sourceData) -or ($sourceData | Select-Object -Unique -Property Name).count -lt $LowerProcessLimit) {
             throw "Check the report file exists, and returns more than $LowerProcessLimit classe(s), then try again."
        }
        Hid-Write-Status -Message "Successfully retrieved $($sourceData.count) records from source data." -Event Information
        Hid-Write-Status -Message "Generating unique teams list..." -Event Information
        # Get unique classes from the source data
        $sourceGroups = $sourceData | Select-Object -Property Name | Sort-Object -Property Name  -Unique -ErrorAction Stop
        Hid-Write-Status -Message "Successfully retrieved $($sourceGroups.count) unique records." -Event Information
}
catch {
        Hid-Write-Status -Message "Error retrieving data from Sims .Net. $($_.Exception.Message)" -Event "Error"
        Hid-Write-Status -Message "Synchronisation finished." -Event "Information"
        Hid-Write-Summary -Message "Error retrieving data from Sims .Net. $($_.Exception.Message)" -Event Error -Icon fa-tasks
        Hid-Write-Summary -Message "Synchronisation finished." -Event Information -Icon fa-tasks 
        exit
}
#endregion Get source data from MIS

#region Get Graph API authorisation
try {
        # Get graph API authorisation code and access token expiry date
        $authorization, $accessTokenExpiry = Get-MSGraphAuthorization -TenantID $AADTenantID -AppID $AADAppId -AppSecret $AADAppSecret
        Hid-Write-Status -Message "Authorisation token retrieved from Microsoft Graph." -Event Information  
}
catch {
        Hid-Write-Status -Message "Error retrieving authorisation from Microsoft Graph. $($_.Exception.Message)" -Event "Error"
        Hid-Write-Status -Message "Synchronisation finished." -Event "Information"
        Hid-Write-Summary -Message "Error retrieving authorisation from Microsoft Graph. $($_.Exception.Message)" -Event Error -Icon fa-tasks
        Hid-Write-Summary -Message "Synchronisation finished." -Event Information -Icon fa-tasks 
        exit
}
#endregion Get Graph API authorization

#region Get target data from Microsoft 365
try {
        Hid-Write-Status -Message "Retrieving Microsoft 365 group data..." -Event Information
        # Retrieve groups from Microsoft 365 where the description conains our group description format (to ensure we work with the correct groups)
        $365Groups = Get-365Groups -groupsUri $groupsUri -searchString $groupDescriptionSearchString -authorization $authorization -ErrorAction Stop
        Hid-Write-Status -Message "Successfully retrieved $($365Groups.count) groups." -Event Information
        Hid-Write-Status -Message "Retrieving Microsoft 365 user data..." -Event Information
        # Retrieve users from Microsoft 365 where the employeeId attribute is not null (managed users)
        $365Users = Get-365Users -Uri $usersUri -authorization $authorization -ErrorAction Stop | where-object {$_.employeeid -ne $null}
        Hid-Write-Status -Message "Successfully retrieved $($365Users.count) users." -Event Information
}
catch {
        Hid-Write-Status -Message "Error retrieving data from Microsoft 365. $($_.Exception.Message)" -Event "Error"
        Hid-Write-Status -Message "Synchronisation finished." -Event "Information"
        Hid-Write-Summary -Message "Error retrieving data from Microsoft 365. $($_.Exception.Message)" -Event Error -Icon fa-tasks
        Hid-Write-Summary -Message "Synchronisation finished." -Event Information -Icon fa-tasks 
        exit
}
#endregion Get target date from Microsoft 365

#region Create new teams
# Get groups from source data that are not in Microsoft
$groupsToCreate = $sourceGroups | Where-Object {$_.Name -notin $365Groups.DisplayName}
Hid-Write-Status -Message "New groups to create: $(($groupsToCreate.Name).Count)" -Event Information
Hid-Write-Summary -Message "$(($groupsToCreate.Name).Count) new teams." -Event Information -Icon fa-tasks      
foreach ($group in $groupstoCreate) {
    # Check if authorisation token is expiring, if so get new authorisation token
    if([int](Get-Date  -UFormat %s) -gt ($accessTokenExpiry - 180)) {
        try {
                Hid-Write-Status -Message "Refreshing access token.." -Event Information
                $authorization, $accessTokenExpiry = Get-MSGraphAuthorization -TenantID $AADTenantID -AppID $AADAppId -AppSecret $AADAppSecret
                Hid-Write-Status -Message "Access token refreshed successfully" -Event Information
        }
        catch {
                Hid-Write-Status -Message "Error refreshing access token. $($_.Exception.Message)" -Event "Error"
                Hid-Write-Status -Message "Process finished." -Event "Information"
                Hid-Write-Summary -Message "Error refreshing access token. $($_.Exception.Message)" -Event Error -Icon fa-tasks
                Hid-Write-Summary -Message "Synchronisation finished." -Event Information -Icon fa-tasks 
                exit    
        }
    }

    try {
            # Create new team.
            # Add first owner in list of owners for initial team creation (cannot create a team without an owner)
            $sourceNewOwnerId = ($sourceData | Where-Object {$_.Name -eq $group.Name}).OwnerId | Select-Object -First 1
            $newOwner = ($365Users | Where-Object { $_.employeeId -eq $sourceNewOwnerId }).id 
            # Build new team request body
            $newTeam = '{
                        "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates(''educationClass'')",
                        "displayName": "' + $($Group.Name) + '",
                        "description": "Automated team for class ' + $($Group.Name) + '",
                          "members": [
                                    {
                                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                                    "roles": ["owner"],
                                    "userId": "' + $newOwner +'"
                                    }
                            ]
             }'
            # Create new team
            $newTeamResponse = Invoke-RestMethod -Uri $teamsUri -Headers $authorization -Method POST -Body $newTeam -ErrorVariable RespErr
            Hid-Write-Status -Message "Successfully created Microsoft 365 class team: $($group.Name)." -Event Information
    }
    catch {
            Hid-Write-Status -Message "Error creating Microsoft 365 class team: $($group.Name). $($_.Exception.Message). Check the class has a teacher and the teacher has an account in Microsoft 365." -Event Error
    }
}

# Sleep for a while to let Microsoft Graph group creations finish.
 Start-Sleep 5

#endregion Create new teams

#region Compare teams owners and members
try {
        Hid-Write-Status -Message "Retrieving Microsoft 365 group data for owner and member comparison..." -Event Information
        # Get Microsoft 365 groups again to retrieve any newly created teams
        $365Groups = Get-365Groups -groupsUri $groupsUri -searchString $groupDescriptionSearchString -authorization $authorization -ErrorAction Stop
        Hid-Write-Status -Message "Successfully retrieved $($365Groups.count) groups for owner and member comparison." -Event Information
}
catch {
        Hid-Write-Status -Message "Error retrieving group data from Microsoft 365 Unable to perform owner and member comparisons. $($_.Exception.Message)" -Event "Error"
        Hid-Write-Status -Message "Synchronisation finished." -Event "Error"
        Hid-Write-Summary -Message "Synchronisation finished." -Event Information -Icon fa-tasks 
        exit
}

Hid-Write-Status -Message "Comparing owners and members for $($365Groups.count) groups." -Event Information
Hid-Write-Summary -Message "$($365Groups.count) teams compared." -Event Information -Icon fa-tasks
foreach($group in $365Groups){
        # Initialise for each variables as null to ensure they are blank before the group is processed 
        $sourceMembers = $null; $targetMembers = $null;$membersToAdd = $null; $membersToRemove = $null;$membersComparison = $null; $groupMember = $null
        $sourceOwners = $null; $targerOwners = $null; $ownersToAdd=$null; $ownersToRemove = $null; $ownersComparison = $null; $groupOwner = $null

        # Check if authorisation token is expiring, if so get a new authorisation token
        if([int](Get-Date  -UFormat %s) -gt ($accessTokenExpiry - 180)) {
            try {
                Hid-Write-Status -Message "Refreshing access token.." -Event Information
                # Get new authorisation token
                $authorization, $accessTokenExpiry = Get-MSGraphAuthorization -TenantID $AADTenantID -AppID $AADAppId -AppSecret $AADAppSecret
                Hid-Write-Status -Message "Access token refreshed successfully" -Event Information
            }
            catch {
                Hid-Write-Status -Message "Error refreshing access token. $($_.Exception.Message)" -Event "Error"
                Hid-Write-Status -Message "Synchronisation finished." -Event "Error"
                Hid-Write-Summary -Message "Synchronisation finished." -Event Information -Icon fa-tasks 
                exit    
            }
        }

        Hid-Write-Status -Message "Checking and updating memberships for 365 team: $($group.DisplayName)..." -Event Information
          # get owners and members from source
          $sourceClassOwners = ($sourceData | Where-Object {$_.Name -eq $group.DisplayName}).OwnerId
          $sourceClassMembers = ($sourceData | Where-Object {$_.Name -eq $group.DisplayName}).MemberId
          # Add extra owners specified in variable
          $AdditionalOwners = $AdditionalOwners -Split ","
          foreach($o in $AdditionalOwners) {
                # Add additional owner to source class owners
                $sourceClassOwners += ($365Users | Where-object {$_.userPrincipalName -eq $o}).employeeId
         }
         
          # Retrieve Microsoft 365 accounts for owners in the source
          $sourceOwners =  $365Users | Where-Object {$_.employeeId -ne $null -and $_.employeeId -in $sourceClassOwners}
          # Get current owners from Microsoft 365
          $targetOwners = (Invoke-RestMethod -Uri "$groupsUri/$($group.Id)/owners" -Method GET -Headers $authorization  -Verbose:$false).Value

          # Retrieve Microsoft 365 accounts for members in the source
          $sourceMembers =  $365Users | Where-Object {$_.employeeId -ne $null -and $_.employeeId -in $sourceClassMembers}
          # Get current members from Microsoft 365
          $targetMembers = (Invoke-RestMethod -Uri "$groupsUri/$($group.Id)/members" -Method GET -Headers $authorization  -Verbose:$false).Value

          
        # If source members and target members exist compare to determine members to add/remove
        if($sourceMembers -and $targetMembers) {
            $membersComparison = Compare-Object -ReferenceObject $targetMembers.id -DifferenceObject $sourceMembers.id 
            $membersToAdd = ($membersComparison | Where-Object {$_.SideIndicator -eq "=>"}).InputObject
            $MembersToRemove = ($membersComparison | Where-Object {$_.SideIndicator -eq "<="}).InputObject
            }
        # If only source members exist, then all should be added to the team
        elseif($sourceMembers -and -not($targetMembers)) {
            $membersToAdd = $sourceMembers.id
        }
        # If only target members exist, then all should be removed from the team
        elseif(-not($sourceMembers -and $targetMembers)) {
            $membersToRemove = $targetMembers.id
        }

        # Add/remove members
        foreach($member in $membersToAdd) {
            $memberUpn = ($365Users | Where-Object {$_.id -eq $member}).userPrincipalName
            Hid-Write-Status -Message "Adding member $memberUpn to team: $($group.DisplayName)." -Event Information
            $groupMember = [PSCustomObject]@{
                        "@odata.id" = "$directoryObjectsUri/$($member)"
                    }
            # Add the owner
            Invoke-RestMethod -Uri "$groupsUri/$($group.id)/members/`$ref" -Method POST -Headers $authorization -Body ($groupMember | ConvertTo-Json -Depth 10) -Verbose:$false
            }


            foreach($member in $membersToRemove) {
                # If the member is not an owner remove the member
                if(-not($member -in $sourceOwners)) {
                    $memberUpn = ($365Users | Where-Object {$_.id -eq $member}).userPrincipalName
                    Hid-Write-Status -Message "Removing member $memberUpn from team: $($group.DisplayName)." -Event Information
                    # Remove the owner
                    Invoke-RestMethod -Uri "$groupsUri/$($group.Id)/members/$($member)/`$ref" -Method DELETE -Headers $authorization  -Verbose:$false     
                }  
            }
             # If source owners and target owners exist compare to determine members to add/remove
            if($sourceOwners -and $targetOwners) {
                $ownersComparison = Compare-Object -ReferenceObject $targetOwners.id  -DifferenceObject $sourceOwners.id
                $ownersToAdd = ($ownersComparison | Where-Object {$_.SideIndicator -eq "=>"}).InputObject
                $ownersToRemove = ($ownersComparison | Where-Object {$_.SideIndicator -eq "<="}).InputObject
            }
            # If only source owners exist, then all should be added to the team
             elseif($sourceOwners -and -not($targetOwners)) {
                $ownerssToAdd = $sourceOwners.id
            
            }
            # If only target owners exist, then all should be removed from the team
            elseif(-not($sourceOwners -and $targetOwners)) {
                $ownersToRemove = $targetOwners.id
            }

            # Add/remove owners
            if($ownersToAdd) {
                 foreach($owner in $ownersToAdd) {
                    $ownerUpn = ($365Users | Where-Object {$_.id -eq $owner}).userPrincipalName
                     Hid-Write-Status -Message "Adding owner $ownerUpn to team: $($group.DisplayName)." -Event Information
                      $groupOwner = [PSCustomObject]@{
                            "@odata.id" = "$usersUri$($owner)"
                        }
                      # Add the owner
                      Invoke-RestMethod -Uri "$groupsUri/$($group.Id)/owners/`$ref" -Method POST -Headers $authorization -body ($groupOwner | ConvertTo-Json -Depth 10)  -Verbose:$false 
                }
            }

            # Check if there are owners to remove
            if($ownersToRemove -and $removeOwners) {
                # If there are more than two owners remove the owner
                if(-not($targetOwners.Count -lt 2)) {
                    foreach($owner in $ownersToRemove) {
                        $ownerUpn = ($365Users | Where-Object {$_.id -eq $owner}).userPrincipalName
                         Hid-Write-Status -Message "removing owner $ownerUpn from team: $($group.DisplayName)." -Event Information      
                         try {
                            # Remove the owner
                            Invoke-RestMethod -Uri "$groupsUri/$($group.Id)/owners/$($owner)/`$ref" -Method DELETE -Headers $authorization  -Verbose:$false
                  
                         }
                         catch {
                              Hid-Write-Status -Message "Error removing owner from team. Check the group has at least one owner once this owner is removed." -Event Error
                         }
                    }
                }
                else {
                    # If only one owner exists, they cannot be removed. Write this to HelloID status.
                    $ownerUpn = ($365Users | Where-Object {$_.id -eq $targetOwners.id}).userPrincipalName
                    Hid-Write-Status -Message "A team requires at least one owner. $ownerUpn cannot be removed from this team." -Event Information
                }
            }
}

#endregion Compare teams owners and members
