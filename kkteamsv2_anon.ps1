################################################################## 
################## TEAM CREATION AND MAINTENANCE ################## 
################## Baseline provided by Stavanger kommune  ########
################################################################### 

### CREDENTIALS ###
#Unsafe way
#$Username = "USERNAME@DOMAIN.NO"
#$cred$secpasswd = ConvertTo-SecureString PASSWORD -AsPlainText -Force
#$Credential = New-Object System.Management.Automation.PSCredential($Username,$secpasswd)
# Safe way:
# wil only work under the user that create the XML
# 
# $myCred = Get-Credentials
# $mycred  |  Export-Clixm pathTolcixml.xml


#Set up the Environment for Karmøy kommune

cd C:\o365\Powershell\Teams
$Credential = Import-Clixml .\mySecureXML.xml 

#########
# Change mytennantname yto your tennant
########
$spSite = "https://mytennantname-admin.sharepoint.com"
$ADBaseOU = "OU=Avdelingsgrupper,ou=my,dc=domain,dc=no"
# Local "database-file" of the groups ExternalDirectoryObjectId that already has been processed
# Want to use a SQLite database
#$alreadyProcessedFile = ".\AlreadyProcessedSharepointSharingTeams.txt"
#$alreadyProcessed = Get-Content $alreadyProcessedFile




### Logger function (write log to a "log" directory in the script relative path)

function Write-Log { param( [string]$logText )
    $logFullPath = "log\" + (Get-Date -Format "yyyy-MM-dd") + ".txt" 
    $logLine = (Get-Date -Format "yyyy-MM-dd HH:mm:ss ") + $logText
    Write-Output $logLine | Out-File $logFullPath  -Append -Encoding utf8
}

Write-Log "INF: Script started"
$Error.Clear()


### MICROSOFT ONLINE ###
Import-Module MSOnline
Connect-MSOLService -Credential $Credential
$so = New-PSSessionOption -IdleTimeout 600000
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection -SessionOption $so
$ImportResults = Import-PSSession $Session

### MICROSOFT TEAMS ###
Connect-MicrosoftTeams -Credential $Credential
Start-Sleep -Seconds 15 # Because of errors with the first New-Team creation, sometimes resulting in: Error occurred while executing Code: GeneralException Message: Failed to start/restart provisioning of Team

### SHAREPOINT ONLINE ###
Connect-SPOService -Url $spSite -Credential $Credential

# Verify we are able to connect to and retrieve Group information
$testGroup = Get-UnifiedGroup -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
if ([string]::IsNullOrEmpty($testGroup.Name))
{
    Write-Log "ERR: Failed to connect to Teams, ending script execution"
    Break # End script
}



Write-Log "INF: Starting Teams creation and maintenance"
#Teams vi skal bruke:
$tempgroups = Get-ADGroup -Filter {info -notlike 0} -SearchBase $ADBaseOU -Properties info,managedby -SearchScope OneLevel

foreach ($grupper in $tempgroups) {

 # Reset
    $Error.Clear()
    $depFromAD = $null
    $depTeamID = $null
    $depName = $null
    $existingTeam = $null
    $newTeamID = $null

    $depName =  $grupper.Name
    $depTeamID = $grupper.info 
    if($depTeamID -eq 1) # If Active Directory object does not contain a Team ObjectID in extensionAttribute3
    {
        # Verify that no un-linked Team with the same name already exists
        $existingTeam = Get-UnifiedGroup -Identity "$depName" -ErrorAction SilentlyContinue
        if($existingTeam -ne $null) # If True Team already exists -> update membership, if False Team does not exist -> create new Team)
        {
            $existingID = $existingTeam.ExternalDirectoryObjectId
            Write-Log "INF: TeamID found for un-linked department $depName. Creating link to Team $depName with TeamID $existingID"
            #TODO: Check if we have any first, and do manual verification before this is enabled
            #Set-ADGroup -Identity $depFromAD -Replace @{extensionAttribute3=$existingTeam.ExternalDirectoryObjectId} # Link the TeamID to the Active Directory department group
        }
        else # No Team exists. Create a new Team and link it to the AD group
        {
            Write-Log "INF: TeamID not found for $depName. Creating new Team: $depName"
            $newTeamID = (New-Team -DisplayName "$depName" -Visibility Private -Description "Alle $depName" -AllowCreateUpdateRemoveConnectors $false -AllowCreateUpdateRemoveTabs $false -AllowDeleteChannels $false -AllowCreateUpdateChannels $false -AllowAddRemoveApps $false).GroupId

            if([string]::IsNullOrEmpty($newTeamID))
            {
                Write-Log "ERR: Team creation for $depName failed"
                Continue
            }
            Set-ADGroup -Identity $grupper -Replace @{info=$newTeamID} # Link the TeamID to the Active Directory department group
            if($Error)
            {
                Write-Log "ERR: Team created but unable to update Active Directory department $depName with TeamID $newTeamID. AD department should update next time this script is run (name matching). Investigate if the problem persist"
            }
        }
    }
    else # Active Directory already contains the Team objectId
    {
        Write-Log "INF: TeamID $depTeamID found for existing department $depName"
        # Verify Team name is correct, and if not, rename Team back to AD depName if any owners renamed the Team. Department Teams should only be renamed by renaming the AD Department group using Windows Adminstrasjon
        $existingTeam = Get-UnifiedGroup -Identity $depTeamID
        if ($existingTeam.DisplayName -ne "$depName") # Team has been renamed by an owner, rename back to Active Directory department group name
        {
            $existingTeamName = $existingTeam.DisplayName
            Write-Log "INF: Setting DisplayName to $depName from $existingTeamName"
            Set-UnifiedGroup -Identity $depTeamID -DisplayName "$depName"

        }
        if (!$existingTeam.AccessType -eq "Private") # All groups that are automatically maintained should be private, check and reset if any owner changed this
        {
            Write-Log "INF: Setting AccessType to Private for $depName"
            Set-UnifiedGroup -Identity "$depName" -AccessType "Private"
        }
        if (!$existingTeam.HiddenFromExchangeClientsEnabled -eq $True) # Do not show Team in Outlook clients
        {
            Write-Log "INF: Setting HiddenFromExchangeClientsEnabled for $depName"
            Set-UnifiedGroup -Identity "$depName" -HiddenFromExchangeClientsEnabled
        }
        if (!$existingTeam.HiddenFromAddressListsEnabled -eq $True) # Do not show Team in Exchange Address Book
        {
            Write-Log "INF: Setting HiddenFromAddressListsEnabled for gr.$depName"
            Set-UnifiedGroup -Identity "$depName" -HiddenFromAddressListsEnabled $True
        }
        # Always set Team picture to the known auto-maintained department picture 
        # (TODO: THIS DOES NOT WORK. The service account needs to be member of the team to change picture, global admin is not enough..gg MS, re-write and use GraphAPI instead)
        #Set-TeamPicture -GroupId $depTeamID -ImagePath .\kkteams.png
        
        # Always set Team description to the known auto-maintained department text
        Set-Team -GroupId $depTeamID -Description "Automatisk vedlikeholdt Team" -AllowCreateUpdateRemoveConnectors $false -AllowCreateUpdateRemoveTabs $false -AllowDeleteChannels $false -AllowCreateUpdateChannels $false -AllowAddRemoveApps $false
    }
#}
Write-Log "INF: Finished Teams creation and policy maintenance"

    $grupper.SamAccountName


}

$grupper = ""

################## TEAM MEMBERSHIPS MAINTENANCE ##################
Write-Log "INF: Starting Teams membership maintenance"
Start-Sleep -Seconds 10 # Allow new Teams to finish intialize before updating memberships




foreach ($grupper in $tempgroups) {
# Init
    $Error.Clear()
    $depFromAD = $null
    $depTeamID = $null
    $existingTeamOwnersUPN = $null
    $existingTeamMembersUPN = $null
    $departmentMembersUPN = $null
    $depManagedBy = $null
    $departmentOwnersUPN = $null
    $ex = $null

    # Get information from Active Directory and Microsoft Teams
#    $depFromAD = Get-ADGroup -Filter {Name -eq $grupper.Name} -SearchBase "OU=Avdelingsgrupper,OU=KarmoyKommune,OU=KKCloud,DC=karmoy,DC=kommune,DC=no" -Properties info,managedby -SearchScope Base
 
 $depFromAD = $grupper

   if($Error)
    {
        $ex = $Error[0]
        Write-Log "ERR: $($grupper.name) department not found in Active Directory. Error message was: $ex"
        Continue
    }
    $depTeamID = $depFromAD.info # When Teams are created, the GroupID is stored in Active Directory to be able to link the objects by ID instead of name (in case owners rename the teams)
    # Get existing Team owners

    $existingTeamOwnersUPN = (Get-UnifiedGroupLinks -Identity $depTeamID -LinkType Owner | Select -ExpandProperty WindowsLiveID).ToLower()
    if($Error)
    {
        $ex = $Error[0]
        Write-Log "ERR: $($grupper.name) department skipped. Error while retrieving Team owners. Team probably does not exists because the AD role department contains no valid O365 users (0 users returned). Error message was: $ex"
        Continue
    }
   
    # Get existing Team members
    $existingTeamMembersUPN = (Get-UnifiedGroupLinks -Identity $depTeamID -LinkType Members | Select -ExpandProperty WindowsLiveID).ToLower()
    if($Error)
    {
        $ex = $Error[0]
        Write-Log "ERR: $($grupper.Name) department skipped. Error while retrieving Team members. Team probably does not exists because the AD department contains no valid O365 users (0 users returned). Error message was: $ex"
        Continue
    }
    
    # Get existing AD members
    $departmentMembersUPN = (Get-ADGroupMember -Identity $grupper | Get-ADUser -Properties UserPrincipalName | Select -ExpandProperty UserPrincipalName).ToLower()
    $departmentMembersUPN
    if($Error)
    {
        $ex = $Error[0]
        Write-Log "ERR: $($grupper.Name) department skipped. Error while retrieving AD department members. Verify the AD security group exists and is linked to the SQL group. Error message was: $ex"
        Continue
    }
     
    # Get existing AD owners (members of .ManagedBy group in AD)
    $depManagedBy = $depFromAD.ManagedBy
    $departmentOwnersUPN = ($depManagedBy | Get-ADUser -Properties UserPrincipalName | Select -ExpandProperty UserPrincipalName).ToLower()
        if($Error)
    {
        $ex = $Error[0]
        Write-Log "ERR: $($grupper.Name) department skipped. Error while retrieving AD role department members. Verify the AD security group has a ROL_ managedBy link to the AD role security group. Error message was: $ex"
        Continue
    }

    
    # Add users that do not exist as members already
    foreach($user in $departmentMembersUPN)
    {
        $Error.Clear()
        if(!$existingTeamMembersUPN.Contains($user))
        {
            Add-TeamUser -GroupId $depTeamID -Role Member -User $user.ToString()
          

            if($Error) 
            {
                $ex = $Error[0]
                Write-Log "ERR: Error adding member $user to $($grupper.name) Exception was: $ex"
            }
            else
            {
                Write-Log "INF: Added member $user to $($grupper.name)"
            }
        }
    }
   
    # Add users that do not exists as owners
    foreach($user in $departmentOwnersUPN)
    {
        if(!$existingTeamOwnersUPN.Contains($user))
        {
            Add-TeamUser -GroupId $depTeamID -Role Owner -User $user.ToString()
            if($Error)
            {
                $ex = $Error[0]
                Write-Log "ERR: Error adding owner $user to $($grupper.Name) Exception was: $ex"
            }
            else
            {
                Write-Log "INF: Added owner $user to $($grupper.name)"
            }
        }
    }
     
    # Remove users that do not exist as member
    foreach($user in $existingTeamMembersUPN)
    {
        $Error.Clear()
        if(!$departmentMembersUPN.Contains($user))
        {
            Remove-UnifiedGroupLinks -Identity $depTeamID -LinkType Members -Links $user -Confirm:$false 
            if($Error) 
            {
                $ex = $Error[0]
                Write-Log "ERR: Error removing member $user from $($grupper.name) Exception was: $ex"
            }
            else
            {
                Write-Log "INF: Removed member $user from $($grupper.name)"
            }
        
        }
    }
   
    # Remove users that do not exist as owner
    foreach($user in $existingTeamOwnersUPN)
    {
        if(!$departmentOwnersUPN.Contains($user))
        {
            Remove-UnifiedGroupLinks -Identity $depTeamID -LinkType Owners -Links $user -Confirm:$false
            if($Error)
            {
                $ex = $Error[0]
                Write-Log "ERR: Error removing owner $user from $($grupper.name) Exception was: $ex"
            }
            else
            {
                Write-Log "INF: Removed owner $user from $($grupper.name)"
            }
        }
    }
     <#

    #>
}






Disconnect-MicrosoftTeams
Disconnect-SPOService
Remove-PSSession $Session