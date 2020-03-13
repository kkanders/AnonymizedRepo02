<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>






if(!$credential ){

$credential = Get-credential myadminteamuseruser@mydomain.no

}

### MICROSOFT ONLINE ###
$tennant = "@mytennant.onmicrosoft.com"
$domain = "@my.domain.no"
$spsite = "https://mytennant-admin.sharepoint.com"

Import-Module MSOnline
Connect-MSOLService -credential $credential
$so = New-PSSessionOption -IdleTimeout 600000
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -credential $credential -Authentication Basic -AllowRedirection -SessionOption $so
$ImportResults = Import-PSSession $Session

### MICROSOFT TEAMS ###
Connect-MicrosoftTeams -credential $credential
Start-Sleep -Seconds 15 # Because of errors with the first New-Team creation, sometimes resulting in: Error occurred while executing Code: GeneralException Message: Failed to start/restart provisioning of Team

### SHAREPOINT ONLINE ###
Connect-SPOService -Url $spsite -credential $credential





Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$TeamForm                        = New-Object system.Windows.Forms.Form
$TeamForm.ClientSize             = '600,400'
$TeamForm.text                   = "KKTeamMaker"
$TeamForm.TopMost                = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "TeamNavn"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(10,12)
$Label1.Font                     = 'Microsoft Sans Serif,10'

$txtTmName                       = New-Object system.Windows.Forms.TextBox
$txtTmName.multiline             = $false
$txtTmName.width                 = 211
$txtTmName.height                = 20
$txtTmName.location              = New-Object System.Drawing.Point(32,40)
$txtTmName.Font                  = 'Microsoft Sans Serif,10'

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "TeamEier"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(10,81)
$Label2.Font                     = 'Microsoft Sans Serif,10'

$txtTmOwne                       = New-Object system.Windows.Forms.TextBox
$txtTmOwne.multiline             = $false
$txtTmOwne.width                 = 208
$txtTmOwne.height                = 20
$txtTmOwne.location              = New-Object System.Drawing.Point(32,120)
$txtTmOwne.Font                  = 'Microsoft Sans Serif,10'



$LblDesc                          = New-Object system.Windows.Forms.Label
$LblDesc.text                     = "Beskrivelse"
$LblDesc.AutoSize                 = $true
$LblDesc.width                    = 25
$LblDesc.height                   = 10
$LblDesc.location                 = New-Object System.Drawing.Point(350,12)
$LblDesc.Font                     = 'Microsoft Sans Serif,10'


$TxtTeamDescription                       = New-Object system.Windows.Forms.TextBox
$TxtTeamDescription.multiline             = $false
$TxtTeamDescription.width                 = 208
$TxtTeamDescription.height                = 20
$TxtTeamDescription.location              = New-Object System.Drawing.Point(350,40)
$TxtTeamDescription.Font                  = 'Microsoft Sans Serif,10'



$tmUserListBox                   = New-Object system.Windows.Forms.ListBox
$tmUserListBox.text              = "listBox"
$tmUserListBox.width             = 522
$tmUserListBox.height            = 182
$tmUserListBox.location          = New-Object System.Drawing.Point(34,166)

$TmBtnFinduser                   = New-Object system.Windows.Forms.Button
$TmBtnFinduser.text              = "Finn eier"
$TmBtnFinduser.width             = 131
$TmBtnFinduser.height            = 30
$TmBtnFinduser.location          = New-Object System.Drawing.Point(249,120)
$TmBtnFinduser.Font              = 'Microsoft Sans Serif,10'

$TmBtnMakeTeam                   = New-Object system.Windows.Forms.Button
$TmBtnMakeTeam.text              = "Lag Team"
$TmBtnMakeTeam.width             = 131
$TmBtnMakeTeam.height            = 36
$TmBtnMakeTeam.location          = New-Object System.Drawing.Point(245,345)
$TmBtnMakeTeam.Font              = 'Microsoft Sans Serif,10'

$TeamForm.controls.AddRange(@($txtTmName,$Label1,$txtTmOwne,$Label2,$tmUserListBox,$TmBtnFinduser,$TmBtnMakeTeam,$TxtTeamDescription,$LblDesc))

$TmBtnFinduser.Add_Click({ TmfnFindOwner $this $txtTmOwne.text })
$TmBtnMakeTeam.Add_Click({ TmBtnMakeTeam $tmUserListBox.SelectedItem $txtTmName.text $TxtTeamDescription.Text })


function TmfnFindOwner ($this,$TeamEier) {
$tmUserListBox.Items.Clear()

foreach ($user in Get-ADUser $TeamEier) {
$tmUserListBox.Items.Add($user.userprincipalname)

}
 }


 


function TmBtnMakeTeam ($uname,$tmName,$tmdesc) {







### Logger function (write log to a "log" directory in the script relative path)
function Write-Log { param( [string]$logText )
    $logFullPath = "log\" + (Get-Date -Format "yyyy-MM-dd") + ".txt" 
    $logLine = (Get-Date -Format "yyyy-MM-dd HH:mm:ss ") + $logText
    $tmUserListBox.Items.Add($logLine)
#    Write-Output $logLine | Out-File $logFullPath  -Append -Encoding utf8
}

Write-Log "INF: Script started"
$Error.Clear()


# Verify we are able to connect to and retrieve Group information
$testGroup = Get-UnifiedGroup -ResultSize 1 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
if ([string]::IsNullOrEmpty($testGroup.Name))
{
    Write-Log "ERR: Failed to connect to Teams, ending script execution"
    Break # End script
}



Write-Log "INF: Starting Teams creation and maintenance"

 $existingTeam = Get-UnifiedGroup -Identity "$tmName" -ErrorAction SilentlyContinue
        if($existingTeam -ne $null) # If True Team already exists -> update membership, if False Team does not exist -> create new Team)
        {
            $existingID = $existingTeam.ExternalDirectoryObjectId
            
            Write-Log "INF: TeamID found for un-linked department $tmName. Creating link to Team $tmName with TeamID $existingID"
            #TODO: Check if we have any first, and do manual verification before this is enabled
            #Set-ADGroup -Identity $depFromAD -Replace @{extensionAttribute3=$existingTeam.ExternalDirectoryObjectId} # Link the TeamID to the Active Directory department group
        }
        else # No Team exists. Create a new Team and link it to the AD group
        {
            Write-Log "INF: TeamID not found for $tmName. Creating new Team: $tmName"
            $newTeamID = (New-Team -DisplayName "$tmName" -Visibility Private -Description "$tmName $tmdesc" -AllowCreateUpdateRemoveConnectors $false -AllowCreateUpdateRemoveTabs $false -AllowDeleteChannels $false -AllowCreateUpdateChannels $false -AllowAddRemoveApps $false).GroupId

            if([string]::IsNullOrEmpty($newTeamID))
            {
                Write-Log "ERR: Team creation for $tmName failed"
                Continue
            }
        }
    

     Add-TeamUser -GroupId $newTeamID -Role Owner -User $uname
     Remove-TeamUser -GroupId $newTeamID -User $credential.UserName 
      Remove-UnifiedGroupLinks -Identity $newTeamID -LinkType Owners -Links $credential.UserName -Confirm:$false    

            if($Error) 
            {
                $ex = $Error[0]
                Write-Log "ERR: Error adding member $uname to $tmName Exception was: $ex"
            }
            else
            {
                Write-Log "INF: Added member $uname to $tmName"
            }




}





#Write your logic code here

[void]$TeamForm.ShowDialog()