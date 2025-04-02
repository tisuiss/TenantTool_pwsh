<#
.DESCRIPTION
    <Brief description of script>
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
    <Inputs if any, otherwise state None - example: File containing list of servers>
.OUTPUTS
    <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>
.NOTES
  Version:        1.0
  Author:         <tisuiss>
  Creation Date:  <16.03.25>
#>

#----------------------------------------------------------[Param]--------------------------------------------------------------
param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUsrFilePath,
    [string]$AdminGrpFilePath
)
#----------------------------------------------------------[Functions]----------------------------------------------------------
function New-RandomPassword {
    param (
        [int]$Length = 12
    )
    
    $lowercase = "abcdefghijklmnopqrstuvwxyz"
    $uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $digits = "0123456789"
    $special = "!@#$%*=:.?"
    
    $validCharacters = $lowercase + $uppercase + $digits + $special
    $passwordArray = @()

    # Ensure at least one of each character type
    $passwordArray += $lowercase[(Get-Random -Maximum $lowercase.Length)]
    $passwordArray += $uppercase[(Get-Random -Maximum $uppercase.Length)]
    $passwordArray += $digits[(Get-Random -Maximum $digits.Length)]
    $passwordArray += $special[(Get-Random -Maximum $special.Length)]

    # Fill the rest of the password length with random characters
    for ($i = 0; $i -lt ($Length - 4); $i++) {
        $passwordArray += $validCharacters[(Get-Random -Maximum $validCharacters.Length)]
    }

    # Shuffle the characters
    $password = -join ($passwordArray | Sort-Object { Get-Random })

    return $password
}
#----------------------------------------------------------[Variables]----------------------------------------------------------
#Clear powershell window
clear-host

$ModuleNeeded = "Microsoft.Graph","Microsoft.Graph.Beta","ImportExcel"
$Debug = "Yes"
$DebutPathFile = "C:\TenantTool\Log\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_AdminUsrCreation.log"
#----------------------------------------------------------[Module]-------------------------------------------------------------
Write-host "Checking if all needed modules are installed"
$MissModule = "0"

foreach ($Module in $ModuleNeeded) {
    $Check = Get-Module -Name $Module -ListAvailable
    if ($Check.count -gt "0") {
        Write-host "Module $Module is installed" -ForegroundColor Green
    } else {
        $MissModule = +1
        Write-host "Module $Module is not installed" -ForegroundColor Red
    }
}

if ($MissModule -gt "0") {
    Read-Host -Prompt "Please install the module $Module and execute again the command."
    exit
}

#----------------------------------------------------------[Debug]--------------------------------------------------------------
if ($Debug -eq "Yes") {
start-transcript -path $DebutPathFile
}
#----------------------------------------------------------[Connexions]---------------------------------------------------------
$scopes = "Agreement.ReadWrite.All, Application.Read.All, Application.ReadWrite.All, Bookings.Manage.All, Bookings.ReadWrite.All, BookingsAppointment.ReadWrite.All, DelegatedPermissionGrant.ReadWrite.All, DeviceManagementManagedDevices.PrivilegedOperations.All, Directory.Read.All, Directory.ReadWrite.All, EntitlementManagement.ReadWrite.All, Group.ReadWrite.All, openid, Policy.Read.All, Policy.ReadWrite.ConditionalAccess, profile, RoleManagement.Read.All, RoleManagement.Read.Directory, RoleManagement.Read.Exchange, RoleManagement.ReadWrite.Directory, RoleManagement.ReadWrite.Exchange, Team.ReadBasic.All, TeamSettings.Read.All, TeamSettings.ReadWrite.All, User.Read, User.Read.All, User.ReadWrite.All"

try {
    $IsGraphSign = Get-MgContext
    $IsGraphSign
    if ($IsGraphSign.count -gt "0") {
        Disconnect-MgGraph | Out-Null
        start-sleep 2
        Connect-MgGraph -NoWelcome -Scopes $scopes
    }
    else {
        Connect-MgGraph -NoWelcome -Scopes $scopes
    }
} Catch {
    Write-Host "Error: $_"
}

Get-MgContext
#----------------------------------------------------------[Automated Variables]------------------------------------------------
$OrgName = (Get-MgOrganization).DisplayName

#----------------------------------------------------------[Execution]----------------------------------------------------------
$ValidAnswers = @("Yes", "No", "Y", "N", "O", "Oui", "Non")
$response = ""
do {
    $UserSignIn = (Get-MgContext).Account
    write-host "Script to create Admin User and Group for $($OrgName)"
    Write-Host "User sign in actually : $($UserSignIn)"
    $response = Read-Host "You are connected on: $($OrgName). Do you want to continue ? (Y/N)"
} while ($ValidAnswers -notcontains $response)

#Answer Yes
if ($response -eq "O" -or $response -eq "Oui" -or $response -eq "Y" -or $response -eq "Yes") {
    #region Group Admin Creation
        #Groups importation (get content)  
        #$groups = Get-Content -Path $JsonGroupsFilePath | ConvertFrom-Json
        Write-host "Import informations from your excel sheet" -ForegroundColor Yellow
        $Groups = Import-Excel -Path $AdminGrpFilePath

        #Groups Creation
        Write-host "Get all security groups in your tenant" -ForegroundColor Yellow
        $AllSecuGroups = Get-MgGroup -All
        # Initialisation des variables pour la progression
        $TotalGroups = $groups.Count
        $CurrentIndex = 0

        #Control Groups
        if ($Debug -eq "Yes") {
            $groups | Format-Table -AutoSize -Wrap
        }

        Write-host "Creation of security groups will start." -ForegroundColor Yellow
        Foreach ($group in $groups) {
            $CurrentIndex++
            $GroupName = $group.DisplayName
            $GroupDescription = $group.Description
            $GroupRole = $group.RBAC_Role
        
            # Mise à jour de la barre de progression
            Write-Progress -Activity "Admin groups creation" `
                           -Status "Creation of : $GroupName ($CurrentIndex/$TotalGroups)" `
                           -PercentComplete (($CurrentIndex / $TotalGroups) * 100)
        
            $GroupCheck = $AllSecuGroups | Where-Object { $_.DisplayName -eq $GroupName }
        
            if (($GroupCheck).count -eq "0") {
                Write-host "Creation of : $($GroupName)"
                New-MgGroup -DisplayName $GroupName -MailEnabled:$false -SecurityEnabled -IsAssignableToRole:$true -MailNickName "group" -Description $GroupDescription
                Start-Sleep -Milliseconds 500  # Pause pour éviter de surcharger les requêtes API

                # Set the owner, assuming the owner is an email address that needs to be converted to an Object ID
                $ownerId = (Get-MgUser -UserId $group.Owner).Id
                $groupId = (Get-MgGroup -All | Where-Object {$_.DisplayName -eq $GroupName}).Id
                New-MgGroupOwner -GroupId $groupId -DirectoryObjectId $ownerId
            
                # Ajout du rôle au groupe
                $Group4Role = Get-MgGroup | Where-Object { $_.DisplayName -eq $GroupName }
                $roleList = $GroupRole -split ',' | ForEach-Object { $_.Trim() }
            
                Foreach ($Role in $roleList) {
                    $roleDefinition = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -eq $Role }
                
                    if ($roleDefinition) {
                        New-MgRoleManagementDirectoryRoleAssignment -DirectoryScopeId '/' -RoleDefinitionId $roleDefinition.Id -PrincipalId $Group4Role.Id | Out-Null
                    } else {
                        Write-Host "Role : '$Role' doesn't exist." -ForegroundColor Red
                    }
                }
            }
        }
        Write-Progress -Activity "Admin groups creation" -Completed

    #endregion Group Admin Creation

    #region User Admin Creation
        ##Domain onmicrosoft information
        $OnMicrosoftDomain =(Get-MgOrganization).VerifiedDomains | where-object {$_.Name -match "onmicrosoft.com" -and $_.IsInitial -eq "True"}
        $AdminDomainAccount = ($OnMicrosoftDomain).Name
        #Admin User importation
        #$users = Get-Content -Path $JsonUsersFilePath | ConvertFrom-Json
        $users = Import-Excel -Path $AdminUsrFilePath

        $empty = $null  # or any other default value you want to assign
        $users | Add-Member -MemberType NoteProperty -Name "MailNickName" -Value $empty
        $users | Add-Member -MemberType NoteProperty -Name "Password" -Value $empty
        $users = foreach ($user in $users) {
            $user.UserPrincipalName = $user.UserPrincipalName -replace '@TenantName.onmicrosoft.com', "@$AdminDomainAccount"
            $user.MailNickName = "$($user.FirstName).$($user.LastName)"
            $user.MailNickName = ($user.MailNickName -replace '\s', '').ToLower()
            $user
        }

        #Users control
        if ($Debug -eq "Yes") {
            $users | Format-Table
        }

        $ClientName = (Get-MgOrganization).DisplayName
        $ClientName = ($ClientName -replace '\s', '').ToLower()

        # Iterate through each row and update UserPrincipalName
        $exportcredentials = @()
        # Initialisation des variables pour la progression
        $TotalUsers = $users.Count
        $CurrentIndex = 0
        $exportcredentials = @()  # Initialise la variable pour stocker les identités des utilisateurs

        foreach ($user in $users) {
            $CurrentIndex++
        
            # Mise à jour de la barre de progression
            Write-Progress -Activity "Admin User Creation" `
                           -Status "Creation of $($user.DisplayName) ($CurrentIndex/$TotalUsers)" `
                           -PercentComplete (($CurrentIndex / $TotalUsers) * 100)
        
            # Génération du mot de passe temporaire
            $Password = New-RandomPassword -Length 12
            $user.Password = $Password
        
            # Création de l'objet utilisateur avec ses infos
            $identity = [PSCustomObject]@{
                FirstName          = $user.FirstName
                LastName           = $user.LastName
                DisplayName        = $user.DisplayName
                UserPrincipalName  = $user.UserPrincipalName
                TemporaryPassword  = $Password
                MailNickName       = $user.MailNickName
            }
        
            # Ajout des informations de l'utilisateur dans la liste d'export
            $exportcredentials += $identity
        
            # Pause pour éviter de surcharger les requêtes API
            Start-Sleep -Milliseconds 500


            #Add Usser to the admin rights group
            $AdminUserSecuGroup = Get-MgGroup -All | Where-Object {$_.DisplayName -eq $user.SecurityGroupsRights}  
            $PasswordProfile = @{
            Password = $user.Password
            }

            if (($user).JobTitle -eq ""){
                $createduser = New-MgUser -DisplayName $user.Displayname -UserPrincipalName $user.UserPrincipalName -Surname $user.LastName -GivenName $user.FirstName -CompanyName $user.Company -Department $user.Department -UsageLocation $user.UsageLocation -AccountEnabled:$true -PasswordProfile $PasswordProfile -MailNickName $user.MailNickName
            } else {
                $createduser = New-MgUser -DisplayName $user.Displayname -UserPrincipalName $user.UserPrincipalName -Surname $user.LastName -GivenName $user.FirstName -CompanyName $user.Company -Department $user.Department -UsageLocation $user.UsageLocation -JobTitle $user.JobTitle -AccountEnabled:$true -PasswordProfile $PasswordProfile -MailNickName $user.MailNickName
            }

            #Add to group
            if ($user.SecurityGroupsRights -ne "") {
                New-MgGroupMember -DirectoryObjectId $createduser.id -GroupId $AdminUserSecuGroup.id
            }
        }

        # Fin de la barre de progression
        Write-Progress -Activity "Admin User Creation" -Completed

            #Export Connexion information + Apply password on account
        $xlsresultpath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path +"\$($ClientName)_AdminUserCreation.xlsx"
        $exportcredentials | Export-Excel -Path $xlsresultpath

            #Control Job
        $UserReportConfirmation = Get-MgBetaUser -All | select-Object DisplayName,UserPrincipalName,CreatedDateTime | Where-Object {$_.UserPrincipalName -like "adm.*"}
        $UserReportConfirmation | ConvertTo-Json -Depth 3 | Out-File "C:\TenantTool\Script\UserReportConfirmation.json"

        $GroupsReportConfirmation = @()
        $GroupsControl = Get-MgGroup -All | Where-Object {$_.Description -match "Admin Right Management"}
        Foreach ($group in $GroupsControl){
            $GroupsReportConfirmation += [PSCustomObject]@{
            "DisplayName" = $group.DisplayName
            "Number_Members" = (Get-MgGroupMember -GroupId $Group.id -All).count
            "Created" = $group.CreatedDateTime
            "ID" = $group.Id
            }
        }
        $GroupsReportConfirmation | ConvertTo-Json -Depth 3 | Out-File "C:\TenantTool\Script\GroupsReportConfirmation.json"

        #MsgBox to inform the user that the import is done and the file is available in the download folder
        $wshell = New-Object -ComObject Wscript.Shell
        $question = "Importation done. The file is available in the download folder."
        $response = $wshell.Popup($question, 0, "Importation done", 64)
    
    #endregion User Admin Creation

}

#Answer No
if ($response -eq "N" -or $response -eq "Non") {
    Write-Host "Disconnection of all tenant. Please relanch the command." -ForegroundColor Red
}

Disconnect-Graph

if ($Debug -eq "Yes") {
    Stop-Transcript
}

exit
