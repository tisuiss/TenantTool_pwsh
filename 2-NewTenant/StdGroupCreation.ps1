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
  Creation Date:  <14.03.25>
#>

#----------------------------------------------------------[Functions]----------------------------------------------------------
#----------------------------------------------------------[Param]--------------------------------------------------------------
param(
    [Parameter(Mandatory=$true)]
    [string]$DataFilePath
)
#----------------------------------------------------------[Variables]----------------------------------------------------------
#Clear powershell window
clear-host
$ModuleNeeded = "Microsoft.Graph","Microsoft.Graph.Beta","ImportExcel"

$Debug = "yes"
$DebutPathFile = "C:\TenantTool\Log\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_StdGrpCreation.log"
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
if ($Debug -eq "yes") {
    start-transcript -path $DebutPathFile
    Write-Debug "Debug mode activated"
    Write-host "DataPath : $DataFilePath"
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
    #Groups importation (get content)
    $XLSContent = Import-Excel -Path $DataFilePath
    Write-Host "Import of Standard Groups will start shortly (Total Groups: $($XLSContent.Count))"
    #Get all existing group
    $AllExistingGroups = Get-MgGroup -All
    #Prepare outview at the end
    $ActionInfo = @()
    #Assign Groups
    Write-Host "Import of Static Groups"
    $StaticGroups = $XLSContent | Where-Object { $_.GroupTypes -eq "Assigned" }
    $MaxStaticGroupToCreate = $StaticGroups.Count
    $CurrentIndex = 0
    
    Foreach ($StaticGroup in $StaticGroups) {
        $CurrentIndex++
        $Name = $StaticGroup.DisplayName
    
        # Update the progress bar
        Write-Progress -Activity "Creating Statics Groups" `
                       -Status "Processing $Name ($CurrentIndex/$MaxStaticGroupToCreate)" `
                       -PercentComplete (($CurrentIndex / $MaxStaticGroupToCreate) * 100)
    
        if (-not (($AllExistingGroups).Displayname -match $name)) {
            $A = New-MgGroup -DisplayName $StaticGroup.DisplayName `
                             -MailEnabled:$False `
                             -MailNickname 'group' `
                             -SecurityEnabled
    
            # Add the owner after creating the group
            $ownerId = (Get-MgUser -UserId $StaticGroup.Owner).Id
            New-MgGroupOwner -GroupId $A.Id -DirectoryObjectId $ownerId
    
            $ActionInfo += $A | Add-Member -MemberType NoteProperty -Name "Result" -Value "Created" -Force
            Start-Sleep 5
        }
    }
    
    # Clear the progress bar when the script completes
    Write-Progress -Activity "Creating Static Groups" -Completed
    Write-Host "Import of Static Groups Completed"

    # Dynamic Groups
    Write-Host "Import of Dynamic Groups"
$DynamicGroups = $XLSContent | Where-Object { $_.GroupTypes -eq "Dynamic" }
$MaxDynamicGroupToCreate = $DynamicGroups.Count  # Corrected to use DynamicGroups.Count
$CurrentIndex = 0

Foreach ($DynamicGroup in $DynamicGroups) {
    $CurrentIndex++
    $Name = $DynamicGroup.DisplayName

    # Update the progress bar
    Write-Progress -Activity "Creating Dynamic Groups" `
                   -Status "Processing $Name ($CurrentIndex/$MaxDynamicGroupToCreate)" `
                   -PercentComplete (($CurrentIndex / $MaxDynamicGroupToCreate) * 100)
    
    if (-not (($AllExistingGroups).Displayname -match $Name)) {
        # Create the dynamic group
        $A = New-MgGroup -DisplayName $DynamicGroup.DisplayName `
                         -MailEnabled:$False `
                         -MailNickname 'group' `
                         -SecurityEnabled `
                         -GroupTypes "DynamicMembership" `
                         -MembershipRule $DynamicGroup.MembershipRule `
                         -MembershipRuleProcessingState "On"
        
        # Set the owner, assuming the owner is an email address that needs to be converted to an Object ID
        if ($DynamicGroup.Owner -eq "") {
            $BreakGlassAcc = Get-MgBetaUser -All | Where-Object {$_.JobTitle -eq "Admin" -and $_.CompanyName -eq "TenantAdmin" -and $_.Department -eq "BreakGlass"}
            if ($BreakGlassAcc.count -gt "0") {
                New-MgGroupOwner -GroupId $A.Id -DirectoryObjectId $BreakGlassAcc.Id
            }
        } else {
            $ownerId = (Get-MgUser -UserId $DynamicGroup.Owner).Id
            New-MgGroupOwner -GroupId $A.Id -DirectoryObjectId $ownerId
        }
        # Optionally, you can log the result or add it to $ActionInfo here if needed
    }
}


# Clear the progress bar when the script completes
Write-Progress -Activity "Creating Dynamic Groups" -Completed
Write-Host "Import of Dynamic Groups Completed"
Write-Host "Import of All Standard Groups Completed"

    #Control Job-GroupsCreation
    Write-Host "Create verification table"
    $GroupsCreationConfirmation = @()
    $GroupsCreationControl = Get-MgGroup -All | Where-Object { ($_.CreatedDateTime).Date -eq (Get-Date).Date } | Sort-Object CreatedDateTime,GroupTypes
    Foreach ($group in $GroupsCreationControl){
        $groupTypeFormatted = "static"
        if ($group.GroupTypes -contains "DynamicMembership") {
            $groupTypeFormatted = "Dynamic"
        }
        elseif ($group.GroupTypes -contains "Unified") {
            $groupTypeFormatted = "static"
        }
        $GroupsCreationConfirmation += [PSCustomObject]@{
            "DisplayName" = $group.DisplayName
            "Number_Members" = (Get-MgGroupMember -GroupId $Group.id -All).count
            "Created" = $group.CreatedDateTime
            "Type" = $groupTypeFormatted
            "ID" = $group.Id
        }
    }

    ConvertTo-Json -InputObject $GroupsCreationConfirmation | Out-File -FilePath "C:\TenantTool\Script\GroupsCreationConfirmation.json"

    Write-Host "Verification table Comleted"
}

Disconnect-Graph

if ($Debug -eq "Yes") {
    Stop-Transcript
}

exit
