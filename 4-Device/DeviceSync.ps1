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
  Creation Date:  <18.03.25>
#>
#----------------------------------------------------------[Param]--------------------------------------------------------------

#----------------------------------------------------------[Debug]--------------------------------------------------------------
$Debug = "No"
$DebutPathFile = "C:\TenantTool\Log\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_DeviceSync.log"
if ($Debug -eq "Yes") {
    start-transcript -path $DebutPathFile
}
#----------------------------------------------------------[Functions]----------------------------------------------------------

#----------------------------------------------------------[Variables]----------------------------------------------------------
#Clear powershell window
clear-host
$ModuleNeeded = "Microsoft.Graph","Microsoft.Graph.Beta"
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

#----------------------------------------------------------[Connexions]---------------------------------------------------------
try {
    $IsGraphSign = Get-MgContext
    $IsGraphSign
    if ($IsGraphSign.count -gt "0") {
        Disconnect-MgGraph | Out-Null
        start-sleep 2
        Connect-MgGraph -scope "DeviceManagementManagedDevices.PrivilegedOperations.All" -NoWelcome
    }
    else {
        Connect-MgGraph -NoWelcome
    }
} Catch {
    Write-Host "Error: $_"
}
#----------------------------------------------------------[Automated Variables]------------------------------------------------
$OrgName = (Get-MgOrganization).DisplayName

#----------------------------------------------------------[Execution]----------------------------------------------------------
#Validate the execution on the tenant
$ValidAnswers = @("Yes", "No", "Y", "N", "O", "Oui", "Non")
$response = ""
do {
    $response = Read-Host "You are connected on: $($OrgName). Do you want to continue ? (Y/N)"
} while ($ValidAnswers -notcontains $response)

#Answer Yes
if ($response -eq "O" -or $response -eq "Oui" -or $response -eq "Y" -or $response -eq "Yes") {
    $Devices = Get-MgDeviceManagementManagedDevice
    
    $TotalDevices = $Devices.Count
    $CurrentDevice = 0
    $DeviceSync =@()

    Foreach ($Device in $Devices){
        $CurrentDevice++

        # Affichage de la barre de progression
        $ProgressPercent = ($CurrentDevice / $TotalDevices) * 100
        Write-Progress -Activity "Sync device : $($Device.DeviceName)" -Status "Processing: $CurrentDevice / $TotalDevices" -PercentComplete $ProgressPercent

        Sync-MgDeviceManagementManagedDevice -ManagedDeviceId $Device.Id
        $LastSync = (Get-MgDeviceManagementManagedDevice | Where-Object {$_.devicename -eq $Device.DeviceName}).Lastsyncdatetime

        $DeviceSync += New-Object PSObject -Property @{
            DeviceName = $Device.DeviceName
            ComplianceState = $Device.ComplianceState
            LastSync = $LastSync
        }

        Write-Host "Sending Intune Sync request to $($Device.DeviceName) - $($Device.ComplianceState), LastSync : $LastSync"
    }

    # Fin de la barre de progression
    Write-Progress -Activity "Creating Users..." -Status "Completed" -Completed

    $DeviceSync | Export-Excel -Path "C:\TenantTool\Script\DeviceSync.xlsx" -AutoSize -TableName "DeviceSync" -Append
} 

#Answer No
if ($response -eq 7) {

}

Disconnect-Graph

if ($Debug -eq "Yes") {
    Stop-Transcript
}

exit
