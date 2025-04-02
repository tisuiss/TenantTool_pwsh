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
  Creation Date:  <17.03.25>
#>

#----------------------------------------------------------[Functions]----------------------------------------------------------
#----------------------------------------------------------[Param]--------------------------------------------------------------
#----------------------------------------------------------[Variables]----------------------------------------------------------
$Debug = "No"
$DebutPathFile = "C:\TenantTool\Log\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_CheckModuleNeeded.log"
#----------------------------------------------------------[Debug]--------------------------------------------------------------
if ($Debug -eq "Yes") {
    start-transcript -path $DebutPathFile
}

#Clear powershell window
clear-host

Invoke-WebRequest -Uri "https://api.bitbucket.org/2.0/repositories/scripttisuiss/Repo_TenantTools/src/main/TenantTools/V3/1-Common/ModuleNeeded.json" -OutFile "C:\TenantTool\Check\CheckModuleNeeded.json"

$Modules = Get-Content "C:\TenantTool\Check\CheckModuleNeeded.json" -Raw | ConvertFrom-Json

$ModuleNeeded = $Modules.modules.name
#----------------------------------------------------------[Module]-------------------------------------------------------------
Write-host "Checking if all needed modules are installed"
$MissModule = @()

foreach ($Module in $ModuleNeeded) {
    $Check = Get-Module -Name $Module -ListAvailable
    if ($Check.count -gt "0") {
        Write-host "Module $Module is installed" -ForegroundColor Green
    } else {
        $MissModule += $Module
        Write-host "Module $Module is not installed" -ForegroundColor Red
    }
}

if ($MissModule.count -gt "0") {
    $ValidAnswers = @("Yes", "No", "Y", "N", "O", "Oui", "Non")
    $response = ""

    do {
        $response = Read-Host -Prompt "Would you like to install all missing module ? (Y/N)"
    } while ($ValidAnswers -notcontains $response)
}

#----------------------------------------------------------[Install Module]-----------------------------------------------------
if ($response -eq "O" -or $response -eq "Oui" -or $response -eq "Y" -or $response -eq "Yes") {
    foreach ($Miss in $MissModule) {
        write-host "Installing module $Miss in CurrentUser scope" -ForegroundColor Yellow
        Install-Module -Name $Miss -Force -AllowClobber -Scope CurrentUser
    }
}


#Final Check
$FinalCheck = @()

foreach ($Module in $ModuleNeeded) {
    $Check = Get-Module -Name $Module -ListAvailable
    if ($Check.count -gt "0") {
        $FinalCheck += [PSCustomObject]@{
            Name = $Module
            Installed = "Yes"
        }
    } else {
        $FinalCheck += [PSCustomObject]@{
            Name = $Module
            Installed = "No"
        }
    }
}

Remove-Item "C:\TenantTool\Check\CheckModuleNeeded.json"

if ($Debug -eq "Yes") {
    Stop-Transcript
}

Exit
#----------------------------------------------------------[Connexions]---------------------------------------------------------

#----------------------------------------------------------[Automated Variables]------------------------------------------------

#----------------------------------------------------------[Execution]----------------------------------------------------------
