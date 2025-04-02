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
$DebutPathFile = "C:\TenantTool\Log\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_PolicyCreation.log"
if ($Debug -eq "Yes") {
    start-transcript -path $DebutPathFile
}
#----------------------------------------------------------[Functions]----------------------------------------------------------
#----------------------------------------------------------[Variables]----------------------------------------------------------
#Clear powershell window
clear-host
$ModuleNeeded = "Microsoft.Graph","Microsoft.Graph.Beta","ImportExcel","MSCommerce","Microsoft.Online.SharePoint.PowerShell","MicrosoftTeams"

#check if file is available
$JsonPath = Test-Path "C:\TenantTool\Script\PolicySettings.json"
if ($JsonPath -eq $false) {
    Read-Host "The file PolicySettings.json does not exist in the folder C:\TenantTool\Script"
    exit
}

$Policies = Get-Content "C:\TenantTool\Script\PolicySettings.json" -Raw | ConvertFrom-Json

#Result table
$ResultTable = @()
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
#----------------------------------------------------------[Connexion]------------------------------------------------------
#Graph   
    try {
        Write-host "Connexion on Graph powershell module" -ForegroundColor Yellow
        $IsGraphSign = Get-MgContext
        if ($IsGraphSign.count -gt "0") {
            Disconnect-MgGraph | Out-Null
            start-sleep 2
            Connect-MgGraph -NoWelcome
        }
        else {
            Connect-MgGraph -NoWelcome
        }
    } Catch {
        Write-Host "Error: $_"
    }
    $OrgName = (Get-MgOrganization).DisplayName

#MsCommerce
    if ($Policies.admin -eq "true"){
        Write-host "Connexion on Admin of the tenant $OrgName" -ForegroundColor Yellow
        Connect-MSCommerce #log in here
    }

#Sharepoint
    if ($Policies.Sahrepoint -eq "true"){
        Write-host "Connexion on the sharepoint of the tenant $OrgName" -ForegroundColor Yellow
            $OnMicrosoftDomain = ((Get-MgOrganization).VerifiedDomains | Where-Object {$_.IsDefault -eq $true}).name
            $OnMicrosoftDomain = $OnMicrosoftDomain -replace ".onmicrosoft.com",""
            Connect-SPOService "https://$OnMicrosoftDomain-admin.sharepoint.com"
    }

#Teams
    if ($Policies.teams -eq "true"){
        Write-host "Connexion on the Teams of the tenant $OrgName" -ForegroundColor Yellow
    }
#----------------------------------------------------------[Admin actions]------------------------------------------------------
#MsCommerce
    if ($Policies.DisableMsCommerce -eq "true") {

        $products = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | Where-Object { $_.PolicyValue -eq "Enabled"}
        foreach ($p in $products){
            Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $p.ProductId -Enabled $False
        }   
        $ResultTable += [PSCustomObject]@{
            "Section" = "Admin"
            "Setting" = "DisableMsCommerce"
            "Action" = "True"
        }
    } 
#----------------------------------------------------------[Intune/Entra actions]-----------------------------------------------


#----------------------------------------------------------[Sharepoint actions]-------------------------------------------------
    if ($Policies.Sharepoint -eq "true") {
    #Disable Shortcut in OneDrive/Sharepoint
        Set-SPOTenant   -DisableAddShortcutsToOneDrive $true `
                        -MajorVersionLimit 50 `
                        -ExpireVersionsAfterDays 30 `
                        -EnableAutoExpirationVersionTrim $false `
                        -SharingCapability ExistingExternalUserSharingOnly `
                        -LegacyAuthProtocolsEnabled $false

        $ShpSettings = Get-SPOTenant | ForEach-Object {
            [PSCustomObject]@{
                Name    = "DisableAddShortcutsToOneDrive"
                Settings = $_.DisableAddShortcutsToOneDrive
            }
            [PSCustomObject]@{
                Name    = "MajorVersionLimit"
                Settings = $_.MajorVersionLimit
            }
            [PSCustomObject]@{
                Name    = "ExpireVersionsAfterDays"
                Settings = $_.ExpireVersionsAfterDays
            }
            [PSCustomObject]@{
                Name    = "SharingCapability"
                Settings = $_.SharingCapability
            }
            [PSCustomObject]@{
                Name    = "LegacyAuthProtocolsEnabled"
                Settings = $_.LegacyAuthProtocolsEnabled
            }
        }

        foreach ($Setting in $ShpSettings) {
            $ResultTable += [PSCustomObject]@{
                "Section" = "Sharepoint"
                "Name" = $Setting.Name
                "Setting" = $Setting.Settings
            }
        }

        Write-Host "Some settings can't be change in Powershell, Please do it manually : "
        $checklist = @(
        "Do you disable site creation by users"
        "Do you set default time zone when create a new site"
        "Do you change default site storage to 200GB"
        )

        foreach ($item in $checklist) {
            do {
                $reponse = Read-Host "$item ? (Y/N)"
            } while ($reponse -notmatch "^[OoNnYy]$") 

            if ($reponse -match "^[OoYy]$") { 
                $ResultTable += [PSCustomObject]@{
                    "Section" = "Sharepoint"
                    "Name" = $item
                    "Setting" = "Manually done"
                } 
            } elseif ($reponse -match "^[Nn]$"){ 
                $ResultTable += [PSCustomObject]@{
                    "Section" = "Sharepoint"
                    "Name" = $item
                    "Setting" = "Not done manually"
                }
            }
        }  
    }
#----------------------------------------------------------[Teams actions]------------------------------------------------------

#----------------------------------------------------------[Verification]-------------------------------------------------------
    $ResultTable | Export-Excel -Path "C:\\TenantTool\\Script\\PoliciesConfigurationConfirmation_$($OrgName).xlsx" -Append

#----------------------------------------------------------[End of Script]------------------------------------------------------
Disconnect-Graph

if ($Debug -eq "Yes") {
    Stop-Transcript
}

exit
