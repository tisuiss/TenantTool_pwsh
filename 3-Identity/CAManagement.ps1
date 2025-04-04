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
  Author:         <MAUG>
  Creation Date:  <10.02.25>
#>

#----------------------------------------------------------[Param]--------------------------------------------------------------
#region Param
    param (
        [Parameter(Mandatory=$true)]
        [string] $DataFilePath
    )

#----------------------------------------------------------[Environment]--------------------------------------------------------
#Change UI Size
$host.UI.RawUI.BufferSize = New-Object Management.Automation.Host.Size(500, 3000)
#Clear powershell window
clear-host

#----------------------------------------------------------[Debug]--------------------------------------------------------------
$Debug = "Yes"
if ($Debug -eq "Yes") {
    start-transcript -path $DebutPathFile
    }

#----------------------------------------------------------[Module]-------------------------------------------------------------
$ModuleNeeded = "Microsoft.Graph","Microsoft.Graph.Beta","ImportExcel"
$DebutPathFile = "C:\TenantTool\Log\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_AdminUsrCreation.log"
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

#----------------------------------------------------------[Variables]----------------------------------------------------------
#region variables
    #xlsx File Path
    $XlsPath = $DataFilePath

    #Initialize CA Settings table
    $CAsSettings = @()

    #Initialize CA results table
    $CAFinalResults = @()

    #Configuration Settings
        #Users
        $DirectoryRoles = @("9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3","c4e39bd9-1100-46d3-8c65-fb160da0071f","7495fdc4-34c4-4d15-a289-98788ce399fd","b0f54661-2d74-4c50-afa3-1ec803f12efe","892c5842-a9a6-463a-8041-72aa08ca3cf6","158c047a-c907-4556-b7ef-446551a6b5f7","7698a772-787b-4ac8-901f-60d6b08affd2","17315797-102d-40b4-93e0-432062caca18","e6d1a23a-da11-4be4-9570-befc86d067a7","b1be1c3e-b65d-4f19-8427-f6fa0d97feb9","29232cdf-9323-42fd-ade2-1d097af3e4de","31392ffb-586c-42d1-9346-e59415a2cc4e","62e90394-69f5-4237-9190-012177145e10","f2ef992c-3afb-46b9-b7cf-a126ee74c451","ac434307-12b9-4fa1-a708-88bf58caabc1","fdd7a751-b60b-444a-984c-02652fe8fa1c","729827e3-9c14-49f7-bb1b-9608f156bbb8","3a2c62db-5318-420d-8d74-23affee5d9d5","4d6ac14f-3453-41d0-bef9-a3e0c569773a","d37c8bed-0711-4417-ba38-b4abe66ce4c2","966707d0-3269-4727-9be2-8c3a10f19b9d","7be44c8a-adaf-4e2a-84d6-ab2649e08a13","e8611ab8-c189-46e8-94e1-60213ab1f814","194ae4cb-b126-40b2-bd5b-6091b380977d","5f2222b1-57c3-48ba-8ad5-d4759f1fde6f","5d6b6bb7-de71-4623-b4af-96380a352509","f28a1f50-f6e7-4571-818b-6a12f2af6b6c","69091246-20e8-4a56-aa4d-066075b2a7a8","baf37b3a-610e-45da-9e62-d9d1e5e8914b","3d762c5a-1b6c-493f-843e-55a3b42923d4","fe930be7-5e62-47db-91af-98c3a49a38b1","11451d60-acb2-45eb-a7d6-43d0f0125c13")
        $Guest6Type = "internalGuest,b2bCollaborationGuest,b2bCollaborationMember,b2bDirectConnectUser,otherExternalUser,serviceProvider"
        $Guest5Type = "internalGuest,b2bCollaborationGuest,b2bCollaborationMember,b2bDirectConnectUser,otherExternalUser"
        $Guest1Type = "serviceProvider"

        $AccessControlEnforcementAuthStrenghMFA = @("windowsHelloForBusiness","fido2","x509CertificateMultiFactor","deviceBasedPush","temporaryAccessPassOneTime","temporaryAccessPassMultiUse","password,microsoftAuthenticatorPush","password,softwareOath","password,hardwareOath","password,sms","password,voice","federatedMultiFactor","microsoftAuthenticatorPush,federatedSingleFactor","softwareOath,federatedSingleFactor","hardwareOath,federatedSingleFactor","sms,federatedSingleFactor","voice,federatedSingleFactor")
        $AccessControlEnforcementAuthStrenghPasswordless = @("windowsHelloForBusiness","fido2","x509CertificateMultiFactor","deviceBasedPush")
        $AccessControlEnforcementAuthStrenghPhishingResistant = @("windowsHelloForBusiness","fido2","x509CertificateMultiFactor")
    
    #region empty .json file
    $EmptyJson = @'
    {
  "Conditions": {
    "Applications": {
      "ApplicationFilter": {
        "Mode": null,
        "Rule": null
      },
      "ExcludeApplications": [],
      "IncludeApplications": [],
      "IncludeAuthenticationContextClassReferences": [],
      "IncludeUserActions": []
    },
    "ClientAppTypes": null,
    "ClientApplications": {
      "ExcludeServicePrincipals": null,
      "IncludeServicePrincipals": null,
      "ServicePrincipalFilter": {
        "Mode": null,
        "Rule": null
      }
    },
    "Devices": {
      "DeviceFilter": {
        "Mode": null,
        "Rule": null
      }
    },
    "InsiderRiskLevels": null,
    "Locations": {
      "ExcludeLocations": [],
      "IncludeLocations": []
    },
    "Platforms": {
      "ExcludePlatforms": [],
      "IncludePlatforms": []
    },
    "ServicePrincipalRiskLevels": [],
    "SignInRiskLevels": [],
    "UserRiskLevels": [],
    "Users": {
      "ExcludeGroups": [],
      "ExcludeGuestsOrExternalUsers": {
        "ExternalTenants": {
          "MembershipKind": []
        },
        "GuestOrExternalUserTypes": []
      },
      "ExcludeRoles": [],
      "ExcludeUsers": [],
      "IncludeGroups": [],
      "IncludeGuestsOrExternalUsers": {
        "ExternalTenants": {
          "MembershipKind": []
        },
        "GuestOrExternalUserTypes": []
      },
      "IncludeRoles": [],
      "IncludeUsers": []
    }
  },
  "Description": null,
  "DisplayName": "EmptyCA",
  "GrantControls": {
    "AuthenticationStrength": {
      "AllowedCombinations": null,
      "CombinationConfigurations": null,
      "CreatedDateTime": null,
      "Description": null,
      "DisplayName": null,
      "Id": null,
      "ModifiedDateTime": null,
      "PolicyType": null,
      "RequirementsSatisfied": null
    },
    "BuiltInControls": null,
    "CustomAuthenticationFactors": [],
    "Operator": null,
    "TermsOfUse": []
  },
  "SessionControls": {
    "ApplicationEnforcedRestrictions": {
      "IsEnabled": null
    },
    "CloudAppSecurity": {
      "CloudAppSecurityType": null,
      "IsEnabled": null
    },
    "DisableResilienceDefaults": null,
    "PersistentBrowser": {
      "IsEnabled": null,
      "Mode": null
    },
    "SignInFrequency": {
      "AuthenticationType": null,
      "FrequencyInterval": null,
      "IsEnabled": null,
      "Type": null,
      "Value": null
    }
  },
  "State": "Disable",
  "TemplateId": null,
  "AdditionalProperties": {}
}
'@
    #endregion empty .json file


#endregion variables
#----------------------------------------------------------[Functions]----------------------------------------------------------
Function Connect-GrpahAPI {
    try {
        $IsGraphSign = Get-MgContext
        if ($IsGraphSign.count -gt "0") {
            Disconnect-MgGraph | Out-Null
            start-sleep 2
            Connect-MgGraph -Scopes 'Policy.Read.All','Application.Read.All','Policy.ReadWrite.ConditionalAccess' -NoWelcome

        }
        else {
            Connect-MgGraph -Scopes 'Policy.Read.All','Application.Read.All','Policy.ReadWrite.ConditionalAccess' -NoWelcome

        }
    } Catch {
        Write-Host "Error: $_"
    }
}

#----------------------------------------------------------[Connexions]---------------------------------------------------------
#region Connexion
    #Connect on Graph
    $YesAnswers = @("Yes","Y","O", "Oui")
    $response = ""
    
    do {
        Connect-GrpahAPI
        $OrgName = (Get-MgOrganization).DisplayName
        $UserSignIn = (Get-MgContext).Account
        Write-Host "`n"
        Write-Host "`n------------------------------`n" -ForegroundColor Cyan
        Write-Host -NoNewline "Connected on : " -ForegroundColor Magenta
        Write-Host "$OrgName" -ForegroundColor White
        Write-Host -NoNewline "Current User Signed In : " -ForegroundColor Magenta
        Write-Host "$UserSignIn" -ForegroundColor White
        Write-Host "`n------------------------------`n" -ForegroundColor Cyan
    
        # Demander la confirmation de l'utilisateur
        $response = Read-Host "You are connected on: $($OrgName). Do you want to continue? (Y/N)"
    } while ($YesAnswers -notcontains $response)
    

#endregion Connexion
#----------------------------------------------------------[Execution]----------------------------------------------------------
#region Excecution
    #region Import Settings from Excel
    write-host "Import Excel Data" -ForegroundColor Green
        #Import Excel file
        $ExcelPackage = Open-ExcelPackage -Path $XlsPath
        $WorksheetsName = (Get-ExcelSheetInfo -Path $XlsPath | Where-Object {$_.Name -notmatch "OLD" -and $_.Name -notmatch "Notes"}).Name

        Foreach ($Name in $WorksheetsName) {
            #Load Worksheet
            $Worksheet = ""
            $Worksheet = $ExcelPackage.Workbook.Worksheets[$Name]

            #Empty Variables
            $IncludeCellsValue = $null
            $ExcludeCellsValue = $null
            $UsersIncludeCellsValue = $null
            $GroupsIncludeCellsValue = $null
            $UsersExcludeCellsValue = $null
            $GroupsExcludeCellsValue = $null
            $Guests5IncludeCellsValue = $null
            $Guests6IncludeCellsValue = $null
            $GuestsServiceProviderInclude = $null
            $RolesExcludeCellsValue = $null
            $Guests5ExcludeCellsValue = $null
            $Guests6ExcludeCellsValue = $null
            $GuestsServiceProviderExclude = $null
            $UsersRolesAdminIncludeCellsValue = $null
            $UsersRolesAdminExcludeCellsValue = $null
            $IncludeGuest = $null
            $ExcludeGuest = $null
            $IncludeLocation = @()
            $ExcludeLocation = @()
            $IncludeCloudApps = @()
            $ExcludeCloudApps = @()
            $IncludeCloudAppList = @()
            $ExcludeCloudAppList = @()
            $IncludeUsersGroupsRoles = @()
            $ExcludeUsersGroupsRoles = @()
            $UsersActionsCellValue = @()

        #region Get Values
            #region User
                #region Include
                $IncludeCellsValue = @('E10', 'I10', 'M10','E11','I11','M11','E12','I12','M12') | ForEach-Object { 
                    $Worksheet.Cells[$_].Value 
                }
                $IncludeCellsValue = $IncludeCellsValue -split "[,\r\n]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
                
                Foreach ($IncludeCell in $IncludeCellsValue) {
                    if ($IncludeCell -eq "All users") {
                        $IncludeUsersGroupsRoles += [PSCustomObject]@{
                            DisplayName = "All users"
                            Id = "All"
                            Type = "user"
                        }
                    } elseif ($IncludeCell -match "Guest or external users") {
                        $IncludeUsersGroupsRoles += [PSCustomObject]@{
                            DisplayName = $IncludeCell
                            Id = $null
                            Type = "Guest"
                        }
                    } elseif ($IncludeCell -match "Directory roles") {
                        $IncludeUsersGroupsRoles += [PSCustomObject]@{
                            DisplayName = $IncludeCell
                            Id = $null
                            Type = "Role"
                        }
                    } else {
                        $UserId = Get-MgUser -Filter "UserPrincipalName eq '$IncludeCell'" -ErrorAction SilentlyContinue
                        $GroupId = Get-MgGroup -Filter "DisplayName eq '$IncludeCell'" -ErrorAction SilentlyContinue
                
                        if ($UserId) {
                            $IncludeUsersGroupsRoles += [PSCustomObject]@{
                                DisplayName = $UserId.Userprincipalname
                                Id = $UserId.Id
                                Type = "User"
                            }
                        } elseif ($GroupId) {
                            $IncludeUsersGroupsRoles += [PSCustomObject]@{
                                DisplayName = $GroupId.DisplayName
                                Id = $GroupId.Id
                                Type = "Group"
                            }
                        }
                    }
                }

                $UsersIncludeCellsValue = ($IncludeUsersGroupsRoles | Where-Object { $_ -ne $null -and $_.Type -eq "User" }).Id
                $GroupsIncludeCellsValue = ($IncludeUsersGroupsRoles | Where-Object { $_ -ne $null -and $_.Type -eq "Group" }).Id
                $RolesIncludeCellsValue = ($IncludeUsersGroupsRoles | Where-Object { $_ -ne $null -and $_.Type -eq "Role" }).DisplayName
                $GuestIncludeCellsValue = ($IncludeUsersGroupsRoles | Where-Object { $_ -ne $null -and $_.Type -eq "Guest" }).DisplayName

                if ($GuestIncludeCellsValue -eq "Guest or external users (6)") {
                    $IncludeGuest = $Guest6Type
                } elseif ($GuestIncludeCellsValue -eq "Guest or external users (5)") {
                    $IncludeGuest = $Guest5Type
                } elseif ($GuestIncludeCellsValue -eq "Guest or external users (1)") {
                    $IncludeGuest = $Guest1Type
                }

                #When no groups or users (ex guest only)
                #if ($UsersIncludeCellsValue.count -eq "0" -and $GroupsIncludeCellsValue.count -eq "0") {
                #    $UsersIncludeCellsValue = @("None")
                #}
                
                #endregion Include

                #region Exclude
            $ExcludeCellsValue = @('E13', 'I13', 'M13','E14','I14','M14','E15','I15','M15') | ForEach-Object {$Worksheet.Cells[$_].Value}
            $ExcludeCellsValue = $ExcludeCellsValue -split "[,\r\n]+" | ForEach-Object {$_.Trim()} | Where-Object { $_ -ne "" }
            Foreach ($ExcludeCell in $ExcludeCellsValue) {
                if ($ExcludeCell -eq "All users") {
                    $ExcludeUsersGroupsRoles += [PSCustomObject]@{
                        DisplayName = "All users"
                        Id = "All"
                        Type = "All Users"
                    }
                } elseif ($ExcludeCell -match "Guest or external users") {
                    $ExcludeUsersGroupsRoles += [PSCustomObject]@{
                        DisplayName = $ExcludeCell
                        Id = $null
                        Type = "Guest"
                    }
                } elseif ($ExcludeCell -match "Directory roles") {
                    $ExcludeUsersGroupsRoles += [PSCustomObject]@{
                        DisplayName = $ExcludeCell
                        Id = $null
                        Type = "Role"
                    }
                } else {
                    $UserId = Get-MgUser -Filter "UserPrincipalName eq '$ExcludeCell'" -ErrorAction SilentlyContinue
                    $GroupId = Get-MgGroup -Filter "DisplayName eq '$ExcludeCell'" -ErrorAction SilentlyContinue
                        if ($UserId.count -gt "0"){
                            $ExcludeUsersGroupsRoles += [PSCustomObject]@{
                                DisplayName = $UserId.Userprincipalname
                                Id = $UserId.Id
                                Type = "User"
                        }
                        } elseif ($GroupId.count -gt "0") {
                            $ExcludeUsersGroupsRoles += [PSCustomObject]@{
                                DisplayName = $GroupId.DisplayName
                                Id = $GroupId.Id
                                Type = "Group"
                            }
                    }
                }
            }
                    $UsersExcludeCellsValue = ($ExcludeUsersGroupsRoles | Where-Object {$_ -ne $null -and $_.Type -eq "User"}).Id
                    $GroupsExcludeCellsValue = ($ExcludeUsersGroupsRoles | Where-Object {$_ -ne $null -and $_.Type -eq "Group"}).Id
                    $RolesExcludeCellsValue = ($ExcludeUsersGroupsRoles | Where-Object {$_ -ne $null -and $_.Type -eq "Role"}).DisplayName
                    $GuestExcludeCellsValue = ($ExcludeUsersGroupsRoles | Where-Object {$_ -ne $null -and $_.Type -eq "Guest"}).DisplayName
                    if ($GuestExcludeCellsValue -eq "Guest or external users (6)") {
                            $ExcludeGuest = $Guest6Type
                        } elseif ($GuestExcludeCellsValue -eq "Guest or external users (5)") {
                            $ExcludeGuest = $Guest5Type
                        } elseif ($GuestExcludeCellsValue -eq "Guest or external users (1)") {
                            $ExcludeGuest = $Guest1Type
                    }
                #endregion Exclude
            #endregion User

            #region Location
                #region Include
            $IncludeLocationSheet = @('E18', 'I18', 'M18','E19','I19','M19','E20','I20','M20') | ForEach-Object {$Worksheet.Cells[$_].Value | Where-Object {$_ -ne $null}}
            if ($IncludeLocationSheet -eq "Any network or location") {
                $IncludeLocation = "All"
            } else {
                foreach ($IncludeLoc in $IncludeLocationSheet) {
                    $LocationObject = ""
                    $LocationObject = Get-MgIdentityConditionalAccessNamedLocation -All | Where-Object { $_.DisplayName -eq $IncludeLoc }
                    if ($LocationObject) {
                        $IncludeLocation += $LocationObject.Id
                    }
                }
            }
                #endregion Include

                #region Exclude
            $ExcludeLocationSheet = @('E21', 'I21', 'M21','E22','I22','M22','E23','I23','M23') | ForEach-Object {$Worksheet.Cells[$_].Value | Where-Object {$_ -ne $null}}
            if ($ExcludeLocationSheet -eq "Any network or location") {
                $ExcludeLocation = "All"
            } else {
                foreach ($ExcludeLoc in $ExcludeLocationSheet) {
                    $LocationObject = ""
                    $LocationObject = Get-MgIdentityConditionalAccessNamedLocation -All | Where-Object { $_.DisplayName -eq $ExcludeLoc }
                    if ($LocationObject) {
                        $ExcludeLocation += $LocationObject.Id
                    }
                }
            } 
                #endregion Exclude
            #endregion Location

            #region Cloud Apps
                #region Include
            $SettingsType = $Worksheet.Cells['B16'].Value
            $IncludeCloudApps = @('E16', 'I16', 'M16') | ForEach-Object {$Worksheet.Cells[$_].Value | Where-Object {$_ -ne $null}}
            if ($SettingsType -eq "Cloud apps") {
                foreach ($IncludeCloudApp in $IncludeCloudApps) {
                    if ($IncludeCloudApp -eq "All cloud apps") {
                        $IncludeCloudAppList += "all"
                    } 
                    elseif ($IncludeCloudApp -match "Microsoft Admin Portals") {
                        $IncludeCloudAppList += "MicrosoftAdminPortals"
                    } 
                    elseif ($IncludeCloudApp -match "Register security information" -or $IncludeCloudApp -match "Register or join devices") {
                        #Do nothing
                    }
                    else {
                        $CloudAppId = Get-MgServicePrincipal -All | Where-Object { $_.DisplayName -eq $IncludeCloudApp }
                        if ($CloudAppId) {
                            $IncludeCloudAppList += $CloudAppId.AppId
                        }
                    }
                }
            } elseif ($SettingsType -eq "User actions") {
                if ($IncludeCloudApps -eq "Register security information"){
                    $UsersActionsCellValue += "urn:user:registersecurityinfo"
                }
                if ($IncludeCloudApps -eq "Register or join devices"){
                    $UsersActionsCellValue += "urn:user:registerdevice"
                }
            }
                #endregion Include

                #region Exclude
            $ExcludeCloudApps = @('E17', 'I17', 'M17') | ForEach-Object {$Worksheet.Cells[$_].Value | Where-Object {$_ -ne $null}}
            Foreach ($ExCloudApp in $ExcludeCloudApps) {
                if ($ExCloudApp -eq "All cloud apps") {
                    $ExcludeCloudAppList += "all"
                }
                else {
                    $CloudAppId = Get-MgServicePrincipal -All | Where-Object {$_.DisplayName -eq $ExCloudApp}
                    $ExcludeCloudAppList += $CloudAppId.AppId
                }
            }
                #endregion Exclude
            #endregion Cloud Apps

        #endregion Get Values

        #region Create Object
        write-host "Add to object : $($Name)" -ForegroundColor Yellow
            $CAsSettings += [PSCustomObject] @{
                CASheetName = $Name
                CAJobAction = $Worksheet.Cells['O6'].Value
                DisplayName = $Worksheet.Cells['B2'].Value
                State = "Disabled"
                #User
                IncludeUsers = $UsersIncludeCellsValue -join ","
                IncludeGroups = $GroupsIncludeCellsValue -join ","
                IncludeRoles = if (($RolesIncludeCellsValue).count -gt "0") {$UsersRolesAdminIncludeCellsValue} else {""}
                IncludeGuests = $IncludeGuest
                ExcludeUsers = $UsersExcludeCellsValue -join ","
                ExcludeGroups = $GroupsExcludeCellsValue -join ","
                ExcludeRoles = $RolesExcludeCellsValue
                ExcludeGuests = $ExcludeGuest

                #Cloud Apps
                IncludeCloudApps = $IncludeCloudAppList -join ","
                ExcludeCloudApps = $ExcludeCloudAppList -join ","
                UsersActions = $UsersActionsCellValue -join ","
                #Location
                IncludeLocations = $IncludeLocation -join ","
                ExcludeLocations = $ExcludeLocation -join ","
                #Access Control Enforcement
                AccessControlEnforcementBlock = $Worksheet.Cells['B28'].Value
                AccessControlEnforcementGrant = $Worksheet.Cells['B29'].Value
                    AccessControlEnforcementMFA = $Worksheet.Cells['D25'].Value
                    AccessControlEnforcementAuthStrengh = $Worksheet.Cells['D26'].Value
                        AccessControlEnforcementAuthStrenghType = $Worksheet.Cells['L26'].Value
                    AccessControlEnforcementAsCompliant = $Worksheet.Cells['D27'].Value
                    AccessControlEnforcementRequireEntraHybridjoinedDevice = $Worksheet.Cells['D28'].Value
                    AccessControlEnforcementRequireApprovedClientApp = $Worksheet.Cells['D29'].Value
                    AccessControlEnforcementRequireAppProtection = $Worksheet.Cells['D30'].Value
                    AccessControlEnforcementRequireAllControls = $Worksheet.Cells['D32'].Value
                    AccessControlEnforcementRequireOneControl = $Worksheet.Cells['D33'].Value
                #Device Platform
                DeviceIncludeAllPlatforms = $Worksheet.Cells['U10'].Value
                DeviceIncludeAndroid = $Worksheet.Cells['W10'].Value
                DeviceIncludeIOS = $Worksheet.Cells['Y10'].Value
                DeviceIncludeWindows = $Worksheet.Cells['AA10'].Value
                DeviceIncludeMacOS = $Worksheet.Cells['AC10'].Value
                DeviceIncludeLinux = $Worksheet.Cells['AE10'].Value
                DeviceExcludeAndroid = $Worksheet.Cells['W11'].Value
                DeviceExcludeIOS = $Worksheet.Cells['Y11'].Value
                DeviceExcludeWindows = $Worksheet.Cells['AA11'].Value
                DeviceExcludeMacOS = $Worksheet.Cells['AC11'].Value
                DeviceExcludeLinux = $Worksheet.Cells['AE11'].Value
                #Client Apps
                ClientBrowser = $Worksheet.Cells['T12'].Value
                ClientMobileAppsAndDesktopClients = $Worksheet.Cells['T13'].Value
                ClientExchangeActiveSync = $Worksheet.Cells['T14'].Value
                ClientOther = $Worksheet.Cells['T15'].Value
                #Filter for device
                DeviceFilterType = $Worksheet.Cells['T16'].Value
                DeviceFilterRegex = $Worksheet.Cells['U16'].Value
                #Authentification Flow
                AuthFlowDeviceCode = $Worksheet.Cells['U18'].Value
                AuthFlowTransfert = $Worksheet.Cells['Y18'].Value
                #User risk
                UserRiskHigh = $Worksheet.Cells['U20'].Value
                UserRiskAverage = $Worksheet.Cells['W20'].Value
                UserRiskLow = $Worksheet.Cells['Y20'].Value
                #Access Control Session
                AccessControlSessionUseAppEnforcedRestrictions = $Worksheet.Cells['T23'].Value
                AccessControlSessionUseConditionalAccessAppControl = $Worksheet.Cells['T24'].Value
                    AccessControlSessionUseConditionalAccessAppControlMonitorOnly = $Worksheet.Cells['U25'].Value
                    AccessControlSessionUseConditionalAccessAppControlBlockDownload = $Worksheet.Cells['AA25'].Value
                AccessControlSessionSignInFrequency = $Worksheet.Cells['T26'].Value
                    AccessControlSessionPeriodicReauthentification = $Worksheet.Cells['U27'].Value
                        AccessControlSessionPeriodicReauthentificationNumber = $Worksheet.Cells['AA27'].Value
                        AccessControlSessionPeriodicReauthentificationFrequency = $Worksheet.Cells['AB27'].Value
                    AccessControlSessionPeriodicEVeryTime = $Worksheet.Cells['AC27'].Value
                AccessControlSessionPersistentbrowsersession = $Worksheet.Cells['T28'].Value	
                    AccessControlSessionPersistentbrowsersessionSettings = $Worksheet.Cells['AB28'].Value
                AccessControlSessionUseCustomizeContinuousAccessEvaluation = $Worksheet.Cells['T29'].Value
                    AccessControlSessionUseCustomizeContinuousAccessEvaluationDisable = $Worksheet.Cells['U30'].Value
                    AccessControlSessionUseCustomizeContinuousAccessEvaluationStrictlyEnforceLocationPolicies = $Worksheet.Cells['AA30'].Value
                AccessControlSessionDisableResilienceDefaults = $Worksheet.Cells['T31'].Value
                AccessControlSessionGlobalSecureAccessSecurityProfile = $Worksheet.Cells['T32'].Value
                    AccessControlSessionGlobalSecureAccessSecurityProfileName = $Worksheet.Cells['U33'].Value
            }
        #endregion Create Object
        #Convertto-Json -InputObject $CAsSettings -Depth 10 | Out-File -FilePath "C:\temp\$($Name).json" -Force
        }
        
        write-Host "Object with ALL CA settings created" -ForegroundColor Green
        #Close Excel File
        $ExcelPackage.Dispose()
    #endregion Import Settings from Excel

    Function New-CA {
        #Empty Variable
        $policyJson = $null
        $policyJson = $EmptyJson | ConvertFrom-Json
        #DisplayName
        $policyJson.DisplayName = $CA.DisplayName
        #State
        $policyJson.State = $CA.State
        #User
        #IncludeUsers
        if ($null -eq $CA.IncludeUsers -or $CA.IncludeUsers -ne "") {
            Foreach ($IncludeUser in $CA.IncludeUsers.Split(",")) {
                $policyJson.Conditions.Users.IncludeUsers += $IncludeUser
            }
        }
        #ExcludeUsers
        if ($null -eq $CA.ExcludeGroups -or $CA.ExcludeUsers -ne "") {
            Foreach ($ExcludeUser in $CA.ExcludeUsers.Split(",")) {
                $policyJson.Conditions.Users.ExcludeUsers += $ExcludeUser
            }
        }
        #IncludeGroups
        if ($null -eq $CA.IncludeGroups -or $CA.IncludeGroups -ne "") {
            Foreach ($IncludeGroup in $CA.IncludeGroups.Split(",")) {
                $policyJson.Conditions.Users.IncludeGroups += $IncludeGroup
            }
        }
        #ExcludeGroups
        if ($null -eq $CA.ExcludeGroups -or $CA.ExcludeGroups -ne "") {
            Foreach ($ExcludeGroup in $CA.ExcludeGroups.Split(",")) {
                $policyJson.Conditions.Users.ExcludeGroups += $ExcludeGroup
            }
        }
        #IncludeRoles
        if ($null -ne $CA.IncludeRoles -and $CA.IncludeRoles -ne "") {
            $policyJson.Conditions.Users.IncludeRoles += $DirectoryRoles
        }
        #ExcludeRoles
        if ($null -ne $CA.ExcludeRoles){
            $policyJson.Conditions.Users.ExcludeRoles += $DirectoryRoles
        }
        #IncludeGuests
        if ($CA.IncludeGuests -ne ""){
            if ($null -ne $CA.IncludeGuests) {
                $policyJson.Conditions.Users.IncludeGuestsOrExternalUsers.ExternalTenants.MembershipKind = "all"
                $policyJson.Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes = $CA.IncludeGuests
            }
        }
        #ExcludeGuests
        if ($CA.ExcludeGuests -ne ""){
            if ($null -ne $CA.ExcludeGuests) {
                $policyJson.Conditions.Users.ExcludeGuestsOrExternalUsers.ExternalTenants.MembershipKind = "all"
                $policyJson.Conditions.Users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes += $CA.ExcludeGuests
            }
        }
        #Cloud Apps
            #Include
            if ($null -ne $CA.IncludeCloudApps -and $CA.IncludeCloudApps -ne "") {
                Foreach ($IncludeCloudAppSplit in $CA.IncludeCloudApps.Split(",")) {
                    $policyJson.Conditions.Applications.IncludeApplications += $IncludeCloudAppSplit
                }
            }
            #Exclude
            $policyJson.Conditions.Applications.ExcludeApplications = @()
            if ($null -eq $CA.ExcludeCloudApps -or $CA.ExcludeCloudApps -ne "") {
                Foreach ($ExcludeCloudApp in $CA.ExcludeCloudApps.Split(",")) {
                    $policyJson.Conditions.Applications.ExcludeApplications += $ExcludeCloudApp
                }
            }
        #User actions
            if ($null -ne $CA.UsersActions -and $CA.UsersActions -ne "") {
                Foreach ($UserAction in $CA.UsersActions.Split(",")) {
                    $policyJson.Conditions.Applications.IncludeUserActions += $UserAction
                }
            }
        #Location
            #Include
            if ($CA.IncludeLocations -ne "") {
                foreach ($IncludeSplit in $CA.IncludeLocations.Split(",")) {
                    $policyJson.Conditions.Locations.IncludeLocations += $IncludeSplit
                }
            }

            #Exclude
            if ($CA.ExcludeLocations -ne "") {
                Foreach ($ExcludeLocation in $CA.ExcludeLocations.Split(",")) {
                    $policyJson.Conditions.Locations.ExcludeLocations = @()
                    $policyJson.Conditions.Locations.ExcludeLocations += $ExcludeLocation
                }
            } else {
                $policyJson.Conditions.Locations.ExcludeLocations = @()
            }
        #Access Control Enforcement
        #Block Access
        if ($CA.AccessControlEnforcementBlock -eq "True") {
            # Block access
            $policyJson.GrantControls.BuiltInControls += @("block")
            $policyJson.GrantControls.Operator = "OR"
        }
        #Grant Access
        if ($CA.AccessControlEnforcementGrant -eq "True") {
            # Grant access
            #$policyJson.GrantControls.BuiltInControls += @("Grant")

            if ($CA.AccessControlEnforcementMFA -eq "True") {
                # Require multi-factor authentication
                $policyJson.GrantControls.BuiltInControls += @("mfa")
                # Require authentication strengh (+ type of strengh)
                if ($CA.AccessControlEnforcementAuthStrengh -eq "True") {
                    $TodayAccessControlEnforcement = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
                    if ($CA.AccessControlEnforcementAuthStrenghType -eq "Multifactor authentication") {
                        $policyJson.GrantControls.AuthenticationStrength.AllowedCombinations = $AccessControlEnforcementAuthStrenghMFA
                        $policyJson.GrantControls.AuthenticationStrength.CreatedDateTime = $TodayAccessControlEnforcement
                        $policyJson.GrantControls.AuthenticationStrength.Description = "Combinations of methods that satisfy strong authentication, such as a password + SMS"
                        $policyJson.GrantControls.AuthenticationStrength.DisplayName = "Multifactor authentication"
                        $policyJson.GrantControls.AuthenticationStrength.Id = "00000000-0000-0000-0000-000000000002"
                        $policyJson.GrantControls.AuthenticationStrength.ModifiedDateTime = $TodayAccessControlEnforcement
                        $policyJson.GrantControls.AuthenticationStrength.PolicyType = "builtIn"
                        $policyJson.GrantControls.AuthenticationStrength.RequirementsSatisfied = "mfa"
                    }
                    if ($CA.AccessControlEnforcementAuthStrenghType -eq "Passwordless MFA") {
                        $policyJson.GrantControls.AuthenticationStrength.AllowedCombinations = $AccessControlEnforcementAuthStrenghPasswordless
                        $policyJson.GrantControls.AuthenticationStrength.CreatedDateTime = $TodayAccessControlEnforcement
                        $policyJson.GrantControls.AuthenticationStrength.Description = "Passwordless methods that satisfy strong authentication, such as Passwordless sign-in with the Microsoft Authenticator"
                        $policyJson.GrantControls.AuthenticationStrength.DisplayName = "Passwordless MFA"
                        $policyJson.GrantControls.AuthenticationStrength.Id = "00000000-0000-0000-0000-000000000003"
                        $policyJson.GrantControls.AuthenticationStrength.ModifiedDateTime = $TodayAccessControlEnforcement
                        $policyJson.GrantControls.AuthenticationStrength.PolicyType = "builtIn"
                        $policyJson.GrantControls.AuthenticationStrength.RequirementsSatisfied = "mfa"
                    } 
                    if ($CA.AccessControlEnforcementAuthStrenghType -eq "Phishing-resistant MFA") {
                        $policyJson.GrantControls.AuthenticationStrength.AllowedCombinations = $AccessControlEnforcementAuthStrenghPhishingResistant
                        $policyJson.GrantControls.AuthenticationStrength.CreatedDateTime = $TodayAccessControlEnforcement
                        $policyJson.GrantControls.AuthenticationStrength.Description = "Phishing-resistant, Passwordless methods for the strongest authentication, such as a FIDO2 security key"
                        $policyJson.GrantControls.AuthenticationStrength.DisplayName = "Phishing-resistant MFA"
                        $policyJson.GrantControls.AuthenticationStrength.Id = "00000000-0000-0000-0000-000000000004"
                        $policyJson.GrantControls.AuthenticationStrength.ModifiedDateTime = $TodayAccessControlEnforcement
                        $policyJson.GrantControls.AuthenticationStrength.PolicyType = "builtIn"
                        $policyJson.GrantControls.AuthenticationStrength.RequirementsSatisfied = "mfa"
                    }
                }
            }

            # Require device to be marked as compliant											
            if ($CA.AccessControlEnforcementAsCompliant -eq "True") {
                $BuiltInControlsCompliant = @("compliantDevice")
                $policyJson.GrantControls.BuiltInControls += $BuiltInControlsCompliant
            }
            #Require Microsoft Entra hybrid joined device											
            if ($CA.AccessControlEnforcementRequireEntraHybridjoinedDevice -eq "True") {
                $BuiltInControlsDomainJoined = @("domainJoinedDevice")
                $policyJson.GrantControls.BuiltInControls += $BuiltInControlsDomainJoined
            }
            #Require approved clients app											
            if ($CA.AccessControlEnforcementRequireApprovedClientApp -eq "True") {
                $policyJson.GrantControls.BuiltInControls += @("approvedApplication")
            }
            #Require app protection policy
            if ($CA.AccessControlEnforcementRequireAppProtection -eq "True") {
                $policyJson.GrantControls.BuiltInControls += @("compliantApplication")
            }
        }
        #Require all the following controls
        if ($CA.AccessControlEnforcementRequireAllControls -eq "True") {
            $policyJson.GrantControls.Operator = "AND"
        }
        #Require at least one of the following controls
        if ($CA.AccessControlEnforcementRequireOneControl -eq "True") {
            $policyJson.GrantControls.Operator = "OR"
        }

        #Device Platform
        #Include
        $policyJson.Conditions.Platforms.IncludePlatforms = @()
        if ($CA.DeviceIncludeAllPlatforms -eq "True") {
            $policyJson.Conditions.Platforms.IncludePlatforms += "All"
        } else {
            if ($CA.DeviceIncludeAndroid -eq "True") {
                $policyJson.Conditions.Platforms.IncludePlatforms += "android"
            }
            if ($CA.DeviceIncludeIOS -eq "True") {
                $policyJson.Conditions.Platforms.IncludePlatforms += "iOS"
            }
            if ($CA.DeviceIncludeWindows -eq "True") {
                $policyJson.Conditions.Platforms.IncludePlatforms += "windows"
            }
            if ($CA.DeviceIncludeMacOS -eq "True") {
                $policyJson.Conditions.Platforms.IncludePlatforms += "macOS"
            }
            if ($CA.DeviceIncludeLinux -eq "True") {
                $policyJson.Conditions.Platforms.IncludePlatforms += "linux"
            }
        }
        #Exclude
        if ($CA.DeviceExcludeAndroid -eq "True") {
            $policyJson.Conditions.Platforms.ExcludePlatforms += "android"
        }
        if ($CA.DeviceExcludeIOS -eq "True") {
            $policyJson.Conditions.Platforms.ExcludePlatforms += "iOS"
        }
        if ($CA.DeviceExcludeWindows -eq "True") {
            $policyJson.Conditions.Platforms.ExcludePlatforms += "windows"
        }
        if ($CA.DeviceExcludeMacOS -eq "True") {
            $policyJson.Conditions.Platforms.ExcludePlatforms += "macOS"
        }
        if ($CA.DeviceExcludeLinux -eq "True") {
            $policyJson.Conditions.Platforms.ExcludePlatforms += "linux"
        }
        #Client Apps
        if ($CA.ClientBrowser -eq "True") {
            $policyJson.Conditions.ClientAppTypes += @("browser")
        }
        if ($CA.ClientMobileAppsAndDesktopClients -eq "True") {
            $policyJson.Conditions.ClientAppTypes += @("mobileAppsAndDesktopClients")
        }
        if ($CA.ClientExchangeActiveSync -eq "True") {
            $policyJson.Conditions.ClientAppTypes += @("exchangeActiveSync")
        }
        if ($CA.ClientOther -eq "True") {
            $policyJson.Conditions.ClientAppTypes += @("other")
        }
        #Filter for device
        #Include
        if ($CA.DeviceFilterType -eq "Include:"){
            $policyJson.Conditions.Devices.DeviceFilter.Mode = "include"
            $policyJson.Conditions.Devices.DeviceFilter.Rule = $CA.DeviceFilterRegex
        }
        #Exclude
        if ($CA.DeviceFilterType -eq "Exclude:"){
            $policyJson.Conditions.Devices.DeviceFilter.Mode = "exclude"
            $policyJson.Conditions.Devices.DeviceFilter.Rule = $CA.DeviceFilterRegex
        }
        #Authentification Flow
        if ($CA.AuthFlowDeviceCode -eq "True") {
            $policyJson.Conditions.ClientApplications.ServicePrincipalFilter.Rule = "deviceCode"
        }
        if ($CA.AuthFlowTransfert -eq "True") {
            $policyJson.Conditions.ClientApplications.ServicePrincipalFilter.Rule = "transfert"
        }
        #User risk
        if ($CA.UserRiskHigh -eq "True") {
            $policyJson.Conditions.UserRiskLevels += "high"
        }
        if ($CA.UserRiskAverage -eq "True") {
            $policyJson.Conditions.UserRiskLevels += "medium"
        }
        if ($CA.UserRiskLow -eq "True") {
            $policyJson.Conditions.UserRiskLevels += "low"
        }
        #Access Control Session
        #Use app enforced restrictions
        if ($CA.AccessControlSessionUseAppEnforcedRestrictions -eq "True") {
            $policyJson.SessionControls.ApplicationEnforcedRestrictions.IsEnabled = $true
        }
        #Use conditional access app control
        if ($CA.AccessControlSessionUseConditionalAccessAppControl -eq "True") {
            $policyJson.SessionControls.CloudAppSecurity.IsEnabled = $true
        }
        #Use conditional access app control monitor only
        if ($CA.AccessControlSessionUseConditionalAccessAppControlMonitorOnly -eq "True") {
            $policyJson.SessionControls.CloudAppSecurity.CloudAppSecurityType = "monitorOnly"
        }
        #Use conditional access app control block download
        if ($CA.AccessControlSessionUseConditionalAccessAppControlBlockDownload -eq "True") {
            $policyJson.SessionControls.CloudAppSecurity.CloudAppSecurityType = "blockDownload"
        }
        #SignInFrequency
        if ($CA.AccessControlSessionSignInFrequency -eq "True") {
            $policyJson.SessionControls.SignInFrequency.IsEnabled = $true
            #Periodic reauthentification
            if ($CA.AccessControlSessionPeriodicReauthentification -eq "True") {
                $policyJson.SessionControls.SignInFrequency.AuthenticationType = "primaryAndSecondaryAuthentication"
                $policyJson.SessionControls.SignInFrequency.FrequencyInterval = "timeBased"
                $policyJson.SessionControls.SignInFrequency.Type = $CA.AccessControlSessionPeriodicReauthentificationFrequency
                $policyJson.SessionControls.SignInFrequency.Value = [int]$CA.AccessControlSessionPeriodicReauthentificationNumber
            }
            #Periodic every time
            if ($CA.AccessControlSessionPeriodicEVeryTime -eq "True") {
                $policyJson.SessionControls.SignInFrequency.AuthenticationType = "primaryAndSecondaryAuthentication"
                $policyJson.SessionControls.SignInFrequency.FrequencyInterval = "everyTime"
            }
        }
        #Persistent browser session
        if ($CA.AccessControlSessionPersistentbrowsersession -eq "True") {
            $policyJson.SessionControls.PersistentBrowser.IsEnabled = $true
            #Always
            if ($CA.AccessControlSessionPersistentbrowsersessionSettings -eq "Always persistent") {
                $policyJson.SessionControls.PersistentBrowser.Mode = "Always"
            }
            #Never persistant
            if ($CA.AccessControlSessionPersistentbrowsersessionSettings -eq "Never persistent") {
                $policyJson.SessionControls.PersistentBrowser.Mode = "Never"
            }
        }
        #Use customize continuous access evaluation
        if ($CA.AccessControlSessionUseCustomizeContinuousAccessEvaluation -eq "True") {
            #Disable
            if ($CA.AccessControlSessionUseCustomizeContinuousAccessEvaluationDisable -eq "True") {
            }
            #Strictly enforce location policies
            if ($CA.AccessControlSessionUseCustomizeContinuousAccessEvaluationStrictlyEnforceLocationPolicies -eq "True") {
            }
        }
        #Disable resilience defaults
        if ($CA.AccessControlSessionDisableResilienceDefaults -eq "True") {
            $policyJson.SessionControls.DisableResilienceDefaults = "true"
        }
        #Global secure access security profile
        if ($CA.AccessControlSessionGlobalSecureAccessSecurityProfile -eq "True") {
            read-host "This policy must have entra ID P2, script not yet ready with this options, please contact maug and create your CA manually"
            break
        }

        #Debug
        #$policyJson | ConvertTo-Json -Depth 10 | Out-File "C:\temp\$($CA.CASheetName).json"
        #Import CA policy on tenant
        $policyJsonString = $null
        # Convert the custom object to JSON with a depth of 10
        $policyJsonString = $policyJson | ConvertTo-Json -Depth 10
        # Create the Conditional Access policy using the Microsoft Graph API
        try {
            New-MgIdentityConditionalAccessPolicy -Body $policyJsonString
        } catch {
            Write-Host "try to create $($CA.DisplayName) -> failed" -ForegroundColor Red
            Read-Host "An error occurred while creating the Conditional Access policy: $_" -ForegroundColor Red
        }
    }
    Function Set-CA {
        $CAtoModify = Get-MgIdentityConditionalAccessPolicy -All | Where-Object {$_.DisplayName -eq $CA.DisplayName}
        Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $CAtoModify.Id -Confirm:$false -ErrorAction SilentlyContinue

        New-CA
    }
        #region Import CA
        Foreach ($CA in $CAsSettings) {
        #Region New CA
            if ($CA.CAJobAction -eq "New") {
                Write-Host "      Creation of : $($CA.DisplayName)" -ForegroundColor Yellow
                $CACreation = New-CA

                $CAFinalResults += [PSCustomObject]@{
                    CAName = $CA.DisplayName
                    CAJobAction = $CA.CAJobAction
                    CAIdCreated = $CACreation.Id
                    CACreatedDateTime = $CACreation.CreatedDateTime
                }
            }
        #endregion New CA

        #region Modify CA
        if ($CA.CAJobAction -eq "Update") {
            Write-Host "      Modification of : $($CA.DisplayName)" -ForegroundColor Green
            Set-CA
        }
        #endregion Modify CA

        #region None
        if ($CA.CAJobAction -eq "None") {
            Write-Host "      No action for : $($CA.DisplayName)" -ForegroundColor Blue
        }
        #endregion None
        start-sleep -seconds 2
    }
    #endregion Import CA

    #region Result
    $ResultsCA = Get-MgIdentityConditionalAccessPolicy -All | Select-Object DisplayName,State,CreatedDateTime,ModifiedDateTime | Sort-Object ModifiedDateTime,CreatedDateTime -Descending
    $ResultsCA | Export-Excel -Path 'C:\TenantTool\Check\CAManagementResult.xlsx' -AutoSize -WorksheetName "CAResult" -TableName "CAResult" -BoldTopRow -FreezeTopRow -AutoFilter -AutoSize -ShowFilterButton -Title "Conditional Access Result" -TitleBold -TitleSize 16 -TitleColor DarkBlue -HeaderColor DarkBlue -HeaderSize 12 -HeaderBold -HeaderTextColor White -TableStyle Medium9
    #endregion Result

    if ($Debug -eq "Yes") {
        stop-transcript
        }

    #region Disconnect 
        try {
            Disconnect-MgGraph | Out-Null
            Write-Host "Disconnected from the Microsoft Graph" -ForegroundColor Green
        }
        catch {
            Write-Host "try to create $($CA.DisplayName) -> failed" -ForegroundColor Red
            Read-Host "An error occurred while disconnecting from the Microsoft Graph: $_" -ForegroundColor Red
            exit
        }
    #endregion Disconnect
#endregion Excecution

