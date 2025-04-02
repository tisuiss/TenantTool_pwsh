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

#----------------------------------------------------------[Param]--------------------------------------------------------------
param(
    [Parameter(Mandatory=$true)]
    [string]$DataFilePath
)
#----------------------------------------------------------[Functions]----------------------------------------------------------
Function Create-DocumentLibrary()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $DocLibraryName
    )    
    Try {
    #Setup Credentials to connect
    #$Cred = Get-Credential
 
    #Set up the context
    $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL) 
    #$Ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.UserName,$Cred.Password)
 
    #Get All Lists from the web
    $Lists = $Ctx.Web.Lists
    $Ctx.Load($Lists)
    $Ctx.ExecuteQuery()
  
    #Check if Library name doesn't exists already and create document library
    if(!($Lists.Title -contains $DocLibraryName))
    { 
        #create document library in sharepoint online powershell
        $ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $ListInfo.Title = $DocLibraryName
        $ListInfo.TemplateType = 101 #Document Library
        $List = $Ctx.Web.Lists.Add($ListInfo)
        $List.Update()
        $Ctx.ExecuteQuery()
   
        write-host  -f Green "New Document Library has been created!"
    }
    else
    {
        Write-Host -f Yellow "List or Library '$DocLibraryName' already exist!"
    }
}
Catch {
    write-host -f Red "Error Creating Document Library!" $_.Exception.Message
}
}

#----------------------------------------------------------[Variables]----------------------------------------------------------
$Debug = "No"
$DebutPathFile = "C:\TenantTool\Log\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_CheckModuleNeeded.log"
#----------------------------------------------------------[Debug]--------------------------------------------------------------
if ($Debug -eq "Yes") {
    start-transcript -path $DebutPathFile
}

#Clear powershell window
clear-host

$ModuleNeeded = "Microsoft.Graph","Microsoft.Graph.Beta","ImportExcel","Microsoft.Online.SharePoint.PowerShell"
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

#----------------------------------------------------------[Connexions]---------------------------------------------------------
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

#Sharepoint
    Write-host "Connexion on the sharepoint of the tenant $OrgName" -ForegroundColor Yellow
    $OnMicrosoftDomain = ((Get-MgOrganization).VerifiedDomains | Where-Object {$_.IsInitial -eq $true}).name
    $OnMicrosoftDomain = $OnMicrosoftDomain -replace ".onmicrosoft.com",""
    Import-Module Microsoft.Online.SharePoint.PowerShell
    Connect-SPOService "https://$OnMicrosoftDomain-admin.sharepoint.com"

#----------------------------------------------------------[Automated Variables]------------------------------------------------
#----------------------------------------------------------[Execution]----------------------------------------------------------
#Data import
    $DataShp = Import-Excel -Path $DataFilePath

#Creation of Sharepoint Site
    $OwnerShpSite = (Get-MgBetaUser | Where-Object {$_.JobTitle -eq "admin"} | Select-Object -First 1).UserPrincipalName
    Foreach ($Data in $DataShp){
        $ShpName = $Data.Name
        $ShpTitle = $Data.Title

        $ShpName = "Test6"
        $ShpTitle = "Test6"
        $SecuGrpShpName = "Grs_Shp_$($ShpName)_RW"

        New-SPOSite -Url https://$OnMicrosoftDomain.sharepoint.com/sites/$ShpName `
                    -Owner $OwnerShpSite `
                    -StorageQuota "204800" `
                    -Title $ShpTitle `
                    -Template STS#3
 
        Set-SpoSite -Identity https://$OnMicrosoftDomain.sharepoint.com/sites/$ShpName `
                    -DisableSharingForNonOwners
        
        New-MgGroup -DisplayName $SecuGrpShpName `
                    -MailEnabled:$False `
                    -MailNickname 'group' `
                    -SecurityEnabled

        New-SPOSiteGroup    -Site hhttps://$OnMicrosoftDomain.sharepoint.com/sites/$ShpName `
                            -Group $SecuGrpShpName`
                            -PermissionLevels "Full Control"

    }
