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
param(
    [Parameter(Mandatory=$true)]
    [string]$DataFilePath
)
#----------------------------------------------------------[Debug]--------------------------------------------------------------
$Debug = "no"
$DebutPathFile = "C:\TenantTool\Log\$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss')_MemberUsrCreation.log"
if ($Debug -eq "Yes") {
    start-transcript -path $DebutPathFile
    Write-Debug "Debug mode activated"
    Write-host "DataPath : $DataFilePath"
}
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
        Connect-MgGraph -NoWelcome
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
    $UserSignIn = (Get-MgContext).Account
    write-host "Script to create Admin User and Group for $($OrgName)"
    Write-Host "User sign in actually : $($UserSignIn)"
    $response = Read-Host "You are connected on: $($OrgName). Do you want to continue ? (Y/N)"
} while ($ValidAnswers -notcontains $response)

#Answer Yes
if ($response -eq "O" -or $response -eq "Oui" -or $response -eq "Y" -or $response -eq "Yes") {
    try {
        #Groups importation (get content)
        $users = Import-excel -Path $DataFilePath
        $UsersNB = (($users).UserPrincipalName).count
        $empty = $null  # or any other default value you want to assign
        $users | Add-Member -MemberType NoteProperty -Name "Password" -Value $empty

        Write-host "Importation of $UsersNB users will start" -ForegroundColor Green
        Write-host "Please wait, the process can take some time" -ForegroundColor Green

        #Users User importation
        $ClientName = (Get-MgOrganization).DisplayName
        $ClientName = ($ClientName -replace '\s', '').ToLower()
        # Iterate through each row and update UserPrincipalName
        $exportcredentials = @()

        foreach ($user in $users) {
        # Add the PasswordProfile to the user
            $Password = New-RandomPassword -Length 12
            $user.Password = $Password
            $identity = [PSCustomObject]@{
                FirstName = $user.FirstName
                LastName = $user.LastName
                DisplayName = $user.Displayname
                UserPrincipalName = $user.UserPrincipalName
                TemporaryPassword = $Password
                MailNickName = $user.UserPrincipalName
            }
            $exportcredentials += $identity
            }

        #Export Connexion information + Apply password on account
        $TotalUsers = $users.Count
        $CurrentUser = 0
        
        foreach ($user in $users) {
            $CurrentUser++
        
            # Affichage de la barre de progression
            $ProgressPercent = ($CurrentUser / $TotalUsers) * 100
            Write-Progress -Activity "Creating Users..." -Status "Processing: $CurrentUser / $TotalUsers" -PercentComplete $ProgressPercent
        
            # Définition du mot de passe
            $PasswordProfile = @{
                Password = $user.Password
            }
        
            # Paramètres de l'utilisateur
            $commonParams = @{
                DisplayName       = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Surname           = $user.LastName
                GivenName         = $user.FirstName
                CompanyName       = $user.Company
                OfficeLocation    = $user.OfficeLocation
                UsageLocation     = $user.UsageLocation
                AccountEnabled    = $true
                PasswordProfile   = $PasswordProfile
                MailNickName      = ($user.UserPrincipalName -split '@')[0]  
            }
        
            if (-not [string]::IsNullOrEmpty($user.Department)) {
                $commonParams.Department = $user.Department
            }
        
            if (-not [string]::IsNullOrEmpty($user.JobTitle)) {
                $commonParams.JobTitle = $user.JobTitle
            }
        
            # Création de l'utilisateur
            New-MgUser @commonParams
        
            # Export des résultats
            $xlsresultpath = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path +"\$($ClientName)_UsersCreation.xlsx"
            $exportcredentials | Export-Excel -Path $xlsresultpath
        }
        
        # Fin de la barre de progression
        Write-Progress -Activity "Creating Users..." -Status "Completed" -Completed

            #Control Job-MailNickName
        $UserReportConfirmation = Get-MgBetaUser -All | Select-Object DisplayName, UserPrincipalName, CreatedDateTime | Where-Object { ($_.CreatedDateTime).Date -eq (Get-Date).Date } | Sort-Object CreatedDateTime
        $userReportConfirmation | Export-Excel -Path "C:\TenantTool\Script\MemberUserReportConfirmation.xlsx" -Append
        start-sleep -Milliseconds 500
    }
    catch {
        Write-Host "Error: $_"
    }
}

#Answer No
if ($response -eq 7) {

}

Disconnect-Graph

if ($Debug -eq "Yes") {
    Stop-Transcript
}

exit
