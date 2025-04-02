
####### EOP ##########
Connect-ExchangeOnline

$DomaineToProtect = (Get-AcceptedDomain).DomainName

Enable-OrganizationCustomization

Write-host "wait time during application of Enable-OrganizationCustomization, please wait (1min), if errors occurs after please relaunch script after some times (2-3h)"
Start-sleep 60

    #Quarantine
        #Create New Quarantine Policy with Limited Access for End Users
        $LimitedAccess = New-QuarantinePermissions  -PermissionToAllowSender $true `
                                                    -PermissionToDelete $true `
                                                    -PermissionToPreview $true `
                                                    -PermissionToRequestRelease $false `
                                                    -PermissionToBlockSender $true `
                                                    -PermissionToDownload $false `
                                                    -PermissionToRelease $true

        $QuarantinePolicyName = "Require_Request_To_Release"

        New-QuarantinePolicy    -Name $QuarantinePolicyName `
                                -EndUserQuarantinePermissions $LimitedAccess `
                                -QuarantineRetentionDays "30" `
                                -EndUserSpamNotificationFrequency "04:00:00" `
                                -ESNEnabled $true
            
        $EOPQuarantineVerif = Get-QuarantinePolicy -Identity $QuarantinePolicyName

        if ($EOPQuarantineVerif.count -gt "0") {
            $ResultTable += [PSCustomObject]@{
                "Panel" = "EOP"
                "Name" = $EOPQuarantineVerif.Identity
                "Setting" = "Quarantine policy created"
            }
        }

        #Set Frequency for Global Quarantine Policy 4h
        Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy | Set-QuarantinePolicy -EndUserSpamNotificationFrequency 04:00:00

        $GlobalQuarantineNotifFrequVerif = (Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy).EndUserSpamNotificationFrequency
        $ResultTable += [PSCustomObject]@{
            "Panel" = "EOP"
            "Name" = "Quarantine Notif Frequency"
            "Setting" = $GlobalQuarantineNotifFrequVerif
        }

    #Malware
        $MalwarePolicyName = "Anti_Malware_Tenant"    

        New-MalwareFilterPolicy -Name "Anti_Malware_Tenant_Policy" `
                                -EnableInternalSenderAdminNotifications $false `
                                -QuarantineTag "Require_Request_To_Release"

        New-MalwareFilterRule   -Name $MalwarePolicyName `
                                -MalwareFilterPolicy "Anti_Malware_Tenant_Policy" `
                                -RecipientDomainIs $DomaineToProtect

        $EOPMalwareVerif = Get-MalwareFilterRule -Identity $MalwarePolicyName

        if ($EOPMalwareVerif.count -gt "0") {
            $ResultTable += [PSCustomObject]@{
                "Panel" = "EOP"
                "Name" = $EOPMalwareVerif.Identity
                "Setting" = "Malware policy created"
            }
        }

    #Spam Inbound
        $EOPSpamInName = "Anti_Spam_Inbound_Tenant"

        New-HostedContentFilterPolicy   -Name "Anti_Spam_Inbound_Tenant_Policy" `
                                        -SpamAction Quarantine `
                                        -SpamQuarantineTag "Require_Request_To_Release" `
                                        -HighConfidenceSpamAction Quarantine `
                                        -HighConfidenceSpamQuarantineTag "Require_Request_To_Release" `
                                        -PhishSpamAction Quarantine `
                                        -PhishQuarantineTag "Require_Request_To_Release" `
                                        -HighConfidencePhishAction Quarantine `
                                        -HighConfidencePhishQuarantineTag "Require_Request_To_Release" `
                                        -BulkSpamAction Quarantine `
                                        -BulkQuarantineTag "Require_Request_To_Release" `
                                        -QuarantineRetentionPeriod "30" `
                                        -BulkThreshold "6" `
                                        -PhishZapEnabled $true `
                                        -SpamZapEnabled $true

        New-HostedContentFilterRule     -Name $EOPSpamIEOPSpamInNamenVerif `
                                        -HostedContentFilterPolicy "Anti_Spam_Inbound_Tenant_Policy" `
                                        -RecipientDomainIs $DomaineToProtect

        $EOPSpamInVerif = Get-HostedContentFilterRule -Identity $EOPSpamInName

        if ($EOPSpamInVerif.count -gt "0") {
            $ResultTable += [PSCustomObject]@{
                "Panel" = "EOP"
                "Name" = $EOPSpamInVerif.Identity
                "Setting" = "Spam IN policy created"
            }
        }

    #Spam Outbound
        $EOPSpamOutName = "Anti_Spam_Outbound_Tenant"
        New-HostedOutboundSpamFilterPolicy  -Name "Anti_Spam_Outbound_Tenant_Policy" `
                                            -RecipientLimitExternalPerHour 400 `
                                            -RecipientLimitInternalPerHour 800 `
                                            -RecipientLimitPerDay 800 `
                                            -ActionWhenThresholdReached BlockUser

        New-HostedOutboundSpamFilterRule    -Name $EOPSpamOutName `
                                            -HostedOutboundSpamFilterPolicy "Anti_Spam_Outbound_Tenant_Policy" `
                                            -SenderDomainIs $DomaineToProtect

        $EOPSpamOutVerif = Get-HostedOutboundSpamFilterRule -Identity $EOPSpamOutName

        if ($EOPSpamOutVerif.count -gt "0") {
            $ResultTable += [PSCustomObject]@{
                "Panel" = "EOP"
                "Name" = $EOPSpamOutVerif.Identity
                "Setting" = "Spam IN policy created"
            }
        }
    #Phishing
            $EOPPhishName = "Anti_Phish_Tenant"
            $UserPhishToProtect = Get-EXOMailbox | 
            Where-Object {$_.RecipientType -eq "UserMailbox" -and $_.UserPrincipalName -notmatch "DiscoverySearch" -and   $_.UserPrincipalName -notstartwith "RA20$_.UserPrincipalName -notstartwith "Richemont$_.UserPrincipalName -notstartwith "Salle$_.UserPrincipalName -notstartwith "Wireless_hdmi$_.UserPrincipalName -notstartwith "ZooobyLoycoRunions$_.UserPrincipalName -notstartwith "Payroll$_.UserPrincipalName -notstartwith "Helpdesk"} | 
            Select-Object DisplayName, UserPrincipalName
            $FormattedList = $UserPhishToProtect | ForEach-Object { "$($_.DisplayName);$($_.UserPrincipalName)" }
            $TargetedUsersArray = $FormattedList

        if ($UserPhishToProtect.count -gt "350"){
            Write-Host "More than 350 users, can't activate user impersonation protection (users : $(($UserPhishToProtect).count))" -ForegroundColor DarkRed

            New-AntiPhishPolicy -Name "Anti_Phish_Tenant_Policy" `
                                -PhishThresholdLevel "3" `
                                -TargetedUserProtectionAction Quarantine `
                                -TargetedUserQuarantineTag "Require_Request_To_Release" `
                                -ImpersonationProtectionState "Automatic" `
                                -EnableTargetedDomainsProtection $true `
                                -EnableOrganizationDomainsProtection $true `
                                -EnableMailboxIntelligenceProtection $true `
                                -MailboxIntelligenceProtectionAction Quarantine `
                                -MailboxIntelligenceQuarantineTag "Require_Request_To_Release" `
                                -SpoofQuarantineTag "Require_Request_To_Release" `
                                -TargetedDomainProtectionAction Quarantine `
                                -TargetedDomainQuarantineTag "Require_Request_To_Release" `
                                -EnableFirstContactSafetyTips $true `
                                -EnableSimilarDomainsSafetyTips $true `
                                -EnableSimilarUsersSafetyTips $true `
                                -EnableUnusualCharactersSafetyTips $true `
                                -EnableUnauthenticatedSender $true `
                                -EnableViaTag $true `
                                -HonorDmarcPolicy $true
        } else {
            New-AntiPhishPolicy -Name "Anti_Phish_Tenant_Policy" `
                                -PhishThresholdLevel "3" `
                                -EnableTargetedUserProtection $true `
                                -TargetedUsersToProtect $TargetedUsersArray `
                                -TargetedUserProtectionAction Quarantine `
                                -TargetedUserQuarantineTag "Require_Request_To_Release" `
                                -ImpersonationProtectionState "Automatic" `
                                -EnableTargetedDomainsProtection $true `
                                -EnableOrganizationDomainsProtection $true `
                                -EnableMailboxIntelligenceProtection $true `
                                -MailboxIntelligenceProtectionAction Quarantine `
                                -MailboxIntelligenceQuarantineTag "Require_Request_To_Release" `
                                -SpoofQuarantineTag "Require_Request_To_Release" `
                                -TargetedDomainProtectionAction Quarantine `
                                -TargetedDomainQuarantineTag "Require_Request_To_Release" `
                                -EnableFirstContactSafetyTips $true `
                                -EnableSimilarDomainsSafetyTips $true `
                                -EnableSimilarUsersSafetyTips $true `
                                -EnableUnusualCharactersSafetyTips $true `
                                -EnableUnauthenticatedSender $true `
                                -EnableViaTag $true `
                                -HonorDmarcPolicy $true
        }

        New-AntiPhishRule   -Name $EOPPhishName `
                            -AntiPhishPolicy "Anti_Phish_Tenant_Policy" `
                            -RecipientDomainIs $DomaineToProtect

        $EOPPhishVerif = Get-HostedOutboundSpamFilterRule -Identity $EOPPhishName

        if ($EOPPhishVerif.count -gt "0") {
            $ResultTable += [PSCustomObject]@{
                "Panel" = "EOP"
                "Name" = $EOPPhishVerif.Identity
                "Setting" = "Spam IN policy created"
            }
        }

    #SafeLinks
        $EOPSafeLinkName = "Safe_Link"

        New-SafeLinksPolicy -Name "Safe_Link_Policy" `
                            -EnableSafeLinksForEmail $true `
                            -EnableForInternalSenders $true `
                            -ScanUrls $true `
                            -DisableUrlRewrite $false `
                            -EnableSafeLinksForTeams $true `
                            -EnableSafeLinksForOffice $true `
                            -TrackClicks $true `
                            -AllowClickThrough $false `
                            -DeliverMessageAfterScan $true `
                            -EnableOrganizationBranding $false


        
        New-SafeLinksRule   -Name $EOPSafeLinkName `
                            -SafeLinksPolicy "Safe_Link_Policy" `
                            -RecipientDomainIs $DomaineToProtect

        $EOPPhishVerif = Get-HostedOutboundSpamFilterRule -Identity $EOPSafeLinkName

        if ($EOPPhishVerif.count -gt "0") {
            $ResultTable += [PSCustomObject]@{
                "Panel" = "EOP"
                "Name" = $EOPPhishVerif.Identity
                "Setting" = "Spam IN policy created"
            }
        }

    #SafeAttachments
        $EOPSafeAttachName = "Safe_Attachment"
        New-SafeAttachmentPolicy    -Name "Safe_Attachment_Policy" `
                                    -Enable $true `
                                    -Action "DynamicDelivery" `
                                    -Redirect $false `
                                    -QuarantineTag "Require_Request_To_Release"

        New-SafeAttachmentRule  -Name $EOPSafeAttachName `
                                -SafeAttachmentPolicy "Safe_Attachment_Policy" `
                                -RecipientDomainIs $DomaineToProtect

        $EOPSafeAttacmentVerif = Get-HostedOutboundSpamFilterRule -Identity $EOPSafeAttachName

        if ($EOPSafeAttacmentVerif.count -gt "0") {
            $ResultTable += [PSCustomObject]@{
                "Panel" = "EOP"
                "Name" = $EOPSafeAttacmentVerif.Identity
                "Setting" = "Spam IN policy created"
            }
        }

    #External email Tagging
        Set-ExternalInOutlook $true
        

####### Sharepoint ##########
Connect-SPOService -Url https://$TenantName-admin.sharepoint.com/

    #Shortcut
    Set-SPOTenant   -DisableAddShortcutsToOneDrive $true 
                    -MajorVersionLimit 50 `
                    -ExpireVersionsAfterDays 30 `
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
        $ResultTable =@()
    foreach ($Setting in $ShpSettings) {
        $ResultTable += [PSCustomObject]@{
            "Panel" = "Sharepoint"
            "Name" = $Setting.Name
            "Setting" = $Setting.Settings
        }
    }
    
    $checklist = @(
    "Disable site creation by users"
    "Set default time zone when create a new site"
    "Change default site storage to 200GB"
    )

    foreach ($item in $checklist) {
        do {
            $reponse = Read-Host "$item ? (Y/N)"
        } while ($reponse -notmatch "^[OoNnYy]$") 

        if ($reponse -match "^[OoYy]$") { 
            $ResultTable += [PSCustomObject]@{
                "Panel" = "Sharepoint"
                "Name" = $item
                "Setting" = "Manually done"
            } 
        } elseif ($reponse -match "^[Nn]$"){ 
            $ResultTable += [PSCustomObject]@{
                "Panel" = "Sharepoint"
                "Name" = $item
                "Setting" = "Not done manually"
            }
        }
    }


