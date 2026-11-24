function Invoke-CIPPStandardEnableMIPLabels {
    <#
    .FUNCTIONALITY
        Internal
    .COMPONENT
        (APIName) EnableMIPLabels
    .SYNOPSIS
        (Label) Enable Sensitivity Labels for groups, Teams and SharePoint sites (containers)
    .DESCRIPTION
        (Helptext) Enables Microsoft Purview sensitivity labels for Microsoft 365 Groups, Teams and SharePoint sites by setting EnableMIPLabels to true on the Group.Unified directory setting. This is the tenant-wide prerequisite before sensitivity labels can be scoped to containers. Requires Entra ID P1. A label sync is still required before the "Groups & sites" scope appears on a label.
        (DocsDescription) Sets `EnableMIPLabels` to `true` on the tenant's `Group.Unified` directory setting (creating the setting from the Microsoft default template if it does not yet exist). This is the tenant-wide prerequisite that lets sensitivity labels be applied to Microsoft 365 Groups, Teams and SharePoint sites (container labeling). After enabling, the published sensitivity labels must still be synchronised to Entra ID (label sync) before the "Groups & sites" scope becomes selectable on a label. Note: enabling container labeling is effectively one-way.
    .NOTES
        CAT
            SharePoint Standards
        TAG
        EXECUTIVETEXT
            Allows the organisation to apply protective classification labels to collaboration spaces (Microsoft Teams, Microsoft 365 groups and SharePoint sites). This is the prerequisite that lets the business govern how those spaces are created, shared and protected.
        ADDEDCOMPONENT
        IMPACT
            Medium Impact
        ADDEDDATE
            2026-06-23
        POWERSHELLEQUIVALENT
            Update-MgBetaDirectorySetting (EnableMIPLabels = true)
        RECOMMENDEDBY
            "Microsoft"
        REQUIREDCAPABILITIES
            "AAD_PREMIUM"
            "AAD_PREMIUM_P2"
        UPDATECOMMENTBLOCK
            Run the Tools\Update-StandardsComments.ps1 script to update this comment block
    .LINK
        https://docs.cipp.app/user-documentation/tenant/standards/list-standards
    #>

    param($Tenant, $Settings)

    $TestResult = Test-CIPPStandardLicense -StandardName 'EnableMIPLabels' -TenantFilter $Tenant -Preset Entra
    if ($TestResult -eq $false) {
        return $true
    }

    # Read the Group.Unified directory setting (EnableMIPLabels lives inside it).
    $CurrentState = (New-GraphGetRequest -Uri 'https://graph.microsoft.com/beta/settings' -tenantid $Tenant -AsApp $true) |
        Where-Object -Property displayName -EQ 'Group.Unified'

    $MIPValue = $null
    if ($CurrentState) {
        $MIPValue = ($CurrentState.values | Where-Object { $_.name -eq 'EnableMIPLabels' }).value
    }
    $StateIsCorrect = ($MIPValue -eq 'true')

    if ($Settings.remediate -eq $true) {
        if ($StateIsCorrect -eq $true) {
            Write-LogMessage -API 'Standards' -tenant $Tenant -message 'Sensitivity labels for containers (EnableMIPLabels) are already enabled.' -sev Info
        } else {
            try {
                if (-not $CurrentState) {
                    # No Group.Unified setting yet: create it from the Microsoft default template, then re-read.
                    $DefaultTemplate = '{"templateId":"62375ab9-6b52-47ed-826b-58e47e0e304b","values":[{"name":"NewUnifiedGroupWritebackDefault","value":"true"},{"name":"EnableMIPLabels","value":"false"},{"name":"CustomBlockedWordsList","value":""},{"name":"EnableMSStandardBlockedWords","value":"false"},{"name":"ClassificationDescriptions","value":""},{"name":"DefaultClassification","value":""},{"name":"PrefixSuffixNamingRequirement","value":""},{"name":"AllowGuestsToBeGroupOwner","value":"false"},{"name":"AllowGuestsToAccessGroups","value":"true"},{"name":"GuestUsageGuidelinesUrl","value":""},{"name":"GroupCreationAllowedGroupId","value":""},{"name":"AllowToAddGuests","value":"true"},{"name":"UsageGuidelinesUrl","value":""},{"name":"ClassificationList","value":""},{"name":"EnableGroupCreation","value":"true"}]}'
                    $null = New-GraphPostRequest -tenantid $Tenant -AsApp $true -Uri 'https://graph.microsoft.com/beta/settings' -Type POST -Body $DefaultTemplate -ContentType 'application/json'
                    $CurrentState = (New-GraphGetRequest -Uri 'https://graph.microsoft.com/beta/settings' -tenantid $Tenant -AsApp $true) |
                        Where-Object -Property displayName -EQ 'Group.Unified'
                }
                ($CurrentState.values | Where-Object { $_.name -eq 'EnableMIPLabels' }).value = 'true'
                $Body = "{values : $($CurrentState.values | ConvertTo-Json -Compress)}"
                $null = New-GraphPostRequest -tenantid $Tenant -AsApp $true -Uri "https://graph.microsoft.com/beta/settings/$($CurrentState.id)" -Type patch -Body $Body -ContentType 'application/json'
                Write-LogMessage -API 'Standards' -tenant $Tenant -message 'Enabled sensitivity labels for containers (EnableMIPLabels = true). Run a label sync so labels can be scoped to Groups & sites.' -sev Info
            } catch {
                $ErrorMessage = Get-CippException -Exception $_
                Write-LogMessage -API 'Standards' -tenant $Tenant -message "Failed to enable sensitivity labels for containers. Error: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
            }
        }
    }

    if ($Settings.alert -eq $true) {
        if ($StateIsCorrect -eq $true) {
            Write-LogMessage -API 'Standards' -tenant $Tenant -message 'Sensitivity labels for containers (EnableMIPLabels) are enabled.' -sev Info
        } else {
            Write-StandardsAlert -message 'Sensitivity labels for containers (EnableMIPLabels) are not enabled.' -object @{ EnableMIPLabels = $MIPValue } -tenant $Tenant -standardName 'EnableMIPLabels' -standardId $Settings.standardId
            Write-LogMessage -API 'Standards' -tenant $Tenant -message 'Sensitivity labels for containers (EnableMIPLabels) are not enabled.' -sev Info
        }
    }

    if ($Settings.report -eq $true) {
        Set-CIPPStandardsCompareField -FieldName 'standards.EnableMIPLabels' -CurrentValue $StateIsCorrect -ExpectedValue $true -TenantFilter $Tenant
        Add-CIPPBPAField -FieldName 'EnableMIPLabels' -FieldValue $StateIsCorrect -StoreAs bool -Tenant $Tenant
    }
}
