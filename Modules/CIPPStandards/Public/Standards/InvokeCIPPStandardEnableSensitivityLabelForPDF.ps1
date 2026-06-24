function Invoke-CIPPStandardEnableSensitivityLabelForPDF {
    <#
    .FUNCTIONALITY
        Internal
    .COMPONENT
        (APIName) EnableSensitivityLabelForPDF
    .SYNOPSIS
        (Label) Enable sensitivity label support for PDF files in SharePoint and OneDrive
    .DESCRIPTION
        (Helptext) Enables sensitivity label support for PDF files in SharePoint Online and OneDrive by setting EnableSensitivityLabelForPDF to true. Lets users apply and view sensitivity labels on PDF files stored in SharePoint and OneDrive.
        (DocsDescription) Sets the SharePoint Online tenant property `EnableSensitivityLabelForPDF` to `true` via the SharePoint admin endpoint. Once enabled, sensitivity labels with predefined permissions can be applied to and are honoured on PDF files in SharePoint and OneDrive, and the Sensitivity button appears in the details pane. Labels that use user-defined permissions or Double Key Encryption (DKE) are not supported for PDFs. This setting is reversible.
    .NOTES
        CAT
            SharePoint Standards
        TAG
        EXECUTIVETEXT
            Extends data-protection labelling to PDF documents stored in SharePoint and OneDrive, so PDFs receive the same classification and protection that already applies to Office files.
        ADDEDCOMPONENT
        IMPACT
            Low Impact
        ADDEDDATE
            2026-06-23
        POWERSHELLEQUIVALENT
            Set-SPOTenant -EnableSensitivityLabelForPDF $true
        RECOMMENDEDBY
            "Microsoft"
        REQUIREDCAPABILITIES
            "SHAREPOINTWAC"
            "SHAREPOINTSTANDARD"
            "SHAREPOINTENTERPRISE"
            "SHAREPOINTENTERPRISE_EDU"
            "ONEDRIVE_BASIC"
            "ONEDRIVE_ENTERPRISE"
        UPDATECOMMENTBLOCK
            Run the Tools\Update-StandardsComments.ps1 script to update this comment block
    .LINK
        https://docs.cipp.app/user-documentation/tenant/standards/list-standards
    #>

    param($Tenant, $Settings)

    $TestResult = Test-CIPPStandardLicense -StandardName 'EnableSensitivityLabelForPDF' -TenantFilter $Tenant -Preset SharePoint
    if ($TestResult -eq $false) {
        return $true
    }

    try {
        $CurrentState = Get-CIPPSPOTenant -TenantFilter $Tenant |
            Select-Object -Property _ObjectIdentity_, TenantFilter, EnableSensitivityLabelForPDF
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        Write-LogMessage -API 'Standards' -tenant $Tenant -message "Could not retrieve SharePoint tenant settings for EnableSensitivityLabelForPDF. Error: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
        return
    }

    $StateIsCorrect = ($CurrentState.EnableSensitivityLabelForPDF -eq $true)

    if ($Settings.remediate -eq $true) {
        if ($StateIsCorrect -eq $true) {
            Write-LogMessage -API 'Standards' -tenant $Tenant -message 'Sensitivity label support for PDF files is already enabled.' -sev Info
        } else {
            try {
                $Properties = @{ EnableSensitivityLabelForPDF = $true }
                $CurrentState | Set-CIPPSPOTenant -Properties $Properties
                Write-LogMessage -API 'Standards' -tenant $Tenant -message 'Enabled sensitivity label support for PDF files (EnableSensitivityLabelForPDF = true).' -sev Info
            } catch {
                $ErrorMessage = Get-CippException -Exception $_
                Write-LogMessage -API 'Standards' -tenant $Tenant -message "Failed to enable sensitivity label support for PDF files. Error: $($ErrorMessage.NormalizedError)" -sev Error -LogData $ErrorMessage
            }
        }
    }

    if ($Settings.alert -eq $true) {
        if ($StateIsCorrect -eq $true) {
            Write-LogMessage -API 'Standards' -tenant $Tenant -message 'Sensitivity label support for PDF files is enabled.' -sev Info
        } else {
            Write-StandardsAlert -message 'Sensitivity label support for PDF files is not enabled.' -object $CurrentState -tenant $Tenant -standardName 'EnableSensitivityLabelForPDF' -standardId $Settings.standardId
            Write-LogMessage -API 'Standards' -tenant $Tenant -message 'Sensitivity label support for PDF files is not enabled.' -sev Info
        }
    }

    if ($Settings.report -eq $true) {
        Set-CIPPStandardsCompareField -FieldName 'standards.EnableSensitivityLabelForPDF' -CurrentValue $StateIsCorrect -ExpectedValue $true -TenantFilter $Tenant
        Add-CIPPBPAField -FieldName 'EnableSensitivityLabelForPDF' -FieldValue $StateIsCorrect -StoreAs bool -Tenant $Tenant
    }
}
