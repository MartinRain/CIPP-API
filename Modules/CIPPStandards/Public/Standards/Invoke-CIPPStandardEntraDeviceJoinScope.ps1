function Invoke-CIPPStandardEntraDeviceJoinScope {
    <#
    .FUNCTIONALITY
        Internal
    .COMPONENT
        (APIName) EntraDeviceJoinScope
    .SYNOPSIS
        (Label) Configure Entra device join scope
    .DESCRIPTION
        (Helptext) Configures which users are allowed to join devices to Microsoft Entra ID.
    #>

    param($Tenant, $Settings)

    try {
        $CurrentInfo = New-GraphGetRequest `
            -uri 'https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy' `
            -tenantid $Tenant
    } catch {
        $ErrorMessage = Get-NormalizedError -Message $_.Exception.Message
        Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Could not get Entra device join scope for $Tenant. Error: $ErrorMessage" -Sev Error
        return
    }

    $ExpectedGroupId = $null
    $ExpectedGroupName = $null

    if ($Settings.appliesTo -eq 'selected') {
        if ([string]::IsNullOrWhiteSpace($Settings.customGroup)) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Custom Group Name is required when scope is set to Custom Group.' -Sev Error
            return
        }

        try {
            $Group = New-GraphGetRequest `
                -Uri "https://graph.microsoft.com/beta/groups?`$top=999&`$select=id,displayName&`$filter=displayName eq '$($Settings.customGroup)'" `
                -tenantid $Tenant `
                -asApp $true

            $ExpectedGroupId = $Group.id
            $ExpectedGroupName = $Group.displayName
        } catch {
            $ErrorMessage = Get-NormalizedError -Message $_.Exception.Message
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Could not find group '$($Settings.customGroup)'. Error: $ErrorMessage" -Sev Error
            return
        }

        if (-not $ExpectedGroupId) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Group '$($Settings.customGroup)' was not found." -Sev Error
            return
        }
    }

    $CurrentAllowed = $CurrentInfo.azureADJoin.allowedToJoin

    $CurrentGroupIds = @()
    if ($CurrentAllowed.groups) {
        $CurrentGroupIds = @($CurrentAllowed.groups)
    } elseif ($CurrentAllowed.groupIds) {
        $CurrentGroupIds = @($CurrentAllowed.groupIds)
    }

    $CurrentAppliesTo = if ($CurrentGroupIds.Count -gt 0) {
        'selected'
    } elseif ($CurrentAllowed.'@odata.type' -like '*noDeviceRegistrationMembership') {
        'none'
    } else {
        'all'
    }

    $CurrentGroupName = if ($Settings.appliesTo -eq 'selected' -and ($CurrentGroupIds -contains $ExpectedGroupId)) {
        $ExpectedGroupName
    } elseif ($CurrentGroupIds.Count -gt 0) {
        $CurrentGroupIds -join ','
    } else {
        ''
    }

    $StateIsCorrect = switch ($Settings.appliesTo) {
        'all' {
            $CurrentAppliesTo -eq 'all'
        }
        'none' {
            $CurrentAppliesTo -eq 'none'
        }
        'selected' {
            ($CurrentAppliesTo -eq 'selected') -and ($CurrentGroupIds -contains $ExpectedGroupId)
        }
        default {
            $false
        }
    }

    $CompareField = [PSCustomObject]@{
        appliesTo   = $CurrentAppliesTo
        customGroup = $CurrentGroupName
    }

    if ($Settings.remediate -eq $true) {
        if ($StateIsCorrect -eq $true) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Entra device join scope already correctly configured' -Sev Info
        } else {
            $AllowedToJoin = switch ($Settings.appliesTo) {
                'all' {
                    @{
                        '@odata.type' = '#microsoft.graph.allDeviceRegistrationMembership'
                    }
                }
                'none' {
                    @{
                        '@odata.type' = '#microsoft.graph.noDeviceRegistrationMembership'
                    }
                }
                'selected' {
                    @{
                        '@odata.type' = '#microsoft.graph.enumeratedDeviceRegistrationMembership'
                        users         = @()
                        groups        = @($ExpectedGroupId)
                    }
                }
            }

            $Body = $CurrentInfo | Select-Object `
                id,
                userDeviceQuota,
                multiFactorAuthConfiguration,
                azureADRegistration,
                azureADJoin,
                localAdminPassword

            $Body.azureADJoin.allowedToJoin = $AllowedToJoin

            $GraphParam = @{
                tenantid    = $Tenant
                Uri         = 'https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy'
                ContentType = 'application/json; charset=utf-8'
                asApp       = $false
                type        = 'PUT'
                Body        = $Body | ConvertTo-Json -Depth 20
            }

            try {
                New-GraphPostRequest @GraphParam
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Successfully configured Entra device join scope to $($Settings.appliesTo)" -Sev Info
                $StateIsCorrect = $true
                $CurrentAppliesTo = $Settings.appliesTo
                $CurrentGroupName = $Settings.customGroup ?? ''
            } catch {
                $ErrorMessage = Get-NormalizedError -Message $_.Exception.Message
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Failed to configure Entra device join scope.' -Sev Error -LogData $ErrorMessage
            }
        }
    }

    if ($Settings.alert -eq $true) {
        if ($StateIsCorrect -eq $true) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Entra device join scope is correctly configured' -Sev Info
        } else {
            Write-StandardsAlert `
                -message 'Entra device join scope is not correctly configured' `
                -object $CompareField `
                -tenant $Tenant `
                -standardName 'EntraDeviceJoinScope' `
                -standardId $Settings.standardId

            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Entra device join scope is not correctly configured' -Sev Info
        }
    }

    if ($Settings.report -eq $true) {
        $CurrentValue = @{
            appliesTo   = $CurrentAppliesTo
            customGroup = $CurrentGroupName
        }

        $ExpectedValue = @{
            appliesTo   = $Settings.appliesTo
            customGroup = $Settings.customGroup ?? ''
        }

        Set-CIPPStandardsCompareField `
            -FieldName 'standards.EntraDeviceJoinScope' `
            -CurrentValue $CurrentValue `
            -ExpectedValue $ExpectedValue `
            -TenantFilter $Tenant

        Add-CIPPBPAField `
            -FieldName 'EntraDeviceJoinScope' `
            -FieldValue $StateIsCorrect `
            -StoreAs bool `
            -Tenant $Tenant
    }
}
