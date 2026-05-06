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

        $Group = New-GraphGetRequest `
            -Uri "https://graph.microsoft.com/beta/groups?`$top=999&`$select=id,displayName&`$filter=displayName eq '$($Settings.customGroup)'" `
            -tenantid $Tenant `
            -asApp $true

        $ExpectedGroupId = ($Group.value ?? $Group | Select-Object -First 1).id
        $ExpectedGroupName = ($Group.value ?? $Group | Select-Object -First 1).displayName

        if (-not $ExpectedGroupId) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Group '$($Settings.customGroup)' was not found." -Sev Error
            return
        }
    }

    $CurrentAllowed = $CurrentInfo.azureADJoin.allowedToJoin

    $StateIsCorrect = switch ($Settings.appliesTo) {
        'all' {
            $CurrentAllowed.appliesTo -eq 'all'
        }
        'none' {
            $CurrentAllowed.appliesTo -eq 'none'
        }
        'selected' {
            ($CurrentAllowed.appliesTo -eq 'selected') -and
            ($CurrentAllowed.groups -contains $ExpectedGroupId -or $CurrentAllowed.groupIds -contains $ExpectedGroupId)
        }
        default {
            $false
        }
    }

    $CompareField = [PSCustomObject]@{
        appliesTo   = $CurrentAllowed.appliesTo
        customGroup = $ExpectedGroupName
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

            $GraphParam = @{
                tenantid    = $Tenant
                Uri         = 'https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy'
                ContentType = 'application/json; charset=utf-8'
                asApp       = $false
                type        = 'PATCH'
                Body        = @{
                    azureADJoin = @{
                        allowedToJoin = $AllowedToJoin
                    }
                } | ConvertTo-Json -Depth 10
            }

            try {
                New-GraphPostRequest @GraphParam
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Successfully configured Entra device join scope to $($Settings.appliesTo)" -Sev Info
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
            Write-StandardsAlert -message 'Entra device join scope is not correctly configured' -object $CompareField -tenant $Tenant -standardName 'EntraDeviceJoinScope' -standardId $Settings.standardId
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Entra device join scope is not correctly configured' -Sev Info
        }
    }

    if ($Settings.report -eq $true) {
        $CurrentValue = @{
            appliesTo = $CurrentAllowed.appliesTo
        }

        $ExpectedValue = @{
            appliesTo   = $Settings.appliesTo
            customGroup = $Settings.customGroup ?? ''
        }

        Set-CIPPStandardsCompareField -FieldName 'standards.EntraDeviceJoinScope' -CurrentValue $CurrentValue -ExpectedValue $ExpectedValue -TenantFilter $Tenant
        Add-CIPPBPAField -FieldName 'EntraDeviceJoinScope' -FieldValue $StateIsCorrect -StoreAs bool -Tenant $Tenant
    }
}
