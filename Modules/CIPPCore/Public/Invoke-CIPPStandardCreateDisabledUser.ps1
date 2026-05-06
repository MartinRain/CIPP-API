function Invoke-CIPPStandardCreateDisabledUser {
    [CmdletBinding()]
    param(
        $TenantFilter,
        $Settings,
        $Headers
    )

    try {
        $DisplayName = $Settings.DisplayName
        $Username = $Settings.Username

        if ([string]::IsNullOrWhiteSpace($DisplayName) -or [string]::IsNullOrWhiteSpace($Username)) {
            return @{
                state   = 'Error'
                message = 'Display name and username are required.'
            }
        }

        $Username = $Username.Trim().ToLower()

        $Domains = New-GraphGetRequest `
            -uri 'https://graph.microsoft.com/beta/domains' `
            -tenantid $TenantFilter

        $OnMicrosoftDomain = (
            $Domains |
            Where-Object { $_.id -like '*.onmicrosoft.com' -and $_.isInitial -eq $true } |
            Select-Object -First 1
        ).id

        if (-not $OnMicrosoftDomain) {
            $OnMicrosoftDomain = (
                $Domains |
                Where-Object { $_.id -like '*.onmicrosoft.com' } |
                Select-Object -First 1
            ).id
        }

        if (-not $OnMicrosoftDomain) {
            return @{
                state   = 'Error'
                message = 'No onmicrosoft.com domain found.'
            }
        }

        $UPN = "$Username@$OnMicrosoftDomain"

        $Existing = New-GraphGetRequest `
            -uri "https://graph.microsoft.com/beta/users?`$filter=userPrincipalName eq '$UPN'&`$select=id,userPrincipalName,accountEnabled" `
            -tenantid $TenantFilter

        if ($Existing.value.Count -gt 0) {
            $ExistingUser = $Existing.value | Select-Object -First 1

            if ($ExistingUser.accountEnabled -eq $false) {
                return @{
                    state   = 'Good'
                    message = "User already exists and is disabled: $UPN"
                }
            }

            $DisableBody = @{
                accountEnabled = $false
            } | ConvertTo-Json -Depth 5 -Compress

            New-GraphPostRequest `
                -uri "https://graph.microsoft.com/beta/users/$($ExistingUser.id)" `
                -tenantid $TenantFilter `
                -type PATCH `
                -body $DisableBody

            return @{
                state   = 'Good'
                message = "User already existed and was disabled: $UPN"
            }
        }

        $UserObj = [pscustomobject]@{
            tenantFilter   = $TenantFilter
            username       = $Username
            Domain         = $OnMicrosoftDomain
            displayName    = $DisplayName
            mailNickname   = $Username
            usageLocation  = 'CA'
            MustChangePass = $true
        }

        $Result = New-CIPPUser `
            -UserObj $UserObj `
            -APIName 'CIPP Standard - Create Disabled Cloud User' `
            -Headers $Headers

        $DisableBody = @{
            accountEnabled = $false
        } | ConvertTo-Json -Depth 5 -Compress

        New-GraphPostRequest `
            -uri "https://graph.microsoft.com/beta/users/$($Result.User.id)" `
            -tenantid $TenantFilter `
            -type PATCH `
            -body $DisableBody

        return @{
            state   = 'Good'
            message = "Created disabled user: $($Result.Username)"
            data    = $Result
        }
    }
    catch {
        $ErrorMessage = Get-CippException -Exception $_

        return @{
            state   = 'Error'
            message = $ErrorMessage.NormalizedError
        }
    }
}
