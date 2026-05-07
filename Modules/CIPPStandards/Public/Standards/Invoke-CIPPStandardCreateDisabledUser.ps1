function Invoke-CIPPStandardCreateDisabledUser {
    param($Tenant, $Settings)

    $DisplayName = $Settings.DisplayName
    $Username = $Settings.Username

    if ([string]::IsNullOrWhiteSpace($DisplayName) -or [string]::IsNullOrWhiteSpace($Username)) {
        Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Display name and username are required.' -Sev Error
        return
    }

    try {
        $Username = $Username.Trim().ToLower()

        $Domains = New-GraphGetRequest -uri 'https://graph.microsoft.com/beta/domains' -tenantid $Tenant

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
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'No onmicrosoft.com domain found.' -Sev Error
            return
        }

        $UPN = "$Username@$OnMicrosoftDomain"

        $Existing = New-GraphGetRequest `
            -uri "https://graph.microsoft.com/beta/users?`$top=999&`$filter=userPrincipalName eq '$UPN'&`$select=id,userPrincipalName,displayName,accountEnabled" `
            -tenantid $Tenant

        $ExistingUser = $Existing | Select-Object -First 1

        if ($ExistingUser.id) {
            if ($ExistingUser.accountEnabled -eq $false) {
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "User '$UPN' already exists and is disabled." -Sev Info
            } else {
                New-GraphPostRequest `
                    -uri "https://graph.microsoft.com/beta/users/$($ExistingUser.id)" `
                    -tenantid $Tenant `
                    -type PATCH `
                    -body (@{ accountEnabled = $false } | ConvertTo-Json -Compress)

                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "User '$UPN' already existed and was disabled." -Sev Info
            }

            Set-CIPPStandardsCompareField `
                -FieldName 'standards.CreateDisabledUser' `
                -CurrentValue @{
                    DisplayName       = $DisplayName
                    UserPrincipalName = $UPN
                    AccountEnabled    = $false
                } `
                -ExpectedValue @{
                    DisplayName       = $DisplayName
                    UserPrincipalName = $UPN
                    AccountEnabled    = $false
                } `
                -TenantFilter $Tenant

            Add-CIPPBPAField -FieldName 'CreateDisabledUser' -FieldValue $true -StoreAs bool -Tenant $Tenant
            return
        }

        $UserObj = [pscustomobject]@{
            tenantFilter   = $Tenant
            username       = $Username
            Domain         = $OnMicrosoftDomain
            displayName    = $DisplayName
            mailNickname   = $Username
            usageLocation  = 'CA'
            MustChangePass = $true
        }

        $Result = New-CIPPUser `
            -UserObj $UserObj `
            -APIName 'CIPP Standard - Create Disabled Cloud User'

        New-GraphPostRequest `
            -uri "https://graph.microsoft.com/beta/users/$($Result.User.id)" `
            -tenantid $Tenant `
            -type PATCH `
            -body (@{ accountEnabled = $false } | ConvertTo-Json -Compress)

        Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Created disabled user '$($Result.Username)'." -Sev Info

        Set-CIPPStandardsCompareField `
            -FieldName 'standards.CreateDisabledUser' `
            -CurrentValue @{
                DisplayName       = $DisplayName
                UserPrincipalName = $UPN
                AccountEnabled    = $false
            } `
            -ExpectedValue @{
                DisplayName       = $DisplayName
                UserPrincipalName = $UPN
                AccountEnabled    = $false
            } `
            -TenantFilter $Tenant

        Add-CIPPBPAField -FieldName 'CreateDisabledUser' -FieldValue $true -StoreAs bool -Tenant $Tenant
    } catch {
        $ErrorMessage = Get-NormalizedError -Message $_.Exception.Message
        Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Failed to process CreateDisabledUser standard. Error: $ErrorMessage" -Sev Error
    }
}
