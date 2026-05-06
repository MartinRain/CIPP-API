function Invoke-CIPPStandardAddUserToGroup {
    <#
    .FUNCTIONALITY
        Internal
    .COMPONENT
        (APIName) AddUserToGroup
    .SYNOPSIS
        (Label) Add user to group
    .DESCRIPTION
        (Helptext) Adds a user to a security group by username/UPN prefix and group display name.
    #>

    param($Tenant, $Settings)

    $Username = $Settings.Username
    $GroupDisplayName = $Settings.GroupDisplayName

    if ([string]::IsNullOrWhiteSpace($Username) -or [string]::IsNullOrWhiteSpace($GroupDisplayName)) {
        Write-LogMessage -API 'Standards' -Tenant $Tenant -Message 'Username and Group Display Name are required.' -Sev Error
        return
    }

    try {
        if ($Username -like '*@*') {
            $User = (New-GraphGetRequest `
                -Uri "https://graph.microsoft.com/beta/users?`$select=id,userPrincipalName&`$filter=userPrincipalName eq '$Username'" `
                -tenantid $Tenant `
                -asApp $true).value
        } else {
            $User = (New-GraphGetRequest `
                -Uri "https://graph.microsoft.com/beta/users?`$select=id,userPrincipalName,mailNickname&`$filter=mailNickname eq '$Username'" `
                -tenantid $Tenant `
                -asApp $true).value
        }

        $Group = (New-GraphGetRequest `
            -Uri "https://graph.microsoft.com/beta/groups?`$select=id,displayName,securityEnabled&`$filter=displayName eq '$GroupDisplayName'" `
            -tenantid $Tenant `
            -asApp $true).value

        if (@($User).Count -ne 1) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "User '$Username' was not found or is ambiguous." -Sev Error
            return
        }

        if (@($Group).Count -ne 1) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Group '$GroupDisplayName' was not found or is ambiguous." -Sev Error
            return
        }

        $User = @($User)[0]
        $Group = @($Group)[0]

        $Members = New-GraphGetRequest `
            -Uri "https://graph.microsoft.com/beta/groups/$($Group.id)/members?`$select=id&`$top=999" `
            -tenantid $Tenant `
            -asApp $true

        $IsMember = @($Members.value).id -contains $User.id

        $StateIsCorrect = $IsMember

        $CompareField = [PSCustomObject]@{
            UserPrincipalName = $User.userPrincipalName
            GroupDisplayName  = $Group.displayName
            IsMember          = $IsMember
        }

        if ($Settings.remediate -eq $true) {
            if ($StateIsCorrect -eq $true) {
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "User '$($User.userPrincipalName)' is already a member of '$($Group.displayName)'." -Sev Info
            } else {
                New-GraphPostRequest `
                    -tenantid $Tenant `
                    -Uri "https://graph.microsoft.com/beta/groups/$($Group.id)/members/`$ref" `
                    -ContentType 'application/json; charset=utf-8' `
                    -asApp $true `
                    -type POST `
                    -Body (@{
                        '@odata.id' = "https://graph.microsoft.com/beta/directoryObjects/$($User.id)"
                    } | ConvertTo-Json)

                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Added user '$($User.userPrincipalName)' to group '$($Group.displayName)'." -Sev Info
                $StateIsCorrect = $true
            }
        }

        if ($Settings.alert -eq $true) {
            if ($StateIsCorrect -eq $true) {
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "User '$($User.userPrincipalName)' is correctly a member of '$($Group.displayName)'." -Sev Info
            } else {
                Write-StandardsAlert -message "User is not a member of the expected group." -object $CompareField -tenant $Tenant -standardName 'AddUserToGroup' -standardId $Settings.standardId
            }
        }

        if ($Settings.report -eq $true) {
            Set-CIPPStandardsCompareField -FieldName 'standards.AddUserToGroup' -CurrentValue $CompareField -ExpectedValue @{
                Username         = $Username
                GroupDisplayName = $GroupDisplayName
                IsMember         = $true
            } -TenantFilter $Tenant

            Add-CIPPBPAField -FieldName 'AddUserToGroup' -FieldValue $StateIsCorrect -StoreAs bool -Tenant $Tenant
        }
    } catch {
        $ErrorMessage = Get-NormalizedError -Message $_.Exception.Message
        Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Failed to process AddUserToGroup standard. Error: $ErrorMessage" -Sev Error
    }
}
