function Invoke-CIPPStandardAddUserToGroup {
    param($Tenant, $Settings)

    $Username = $Settings.Username
    $GroupDisplayName = $Settings.GroupDisplayName

    try {
        if ($Username -like '*@*') {
            $User = New-GraphGetRequest -Uri "https://graph.microsoft.com/beta/users?`$top=999&`$select=id,userPrincipalName,mailNickname&`$filter=userPrincipalName eq '$Username'" -tenantid $Tenant -asApp $true
        } else {
            $User = New-GraphGetRequest -Uri "https://graph.microsoft.com/beta/users?`$top=999&`$select=id,userPrincipalName,mailNickname&`$filter=mailNickname eq '$Username'" -tenantid $Tenant -asApp $true
        }

        $Group = New-GraphGetRequest -Uri "https://graph.microsoft.com/beta/groups?`$top=999&`$select=id,displayName,securityEnabled&`$filter=displayName eq '$GroupDisplayName'" -tenantid $Tenant -asApp $true

        if (-not $User.id) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "User '$Username' was not found." -Sev Error
            return
        }

        if (-not $Group.id) {
            Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Group '$GroupDisplayName' was not found." -Sev Error
            return
        }

        $Members = New-GraphGetRequest -Uri "https://graph.microsoft.com/beta/groups/$($Group.id)/members?`$top=999&`$select=id" -tenantid $Tenant -asApp $true
        $StateIsCorrect = @($Members).id -contains $User.id

        if ($Settings.remediate -eq $true) {
            if ($StateIsCorrect -eq $true) {
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "User '$($User.userPrincipalName)' is already a member of '$($Group.displayName)'." -Sev Info
            } else {
                $GraphParam = @{
                    tenantid     = $Tenant
                    Uri          = "https://graph.microsoft.com/beta/groups/$($Group.id)/members/`$ref"
                    ContentType  = 'application/json; charset=utf-8'
                    asApp        = $true
                    type         = 'POST'
                    Body         = @{
                        '@odata.id' = "https://graph.microsoft.com/odata/directoryObjects('$($User.id)')"
                    } | ConvertTo-Json
                }

                New-GraphPostRequest @GraphParam
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "Added user '$($User.userPrincipalName)' to group '$($Group.displayName)'." -Sev Info
                $StateIsCorrect = $true
            }
        }

        if ($Settings.alert -eq $true) {
            if ($StateIsCorrect -eq $true) {
                Write-LogMessage -API 'Standards' -Tenant $Tenant -Message "User '$($User.userPrincipalName)' is correctly a member of '$($Group.displayName)'." -Sev Info
            } else {
                Write-StandardsAlert -message "User is not a member of the expected group." -tenant $Tenant -standardName 'AddUserToGroup' -standardId $Settings.standardId
            }
        }

        if ($Settings.report -eq $true) {
            Set-CIPPStandardsCompareField -FieldName 'standards.AddUserToGroup' -CurrentValue @{
                UserPrincipalName = $User.userPrincipalName
                GroupDisplayName  = $Group.displayName
                IsMember          = $StateIsCorrect
            } -ExpectedValue @{
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
