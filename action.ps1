# HelloID-Task-SA-target-AzureActiveDirectory-GroupUpdateAttributes
###################################################################
# Form mapping
$formObject = @{
    description = $form.description
    displayName = $form.displayName
}
#GroupId is seperated from the form object becasue it is not send in the body of the update call
$groupIdentity = $form.groupIdentity

try {
    Write-Information "Executing AzureActiveDirectory action: [GroupUpdateAttributes] for: [$($formObject.DisplayName)]"
    Write-Information "Retrieving Microsoft Graph AccessToken for tenant: [$AADTenantID]"
    $splatTokenParams = @{
        Uri         = "https://login.microsoftonline.com/$($AADTenantID)/oauth2/token"
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = @{                                                                                                                         
            grant_type    = 'client_credentials'
            client_id     = $AADAppID
            client_secret = $AADAppSecret
            resource      = 'https://graph.microsoft.com'
        }
    }
    $accessToken = (Invoke-RestMethod @splatTokenParams).access_token

    $headers = [System.Collections.Generic.Dictionary[string, string]]::new()
    $headers.Add("Authorization", "Bearer $($accessToken)")
    $headers.Add("Content-Type", "application/json")

    $splatUpdateGroupParams = @{
        Uri         = "https://graph.microsoft.com/v1.0/groups/$($groupIdentity)"
        ContentType = 'application/json'
        Method      = 'PATCH'
        Headers     = $headers
        Body        = $formObject | ConvertTo-Json
    }
    $azureADGroup = Invoke-RestMethod @splatUpdateGroupParams

    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'AzureActiveDirectory'
        TargetIdentifier  = $groupIdentity
        TargetDisplayName = $formObject.displayName
        Message           = "AzureActiveDirectory action: [GroupUpdateAttributes] for: [$($groupIdentity)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "AzureActiveDirectory action: [GroupUpdateAttributes] for: [$($groupIdentity)] executed successfully"
}
catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'AzureActiveDirectory'
        TargetIdentifier  = $groupIdentity
        TargetDisplayName = $formObject.displayName
        Message           = "Could not execute AzureActiveDirectory action: [GroupUpdateAttributes] for: [$($groupIdentity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    if ($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') {
        $auditLog.Message = "Could not execute AzureActiveDirectory action: [GroupUpdateAttributes] for: [$($groupIdentity)], error: $($ex.ErrorDetails)" 
    } elseif ($ex.Exception.Response.StatusCode -eq 404) { 
        $auditLog.Message = "Could not execute AzureActiveDirectory action: [GroupUpdateAttributes] for: [$($groupIdentity)], the specified group does not exist in the Azure Active Directory." 
    } else { 
        $auditLog.Message = "Could not execute AzureActiveDirectory action: [GroupUpdateAttributes] for: [$($groupIdentity)], error: $($ex.Exception.Message)" 
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "$($auditLog.Message)"
}
###################################################################
