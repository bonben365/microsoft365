#Connect-MgGraph -Scopes "Application.Read.All","Application.ReadWrite.All","User.Read.All"

$appName =  "testapp"
$app = New-MgApplication -DisplayName $appName
$appObjectId = $app.Id

Get-MgApplication -ApplicationId $appObjectId | select DisplayName, Id, AppId

$passwordCred = @{
    "displayName" = "DemoClientSecret"
    "endDateTime" = (Get-Date).AddMonths(+12)
}
$clientSecret = Add-MgApplicationPassword `
    -ApplicationId $appObjectId `
    -PasswordCredential $passwordCred

$clientSecret | Format-List

#Add Application Permission
#User.ReadBasic.All    Application    97235f07-e226-4f63-ace3-39588e11d3a1
$permissionParams = @{
    RequiredResourceAccess = @(
        @{
            ResourceAppId = "00000003-0000-0000-c000-000000000000"
            ResourceAccess = @(
                @{
                    Id = "97235f07-e226-4f63-ace3-39588e11d3a1"
                    Type = "Role"
                }
            )
        }
    )
}
Update-MgApplication -ApplicationId $appObjectId -BodyParameter $permissionParams

Write-Host "Client ID: $($app.AppID)"
Write-Host "Tenent ID: $((Get-MgOrganization).Id)"
Write-Host "Client Secret: $($clientSecret.SecretText)"


#Configuration
#$ClientId = ''
#$TenantId = ''
#$ClientSecret = ''

#Convert the client secret to a secure string
#$ClientSecretPass = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force

#Create a credential object using the client ID and secure string
#$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ClientId, $ClientSecretPass

#Connect to Microsoft Graph with Client Secret
#Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential
