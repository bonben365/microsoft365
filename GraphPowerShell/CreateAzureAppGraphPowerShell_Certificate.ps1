$appName =  "testapp1"
$app = New-MgApplication -DisplayName $appName
$appObjectId = $app.Id

Get-MgApplication -ApplicationId $appObjectId | select DisplayName, Id, AppId

$certParams = @{
    Subject = 'CN=GraphApi'
    CertStoreLocation = 'cert:\LocalMachine\My'
    NotAfter = (Get-Date).AddYears(5)
    KeySpec = 'KeyExchange'
}
$mycert = New-SelfSignedCertificate @certParams

$cert = Get-ChildItem -Path Cert:\LocalMachine\my | ? {$_.Subject -eq "$($certParams.Subject)"}
$CertCredentials = @(
    @{
        Type = "AsymmetricX509Cert"
        Usage = "Verify"
        Key = [byte[]]$cert.RawData
    }
)
Update-MgApplication -ApplicationId $appObjectId -KeyCredentials $CertCredentials

Write-Host "Client ID: $($app.AppID)"
Write-Host "Tenent ID: $((Get-MgOrganization).Id)"
Write-Host "Cert Thumbprint: $($mycert.Thumbprint)"