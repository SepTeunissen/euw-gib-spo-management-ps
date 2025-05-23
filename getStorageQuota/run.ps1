using namespace System.Net
<# Voorbeeld input:
{
    "tenant": "giantict",
    "clientID": "123c0119-9d9b-4028-a142-b61b65b76495",
    "keyVault": "euw-gib-kv-tst-autom"
}
#>
# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Initialize the response object
$response = @{
    status  = "success"
    message = ""
    data    = $null   
}

#inputs
$tenant = $Request.Body.tenant 
$clientID = $Request.Body.clientID 
$keyVault = $Request.Body.keyVault 
$certificateName = "$tenant-SPManagementCertificate"
$tenantFull = $tenant + ".onmicrosoft.com"
$siteurl = "https://$tenant-admin.sharepoint.com/"

$certSecret = Get-AzKeyVaultSecret -VaultName $keyVault -Name $certificateName
$certsecretValueText = ($certSecret.SecretValue | ConvertFrom-SecureString -AsPlainText )

write-host "tenant: $tenant"
write-host "clientID: $clientID"
write-host "keyVault: $keyVault"   
write-host "certificateName: $certificateName"
write-host "tenantFull: $tenantFull"
write-host "siteurl: $siteurl" 

write-host "certSecret: $certSecret"
write-host "certsecretValueText: $certsecretValueText"

try {
    write-host "Connecting to SharePoint Online"
    write-host "Connecting to $siteurl with clientID $clientID and tenant $tenantFull and $certsecretValueText"

    $env:PNPPOWERSHELL_UPDATECHECK = "off"
    Connect-PnPOnline -Url $siteUrl -ClientId $clientID -Tenant $tenantFull -CertificateBase64Encoded $certsecretValueText 
    write-host "Connected to SharePoint Online"

    $result = (Invoke-PnPSPRestMethod -Url "/_api/StorageQuotas()?api-version=1.3.2").value
    $result.value

    Write-Host "result $result"
    $response.status = "success"    
    $response.message = "Storage Quota retrieved successfully"
    $response.data = $result    
    Write-Host "Succes"
}
catch {    
    $err = $_.Exception.Message
    $line = $_.InvocationInfo.ScriptLineNumber

    $response.status = "failure"
    $response.message = "$err at linenumber $line"
    $response.data = $null   
}

# Convert the response to JSON
$responseJson = $response | ConvertTo-Json

# Return the response
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = $responseJson
    })


 