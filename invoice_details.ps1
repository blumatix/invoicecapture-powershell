[CmdletBinding()]
param (
    [string]$filename,
    [string]$apiKey,
    [string]$url,
    [string]$version
)

$invoice=[System.IO.File]::ReadAllBytes($filename)

$request = @{
    "Filter" = 0;
    "Invoice" = [Convert]::ToBase64String($invoice);
    "Version" = $version;
} | ConvertTo-Json

$response = Invoke-WebRequest `
            -Uri $url `
            -Method POST `
            -Body $request `
            -ContentType "application/json" `
            -Headers @{"accept"="application/json"; "X-ApiKey"= $apiKey}

$result = $response.Content | ConvertFrom-Json | ConvertTo-Json 

Write-Host $result
