<#
    .Synopsis
    Requests the InvoiceDetails prediction for a single invoice

    .Description
    This script can be used to request InvoiceDetails for a given invoice. The response of an
    InvoiceDetail request is written in json format to an invoiceResult.json file. Moreover
    it is also able to write a result invoice pdf file with markers for each detected InvoiceDetail
    as well as a csv file containing the InvoiceDetail prediction results.

    .Parameter filename
    The path to the invoice file that shall be processed. Currently we support the following formats:
        - pdf
        - png
        - jpeg
        - tiff

    .Parameter apiKey
    Your apiKey which is needed for authentication and authorisation

    .Parameter url
    The base url to our service

    .Parameter version
    The current service version.

    .Parameter resultPdf
    If set then a result pdf file is generated

    .Parameter csv
    If set then a csv file is generated

#>
[CmdletBinding()]
param (
    [string]$filename,
    [string]$apiKey,
    [string]$url,
    [string]$version,
    [switch]$resultPdf,
    [switch]$csv
)

function WriteJson {
    param (
        $resultObject
    )

    $result = $resultObject | ConvertTo-Json -Depth 10
    $jsonFile = Join-Path -Path $currentLocation -ChildPath "invoiceResult.json"
    write-host "Write InvoiceDetails predictions to $jsonFile"
    [System.IO.File]::WriteAllText($jsonFile, $result)
}

function WritePdf {
    param(
        [string]$resultPdfBase64
    )

    $resultPdfInvoice = [Convert]::FromBase64String($resultPdfBase64)
    $pdfFile = Join-Path -Path $currentLocation -ChildPath "invoiceResult.pdf"
    write-host "Write InvoiceResultPdf to $pdfFile"
    [System.IO.File]::WriteAllBytes($pdfFile, $resultPdfInvoice)
}

function WriteCsv {
    param(
        $predictionResult
    )

    $csvFile = Join-Path -Path $currentLocation -ChildPath "invoiceResult.csv"
    write-host "Write InvoiceResult csv to $csvFile"
    $singlePredictions = $predictionResult   | Select-Object -expand InvoiceDetailTypePredictions  `
                                                | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
                                                | ForEach-Object {$_.Replace('"','')}

    # VatGroups: VatRate, NetAmount, VatAmount 
    $predictionGroups = ($predictionResult | Select-Object -expand PredictionGroups) `
                                | Select-Object -expand InvoiceDetailTypePredictions `
                                | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
                                | ForEach-Object {$_.Replace('"','')}

    # Merge predictions, skip header of prediction groups                                
    $csvResult = $singlePredictions + ($predictionGroups | Select-Object -skip 1)

    $csvResult | Set-Content $csvFile
}

Set-Location -Path "."
$currentLocation = Get-Location
$invoice=[System.IO.File]::ReadAllBytes($filename)

# Prepare request
$request = @{
    "Filter" = 0;
    "Invoice" = [Convert]::ToBase64String($invoice);
    "Version" = $version;
    "CreateResultPdf" = if ($resultPdf.IsPresent) {1} else {0};
} | ConvertTo-Json

# Send request
$response = Invoke-WebRequest `
            -Uri "$url/invoicedetail/detect" `
            -Method POST `
            -Body $request `
            -ContentType "application/json" `
            -Headers @{"accept"="application/json"; "X-ApiKey"= $apiKey}

# On success            
if ($response.statuscode -eq 200) {
    $resultObject = $response.Content | ConvertFrom-Json

    WriteJson $resultObject
    
    if ($resultPdf -eq $true) {
        WritePdf $resultObject.ResultPdf
    }

    if ($csv -eq $true)
    {
        WriteCsv $resultObject
    }
}
else {
    Write-Host "Http status code $($response.statuscode), status descripion $($response.statusdescription)"
}
