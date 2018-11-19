<#
    .Synopsis
    Requests InvoiceDetails for a single invoice

    .Description
    This script can be used to request InvoiceDetails for a given invoice. The response of an
    InvoiceDetail request is written in json format to an invoiceResult.json file. Moreover
    it is also able to write a result invoice pdf file with markers for each detected InvoiceDetail
    as well as a csv file containing the InvoiceDetail prediction results.

    .Parameter folderPath
    A path to a directory which contains invoice to be processed

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

    .Parameter mergeCsv
    If set then all produced csv files are merged into single merged.csv file.

    .Parameter invoiceDetails
    A list of InvoiceDetails which shall be returned
        - DeliveryDate: 8
        - GrandTotalAmount: 16
        - InvoiceDate: 64
        - NetTotalAmount: 256
        - InvoiceId: 1024
        - DocumentType: 8192
        - Iban: 16384
        - InvoiceCurrency: 524288
        - DeliveryNoteId: 1048576
        - CustomerId: 2097152
        - TaxNo: 4194304
        - UId: 8388608
        - SenderOrderId: 16777216
        - ReceiverOrderId: 33554432
        - SenderOrderDate: 67108864
        - ReceiverOrderDate: 134217728
        - VatGroup: 536870912
        - VatTotalAmount: 1073741824    
#>
[CmdletBinding()]
param (
    [string]$folderPath="",
    [string]$filename="",
    [string]$apiKey=,
    [string]$url,
    [string]$version,
    [switch]$resultPdf,
    [switch]$csv,
    [switch]$mergeCsv,
    [ValidateNotNullOrEmpty()]
    [ValidateSet('DeliveryDate','GrandTotalAmount','InvoiceDate','InvoiceId','DocumentType','Iban',`
                'InvoiceCurrency','DeliveryNoteId','CustomerId','UId','SenderOrderId','ReceiverOrderId','SenderOrderDate',`
                'ReceiverOrderDate','VatGroup','CustomInvoiceDetail')]
    [string[]]$invoiceDetails
)

function InvoiceDetailsToFilterFlags {
    param (
        [string[]]$invoiceDetails
    )

    $map =@{DeliveryDate = 8;
            GrandTotalAmount = 16;
            InvoiceDate = 64;
            NetTotalAmount = 256;
            InvoiceId = 1024;
            DocumentType= 8192;
            Iban = 16384;
            InvoiceCurrency = 524288;
            DeliveryNoteId = 1048576;
            CustomerId = 2097152;
            TaxNo = 4194304;
            UId = 8388608;
            SenderOrderId = 16777216;
            ReceiverOrderId = 33554432;
            SenderOrderDate = 67108864;
            ReceiverOrderDate = 134217728;
            VatGroup = 536870912;
            VatTotalAmount = 1073741824;
            CustomInvoiceDetail = -2147483648}   

    $filterMask = 0

    foreach ($invoiceDetail in $invoiceDetails) {
        $filterMask = $filterMask -bor $map[$invoiceDetail]
    }

    return $filterMask
}

function WriteJson {
    param (
        [string]$baseFilename,
        $resultObject
    )

    $result = $resultObject | ConvertTo-Json -Depth 10
    $jsonFile = Join-Path -Path $currentLocation -ChildPath "$($baseFilename).json"
    write-host "Write InvoiceDetails predictions to $jsonFile"
    [System.IO.File]::WriteAllText($jsonFile, $result)
}

function WritePdf {
    param(
        [string]$baseFilename,
        [string]$resultPdfBase64
    )

    $resultPdfInvoice = [Convert]::FromBase64String($resultPdfBase64)
    $pdfFile = Join-Path -Path $currentLocation -ChildPath "$($baseFilename).pdf"    
    write-host "Write InvoiceResultPdf to $pdfFile"    
    [System.IO.File]::WriteAllBytes($pdfFile, $resultPdfInvoice)
}

function WriteCsv {
    param(
        [string]$baseFilename,
        $predictionResult
    )

    $csvFile = Join-Path -Path $currentLocation -ChildPath "$($baseFilename).csv"
    
    write-host "Write InvoiceResult csv to $csvFile"

    $singlePredictions = $predictionResult `
        | Select-Object -expand InvoiceDetailTypePredictions  `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | ForEach-Object {$_.Replace('"','')}

    # VatGroups: VatRate, NetAmount, VatAmount 
    $predictionGroups = ($predictionResult `
        | Select-Object -expand PredictionGroups) `
        | Select-Object -expand InvoiceDetailTypePredictions `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | ForEach-Object {$_.Replace('"','')}

    # Merge predictions, skip header of prediction groups                                
    $csvResult = $singlePredictions + ($predictionGroups | Select-Object -skip 1)
    $csvResult | Set-Content $csvFile
}

function MergeCsvFiles {
    $result = Join-Path -Path $currentLocation -ChildPath "merged.csv"
    $csvs = Get-ChildItem "*.csv" 
    
    #read and write CSV header
    $header = [System.IO.File]::ReadAllLines($csvs[0])[0] + "`t" + "filename"
    [System.IO.File]::WriteAllLines($result, $header)
    
    #read and append file contents minus header
    $sb = [System.Text.StringBuilder]::new()
    foreach ($csvFile in $csvs)  {
        # skip header        
        $lines = [System.IO.File]::ReadAllLines($csvFile) | Select-object -Skip 1
        
        foreach ($line in $lines) {
            $newLine = $line + "`t" + $csvFile.Name
            $sb.AppendLine($newLine)
        }
    }

    [System.IO.File]::AppendAllText($result, $sb.ToString())
}

function PostRequest {
    param (
        $invoice,
        $version,
        $filter
    )

    # Prepare request
    $request = @{
        "Filter" = $filter;
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
    
    return $response
}

function ProcessInvoice {
    param (
        [string]$filename,
        $filter
    )

    $invoice=[System.IO.File]::ReadAllBytes($filename)
    $response = PostRequest $invoice $version $filter
    
    # On success            
    if ($response.statuscode -eq 200) {
        $baseFileName = (Get-ChildItem $filename).BaseName
        $resultObject = $response.Content | ConvertFrom-Json

        # Check invoice state
        if ($resultObject.InvoiceState -eq 'Failed') {
            write-host "InvoiceDetails processing failed - InvoiceState == Failed"
            return
        }
    
        WriteJson $baseFileName $resultObject
        
        if ($resultPdf -eq $true) {
            WritePdf $baseFileName $resultObject.ResultPdf
        }
    
        if ($csv -eq $true) {
            WriteCsv $baseFileName $resultObject
        }
    }
    else {
        Write-Host "Http status code $($response.statuscode), status descripion $($response.statusdescription)"
    }    
}

# Results are written into the current folder
Set-Location -Path "."
$currentLocation = Get-Location

$filter = 0
if ($invoiceDetails.Length -gt 0) {
    $filter = InvoiceDetailsToFilterFlags -invoiceDetails $invoiceDetails
    write-host $filter
}

if (-Not [string]::IsNullOrEmpty($folderPath) -and (Test-Path $folderPath)) {
    $files = Get-ChildItem $folderPath | Where-Object {$_.extension -match 'pdf|tiff|tif|png|jpeg|jpg'} | ForEach-Object {$_.FullName}

    foreach ($filename in $files) {        
        write-host $filename
        ProcessInvoice -filename $filename -filter $filter
    }
}
elseif (-Not [string]::IsNullOrEmpty($filename) -and (Test-Path $filename)) {
    ProcessInvoice -filename $filename -filter $filter
}

if ($mergeCsv -eq $true)
{
    MergeCsvFiles
}