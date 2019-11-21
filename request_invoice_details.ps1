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
			- None = 0;
			- Sender = 2;
            - DeliveryDate = 8;
            - GrandTotalAmount = 16;
			- VatRate = 32;
            - InvoiceDate = 64;
			- Receiver = 128;
            - NetTotalAmount = 256;
            - InvoiceId = 1024;
            - DocumentType= 8192;
            - Iban = 16384;
			- Bic = 32768;
            - LineItem = 65536;
			- VatAmount = 131072;
            - InvoiceCurrency = 524288;
            - DeliveryNoteId = 1048576;
            - CustomerId = 2097152;
            - TaxNumber = 4194304;
            - UId = 8388608;
            - SenderOrderId = 16777216;
            - ReceiverOrderId = 33554432;
            - SenderOrderDate = 67108864;
            - ReceiverOrderDate = 134217728;
			- NetAmount = 268435456;
            - VatGroup = 536870912;
            - VatTotalAmount = 1073741824;

    .Parameter outputPath
    The path to the output directory, i.e. where the result filtes will be written to. The default path is './'

    .Parameter proxyUri
    An uri to a proxy server. If your requests are routed through a proxy than you must set the uri to it. Note:
    You will be requested for your Credentials

    .Example
    .\request_invoice_details.ps1 -filename 'PathToInvoice\invoice.tif' -outputPath Outputpath -resultPdf -csv

    In this example a single invoice is processed.  A result pdf and a csv file is created in addition to the json result file.

    .Example
    .\request_invoice_details.ps1 -folderPath PathToInvoices -outputPath OutputPath -resultPdf -csv -mergeCsv

    In this example a folder with invoices is processed. For each invoice a json, a result pdf and a csv file is created. Finally
    all csv files are merged into a single merged.csv file.

    .Example
    .\request_invoice_details.ps1 -folderPath PathToInvoices -outputPath OutputPath -resultPdf -csv -mergeCsv -invoiceDetails GrandTotalAmount, VatGroup

    In this example a folder with invoices is processed. Furthermore, only two InvoiceDetails are requested.

    .Example
    .\request_invoice_details.ps1 -folderPath PathToInvoices -apiKey YOURAPIKEY -url capturesdkurl -v capturesdkversion -outputPath YOUROUTPUTPATH -proxyUri http://myproxy:3128


#>
[CmdletBinding()]
param (
    [string]$folderPath="",
    [string]$filename="",
    [string]$apiKey,
    [string]$url,
    [string]$version,
    [switch]$resultPdf,
    [switch]$csv,
    [switch]$mergeCsv,
    [ValidateNotNullOrEmpty()]
    [ValidateSet(
	'None',
	'Sender',
	'DeliveryDate',
	'GrandTotalAmount',
	'VatRate',
	'InvoiceDate',
	'Receiver',
	'NetTotalAmount',
	'InvoiceId',
	'DocumentType',
	'Iban',
	'Bic',
	'LineItem',
	'VatAmount',
	'InvoiceCurrency',
	'DeliveryNoteId',
	'CustomerId',
	'TaxNumber',
	'UId',
	'SenderOrderId',
	'ReceiverOrderId',
	'SenderOrderDate',
	'ReceiverOrderDate',
	'NetAmount',
	'VatGroup',
	'VatTotalAmount')]
    [string[]]$invoiceDetails,
    [string]$outputPath=".",
    [string]$proxyUri=""
)

function InvoiceDetailsToFilterFlags {
    param (
        [string[]]$invoiceDetails
    )

    $map =@{
			None = 0;
			Sender = 2;
            DeliveryDate = 8;
            GrandTotalAmount = 16;
			VatRate = 32;
            InvoiceDate = 64;
			Receiver = 128;
            NetTotalAmount = 256;
            InvoiceId = 1024;
            DocumentType= 8192;
            Iban = 16384;
			Bic = 32768;
            LineItem = 65536;
			VatAmount = 131072;
            InvoiceCurrency = 524288;
            DeliveryNoteId = 1048576;
            CustomerId = 2097152;
            TaxNumber = 4194304;
            UId = 8388608;
            SenderOrderId = 16777216;
            ReceiverOrderId = 33554432;
            SenderOrderDate = 67108864;
            ReceiverOrderDate = 134217728;
			NetAmount = 268435456;
            VatGroup = 536870912;
            VatTotalAmount = 1073741824;}

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

    $csvResult = $singlePredictions + ($predictionGroups | Select-Object -skip 1) + $senderString + $receiverString + ($lineItemString | Select-Object -skip 1)

    # Sender, Receiver
    if($invoiceDetail -contains "Sender")
    {
        $senderString = BuildSenderReceiverStringFromObject $true $predictionResult
        $csvResult += $senderString
    }

    if($invoiceDetail -contains "Receiver")
    {
        $receiverString = BuildSenderReceiverStringFromObject $false $predictionResult
        $csvResult += $receiverString
    }

    # Line Items
    if($invoiceDetails -contains "LineItem")
    {
        $lineItemString = BuildLineItemString($predictionResult)
        $csvResult += ($lineItemString | Select-Object -skip 1)
    }

    $csvResult | Set-Content $csvFile
}

function BuildLineItemString{
    param($predictionResult)

    # Line Items - Consist of Properties and Objects that need to be expanded - explicitly handled:
    $lineItems = ($predictionResult `
        | Select-Object -expand LineItemTable `
        | Select-Object -expand LineItems) `

    $itemIds = $lineItems `
        | Select-Object -expand ItemId `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $descriptions = $lineItems `
        | Select-Object -expand Description `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')} `
        | ForEach-Object {$_.Replace("`n", '')}

    $quantities = $lineItems `
        | Select-Object -expand Quantity `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $unitPrices = $lineItems `
        | Select-Object -expand UnitPrice `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $totalAmounts = $lineItems `
        | Select-Object -expand TotalAmount `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $orderIds = $lineItems `
        | Select-Object -expand OrderId `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $deliveryIds = $lineItems `
        | Select-Object -expand DeliveryId `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $positionNumbers = $lineItems `
        | Select-Object -Property PositionNumber `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $scores = $lineItems `
        | Select-Object -Property Score `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}
  
    $xs = $lineItems `
        | Select-Object -Property X `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}
  
    $ys = $lineItems `
        | Select-Object -Property Y `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}
  
    $widths = $lineItems `
        | Select-Object -Property Width `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}
  
    $heights = $lineItems `
        | Select-Object -Property Height `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}
    
    # Build Line Item Strings
    $propertySeparator = ";"
    $finalLineItemString = New-Object 'string[]' $itemIds.Length; $arr.length
    for ($i=0; $i -lt $itemIds.Count; $i++) {
        $finalLineItemString[$i] = 
        "65536`t" + "LineItem`t" +
        
        #Text
        $itemIds[$i] + $propertySeparator +
        $descriptions[$i] + $propertySeparator +
        $quantities[$i] + $propertySeparator +
        $unitPrices[$i] + $propertySeparator +
        $totalAmounts[$i] + $propertySeparator +
        $orderIds[$i] + $propertySeparator +
        $deliveryIds[$i] + $propertySeparator +
        "PositionNumber|" + $positionNumbers[$i] + "`t" +

        #Value
        $itemIds[$i] + $propertySeparator +
        $descriptions[$i] + $propertySeparator +
        $quantities[$i] + $propertySeparator +
        $unitPrices[$i] + $propertySeparator +
        $totalAmounts[$i] + $propertySeparator +
        $orderIds[$i] + $propertySeparator +
        $deliveryIds[$i] + $propertySeparator +
        "PositionNumber|" + $positionNumbers[$i] + "`t" +

        $scores[$i] + "`t" +
        $xs[$i] + "`t" +
        $ys[$i] + "`t" +
        $widths[$i] + "`t" +
        $heights[$i] + "`t"
    }

    return $finalLineItemString
}

function BuildSenderReceiverStringFromObject{
    param([bool]$isSender, [System.Object]$predictionResult)

    $senderReceiver = $predictionResult | Select-Object -expand Sender
    $invoiceDetailId = 2
    $invoiceDetailName = "Sender"

    if(!$isSender)
    {
        $senderReceiver = $predictionResult | Select-Object -expand Receiver
        $invoiceDetailId = 128
        $invoiceDetailName = "Receiver"
    }

    $name = $senderReceiver `
        | Select-Object -expand Name `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $address = $senderReceiver `
        | Select-Object -expand Address `

    $street = $address `
        | Select-Object -expand Street `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $zipCode = $address `
        | Select-Object -expand ZipCode `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $city = $address `
        | Select-Object -expand City `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $country = $address `
        | Select-Object -expand Country `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $websiteUrls = $senderReceiver `
        | Select-Object -expand WebsiteUrl `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $emails = $senderReceiver `
        | Select-Object -expand Email `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')}

    $phone = $senderReceiver `
        | Select-Object -expand Phone `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')} `

    $fax = $senderReceiver `
        | Select-Object -expand Fax `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "|" `
        | Select-Object -skip 1 `
        | ForEach-Object {$_.Replace('"','')} `

    #Text/Value
    $propertySeparator = ";"
    $textValueString = 
        $name + $propertySeparator +
        "Address" + "|" +
        "{" + 
            $street + $propertySeparator +
            $zipCode + $propertySeparator +
            $city + $propertySeparator +
            $country + $propertySeparator +
        "}" + "|" +
        "{" + 
            $street + $propertySeparator +
            $zipCode + $propertySeparator +
            $city + $propertySeparator +
            $country + $propertySeparator +
        "}" + "|" +
        $address."Score" + "|" +
        $address."X" + "|" +
        $address."Y" + "|" +
        $address."Width" + "|" +
        $address."Height" + $propertySeparator
            
    for ($i=0; $i -lt $websiteUrls.Count; $i++) {
        $textValueString += $webSiteUrls[$i] + $propertySeparator
    }

    for ($i=0; $i -lt $emails.Count; $i++) {
        $textValueString += $emails[$i] + $propertySeparator
    }
    
    $textValueString += 
        $phone + $propertySeparator +
        $fax `

    $finalAddressString =
        $invoiceDetailId.ToString() + "`t" +
        $invoiceDetailName + "`t" +
        $textValueString + "`t" + 
        $textValueString + "`t" + 
        $senderReceiver."Score" + "`t" +
        $senderReceiver."X" + "`t" +
        $senderReceiver."Y" + "`t" +
        $senderReceiver."Width" + "`t" +
        $senderReceiver."Height" + "`t"
    
    return $finalAddressString
}

function MergeCsvFiles {
    $result = Join-Path -Path $currentLocation -ChildPath "merged.csv"
    $csvs = Get-ChildItem "$currentLocation\*.csv" -Exclude "merged.csv"
    
    $origFiles = Get-ChildItem -Path $folderPath -Name

    $write_delimiter = "`t"

    write-host "Merge all CSV to $result"

    # Definition header
    $header = 
    "FileName",
    "DeliveryDate",
    "InvoiceDate",
    "InvoiceId",
    "DocumentType",
    "Iban",
    "Bic",
    "InvoiceCurrency",
    "DeliveryNoteId",
    "CustomerId",
    "TaxNumber",
    "UId",
    "SenderOrderId",
    "ReceiverOrderId",
    "SenderOrderDate",
    "ReceiverOrderDate",
    "GrandTotalAmount",
    "NetTotalAmount",
    "NetAmount",
    "VatGroup",
    "VatAmount",
    "VatRate",
    "VatTotalAmount",
    "Sender",
    "Receiver",
    "LineItem"

    # Write header
    try
    {
        $stream = [System.IO.StreamWriter]::new( $result )
        $header | ForEach-Object{ $stream.Write( $_ + $write_delimiter) }
        $stream.WriteLine()
    }
    finally
    {
        $stream.close()
    }

    #read and append file contents minus header
    $sb = [System.Text.StringBuilder]::new()
    foreach ($csvFile in $csvs)  {
        # skip header        
        $lines = [System.IO.File]::ReadAllLines($csvFile) | Select-object -Skip 1
        
        $column = 2
        $read_delimiter = "`t"
        $property_delimiter = "|"
        $write_delimiter2 = ";"

        $line_data = @{}

        $newFile = [io.path]::GetFileNameWithoutExtension($csvFile.Name)

        write-host $newFile

        $line_data.add("file_name", $newFile)

        # read all the data into a hashtable key: Type, val: Value
        foreach($line in $lines){
            $new_key = $line.Split($read_delimiter)[1]
            $new_val = $line.Split($read_delimiter)[2]
            if(!$line_data.ContainsKey($new_key)){               
                $new_list = New-Object System.Collections.Generic.List[String]
                $line_data.$new_key = $new_list

            }            
            $line_data.$new_key.Add($new_val)
            
        }

        # VatRate, VatAmount, NetAmount --> VatGroup
        $vat_max = (@($line_data."VatRate".Count,$line_data."VatAmount".Count,$line_data."NetAmount".Count) | measure -Max).Maximum
        $new_list = New-Object System.Collections.Generic.List[String]
        $line_data."VatGroup" = $new_list

        for ($i=0; $i -le $vat_max-1; $i++) {
            if(!$line_data."VatRate"[$i]){
                $line_data."VatRate"[$i] = "NA"
            }
            if(!$line_data."VatAmount"[$i]){
                $line_data."VatAmount"[$i] = "NA"
            }
            if(!$line_data."NetAmount"[$i]){
                $line_data."NetAmount"[$i] = "NA"
            }
            $line_data."VatGroup".add($line_data."VatRate"[$i]+$property_delimiter + $line_data."VatAmount"[$i] + $property_delimiter + $line_data."NetAmount"[$i])
        }

        #Sender
        $properties = $line_data."Sender" -split ";"
        $senderString = BuildSenderReceiverStringForMerged $properties
        $line_data."Sender" = New-Object System.Collections.Generic.List[String]
        $line_data."Sender".add($senderString)

        #Receiver
        $properties = $line_data."Receiver" -split ";"
        $receiverString = BuildSenderReceiverStringForMerged $properties
        $line_data."Receiver" = New-Object System.Collections.Generic.List[String]
        $line_data."Receiver".add($receiverString)

        #LineItems
        $lineItemString = ""
        $lineItems = $line_data."LineItem" -split "ItemId\|" | Select-Object -skip 1
        for ($i=0; $i -lt $lineItems.Count; $i++)
        {
            if($i %2 -eq 1)
            {
                continue
            }
            $lineItems[$i] = "ItemId|" + $lineItems[$i]
            
            $properties = $lineItems[$i] -split ";"

            for ($j=0; $j -lt $properties.Count; $j++)
            {
                $propertyValues = $properties[$j] -split "\|"
                if($propertyValues[1] -eq "")
                {
                    $lineItemString += "NA" + "|"
                }
                else 
                {
                    $lineItemString += $propertyValues[1] + "|"
                }
            }
            $lineItemString = $lineItemString -replace ".$"
            $lineItemString += ";"
        }
        $lineItemString = $lineItemString -replace ".$"
        $line_data."LineItem" = New-Object System.Collections.Generic.List[String]
        $line_data."LineItem".add($lineItemString)

        # Write Filename
        [void]$sb.Append($line_data."file_name" + $write_delimiter)

        # Write all the fields
        $write_in_order = 
		"DeliveryDate",
		"InvoiceDate",
		"InvoiceId",
		"DocumentType",
		"Iban",
		"Bic",
		"InvoiceCurrency",
		"DeliveryNoteId",
		"CustomerId",
		"TaxNumber",
		"UId",
		"SenderOrderId",
		"ReceiverOrderId",
		"SenderOrderDate",
        "ReceiverOrderDate",
        "GrandTotalAmount",
        "NetTotalAmount",
		"NetAmount",
        "VatGroup",
        "VatAmount",
        "VatRate",
        "VatTotalAmount",
        "Sender",
		"Receiver",
		"LineItem"
		
        foreach ($my_type in $write_in_order){
            for ($i=0; $i -le $line_data.$my_type.Count-1; $i++) {
                $my_word = $line_data.$my_type[$i]
                if(!$my_word){
                    $my_word = "NA"
                }
                if($i -gt 0){
                    [void]$sb.Append($write_delimiter2)
                }
                [void]$sb.Append($my_word)
            }
            [void]$sb.Append($write_delimiter)
        }

        [void]$sb.AppendLine()
    }

    [System.IO.File]::AppendAllText($result, $sb.ToString())
}

function BuildSenderReceiverStringForMerged{
    param ($properties)
    $properties = $properties | Select-Object -skip 1
        
    $senderString = ""
    $skipToggle = $false

    for ($i=0; $i -lt $properties.Count; $i++)
    {
        $values = $properties[$i] -split "\|"

        if($values[0] -eq "}")
        {
            $skipToggle = !$skipToggle
            continue
        }

        if($skipToggle)
        {
            continue
        }

        if($values[0] -eq "Address")
        {
            $hasAddress = $true
            $values[1] = $values[2]
        }
        
        if($values[1] -eq "")
        {
            $senderString += "NA" + $property_delimiter
        }
        else 
        {
            $senderString += $values[1] + $property_delimiter   
        }
    }
    $senderString = $senderString -replace ".$"
    return $senderString
}

function ResolvePath {
    param (
        [string]$outputPath
    )

    if ((Test-Path $outputPath) -eq $true) {
        $outputPath = Resolve-Path $outputPath
        return $outputPath
    }
    else{
        # Invalid path
        Write-Error -Message "Invalid outputPath: $outputPath"
        exit 1
    }
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

    if (!$proxyUri) {
        $response = Invoke-WebRequest `
                    -Uri "$url/invoicedetail/detect" `
                    -Method POST `
                    -Body $request `
                    -ContentType "application/json" `
                    -Headers @{"accept"="application/json"; "X-ApiKey"= $apiKey} `
					-UseBasicParsing
    }
    else 
    {
        $response = Invoke-WebRequest `
            -Uri "$url/invoicedetail/detect" `
            -Method POST `
            -Body $request `
            -ContentType "application/json" `
            -Headers @{"accept"="application/json"; "X-ApiKey"= $apiKey} `
            -Proxy $proxyUri `
            -ProxyCredential $psCredentials `
			-UseBasicParsing
    }
    
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
        $baseFileName = (Get-ChildItem $filename).Name
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

# Results are written into the current folder by default
$currentLocation = ResolvePath $outputPath

# Check if we should only predict some InvoiceDetails or all available ones
$filter = if ($invoiceDetails.Length -gt 0) { InvoiceDetailsToFilterFlags -invoiceDetails $invoiceDetails } else { $filter }

if ($proxyUri) {
    $psCredentials = Get-Credential    
}


if (-Not [string]::IsNullOrEmpty($folderPath)) {
    $folderPath = ResolvePath $folderPath
    $files = Get-ChildItem (ResolvePath $folderPath)  | Where-Object {$_.extension -match 'pdf|tiff|tif|png|jpeg|jpg'} | ForEach-Object {$_.FullName}

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