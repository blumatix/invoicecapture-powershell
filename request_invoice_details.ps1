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
	- BankCode = 4294967296;
	- BankAccount = 8589934592;
	- BankGroup = 17179869184;
    - IsrReference = 34359738368;
    - DiscountDate = 68719476736;
    - DiscountStart = 137438953472;
    - DiscountDuration = 274877906944;
    - DiscountPercent = 549755813888;
    - DiscountGroup = 1099511627776;
    - DueDateDate = 2199023255552;
    - DueDateStart = 4398046511104;
    - DueDateDuration = 8796093022208;
    - DueDateGroup = 17592186044416;
    - IsrSubscriber = 35184372088832;
    - KId = 70368744177664;
    - CompanyRegistrationNumber = 140737488355328;
    - Contacts = 281474976710656;

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
	'VatTotalAmount',
	'BankCode',
	'BankAccount',
	'BankGroup',
    'IsrReference',
    'DiscountDate',
    'DiscountStart',
    'DiscountDuration',
    'DiscountPercent',
    'DiscountGroup',
    'DueDateDate',
    'DueDateStart',
    'DueDateDuration',
    'DueDateGroup',
    'IsrSubscriber',
    'KId',
    'CompanyRegistrationNumber',
    'Contacts'
    )]
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
    VatTotalAmount = 1073741824;
	BankCode = 4294967296;
	BankAccount = 8589934592;
	BankGroup = 17179869184;
    IsrReference = 34359738368;
    DiscountDate = 68719476736;
    DiscountStart = 137438953472;
    DiscountDuration = 274877906944;
    DiscountPercent = 549755813888;
    DiscountGroup = 1099511627776;
    DueDateDate = 2199023255552;
    DueDateStart = 4398046511104;
    DueDateDuration = 8796093022208;
    DueDateGroup = 17592186044416;
    IsrSubscriber = 35184372088832;
    KId = 70368744177664;
    CompanyRegistrationNumber = 140737488355328;
    Contacts = 281474976710656;
    }

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

    # VatGroups: VatRate, NetAmount, VatAmount & BankCode, BankAccount & DiscountDate, DiscountStart, DiscountDuration, DiscountPercent & DueDateDate, DueDateStart, DueDateDuration
    $predictionGroups = ($predictionResult `
        | Select-Object -expand PredictionGroups) `
        | Select-Object -expand InvoiceDetailTypePredictions `
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" `
        | ForEach-Object {$_.Replace('"','')}

    # Contacts
    $convertedContacts = @()

    $predictionResult `
        | Select-Object -expand Contacts `
        | ForEach-Object {$convertedContacts += ConvertContactToDetailTypeResponse($_)}

    $convertedContacts = ($convertedContacts`
        | ConvertTo-Csv -NoTypeInformation -Delimiter "`t"`
        | ForEach-Object {$_.Replace('"','')})

    $csvResult = $singlePredictions + ($predictionGroups | Select-Object -skip 1) + ($convertedContacts | Select-Object -skip 1)
    $csvResult | Set-Content $csvFile
}

function ConvertContactToDetailTypeResponse ($contact) {
    $obj = New-Object -TypeName psobject
    $obj | Add-Member -MemberType NoteProperty -Name Type -Value '281474976710656'
    $obj | Add-Member -MemberType NoteProperty -Name TypeName -Value 'Contacts'    
    $obj | Add-Member -MemberType NoteProperty -Name Text -Value $contact."Name"."Text"
    $obj | Add-Member -MemberType NoteProperty -Name Value -Value $contact."Name"."Value" 
    $obj | Add-Member -MemberType NoteProperty -Name Score -Value $contact."Name"."Score"
    $obj | Add-Member -MemberType NoteProperty -Name X -Value $contact."Name"."X"   
    $obj | Add-Member -MemberType NoteProperty -Name Y -Value $contact."Name"."Y"
    $obj | Add-Member -MemberType NoteProperty -Name Width -Value $contact."Name"."Width" 
    $obj | Add-Member -MemberType NoteProperty -Name Height -Value $contact."Name"."Height"
    $obj | Add-Member -MemberType NoteProperty -Name Confidence -Value '-1'
    
    return $obj
}

function MergeCsvFiles {
    $result = Join-Path -Path $currentLocation -ChildPath "merged.csv"
    $csvs = Get-ChildItem "$currentLocation\*.csv" -Exclude "merged.csv"

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
    "BankCode",
	"BankAccount",
	"BankGroup",
    'IsrReference',
    'DiscountDate',
    'DiscountStart',
    'DiscountDuration',
    'DiscountPercent',
    'DiscountGroup',
    'DueDateDate',
    'DueDateStart',
    'DueDateDuration',
    'DueDateGroup',
    'IsrSubscriber',
    'KId',
    'CompanyRegistrationNumber',
    'Contacts'

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
        $lines = [System.IO.File]::ReadAllLines($csvFile)
        if($lines.Length -gt 0 -and $lines[0].StartsWith("Type"))
        {
            $lines = $lines | Select-object -Skip 1
        }
        
        $read_delimiter = "`t"
        $property_delimiter = "|"

        # This used to be ';' - i see no reason why we should not unify it with the property_delimiter
        $write_delimiter2 = "|"

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
        $new_list_vat = New-Object System.Collections.Generic.List[String]
        $line_data."VatGroup" = $new_list_vat
        $length = $vat_max-1

        $line_data = AddEmptyEntriesInLineData -key "VatRate" -length $length -line_data $line_data
        $line_data = AddEmptyEntriesInLineData -key "VatAmount" -length $length -line_data $line_data
        $line_data = AddEmptyEntriesInLineData -key "NetAmount" -length $length -line_data $line_data

        $line_data = ReplaceNullEntriesInLineData -key "VatRate" -length $length -line_data $line_data
        $line_data = ReplaceNullEntriesInLineData -key "VatAmount" -length $length -line_data $line_data
        $line_data = ReplaceNullEntriesInLineData -key "NetAmount" -length $length -line_data $line_data

        for ($i=0; $i -le $length; $i++) {
            $line_data."VatGroup".add($line_data."VatRate"[$i]+$property_delimiter + $line_data."VatAmount"[$i] + $property_delimiter + $line_data."NetAmount"[$i])
        }

        # BankCode, BankAccount --> BankGroup
        $bank_max = (@($line_data."BankCode".Count,$line_data."BankAccount".Count) | measure -Max).Maximum
        $new_list_bank = New-Object System.Collections.Generic.List[String]
        $line_data."BankGroup" = $new_list_bank
        $length = $bank_max-1

        $line_data = AddEmptyEntriesInLineData -key "BankCode" -length $length -line_data $line_data
        $line_data = AddEmptyEntriesInLineData -key "BankAccount" -length $length -line_data $line_data

        $line_data = ReplaceNullEntriesInLineData -key "BankCode" -length $length -line_data $line_data
        $line_data = ReplaceNullEntriesInLineData -key "BankAccount" -length $length -line_data $line_data

        for ($i=0; $i -le $length; $i++) {
            $line_data."BankGroup".add($line_data."BankCode"[$i]+$property_delimiter + $line_data."BankAccount"[$i])
        }

        # DiscountDate, DiscountStart, DiscountDuration, DiscountPercent --> DiscountGroup
        $discount_max = (@($line_data."DiscountDate".Count,$line_data."DiscountStart".Count,$line_data."DiscountDuration".Count,$line_data."DiscountPercent".Count) | measure -Max).Maximum
        $new_list_discount = New-Object System.Collections.Generic.List[String]
        $line_data."DiscountGroup" = $new_list_discount
        $length = $discount_max-1

        $line_data = AddEmptyEntriesInLineData -key "DiscountDate" -length $length -line_data $line_data
        $line_data = AddEmptyEntriesInLineData -key "DiscountStart" -length $length -line_data $line_data
        $line_data = AddEmptyEntriesInLineData -key "DiscountDuration" -length $length -line_data $line_data
        $line_data = AddEmptyEntriesInLineData -key "DiscountPercent" -length $length -line_data $line_data

        $line_data = ReplaceNullEntriesInLineData -key "DiscountDate" -length $length -line_data $line_data
        $line_data = ReplaceNullEntriesInLineData -key "DiscountStart" -length $length -line_data $line_data
        $line_data = ReplaceNullEntriesInLineData -key "DiscountDuration" -length $length -line_data $line_data
        $line_data = ReplaceNullEntriesInLineData -key "DiscountPercent" -length $length -line_data $line_data

        for ($i=0; $i -le $length; $i++) {
            $line_data."DiscountGroup".add($line_data."DiscountDate"[$i] + $property_delimiter + $line_data."DiscountStart"[$i] + $property_delimiter + $line_data."DiscountDuration"[$i] + $property_delimiter + $line_data."DiscountPercent"[$i])
        }

        # DueDateDate, DueDateStart, DueDateDuration --> DueDateGroup
        $duedate_max = (@($line_data."DueDateDate".Count,$line_data."DueDateStart".Count,$line_data."DueDateDuration".Count) | measure -Max).Maximum
        $new_list_duedate = New-Object System.Collections.Generic.List[String]
        $line_data."DueDateGroup" = $new_list_duedate
        $length = $duedate_max-1

        $line_data = AddEmptyEntriesInLineData -key "DueDateDate" -length $length -line_data $line_data
        $line_data = AddEmptyEntriesInLineData -key "DueDateStart" -length $length -line_data $line_data
        $line_data = AddEmptyEntriesInLineData -key "DueDateDuration" -length $length -line_data $line_data

        $line_data = ReplaceNullEntriesInLineData -key "DueDateDate" -length $length -line_data $line_data
        $line_data = ReplaceNullEntriesInLineData -key "DueDateStart" -length $length -line_data $line_data
        $line_data = ReplaceNullEntriesInLineData -key "DueDateDuration" -length $length -line_data $line_data

        for ($i=0; $i -le $length; $i++) {
            $line_data."DueDateGroup".add($line_data."DueDateDate"[$i] + $property_delimiter + $line_data."DueDateStart"[$i] + $property_delimiter + $line_data."DueDateDuration"[$i])
        }

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
        "BankCode",
        "BankAccount",
        "BankGroup",
        'IsrReference',
        'DiscountDate',
        'DiscountStart',
        'DiscountDuration',
        'DiscountPercent',
        'DiscountGroup',
        'DueDateDate',
        'DueDateStart',
        'DueDateDuration',
        'DueDateGroup',
        'IsrSubscriber',
        'KId',
        'CompanyRegistrationNumber',
        'Contacts'
		
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


function AddEmptyEntriesInLineData {
    param (
        [string]$key,
        $length,
        $line_data
    )

    if(!$line_data.ContainsKey($key))
    {
        $line_data.$key = New-Object System.Collections.Generic.List[String]
        for($i=0; $i -le $length; $i++)
        {
            $line_data.$key.add("NA")
        }
    }

    return $line_data
}

function ReplaceNullEntriesInLineData {
    param (
        [string]$key,
        $length,
        $line_data
    )

    for ($i=0; $i -le $length; $i++) {
        if(!$line_data.$key[$i]){
            $line_data.$key[$i] = "NA"
        }    
    }

    return $line_data
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