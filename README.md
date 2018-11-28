# invoicecapture-powershell
Contains powershell scripts to access our capture client

__NOTE__: In order to being able to execute this script you may have to update the user preference for the PowerShell execution policy.
For further information please refer to [powershell execution policy settings](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-6)

## Usage Examples
The following example processes a single invoice.
```sh
.\request_invoice_details.ps1 -filename PATH_TO_INVOICE_FILE -apiKey YOUR_API_KEY -version CAPTURE_VERSION -url CAPTURE_SDK_URL -resultPdf -csv
```

This script processes all invoices within a certain folder. Only file with pdf, tiff, tif, jpeg, jpg or png extensions
will be taken into account. It outputs json, csv and pdf files. Finally the all produced csv files are merged into
a single merged.csv file
```sh
.\request_invoice_details.ps1 -folderPath PATH_TO_INVOICE_Folder -apiKey YOUR_API_KEY -version CAPTURE_VERSION -url CAPTURE_SDK_URL -resultPdf -csv -mergeCsv
```

This script processes all invoices within a certain folder. Furthermore only the GrandTotalAmount and VatGroup InvoiceDetails will be returned.
```sh
.\request_invoice_details.ps1 -folderPath PATH_TO_INVOICE_Folder -apiKey YOUR_API_KEY -version CAPTURE_VERSION -url CAPTURE_SDK_URL -resultPdf -csv -mergeCsv -invoiceDetails GrandTotalAmount, VatGroup
```


This script processes all invoices within a certain folder and writes the result files (json, csv and pdf) into a provided output folder
```sh
.\request_invoice_details.ps1 -folderPath PATH_TO_INVOICE_Folder -apiKey YOUR_API_KEY -version CAPTURE_VERSION -url CAPTURE_SDK_URL -resultPdf -csv -mergeCsv -invoiceDetails GrandTotalAmount, VatGroup -outputPath C:\tmp\output
```

## Result
This script produces the following files for each invoice
- invoiceResult.csv: The InvoiceDetails predictions in csv format
- invoiceResult.json: The prediction result in json format
- invoiceResult.pdf: An invoice result pdf which contains markers for all detected invoice details
