# invoicecapture-powershell
Contains powershell scripts to access our capture client

## Usage:
```sh
.\invoice_details.ps1 -filename PATH_TO_INVOICE_FILE -apiKey YOUR_API_KEY -version CAPTURE_VERSION -url CAPTURE_SDK_URL -resultPdf -csv
```
## Result
This script produces the following files
- invoiceResult.csv: The InvoiceDetails predictions in csv format
- invoiceResult.json: The prediction result in json format
- invoiceResult.pdf: An invoice result pdf which contains markers for all detected invoice details
