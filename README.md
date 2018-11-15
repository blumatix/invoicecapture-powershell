# invoicecapture-powershell
Contains powershell scripts to access our capture client

__NOTE__: In order to being able to execute this script you may have to update the user preference for the PowerShell execution policy.
For further information please refer to [powershell execution policy settings](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-6)


## Usage:
```sh
.\invoice_details.ps1 -filename PATH_TO_INVOICE_FILE -apiKey YOUR_API_KEY -version CAPTURE_VERSION -url CAPTURE_SDK_URL -resultPdf -csv
```
## Result
This script produces the following files
- invoiceResult.csv: The InvoiceDetails predictions in csv format
- invoiceResult.json: The prediction result in json format
- invoiceResult.pdf: An invoice result pdf which contains markers for all detected invoice details
