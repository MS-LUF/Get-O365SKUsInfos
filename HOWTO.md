# Get-O365SKUsInfos - HOW TO

## version 1.0.0
- first document version

## install module
`Install-Module -Name Get-O365SKUsInfos`

## import module in your powershell environment
`Import-Module Get-O365SKUsInfos`
### when installed out of default PowerShell modules path
`Import-Module c:\mypath\Use-UptoBox\Get-O365SKUsInfos.psd1`

## Manage your internet access
### using System.Net.WebProxy
- you can use `System.Net.WebProxy` object to set your proxy environment globally in your current powershell session.
- for instance use the following lines of code to define a proxy url *http://myproxy.tld:8080*
```
    $proxyobj = New-Object System.Net.WebProxy "http://myproxy.tld:8080"
    [System.Net.WebRequest]::DefaultWebProxy = $proxyobj

```
## PS Custom Object properties
### Product name
- SKU property giving the commercial name of the SKU
### String ID
- SKU or Service Plan property giving the technical name of a Service Plan or SKU
### GUID
- SKU or Service Plan property giving the GUID of a Service Plan or SKU
### Service plans included
- SKU property giving the Services Plan included in a SKU (as **PSCustom Objects**)
### Plan Name
- Service Plan property giving the commercial name of the Service Plan

## Get last SKUs / Services Plan catalog from Microsoft website
### get information online and a powershell custom object output
`Get-O365SKUCatalog`
### get information online and cache the result in a global variable
- can be usefull if you want to save the result and import it back later for instance or just optimize your internet outgoing flows  
`Get-O365SKUCatalog -AsGlobalVariable`  
- global variable used : `$global:O365SKUsInfos`
### get Service Plan information as string instead of PSObject
- can be useful if you want to export data in to a CSV file for instance  
`Get-O365SKUCatalog -ServicePlansInfoAsStrings`

## Get SKU information
### from a GUID
`Get-O365SKUinfo -GUID 8f0c5670-4e56-4892-b06d-91c085d7004f`
### from a product name (aka commercial name)
`Get-O365SKUinfo -ProductName "Microsoft 365 F1"`
### from a technical name (aka String ID)
`Get-O365SKUinfo -StringID M365_F1`

## Get Service Plan information
### from GUID
`Get-O365Planinfo -GUID 41781fb2-bc02-4b7c-bd55-b576c07bb09d`
### from a plan name (aka commercial name)
`Get-O365Planinfo -PlanName "AZURE ACTIVE DIRECTORY PREMIUM P1"`
### from a technical name (aka String ID)
`Get-O365Planinfo -StringID AAD_PREMIUM`

## Get SKU information from a Service Plan
### from a Service Plan GUID
`Get-O365SKUInfoFromPlan -GUID 41781fb2-bc02-4b7c-bd55-b576c07bb09d`
### from a plan name (aka commercial name)
`Get-O365SKUInfoFromPlan -PlanName "AZURE ACTIVE DIRECTORY PREMIUM P1"`
### from a Service Plan technical name (aka String ID)
`Get-O365SKUInfoFromPlan -StringID AAD_PREMIUM`
