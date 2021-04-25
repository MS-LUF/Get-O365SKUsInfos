![image](http://www.lucas-cueff.com/files/gallery.png)

# Get-O365SKUsInfos
Get last Microsoft Office 365 SKU / Service plans info (GUID, String ID, Product Name).
- You can use this module to :
    - resolve GUID service plan / SKU and find the product name / plan name linked to it
    - find in what SKU a service plan is covered
    - get general informations about a Service Plan or SKU 

(c) 2021 lucas-cueff.com Distributed under Artistic Licence 2.0 (https://opensource.org/licenses/artistic-license-2.0).

## Notes
- SKU / Service Plan informations updated directly from Microsoft Website and send back as powershell object.
- this module requires an internet access to get online information from Microsoft Website.
    - if you are using an indirect internet access through a proxy, please use `System.Net.WebProxy` object to set your environment before using the functions
- this module is designed to be used on **Windows system**, compatible with both **Windows PowerShell** and **PowerShell**
    - `HTMLFile` com object is used to parse HTML table

## Notes version :
### 1.0.0 first public release
 - cmdlet to list all SKUs / Services Plan available in Microsoft catalog : *Get-O365SKUCatalog*
     - information are downloaded from [microsoft github doc repository](https://github.com/MicrosoftDocs/azure-docs/blob/master/articles/active-directory/enterprise-users/licensing-service-plan-reference.md)
 - cmdlet to get all availabe information about a SKU : *Get-O365SKUinfo*
     - you can find a SKU based on its GUID, String ID (technical name), or friendly name (aka commercial name)
 - cmdlet to get all availabe information about a Service Plan : *Get-O365Planinfo*
     - you can find a Service Plan based on its GUID, String ID (technical name), or friendly name (aka commercial name)
 - cmdlet to find in the SKUs containing a Service Plan : *Get-O365SKUInfoFromPlan*
     - you can use a Service Plan based on its GUID, String ID (technical name), or friendly name (aka commercial name)
### 1.1.0
 - fix IE com object issue (Windows pwsh crash with invoke-webrequest)
 - using basic parsing and HTMLFile com object
 - now compatible with Powershell Core on Windows system
 - use now Github Markdown document as source instead of Microsoft website
### 1.1.1 last public release
 - minor update, replace `$host` with `$psversiontable`

## Why this PowerShell Module
- When you often deal with Office 365 SKUs and Services Plan (technically speaking) it's a nightmare to resolve name to technical GUID to be sure the proper SKU and Service Plan is linked to the right Azure AD user or Azure AD group.
- Moreover the licensing Graph API is built on a "black list" system regarding the Services Plans it means you have to specify all Services Plan to be disabled instead of just giving the one to be enabled...
- I hope it could help someone to deal with word instead of GUID in your scripts :)
- Also, sometimes some guys are asking me what SKU should be ordered in order to cover X ou Y Services Plan... So instead of looking for markdown or html document from Microsoft, you can quickly have your answer using a PowerShell oneliner command :)
- Last but not least... those informations (SKUs & Service Plans) could be updated several times a year ! it means you cannot have a static catalog file somewhere, you have to follow the updates... (*you know... those famous O365 / Azure Microsoft mails you receive every day and never read... ^^*)

## How To
[Simple How TO](https://github.com/MS-LUF/Get-O365SKUsInfos/blob/main/HOWTO.md)

## install Get-O365SKUsInfos from PowerShell Gallery repository
You can easily install it from [powershell gallery repository](https://www.powershellgallery.com/packages/Get-O365SKUsInfos/) using a simple powershell command and an internet access :-) 
```
	Install-Module -Name Get-O365SKUsInfos
```

## import module from PowerShell 
```
	C:\PS> import-module Get-O365SKUsInfos
```

## module content
### function
- Get-O365SKUCatalog
- Get-O365SKUinfo
- Get-O365SKUInfoFromPlan
- Get-O365Planinfo
