#
## Created by: lucas.cueff[at]lucas-cueff.com
#
## released on 01/2021
#
# v1.0.0 : first public release - beta version
# v1.1.0 : last public release - beta version - fix IE com object (windows pwsh crash with invoke-webrequest)
#
#'(c) 2020-2021 lucas-cueff.com - Distributed under Artistic Licence 2.0 (https://opensource.org/licenses/artistic-license-2.0).'
<#
	.SYNOPSIS 
    Get last Microsoft Office 365 SKU / Service plans info (GUID, String ID, Product Name).

	.DESCRIPTION
    Get last Microsoft Office 365 SKU / Service plans info (GUID, String ID, Product Name).
    Resolve SKU and Service Plans from GUID, String ID, Name.
    Get all SKUs that include a Service Plan.
    Cache last catalog from Microsoft locally.
    
    .EXAMPLE
    Get-O365SKUCatalog
    Get last Microsoft Office 365 SKU / Service Plans catalog from Microsoft website and return an array of custom psobjects

    .EXAMPLE
    Get-O365SKUinfo
    Get SKU info using a GUID, String ID or Product Name

    .EXAMPLE
    Get-O365SKUInfoFromPlan
    Get all SKU and related info including a Service Plan (using GUID, String ID or Plan Name)

    .EXAMPLE
    Get-O365Planinfo
    Get Service Plan info using a GUID, String ID or Plan Name
#>
function Convert-HTMLTableToArray {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]
            $HTMLObject,
        [Parameter(Mandatory = $true)]
            [int]$TableNumber
    )
    process {
        $tables = @($HTMLObject.getElementsByTagName("TABLE"))
        $table = $tables[$TableNumber]
        $titles = @()
        $rows = @($table.Rows)
        foreach($row in $rows) {
            $cells = @($row.Cells)
            if($cells[0].tagName -eq "TH"){
                $titles = @($cells | foreach-object { ("" + $_.InnerText).Trim() })
                continue
            }
            if (!($titles)) {
                $titles = @(1..($cells.Count + 2) | foreach-object { "P$_" })
            }
            $resultObject = [Ordered] @{}
            for($counter = 0; $counter -lt $cells.Count; $counter++) {
                $title = $titles[$counter]
                if(!($title)) { continue }
                $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
            }
            [PSCustomObject]$resultObject
        }
    }
}
function Get-O365SKUCatalog {
<#
	.SYNOPSIS 
	Get last Microsoft Office 365 SKU / Service Plans information from Microsoft Website

	.DESCRIPTION
    Get last Microsoft Office 365 SKU / Service Plans information from Microsoft Website
    Downloaded from https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
	
	.PARAMETER ServicePlansInfoAsStrings
	-ServicePlansInfoAsStrings [switch]
    add a new property 'Service plans included as strings' so the object could be easily exported to an external file like a CSV file.

    .PARAMETER AsGlobalVariable
    -AsGlobalVariable [switch]
    save objets into global variable O365SKUsInfos
        
	.OUTPUTS
   	TypeName : pscustomobject
		
	.EXAMPLE
    Get-O365SKUCatalog
    Get all MS O365 SKUs / Service Plans info
    
    .EXAMPLE
    Get-O365SKUCatalog -AsGlobalVariable
    Get all MS O365 SKUs / Service Plans info and save psobjets to $global:O365SKUsInfos
#>
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$false)]
            [switch]$ServicePlansInfoAsStrings,
        [parameter(Mandatory=$false)]
            [switch]$AsGlobalVariable
    )
    process {
        #$script:URICatalog =  "https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference"
        $script:URICatalog = "https://github.com/MicrosoftDocs/azure-docs/blob/master/articles/active-directory/enterprise-users/licensing-service-plan-reference.md"
        $script:tempfile = [System.IO.Path]::GetTempFileName()
        write-verbose "Microsoft O365 SKU catalog URL : $($script:URICatalog)"
        write-verbose "Temporary html file : $($script:tempfile)"
        try {
            $request = Invoke-WebRequest -Uri $script:URICatalog -OutFile $script:tempfile -UseBasicParsing
        } catch {
            throw "Microsoft O365 SKU online catalog $($script:URICatalog) is not available. Please check your network / internet connexion."
        }
        if (!(test-path $script:tempfile)) {
            throw "not able to dowload licensing-service-plan-reference HTML content to $($script:tempfile)"
        } else {
            write-verbose "Temporary html file $($script:tempfile) created successfully"
        }
        $htmlcontent = get-content -Raw -Path $script:tempfile
        try {
            $htmlobj = New-Object -ComObject "HTMLFile"
        } catch {
            throw "not able to create HTMLFile com object"
        }
        try {
            if ($PSVersionTable.PSVersion.Major -gt 5) {
                $encodedhtmlcontent = [System.Text.Encoding]::Unicode.GetBytes($htmlcontent)
                $htmlobj.write($encodedhtmlcontent)
            } else {
                $htmlobj.IHTMLDocument2_write($htmlcontent)
            }
        }
        catch {
            throw "not able to create Com HTML object from temporary file $($script:tempfile)"
        }
        if ($htmlobj) {
            $skuinfo = Convert-HTMLTableToArray -HTMLObject $htmlobj -TableNumber 1
        }
        $skuinfo
        foreach ($sku in $skuinfo) {
            if ($sku.'Service plans included') {
                $tmpserviceplan = $sku.'Service plans included'.split("`n")
                $tmpserviceplanname = $sku.'Service plans included (friendly names)'.split("`n")
                $resultserviceplan = @()
                for ($i=0;$i -le ($tmpserviceplan.count -1);$i++) {
                    $tmpstringid = ($tmpserviceplan[$i]).substring(0,$tmpserviceplan[$i].length - 39)
                    $tmpguid = ($tmpserviceplan[$i]).substring($tmpstringid.length,$tmpserviceplan[$i].length - $tmpstringid.length)
                    $tmpplanname = ($tmpserviceplanname[$i]).substring(0,$tmpserviceplanname[$i].length - 39)
                    $resultserviceplan += [PSCustomObject]@{
                        'String ID' = $tmpstringid.replace(" ","")
                        'GUID' = ((($tmpguid.replace("(","")).replace(")","")).replace(" ","")).replace("`r","")
                        'Plan Name' = if (($tmpplanname.substring($tmpplanname.length - 1,1)) -eq " ") {
                            $tmpplanname.substring(0, $tmpplanname.length - 1)
                        } else {
                            $tmpplanname
                        }
                    }
                }
                if ($ServicePlansInfoAsStrings.IsPresent) {
                    $sku | Add-Member -NotePropertyName "Service plans included as strings" -NotePropertyValue $sku.'Service plans included'
                }
                $sku.'Service plans included' = $resultserviceplan
                $sku.PSObject.Properties.Remove('Service plans included (friendly names)')
            }
        }
        if ($AsGlobalVariable.IsPresent) {
            $global:O365SKUsInfos = $skuinfo
            write-verbose -message "Global Variable O365SKUsInfos set with SKUs Infos"
        }
        return $skuinfo
    } 
}
function Get-O365SKUinfo {
<#
	.SYNOPSIS 
	Get last Microsoft Office 365 SKU information from Microsoft Website.

	.DESCRIPTION
    Get last Microsoft Office 365 SKU information from Microsoft Website.
    Downloaded from https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
    You can search the SKU info based on its GUID, String ID or Product Name
	
	.PARAMETER GUID
	-GUID [GUID]
    search a SKU using its GUID

    .PARAMETER StringID
    -StringID [string]
    search a SKU using its StringID

    .PARAMETER ProductName
    -ProductName [string]
    search a SKU using its Product Name
        
	.OUTPUTS
   	TypeName : pscustomobject
		
	.EXAMPLE
    Get-O365SKUinfo -GUID 8f0c5670-4e56-4892-b06d-91c085d7004f
    Get SKU info based on GUID 8f0c5670-4e56-4892-b06d-91c085d7004f
    
    .EXAMPLE
    Get-O365SKUinfo -ProductName "Microsoft 365 F1"
    Get SKU info of "Microsoft 365 F1"
#>
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$false)]
            [guid]$GUID,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
            [string]$StringID,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
            [string]$ProductName
    )
    process {
        if (!($global:O365SKUsInfos)) {
            Get-O365SKUCatalog -AsGlobalVariable | out-null
        }
        if ($GUID) {
            $global:O365SKUsInfos | Where-Object {$_.GUID -eq $GUID}
        } elseif ($StringID) {
            $global:O365SKUsInfos | Where-Object {$_.'String ID' -eq $StringID}
        } elseif ($ProductName) {
            $global:O365SKUsInfos | Where-Object {$_.'Product Name' -eq $ProductName}
        } else {
            throw "please use GUID or StringID or ProductName parameters"
        }
    }
}
function Get-O365SKUInfoFromPlan {
<#
	.SYNOPSIS 
	Get last Microsoft Office 365 SKU information that included a specific Service Plan.

	.DESCRIPTION
    Get last Microsoft Office 365 SKU information that included a specific Service Plan.
    Downloaded from https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
    You can search the Service Plan info based on its GUID, String ID or Plan Name
	
	.PARAMETER GUID
	-GUID [GUID]
    search a SP using its GUID

    .PARAMETER StringID
    -StringID [string]
    search a SP using its StringID

    .PARAMETER PlanName
    -PlanName [string]
    search a SP using its Plan Name
        
	.OUTPUTS
   	TypeName : pscustomobject
		
	.EXAMPLE
    Get-O365SKUInfoFromPlan -GUID 41781fb2-bc02-4b7c-bd55-b576c07bb09d
    Get all SKU including Service Plan GUID 41781fb2-bc02-4b7c-bd55-b576c07bb09d
    
    .EXAMPLE
    Get-O365SKUInfoFromPlan -PlanName "AZURE ACTIVE DIRECTORY PREMIUM P1"
    Get all SKU including Service Plan "AZURE ACTIVE DIRECTORY PREMIUM P1"
#>
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$false)]
            [guid]$GUID,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
            [string]$StringID,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
            [string]$PlanName
    )
    process {
        if (!($global:O365SKUsInfos)) {
            Get-O365SKUCatalog -AsGlobalVariable | out-null
        }
        if ($GUID) {
            $global:O365SKUsInfos | Where-Object {$_.'Service plans included'.GUID -contains $GUID}
        } elseif ($StringID) {
            $global:O365SKUsInfos | Where-Object {$_.'Service plans included'.'String ID' -contains $StringID}
        } elseif ($PlanName) {
            $global:O365SKUsInfos | Where-Object {$_.'Service plans included'.'Plan Name' -contains $PlanName}
        } else {
            throw "please use GUID or StringID or PlanName parameters"
        }
    }
}
function Get-O365Planinfo {
<#
	.SYNOPSIS 
	Get last Microsoft Office 365 Service Plan information from Microsoft Website.

	.DESCRIPTION
    Get last Microsoft Office 365 Service Plan information from Microsoft Website.
    Downloaded from https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
    You can search the Service Plan info based on its GUID, String ID or Plan Name
	
	.PARAMETER GUID
	-GUID [GUID]
    search a SP using its GUID

    .PARAMETER StringID
    -StringID [string]
    search a SP using its StringID

    .PARAMETER PlanName
    -PlanName [string]
    search a SP using its Plan Name
        
	.OUTPUTS
   	TypeName : pscustomobject
		
	.EXAMPLE
    Get-O365Planinfo -GUID 41781fb2-bc02-4b7c-bd55-b576c07bb09d
    Get Service Plan info based on GUID 41781fb2-bc02-4b7c-bd55-b576c07bb09d
    
    .EXAMPLE
    Get-O365Planinfo -PlanName "AZURE ACTIVE DIRECTORY PREMIUM P1"
    Get Service Plan info of "AZURE ACTIVE DIRECTORY PREMIUM P1"
#>
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$false)]
            [guid]$GUID,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
            [string]$StringID,
        [parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
            [string]$PlanName
    )
    process {
        if (!($global:O365SKUsInfos)) {
            Get-O365SKUCatalog -AsGlobalVariable | out-null
        }
        if ($GUID) {
            ($global:O365SKUsInfos | Where-Object {$_.'Service plans included'.GUID -contains $GUID})[0].'Service plans included' | Where-Object {$_.GUID -eq $GUID}
        } elseif ($StringID) {
            ($global:O365SKUsInfos | Where-Object {$_.'Service plans included'.'String ID' -contains $StringID})[0].'Service plans included' | Where-Object {$_.'String ID' -eq $StringID}
        } elseif ($PlanName) {
            ($global:O365SKUsInfos | Where-Object {$_.'Service plans included'.'Plan Name' -contains $PlanName})[0].'Service plans included' | Where-Object {$_.'Plan Name' -eq $PlanName}
        } else {
            throw "please use GUID or StringID or PlanName parameters"
        }
    }
}

Export-ModuleMember -Function Get-O365SKUCatalog, Get-O365SKUinfo, Get-O365SKUInfoFromPlan, Get-O365Planinfo