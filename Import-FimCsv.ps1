<#
.NAME
    Import-FimCsv

.SYNOPSIS
    Perform manipulation of data in the FIM Service database

.DESCRIPTION
    Perform manipulation of data in the FIM Service database using 
    data supplied in a delimited file

.NOTES
    File Name   : Import-FimCsv.ps1
    Author      : SCGriffiths (Oxford Computer Group)
    Requires    : PowerShell v2
                  Microsoft PowerShell cmdlets in FIMPowerShell.ps1 (https://technet.microsoft.com/en-us/library/ff720152(v=ws.10).aspx)

.EXAMPLE
    Import-FimCsv -File NewUsers.txt

.PARAMETER File
    Identifies the import file
    

#>

[CmdletBinding()]
PARAM
(
    [parameter(Mandatory=$false, Position=0)] [string] $File = 'C:\FIMProject\Work\Default.txt'
   ,[parameter(Mandatory=$false)]             [string] $Delimiter = ','
   ,[parameter(Mandatory=$false)]             [string] $MVDelimiter = ';'
   ,[parameter(Mandatory=$false)]             [string] $ObjectType = 'Fruit'
   ,[parameter(Mandatory=$false)]             [string] $State = 'Create'
   ,[parameter(Mandatory=$false)]             [string] $Operation = 'None'
   ,[parameter(Mandatory=$false)]             [string] $MatchAttribute = 'IntegerValue'
   ,[parameter(Mandatory=$false)]             [string] $Uri = 'http://localhost:5725'
   ,[parameter(Mandatory=$false)]             [switch] $UseCachedSchema = $false
   ,[parameter(Mandatory=$false)]             [string] $MSPSCmdletScript = 'FIMPowerShell.ps1'

)

Set-Variable DEFINED_IN_FILE -Value '#File'
Set-Variable SPECIAL_ATTRIBUTES_LIST -Value '!ObjectType,!State,!Operation'
Set-Variable SPECIAL_ATTRIBUTE_OBJECT_TYPE -Value '!ObjectType'
Set-Variable SPECIAL_ATTRIBUTE_STATE -Value '!State'
Set-Variable SPECIAL_ATTRIBUTE_OPERATION -Value '!Operation'


if (-not(Test-Path $File)) {
    throw "Input file not found: $File"
}


$scriptHome = (Split-Path -Parent $MyInvocation.MyCommand.Definition)

if (-not(Test-Path $scriptHome\$MSPSCmdletScript))
{
    throw "Supporting module not found: $scriptHome\$MSPSCmdletScript"
}
else
{
    . "$scriptHome\$MSPSCmdletScript"
}


# F U N C T I O N S

function Convert-FimExportToPSObject 
{ 
    PARAM 
    ( 
        [parameter(Mandatory=$true, ValueFromPipeline = $true)] 
        [Microsoft.ResourceManagement.Automation.ObjectModel.ExportObject] 
        $ExportObject 
    )
    
    PROCESS
    {         
        $psObject = New-Object PSObject 
        $ExportObject.ResourceManagementObject.ResourceManagementAttributes | ForEach-Object{ 
            if ($_.Value -ne $null) 
            { 
                $value = $_.Value 
            } 
            elseif($_.Values -ne $null) 
            { 
                $value = $_.Values 
            } 
            else 
            { 
                $value = $null 
            } 
            $psObject | Add-Member -MemberType NoteProperty -Name $_.AttributeName -Value $value 
        } 
        Write-Output $psObject 
    } 
}

function GetFileHeader
{
    try {
        $header = (Get-Content -Path $File | Select-Object -First 1).Replace('"','')
    }
    catch [System.Exception] {
        throw "Unable to read  header for file: $File"
    }
    
    $attributes = $header.Split($Delimiter)
    $attributes
}

function GetObjectSchema
{
    $schema = GetSpecialObjectSchema
    $schema += GetFimObjectSchema
    $schema
}

function GetSpecialObjectSchema
{
    $templateProperties = @{DataType='String'; Multivalued='$false'}
    $attributeTemplate = New-Object -TypeName PSObject -Property $templateProperties
    $schema = @{}
    
    foreach ($attributeName in $SPECIAL_ATTRIBUTES_LIST.Split(',')) {
        $thisAttribute = $attributeTemplate.PSObject.Copy()
        $schema.Add($attributeName, $thisAttribute)
    }
    
    $schema
}

function GetFimObjectSchema
{
    $templateProperties = @{DataType=''; Multivalued=''}
    $attributeTemplate = New-Object -TypeName PSObject -Property $templateProperties
    $schema = @{}
    
    if (-not($ObjectType -eq $DEFINED_IN_FILE)) {
    
        foreach ($attribute in GetObjectAttributes) {
            $thisAttribute = $attributeTemplate.PSObject.Copy()
            $thisAttribute.DataType = $attribute.DataType
            $thisAttribute.Multivalued = $attribute.Multivalued
            $schema.Add("$($attribute.Name)", $thisAttribute)
        }
        
        $schema
    }
}

function GetObjectAttributes
{
    $attributes = @(Export-FIMConfig `
                    -Uri $DefaultUri  `
                    -OnlyBaseResources `
                    -CustomConfig "/BindingDescription[BoundObjectType=ObjectTypeDescription[Name='$ObjectType']]/BoundAttributeType" |
        Convert-FimExportToPSObject |
        Select Name,DataType,Multivalued)
    
    if ($attributes.Count -eq 0) {
        throw "Unknown object type: $ObjectType"
    }
    
    $attributes
}

function ValidateHeader
{
    foreach ($attributeName in $fileAttributes) {
        if (-not($schemaAttributes.ContainsKey($attributeName))) {
            throw "An attribute in the file header, $attributeName, cannot be found in the object schema for $ObjectType"
        }
    }
}

function ProcessFile
{
    Import-Csv -Path $File -Delimiter $Delimiter |
        ForEach-Object {
            ProcessRow $_
        }
}

function ProcessRow([PSCustomObject] $row)
{
    $rowObject = GetObjectType $row
    $rowState = GetState $row
    $rowOperation = GetOperation $row
        
    switch ($rowState.ToLower()) {
        create {CreateFIMObject $row}
        delete {DeleteFIMObject $row}
        put {ModifyFIMObject $row}
        default {throw "Unknown object state: $rowState"}
    }
}

function GetObjectType([PSCustomObject] $row)
{
    if ($fileAttributes -contains $SPECIAL_ATTRIBUTE_OBJECT_TYPE) {
        $row."$SPECIAL_ATTRIBUTE_OBJECT_TYPE"
    } else {
        $ObjectType
    }
}

function GetState([PSCustomObject] $row)
{
    if ($fileAttributes -contains $SPECIAL_ATTRIBUTE_STATE) {
        $row."$SPECIAL_ATTRIBUTE_STATE"
    } else {
        $State
    }
}

function GetOperation([PSCustomObject] $row)
{
    if ($fileAttributes -contains $SPECIAL_ATTRIBUTE_OPERATION) {
        $row."$SPECIAL_ATTRIBUTE_OPERATION"
    } else {
        $Operation
    }
}

function CreateFIMObject([PSCustomObject] $row)
{
    $object = CreateImportObject -ObjectType $rowObject
 
    <# - Associated functions are incomplete. Need to finish functions that manage reference types
    foreach ($attribute in $row.PSObject.Properties) {
        AddAttributesToObject $attribute $object
    }
    #>
    
    foreach ($attribute in $row.PSObject.Properties) {
        if ($schemaAttributes."$($attribute.Name)".Multivalued -eq $false -and (-not($attribute.Name.StartsWith('!')))) {
            SetSingleValue $object $attribute.Name $row."$($attribute.Name)"
        } elseif ($schemaAttributes."$($attribute.Name)".Multivalued -eq $true -and (-not($attribute.Name.StartsWith('!')))) {
            foreach ($mvAttribute in $attribute.Value.Split($MVDelimiter)) {
                AddMultiValue $object $attribute.Name $mvAttribute
            }
        }
    }
    
    $object | Import-FIMConfig -Uri $Uri
}

function AddAttributesToObject($attribute, $object)
{
    if ($schemaAttributes."$($attribute.Name)".Multivalued -eq $false) {
        AddSVAttributeToObject $object $attribute.Name $attribute.value
    } elseif ($schemaAttributes."$($attribute.Name)".Multivalued -eq $true) {
        AddMVAttributeToObject $object $attribute.Name $attribute.Value
    }
}

function AddSVAttributeToObject($object, $attributeName, $attributeValue)
{
    if ($schemaAttributes."$attributeName".DataType -eq 'Reference') {
        AddSVReferenceAttributeToObject $object $attributeName $attributeValue
    } else {
        AddSVSimpleAttributeToObject $object $attributeName $attributeValue
    }
}

function AddSVReferenceValueToObject($object, $attributeName, $attributeValue)
{
    $referenceData = @($attributeValue.Split($MVDelimeter))
    if ($referenceData.Count -ne 3) {
        throw "Invalid representation of a reference type: $referenceData"
    }
    
    $filter = "/{0}[{1}='{2}']" -f $referenceData[0], $referenceData[1], $referenceData[2],
    $referencedObject = @(QueryResource $filter $Uri)
    if ($referencedObject[0] -eq $null) {
        throw "Unable to find object for reference: $attributeValue"
    } elseif ($referencedObject.Count -eq 1) {
        SetSingleValue $object $attributeName $referencedObject.ResourceManagementObject.ObjectIdentifier
    } else {
        throw "Unable to set reference because > 1 object satisfies the criteria: $attributeValue"
    }
}

function AddSVSimpleAttributeToObject($object, $attributeName, $attributeValue)
{
    SetSingleValue $object $attributeName $attributeValue
}

function DeleteFIMObject([PSCustomObject] $row)
{
    $objectID = GetFIMObjectID $row $rowObject
    if ($objectID -ne $null) {
        DeleteObject $objectID | Import-FIMConfig -Uri $Uri
    }
    
}

function GetFIMObjectID([PSCustomObject] $row)
{
    if (-not($fileAttributes -contains $MatchAttribute)) {
        throw "The attribute chosen to find an object is not present in the file: $MatchAttribute"
    }
    
    if ($MatchAttribute -eq 'ObjectID') {
        GetResource $row.ObjectID $Uri
    } else {
        $filter = "/{0}[{1}='{2}']" -f $rowObject, $MatchAttribute, $row.$MatchAttribute
        $result = @(QueryResource $filter $Uri)
        
        if ($result[0] -eq $null) {
            Write-Verbose([string]((Get-Date -Format u) + " WARNING: Unable to delete object because object not found for MatchAttribute: $MatchAttribute"))
            $null
        } elseif ($result.Count -eq 1) {
            $result[0].ResourceManagementObject.ObjectIdentifier -replace 'urn:uuid:', ''
        } else {
            Write-Verbose([string]((Get-Date -Format u) + " WARNING: Unable to delete object because > 1 object found for MatchAttribute: $MatchAttribute"))
            $null
        }
    }
}

function ModifyFIMObject
{
}

function CacheAttributeProperties
{
<#
    Provide the name of an object type
    Read the attribute details from the FIM schema
    Write the details of each attribute to a CSV file named after the object
    Or perhaps just a single XML file such as:
    <CachedSchema>
      <Object>
        <LastUpdate>20150311T12:56:03</LastUpdate>
        <Name>Person</Name>
        <Attributes>
            <Attribute>
                <Name>DisplayName</Name>
                <DataType>String</DataType>
                <MultiValued>False</MultiValued>
            </Attribute>
            <Attribute>
                <Name>AccountName</Name>
                <DataType>String</DataType>
                <MultiValued>False</MultiValued>
            </Attribute>
        </Attributes>
      </Object>
    </CachedSchema>
#>
}



# M A I N

$schemaAttributes=@{}

$fileAttributes = @(GetFileHeader)
$schemaAttributes = GetObjectSchema

ValidateHeader
ProcessFile



# C A T C H - A L L   E R R O R   H A N D L E R

trap 
{ 
    Write-Host "`nError: $($_.Exception.Message)`n" -foregroundcolor white -backgroundcolor darkred
    Exit 1
}