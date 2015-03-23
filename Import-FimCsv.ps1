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
				  which needs some minor modification to the GetResource function to use -OnlyBaseResources when calling Export-FIMConfig

.EXAMPLE
	The following example creates Person objects from the import file NewUsers.txt. The
	header of the import file must contain the system name of the FIM attribute to be imported.
	Reference (Manager) and multi-valued (ProxyAddressCollection) attributes are also
	illustrated in the sample input file.

	NewUsers.txt
	------------
	EmployeeID,FirstName,LastName,Manager,ProxyAddressCollection
	100123,Alice,Roberts,(Person|EmployeeID|757011),SMTP:alice@acme.com;smtp:alice.roberts@acmecorp.com

    Import-FimCsv -File NewUsers.txt -ObjectType Person -State Create

.EXAMPLE
	The following example creates Person objects from the import file NewUsers.txt and
	includes the State as a column header rather than as a parameter:

	NewUsers.txt
	------------
	!State,EmployeeID,DisplayName,FirstName,LastName
	Create,100123,Alice Roberts,Alice,Roberts

	Import-FimCsv -File NewUsers.txt -ObjectType Person

.EXAMPLE
	The following example deletes Person objects identified in the import file OldUsers.txt.
	The Person objects to be deleted are identified through the MatchAttribute parameter. A 
	single match is required for the object to be deleted.

	OldUsers.txt
	------------
	EmployeeID,FirstName,LastName
	7001345,Bob,Jones
	8561120,Charlie,Smith

	Import-FimCsv -File OldUsers.txt -ObjectType Person -State Delete -MatchAttribute EmployeeID

.EXAMPLE
	The following example modifies the Person objects identified in the import file
	UpdatedUsers.txt. The Person objects to be modified are identified through the MatchAttribute
	parameter. A single match is required for the object to be updated.

	UpdatedUsers.txt
	----------------
	EmployeeID,FirstName,LastName,Manager
	100123,Alex,Robins,(Person|EmployeeID|977541)

	Import-FimCsv -File UpdatedUsers.txt -ObjectType Person -State Put -MatchAttribute EmployeeID

.PARAMETER File
    Identifies the import file

.PARAMETER Delimiter
	Identifies the field separator for the attributes represented in the import file.

	The default is a comma (,).
    
.PARAMETER MVDelimiter
	Identifies the separator within a field that represents a multi-valued attribute.

	The default is a semi-colon (;).

.PARAMETER
	Identifies the separator within a field that represents a reference attribute.

	The default is the pipe symbol (|).
	
	The format of a field representing a reference attribute is (ObjectType|Attribute|Value).

.PARAMETER ObjectType
	Identifies the type of object represented in the import file.

	The default is Person.

.PARAMETER State
	Identifies the type of activity to be performed and is one of Create, Put, or Delete

	If the header of the import file contains a column called !State, the value in this
	column is used instead of that provided by this parameter. This is to allow create,
	update and delete activities to be included in the same import file.

.PARAMETER Operation
	Identifies the type of operation to be performed and is one of Add, Replace, or Delete.

	The Operation parameter only needs to be specified when modifying a	multi-valued
	attribute (State = Put) and must be either Add (to add to the list) or Delete (to
	remove from the list).

	The default is Replace and is used for all single-valued attributes.

.PARAMETER MatchAttribute
	Identifies the attribute to use when locating an object to update or delete

.PARAMETER Uri
	Identifies the URI of the FIM Service

	The default is http://localhost:5725

.PARAMETER UseCachedSchema
	Indicates that a saved copy of the schema for the object type indicated by the ObjectType
	parameter should be used instead of querying FIM for it. NOT IMPLEMENTED

.TODO
	- Include a !MatchParameter special attribute to allow the attribute used in matching for
	  Put and Delete activities to be specified on an object by object basis
	- Add schema cache. Maybe consider always using the cache if available and replacing the
	  UseCachedSchema switch with IgnoreCachedSchema and RefreshCachedSchema switches
	- Add option to test for uniqueness on create, e.g. -CheckUniqueAttribute DisplayName will
	  look for any object of the same type with a matching DisplayName. Not that the named
	  attribute must be checked against the object schema
	- Allow delete of multiple objects based on filter. Could set a delete limit.
	- Allow update of multiple objects based on filter. Probably a really bad idea!
	- Test bad data representation
	- Test multiple rows
	- Add more examples including the use of Operation when modifying multi-valued attributes
	- Add special attributes !AttributeName,!AttributeValue and !Action to allow more fine-grained
	  updating abilities
	- Thoughts on empty strings, i.e. when no data is provided for a column, e.g. ...,data1,,data2,...
	  Currently if a column entry contains no data, the attribute is ignored. However, having the
	  capability to clear an attribute would be useful, but...how to differentiate between the example
	  above meaning 'no data, do nothing' and 'no data, clear the attribute'? The Operation parameter
	  is really an attribute-level directive, so some options:
	  - Use the !Operation special column header to define how emply columns are treated on a row-by-row
	    basis. OK, but might want to clear and ignore in the same row
	  - Integrate the directive with the data, e.g. ...,data,<some directive>,data,... This provides
	    fine-grained control, but if the data is being sourced from a spreadsheet, it may make the
	    creation of the spreadsheet problematic. The directive approach could also be used for managing
	    activities on multi-valied attributes, e.g. a prefix of + directs the script to add to the MV attribute,
	    a prefix of - directs the script to remove from the MV attribute, while * directs the MV attribute
	    to be cleared, which would require the current values to be read and removed in a loop
	  - Clearing attributes and managing (specifically removing and clearing) MV attributes could be addressed
	    using the alternative header format !Anchor,!Operation,!Attribute,!Value. This approach would leave the
	    more open header format that uses attribute names to three basic activities and where an empty column
	    means that the attribute is ignored: create object, delete object, modify object. The modify activity
	    would treat all MV attributes as an Add operation. The processing pattern for the alternative header
	    format is then:
	    - Locate object via anchor
	    - If one object returned, add/remove/replace/clear !Value to/from !Attribute
	    - If zero or > 1 objects are returned, report, log and continue
	     


#>

[CmdletBinding()]
PARAM
(
    [parameter(Mandatory=$false, Position=0)] [string] $File = 'C:\FIMProject\Work\ModifyAll.txt'
   ,[parameter(Mandatory=$false)]             [string] $Delimiter = ','
   ,[parameter(Mandatory=$false)]             [string] $MVDelimiter = ';'
   ,[parameter(Mandatory=$false)]             [string] $RefDelimiter = '|'
   ,[parameter(Mandatory=$false)]             [string] $ObjectType = 'Fruit'
   ,[parameter(Mandatory=$false)]             [string] $State = 'None'
   ,[parameter(Mandatory=$false)]             [string] $Operation = 'Replace'
   ,[parameter(Mandatory=$false)]             [string] $MatchAttribute = 'None'
   ,[parameter(Mandatory=$false)]             [string] $Uri = 'http://srv.corp.contoso.com:5725'
   ,[parameter(Mandatory=$false)]             [switch] $UseCachedSchema = $false
   ,[parameter(Mandatory=$false)]             [string] $MSPSCmdletScript = 'FIMPowerShell.ps1'

)

Set-Variable SPECIAL_ATTRIBUTES_LIST -Value '!State,!Operation'
Set-Variable SPECIAL_ATTRIBUTE_STATE -Value '!State'
Set-Variable SPECIAL_ATTRIBUTE_OPERATION -Value '!Operation'

Set-Variable OPERATION_ADD -Value 'Add'
Set-Variable OPERATION_DELETE -Value 'Delete'
Set-Variable OPERATION_REPLACE -Value 'Replace'
Set-Variable OPERATION_NONE -Value 'None'


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
		$ExportObject.ResourceManagementObject.ResourceManagementAttributes | ForEach-Object { 
			if ($_.Value -ne $null) { 
				$value = $_.Value 
			} elseif($_.Values -ne $null) { 
				$value = $_.Values 
			} else { 
				$value = $null 
			} 
			$psObject | Add-Member -MemberType NoteProperty -Name $_.AttributeName -Value $value 
		} 
		$psObject 
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
    $schema  = GetSpecialObjectSchema
    $schema += GetFimObjectSchema
    $schema
}

function GetSpecialObjectSchema
{
    $templateProperties = @{DataType='String'; Multivalued='False'}
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
    
    foreach ($attribute in GetObjectAttributes) {
        $thisAttribute = $attributeTemplate.PSObject.Copy()
        $thisAttribute.DataType = $attribute.DataType
        $thisAttribute.Multivalued = $attribute.Multivalued
        $schema.Add("$($attribute.Name)", $thisAttribute)
    }
        
    $schema
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
    $rowState = GetState $row
        
    switch ($rowState.ToLower()) {
        create {CreateFIMObject $row}
        delete {DeleteFIMObject $row}
        put {ModifyFIMObject $row}
        default {throw "Unknown object state: $rowState`nUse the -State parameter to specify one of Create, Put or Delete"}
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

# C R E A T E   R E G I O N

function CreateFIMObject([PSCustomObject] $row)
{
	$object = CreateImportObject -ObjectType $ObjectType
 
    foreach ($attribute in $row.PSObject.Properties) {
        AddAttributesToObject $attribute $object
    }
    
    $object | Import-FIMConfig -Uri $Uri
}

function AddAttributesToObject($attribute, $object)
{
	$isSpecialAttribute = ($attribute.Name.StartsWith('!'))
	$isPresent = (-not [string]::IsNullOrEmpty($attribute.Value))

	if (-not $isSpecialAttribute -and $isPresent) {
		$isMultiValued = ($schemaAttributes."$($attribute.Name)".Multivalued -eq 'True')
	
		if ($isMultiValued) {
	        AddMultivaluedAttributeToObject $object $attribute.Name $attribute.Value
	    } else {
			AddSinglevaluedAttributeToObject $object $attribute.Name $attribute.Value
		}
	}
}

function AddMultivaluedAttributeToObject($object, $attributeName, $attributeValue)
{
    $isReferenceType = ($schemaAttributes."$attributeName".DataType -eq 'Reference')

	if ($isReferenceType) {
        AddMultivaluedReferenceAttributeToObject $object $attributeName $attributeValue
    } else {
        AddMultivaluedSimpleAttributeToObject $object $attributeName $attributeValue
    }
}

function AddMultivaluedReferenceAttributeToObject($ObjectType, $attributeName, $attributeValue)
{
	foreach ($mvAttributeValue in $attributeValue.Split($MVDelimiter)) {
		if (IsValidReferenceRepresentation $mvAttributeValue) {
			AddMultiValue $object $attributeName (QueryFIMResource (BuildFilter $mvAttributeValue))
		}
	}
}

function AddMultivaluedSimpleAttributeToObject($ObjectType, $attributeName, $attributeValue)
{
	foreach ($mvAttributeValue in $attributeValue.Split($MVDelimiter)) {
		AddMultiValue $object $attribute.Name $mvAttributeValue
	}
}

function AddSinglevaluedAttributeToObject($object, $attributeName, $attributeValue)
{
    $isReferenceType = ($schemaAttributes."$attributeName".DataType -eq 'Reference')

	if ($isReferenceType) {
        AddSinglevaluedReferenceAttributeToObject $object $attributeName $attributeValue
    } else {
        AddSinglevaluedSimpleAttributeToObject $object $attributeName $attributeValue
    }
}

function AddSinglevaluedReferenceAttributeToObject($object, $attributeName, $attributeValue)
{
    if (IsValidReferenceRepresentation $attributeValue) {
		SetSingleValue $object $attributeName (QueryFIMResource (BuildFilter $attributeValue))
	}
}

function IsValidReferenceRepresentation($attributeValue)
{
	$isValidRepresentation = ($attributeValue -match "^\((.+\$RefDelimiter){2}.+\)$")
    if (-not $isValidRepresentation) {
        throw "Invalid representation of a reference type: $attributeValue"
    }

	$isValidRepresentation
}

function BuildFilter($attributeValue)
{
	$referenceData = ($attributeValue -replace '\(|\)','').Split($RefDelimiter)
    $filter = "/{0}[{1}='{2}']" -f $referenceData[0], $referenceData[1], $referenceData[2]
	$filter
}

function QueryFIMResource($filter)
{
    $referencedObject = @(QueryResource $filter $Uri)
    if ($referencedObject[0] -eq $null) {
        throw "Unable to find object for reference: $attributeValue"
    } elseif ($referencedObject.Count -eq 1) {
        $referencedObject.ResourceManagementObject.ObjectIdentifier
    } else {
        throw "Unable to set reference because > 1 object satisfies the criteria: $attributeValue"
    }
}

function AddSinglevaluedSimpleAttributeToObject($object, $attributeName, $attributeValue)
{
    SetSingleValue $object $attributeName $attributeValue
}

# D E L E T E   R E G I O N

function DeleteFIMObject([PSCustomObject] $row)
{
    $objectID = GetFIMObjectID $row $ObjectType
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
        $filter = "/{0}[{1}='{2}']" -f $ObjectType, $MatchAttribute, $row.$MatchAttribute
        $result = @(QueryResource $filter $Uri)
        
        if ($result[0] -eq $null) {
            Write-Verbose([string]((Get-Date -Format u) + " WARNING: Object not found for MatchAttribute: $MatchAttribute = $($row.$MatchAttribute)"))
            $null
        } elseif ($result.Count -eq 1) {
            $result[0].ResourceManagementObject.ObjectIdentifier -replace 'urn:uuid:', ''
        } else {
            Write-Verbose([string]((Get-Date -Format u) + " WARNING: > 1 object found for MatchAttribute: $MatchAttribute = $($row.$MatchAttribute)"))
            $null
        }
    }
}

# M O D I F Y   R E G I O N

function ModifyFIMObject([PSCustomObject] $row)
{
	$objectID = GetFIMObjectID $row $ObjectType
    if ($objectID -ne $null) {
		$object = ModifyImportObject -TargetIdentifier $objectID -ObjectType $ObjectType
	    foreach ($attribute in $row.PSObject.Properties) {
			AddAttributesToObject $attribute $object
		}
    
		$object | Import-FIMConfig -Uri $Uri
	}
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