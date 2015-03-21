# Copyright 2010 Microsoft Corporation

if(@(get-pssnapin | where-object {$_.Name -eq "FIMAutomation"} ).count -eq 0) {add-pssnapin FIMAutomation}

# States
# 0 = Create
# 1 = Put
# 2 = Delete
# 3 = Resolve
# 4 = None

# Operations
# 0 = Add
# 1 = Replace
# 2 = Delete

# Low-level operations

$DefaultUri = "http://srv.corp.contoso.com:5725"

function CreateImportObject
{
    PARAM([string]$ObjectType)
    END
    {
        $importObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
        $importObject.SourceObjectIdentifier = [System.Guid]::NewGuid().ToString()
        $importObject.ObjectType = $ObjectType
        $importObject
    }
}

function ModifyImportObject
{
    PARAM([string]$TargetIdentifier, $ObjectType = "Resource")
    END
    {
        $importObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
        $importObject.ObjectType = $ObjectType
        $importObject.TargetObjectIdentifier = $TargetIdentifier
        $importObject.SourceObjectIdentifier = $TargetIdentifier
        $importObject.State = 1 # Put
        $importObject
    }
}

function DeleteObject
{
    PARAM([string]$TargetIdentifier, $ObjectType = "Resource")
    END
    {
        $importObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
        $importObject.ObjectType = $ObjectType
        $importObject.TargetObjectIdentifier = $TargetIdentifier
        $importObject.SourceObjectIdentifier = $TargetIdentifier
        $importObject.State = 2 # Delete
        $importObject
    }
}

function ResolveObject
{
    PARAM([string] $ObjectType, [string]$AttributeName, [string]$AttributeValue)
    END
    {
        $importObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportObject
        $importObject.TargetObjectIdentifier = $TargetIdentifier
        $importObject.ObjectType = $ObjectType
        $importObject.State = 3 # Resolve
        $importObject.SourceObjectIdentifier = [System.String]::Format("urn:uuid:{0}", [System.Guid]::NewGuid().ToString())
        $importObject.AnchorPairs = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.JoinPair
        $importObject.AnchorPairs[0].AttributeName = $AttributeName
        $importObject.AnchorPairs[0].AttributeValue = $AttributeValue
        $importObject
    }
}

function SetSingleValue
{
    PARAM($ImportObject, $AttributeName, $NewAttributeValue, $FullyResolved=1)
    END
    {
        $ImportChange = CreateImportChange -AttributeName $AttributeName -AttributeValue $NewAttributeValue -Operation 1
        $ImportChange.FullyResolved = $FullyResolved
        AddImportChangeToImportObject $ImportChange $ImportObject
    }
}

function AddMultiValue
{
    PARAM($ImportObject, $AttributeName, $NewAttributeValue, $FullyResolved=1)
    END
    {
        $ImportChange = CreateImportChange -AttributeName $AttributeName -AttributeValue $NewAttributeValue -Operation 0
        $ImportChange.FullyResolved = $FullyResolved
        AddImportChangeToImportObject $ImportChange $ImportObject
    }
}

function RemoveMultiValue
{
    PARAM($ImportObject, $AttributeName, $NewAttributeValue, $FullyResolved=1)
    END
    {
        $ImportChange = CreateImportChange -AttributeName $AttributeName -AttributeValue $NewAttributeValue -Operation 2
        $ImportChange.FullyResolved = $FullyResolved
        AddImportChangeToImportObject $ImportChange $ImportObject
    }
}

function AddImportChangeToImportObject
{
    PARAM($ImportChange, $ImportObject)
    END
    {
        if ($ImportObject.Changes -eq $null)
        {
            $ImportObject.Changes = (,$ImportChange)
        }
        else
        {
            $ImportObject.Changes += $ImportChange
        }
    }
}

function CreateImportChange
{
    PARAM($AttributeName, $AttributeValue, $Operation)
    END
    {
        $importChange = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ImportChange
        $importChange.Operation = $Operation
        $importChange.AttributeName = $AttributeName

		if ($AttributeValue) {
			$importChange.AttributeValue = $AttributeValue
		}
        
		$importChange.FullyResolved = 1
        $importChange.Locale = "Invariant"
        $importChange
    }
}

function GetSidAsBase64
{
    PARAM($AccountName, $Domain)
    END
    {
        $sidArray = [System.Convert]::FromBase64String("AQUAAAAAAAUVAAAA71I1JzEyxT2s9UYraQQAAA==") # This sid is a random value to allocate the byte array.
        $args = (,$Domain)
        $args += $AccountName
        $ntaccount = New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList $args
        $desiredSid = $ntaccount.Translate([System.Security.Principal.SecurityIdentifier])
        $desiredSid.GetBinaryForm($sidArray,0)
        $desiredSidString = [System.Convert]::ToBase64String($sidArray)
        $desiredSidString
    }
}

# Diagnostic operations

function ConvertResourceToHashtable
{
    PARAM([Microsoft.ResourceManagement.Automation.ObjectModel.ExportObject]$ExportObject)
    END
    {
        $hashtable = @{"ObjectID" = "Not found"}
        foreach($attribute in $exportObject.ResourceManagementObject.ResourceManagementAttributes)
        {
            if ($attribute.IsMultiValue -eq 1)
            {
                $hashtable[$attribute.AttributeName] = $attribute.Values
            }
            else
            {
                $hashtable[$attribute.AttributeName] = $attribute.Value
            }
        }
        $hashtable
    }
}

function ConvertHashtableToResource
{
    PARAM($Hashtable)
    END
    {
        $ExportObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ExportObject
        $ExportObject.Source = $DefaultUri
        $ExportObject.ResourceManagementObject = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ResourceManagementObject
        foreach($key in $Hashtable.Keys)
        {
            $value = $Hashtable[$key]
            $newAttribute = New-Object Microsoft.ResourceManagement.Automation.ObjectModel.ResourceManagementAttribute
            $newAttribute.AttributeName = $key
            $newAttribute.Value = $value
            $newAttribute.HasReference = 0
            $newAttribute.IsMultiValue = 0
            
            if($newAttribute.AttributeName -eq "ObjectID")
            {
                $ExportObject.ResourceManagementObject.ObjectIdentifier = $newAttribute.Value
            }
            
            if($newAttribute.AttributeName -eq "ObjectType")
            {
                $ExportObject.ResourceManagementObject.ObjectType = $newAttribute.Value
            }
            
            $ExportObject.ResourceManagementObject.IsPlaceholder = 0
            
            if($ExportObject.ResourceManagementObject.ResourceManagementAttributes -eq $null)
            {
                $ExportObject.ResourceManagementObject.ResourceManagementAttributes = (,$newAttribute)
            }
            else
            {
                $ExportObject.ResourceManagementObject.ResourceManagementAttributes += $newAttribute
            }
        }
        $ExportObject
    }
}

function GetResource
{
    PARAM($ObjectIdentifier, $Uri = $DefaultUri)
    END
    {
        $object = Export-FIMConfig -CustomConfig [System.String]::Format("*[ObjectID='{0}']", $ObjectIdentifier) -Uri $Uri
        if($object -eq $null)
        {
            Write-Host "Object was not found."
        }
        else
        {
            $object
        }
    }
}

function QueryResource
{
    PARAM($Filter, $Uri = $DefaultUri)
    END
    {
        $resources = Export-FIMConfig -OnlyBaseResources -CustomConfig $Filter -Uri $Uri
        $resources
    }
}

function QueryResourceAll
{
    PARAM($Filter, $Uri = $DefaultUri)
    END
    {
        $resources = Export-FIMConfig -CustomConfig $Filter -Uri $Uri
        $resources
    }
}

# High-level operations

function AddMembersToGroup
{
    PARAM($GroupIdentifier, $PersonIdentifiers, $IdentifierName="Email", $Uri = $DefaultUri)
    END
    {
        $ResolveGroup = ResolveObject -ObjectType "Group" -AttributeName $IdentifierName -AttributeValue $GroupIdentifier
        $ResolveGroup | Import-FIMConfig -Uri $Uri
        $ImportObjects = $NULL
        $AddedMembers = $NULL
        foreach($PersonIdentifier in $PersonIdentifiers)
        {
            $ImportObject = ResolveObject -ObjectType "Person" -AttributeName $IdentifierName -AttributeValue $PersonIdentifier
            if($AddedMembers -eq $NULL)
            {
                $AddedMembers = @($ImportObject.SourceObjectIdentifier)
            }
            else
            {
                $AddedMembers += $ImportObject.SourceObjectIdentifier
            }
            if($ImportObjects -eq $NULL)
            {
                $ImportObjects = @($ImportObject)
            }
            else
            {
                $ImportObjects += $ImportObject
            }
        }
        
        $ModifyImportObject = ModifyImportObject -TargetIdentifier $ResolveGroup.TargetObjectIdentifier -ObjectType "Group"
        $ModifyImportObject.SourceObjectIdentifier = $ResolveGroup.SourceObjectIdentifier
        
        foreach($AddedMember in $AddedMembers)
        {
            $newValue = $AddedMember
            AddMultiValue -ImportObject $ModifyImportObject -AttributeName "ExplicitMember" -NewAttributeValue $newValue -FullyResolved 0
            #RemoveMultiValue -ImportObject $ModifyImportObject -AttributeName "ExplicitMember" -NewAttributeValue $newValue -FullyResolved 0
        }
        $ImportObjects += $ModifyImportObject
        #$ImportObjects | Import-FIMConfig -Uri $Uri
        $ImportObjects
    }
}

function EnableAllMPRs
{
    PARAM($Uri = $DefaultUri)
    END
    {
        $AllMPRs = QueryResource -Filter "/ManagementPolicyRule[Disabled='true']" -Uri $Uri
        $ImportObjects = $null
        foreach($mpr in $AllMPRs)
        {
            $ModifyImportObject = ModifyImportObject -TargetIdentifier $mpr.ResourceManagementObject.ObjectIdentifier -ObjectType "ManagementPolicyRule"
            SetSingleValue $ModifyImportObject "Disabled" "false"
            if($ImportObjects -eq $null)
            {
                $ImportObjects = (,$ModifyImportObject)
            }
            else
            {
                $ImportObjects += $ModifyImportObject
            }
            
        }
        if($ImportObjects -ne $null)
        {
            $ImportObjects | Import-FIMConfig -Uri $Uri
        }
    }
}

function CreatePerson
{
    PARAM($FirstName, $LastName, $AccountName, $Domain, $Email, $Uri = $DefaultUri)
    END
    {
        $DisplayName = [System.String]::Format("{0} {1}", $FirstName, $LastName)
        
        $NewPerson = CreateImportObject -ObjectType "Person"
        SetSingleValue $NewPerson "FirstName" $FirstName
        SetSingleValue $NewPerson "LastName" $LastName
        SetSingleValue $NewPerson "DisplayName" $DisplayName
        SetSingleValue $NewPerson "AccountName" $AccountName
        SetSingleValue $NewPerson "Domain" $Domain
        SetSingleValue $NewPerson "Email" $Email
        
        $UserSid = GetSidAsBase64 -AccountName $AccountName -Domain $Domain
        SetSingleValue $NewPerson "ObjectSID" $UserSid
        
        $NewPerson | Import-FIMConfig -Uri $Uri
    }
}

function CreateGroup
{
    PARAM($DisplayName, $Owner, $AccountName, $Domain, $Email, $GroupType, $GroupScope, $Uri = $DefaultUri)
    END
    {
        $NewGroup = CreateImportObject -ObjectType "Group"
        SetSingleValue $NewGroup "DisplayName" $DisplayName
        SetSingleValue $NewGroup "AccountName" $AccountName
        SetSingleValue $NewGroup "Domain" $Domain
        if($Email -ne $null)
        {
            SetSingleValue $NewGroup "Email" $Email
        }
        if($GroupScope -ne $null)
        {
            SetSingleValue $NewGroup "Scope" $GroupScope
        }
        SetSingleValue $NewGroup "Type" $GroupType
        
        $ResolveOwner = ResolveObject -ObjectType "Person" -AttributeName "Email" -AttributeValue $Owner
        SetSingleValue -ImportObject $NewGroup -AttributeName "Owner" -AttributeValue $ResolveOwner.SourceObjectIdentifier -FullyResolved 0
        
        $ImportObjects = (,$ResolveOwner)
        $ImportObjects += $NewGroup
        $ImportObjects
        $ImportObjects | Import-FIMConfig -Uri $Uri
    }
}

# Value values for scope:
# Universal
# Global
# Domain
function CreateSecurityGroup
{
    PARAM($DisplayName, $Owner, $AccountName, $Domain, $Email, $GroupScope, $Uri = $DefaultUri)
    END
    {
        CreateGroup -DisplayName $DisplayName -Owner $Owner -AccountName $AccountName -Domain $Domain -Email $Email -GroupScope $GroupScope -GroupType "SecurityGroup" -Uri $Uri
    }
}

function CreateDistributionGroup
{
    PARAM($DisplayName, $Owner, $AccountName, $Domain, $Email, $Uri = $DefaultUri)
    END
    {
        CreateGroup -DisplayName $DisplayName -Owner $Owner -AccountName $AccountName -Domain $Domain -Email $Email -GroupType "Distribution" -Uri $Uri    
    }
}

function AllRequestsToElevatedSecurityGroup
{
    END
    {
        $resources = QueryResource -Filter "/Request[Creator=/Group[DisplayName='Elevated Access Security Group']/ComputedMember and Operation='Put' ' and Status='Completed']"
        $resources
    }
}
