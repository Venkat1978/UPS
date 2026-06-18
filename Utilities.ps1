<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.142
	 Created on:   	7/26/2017 5:01 PM	 
	 Organization: 	
	 Filename:     	Utilities.ps1
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

function Read-Configurations
{
	Process
	{
		[xml]$global:ConfigFile = Get-Content ".\Config.xml"
        $global:VideoFieldMappings = Get-Content ".\FieldMapping.json" | ConvertFrom-Json
        
	}
}

function Get-SiteObject($splistItem)
{
	$SiteObject = New-Object -TypeName PSObject
	foreach ($column  in $global:JsonFile.Columns)
	{
		if ($splistItem.FieldValues.ContainsKey($column.ColumnInternalName))
		{
			$SiteObject | Add-Member -Type NoteProperty -Name $column.ColumnInternalName -Value $splistItem.FieldValues[$column.ColumnInternalName]
		}
	}
	return $SiteObject
}

 
function Get-PublicSiteObject($splistItem)
{
	$SiteObject = New-Object -TypeName PSObject
	foreach ($column  in $global:JsonFile.Columns)
	{
		if (![string]::IsNullOrEmpty($column.PublicSiteInventoryName))
		{
			if ($splistItem.FieldValues.ContainsKey($column.PublicSiteInventoryName))
			{
				$SiteObject | Add-Member -Type NoteProperty -Name $column.PublicSiteInventoryName -Value $splistItem.FieldValues[$column.PublicSiteInventoryName]
			}
		}
	}
	#Include All Mandatory Fields
	$SiteObject | Add-Member -Type NoteProperty -Name "ID" -Value $splistItem.ID
	$SiteObject | Add-Member -Type NoteProperty -Name "Url" -Value $splistItem.FieldValues["Url"] -Force
	$SiteObject | Add-Member -Type NoteProperty -Name "Title" -Value $splistItem.FieldValues["Title"] -Force
	return $SiteObject
}

<#
	.SYNOPSIS
		A brief description of the Get-StoredCredentials function.

	.DESCRIPTION
		A detailed description of the Get-StoredCredentials function.

	.PARAMETER  UserName
		The description of a the UserName parameter.

	.PARAMETER  Password
		The description of a the Password parameter.

	.EXAMPLE
		PS C:\> Get-StoredCredentials -UserName 'One value' -Password 'PASS'
		System.Management.Automation.PSCredential Object
		This example shows how to call the Get-StoredCredentials function with named parameters.


	.INPUTS
		System.String,System.String

	.OUTPUTS
		System.Management.Automation.PSCredential

	.NOTES
		For more information about advanced functions, call Get-Help with any
		of the topics in the links listed below.
#>

function Get-StoredCredentials
{
	[CmdletBinding()]
	param (
		[Parameter(Position = 0, Mandatory = $true, ParameterSetName = "UserPass")]
		[System.String]$UserName,
		[Parameter(Position = 1, Mandatory = $true, ParameterSetName = "UserPass")]
		[System.String]$Password,
		[Parameter(Position = 0, Mandatory = $true, ParameterSetName = "EncryptedPasswordFilePath")]
		[System.String]$EncryptedPasswordFilePath
	)
	
	switch ($PsCmdlet.ParameterSetName)
	{
		"UserPass"
		{
			$cred = [System.Net.CredentialCache]::DefaultCredentials
			[System.Net.WebRequest]::DefaultWebProxy.Credentials = $cred
			$creds = New-Object System.Management.Automation.PSCredential($Username, (ConvertTo-SecureString $Password -AsPlainText -Force));
			return $creds
		}
		"EncryptedPasswordFilePath"
		{
			GetCredentials $EncryptedPasswordFilePath
		}
	}
	
}

#Adding Site object with SiteUrl as a Key
function Add-SitePropertyObject ($siteUrl, $object)
{
	if (!$global:SitePropertiesCache.ContainsKey($siteUrl))
	{
		$global:SitePropertiesCache.Add($siteUrl, $object)
	}
}

function Get-SitePropertyObject ($siteUrl)
{
	if ($global:SitePropertiesCache.ContainsKey($siteUrl))
	{
		return $global:SitePropertiesCache.Get_Item($siteUrl)
	}
	throw New-Object System.Exception "can't find site with Url " $siteUrl
}

function clone($obj)
{
    $newobj = New-Object PsObject
    $obj.psobject.Properties | % {Add-Member -MemberType NoteProperty -InputObject $newobj -Name $_.Name -Value $_.Value}
    return $newobj
}
#Adding Deleted Site object with SiteUrl as a Key
function Add-DeletedSitePropertyObject ($siteUrl, $object)
{
	if (!$global:DeletedSitesCache.ContainsKey($siteUrl))
	{
		$global:DeletedSitesCache.Add($siteUrl, $object)
	}
}

function Get-DeletedSitePropertyObject ($siteUrl)
{
	return $global:DeletedSitesCache.Get_Item($siteUrl)
}

function GetCredentials($securePwFile)
{
	# Check if credentials are already existing, if not you will be asked to provide them
	if (!(Test-Path -Path $securePwFile -PathType Any))
	{
		Log -Message "*******Can't find File  $($securePwFile), provide UserName and password to be created.******" -LogLevel "Info"
		ExportSecureCredentials $securePwFile;
	}
	$cred = ImportSecureCredentials $securePwFile;
	return $cred
}

# Provide Credentials (Username and Password), they'll be stored into the given file - with an encrypted password
function ExportSecureCredentials($securePwFile)
{
	$credreadonly = get-credential
	$cred = $credreadonly | select-object *
	$cred.password = $cred.password | ConvertFrom-SecureString
	$cred | export-clixml $securePwFile
}

# Load the credentials from a given file and create a PSCredential Object
function ImportSecureCredentials()
{
	$account = import-clixml $securePwFile
	$account.password = $account.Password | ConvertTo-SecureString
	$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $account.username, $account.password
	return $cred
}


Function TestFolderExists
{
	<#
		.SYNOPSIS
			Verifies that the specified folder/path exists.

		.DESCRIPTION
			Verifies that the specified folder/path exists.

		.PARAMETER Folder
			Specifies the absolute or relative path to the file.

		.EXAMPLE
			PS> TestFolderExists -Folder "C:\Folder\Sub Folder\File name.csv"

		.EXAMPLE
			PS> TestFolderExists -Folder "File name.csv"

		.EXAMPLE
			PS> TestFolderExists -Folder "C:\Folder\Sub Folder"

		.EXAMPLE
			PS> TestFolderExists -Folder ".\Folder\Sub Folder"

		.INPUTS
			System.String

		.OUTPUTS
			System.Boolean

		.NOTES

	#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True)]
		[string]$Folder
	)
	
	If ([System.IO.Path]::HasExtension($Folder)) { $PathToFile = ([System.IO.Directory]::GetParent($Folder)).FullName }
	Else { $PathToFile = [System.IO.Path]::GetFullPath($Folder) }
	If ([System.IO.Directory]::Exists($PathToFile)) { Return $True }
	Return $False
}


Function GetElapsedTime
{
	<#
		.SYNOPSIS
			Calculates a time interval between two DateTime objects.

		.DESCRIPTION
			Calculates a time interval between two DateTime objects.

		.PARAMETER Start
			Specifies the start time.

		.PARAMETER End
			Specifies the end time.

		.EXAMPLE
			PS> GetElapsedTime -Start "1/1/2011 12:00:00 AM" -End "1/2/2011 2:00:00 PM"

		.EXAMPLE
			PS> GetElapsedTime -Start ([datetime]"1/1/2011 12:00:00 AM") -End ([datetime]"1/2/2011 2:00:00 PM")

		.INPUTS
			System.String

		.OUTPUTS
			System.Management.Automation.PSObject

		.NOTES

	#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[DateTime]$Start,
		[Parameter(Mandatory = $True, Position = 1)]
		[DateTime]$End
	)
	
	$TotalSeconds = ($End).Subtract($Start).TotalSeconds
	$objElapsedTime = New-Object PSObject
	
	# less than 1 minute
	If ($TotalSeconds -lt 60)
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $($TotalSeconds)
	}
	
	# more than 1 minute, less than 1 hour
	If (($TotalSeconds -ge 60) -and ($TotalSeconds -lt 3600))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate($TotalSeconds / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}
	
	# more than 1 hour, less than 1 day
	If (($TotalSeconds -ge 3600) -and ($TotalSeconds -lt 86400))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value 0
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value $([Math]::Truncate($TotalSeconds / 3600))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate(($TotalSeconds % 3600) / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}
	
	# more than 1 day, less than 1 year
	If (($TotalSeconds -ge 86400) -and ($TotalSeconds -lt 31536000))
	{
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Days -Value $([Math]::Truncate($TotalSeconds / 86400))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Hours -Value $([Math]::Truncate(($TotalSeconds % 86400) / 3600))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Minutes -Value $([Math]::Truncate((($TotalSeconds - 86400) % 3600) / 60))
		Add-Member -InputObject $objElapsedTime -MemberType NoteProperty -Name Seconds -Value $([Math]::Truncate($TotalSeconds % 60))
	}
	
	Return $objElapsedTime
}


function Get-AzureADUserByID($userID)
{
    return Get-AzureADUser -ObjectId $userID
}


function Get-UnifiedGroupByID($groupID)
{
    return Get-AzureADGroup -ObjectId $groupID -ErrorAction Continue

    #return Get-UnifiedGroup -Identity $groupID 
}




