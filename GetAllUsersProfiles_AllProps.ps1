### ----------------------------------------------------------------------------------------------------------- ###
###                                                                                                             ###
### Purpose: This script is used to iterate through all the SharePoint Online User Profiles and bring the       ###
###          number of users using each Hk Hub Property Value, e.g. 40 users with 'Hk News' as Sender.  ###
###          The output will be saved in a SharePoint library                                                   ###
###                                                                                                             ###
### Date:   November, 2020                                                                                      ###
### Version: 1.7.0.0                                                                                            ###
###                                                                                                             ###
### Required Modules:                                                                                           ###
###        - SharePoint PNP module (Install-Module SharePointPnPPowerShellOnline)                               ###
###        - Azure AD Module (Install-Module AzureAD)                                                           ###
###                                                                                                             ###
###                                                                                                             ###
### Release Notes:                                                                                              ###
###        - 1.1.0.0: Report also for the two properties: Hk-Region and Hk-Country                      ###
###        - 1.2.0.0: Report to include the two properties: Hk-Region and Hk-Country                    ###
###                   as part of Sender/Topic                                                                   ###
###        - 1.5.0.0: Report also includes number of users completed the wizard                                 ###
###        - 1.6.0.0: Report also includes Custom Bookmarks Statistics                                          ###
###        - 1.7.0.0: Fixing throttling while calling GetUserProfileByName                                      ###
###        - 1.8.0.0: Report also includes App Bookmarks Statistics                                             ###
###        - 1.9.0.0: Store statistics in welcome wizard with the below updates                                 ###
###                    - Store records in folders: Public and Private                                           ###
###                    - Store public items with term GUID                                                      ###
###        - 1.9.0.1: Fix getting Terms                                                                         ###
###                    - Removing TermStore from command, as MS keeps changing it :)                            ###
###                    - adding more exception handling                                                         ###
### ----------------------------------------------------------------------------------------------------------- ###
### ----------------------------------------------------------------------------------------------------------- ###
### ----------------------------------------------------------------------------------------------------------- ###
### ----------------------------------------------------------------------------------------------------------- ###


################ Variables #################
#$AdminSiteUrl = "https://Hkgroup-admin.sharepoint.com"

$AdminSiteUrl = "https://Hkgroup-admin.sharepoint.com"
$CSVPath = ".\output\UserProfiles.csv"
$usersFilePath = ".\output\AllUsers.txt"
$outputFilePath = ".\output\Report.txt"
$errorLogFile = ".\output\Error.log"
$logFile = ".\output\Log.log"
$ConfigFile = [xml](Get-Content "C:\Projects\PSJobs\Config.xml")
$siteCatalogConfig = $ConfigFile.Configuration.appSettings.SiteCatalog

#To fix the TLS error "Connect-PnPOnline : The underlying connection was closed: An unexpected error occurred on a send."
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

#$Proxy = New-Object System.Net.WebProxy http://ibm-proxy-de.Hkgroup.net:9090
#$Proxy.UseDefaultCredentials=$true
#[VOID]$Proxy.BypassArrayList.Add("[a-z]+\.Hkgroup\.net$")
#[System.Net.WebRequest]::DefaultWebProxy=$Proxy 


#SharePoint details
#$spSiteUrl = "https://Hkgroup.sharepoint.com/sites/smarthub"
#$spSiteUrl = "https://Hkgroup.sharepoint.com/teams/CCSPplay/NewsHub"

$spSiteUrl = "https://Hkgroup.sharepoint.com/sites/Welcome-Wizard"
$spListTitle = "Hk HUB Users Statistics"
$spUserProfileEndPoint = "_vti_bin/UserProfileService.asmx?wsdl"
#$termStore = "Taxonomy_ZvPF2vvY20KHY2X6i4IzJg=="
$termGroup = "Hk Hub"

$spList_PropNameColumn = "HH_TitleType"
$spList_PropValueColumn = "Title"
$spList_SubPropValueColumn = "HH_Location"
$spList_UsersCountColumn = "HH_UsersCount"
$spList_TermGUID = "HH_TermGUID"
$spList_PublicFolderName = "Public"
$spList_PrivateFolderName = "Private"

C:\Projects\PSJobs\Utilities.ps1

$uPSField_WelcomeWizardCompleted = "Hk-DoesWizardCompleted"
$uPSField_CustomBookmarks = "Hk-CustomBookmarks"
$uPSField_AppsBookmarks = "Hk-AppBookmarks"


$uPSFields = New-Object System.Collections.ArrayList
$sender = @{Name = "Hk-Sender"; DisplayName = "Sender"; Level = "Parent"; OutputListTitle = "$spListTitle"}
$topic = @{Name = "Hk-Topic"; DisplayName = "Topic"; Level = "Parent"; OutputListTitle = "$spListTitle"}
$country = @{Name = "Hk-Country"; DisplayName = "Country"; Level = "Child"; OutputListTitle = "$spListTitle"}
$region = @{Name = "Hk-Region"; DisplayName = "Region"; Level = "Child"; OutputListTitle = "$spListTitle"}
$completedWizardSPListTitle = "$spListTitle" #Used to store the completed wizard users count value

$index = $uPSFields.Add($sender)
$index = $uPSFields.Add($topic)
$index = $uPSFields.Add($country)
$index = $uPSFields.Add($region)

$uPSPropertiesValueCollections = @{}      #This is a collection of collections in the below form
<#
    @{
        'Sender' = @{'News\Germany'= '5' 'Adhesive\Europe'= '20'}
        'Topic' = @{'Strategy\Egypt'= '2' 'People\Europe'='13'}
    }
#> 

$completedWizardUsersCount = 0
$customBookmarkingUsersCount = 0 #The number of users using the custom bookmarking feature
$appsBookmarkingUsersCount = 0 #The number of users using the apps bookmarking feature

################ End of Variables #################

 
$Cred = Get-StoredCredentials -EncryptedPasswordFilePath $siteCatalogConfig.SecuredPwFile
Write-Output "Connecting to SharePoint site '$spSiteUrl'..."  | Out-File $logFile

#Adding the connection variable, because after some time the queries fails because of losing connection
$currentConnection = $null
try
{
    #Connect-PnPOnline $AdminSiteUrl -Credentials $Cred
    #$currentConnection = Connect-PnPOnline -Url $spSiteUrl -Credentials $Cred -ReturnConnection
    Connect-PnPOnline -Url $spSiteUrl -Credentials $Cred -ClientId 6fcef744-d7d9-413c-be6b-f14b8af069b1
    $currentConnection = Get-PnPConnection
    Write-Output "Connected successfully to the SharePoint Online site" | Out-File $logFile -Append
}
catch
{
    Write-Output "An error has occurred while connecting to the SharePoint Online site. Aborting..."  | Out-File $errorLogFile
    Write-Output ($_ | Out-String)  | Out-File $errorLogFile -Append
    return
}

try
{
    # Connecting to User Profile Service
    Write-Output "Connecting to SharePoint Online User Profile Service '$AdminSiteUrl/$spUserProfileEndPoint'..."  | Out-File $logFile -Append
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Cred.UserName, $Cred.Password)
	
    $upService = New-WebServiceProxy -Uri ("$AdminSiteUrl/$spUserProfileEndPoint") -UseDefaultCredential False
    $upService.Credentials = $credentials
    $uri = New-Object System.Uri($AdminSiteUrl)
    $container = New-Object System.Net.CookieContainer
    $container.SetCookies($uri, $credentials.GetAuthenticationCookie($AdminSiteUrl))
    $upService.CookieContainer = $container
	
    Write-Output "Connected Successfully to SharePoint Online User Profile Service"  | Out-File $logFile -Append
}
catch
{
    Write-Output "An error has occurred while connecting to the SharePoint User Profile Service. Aborting..."  | Out-File $errorLogFile
    Write-Output ($_ | Out-String)  | Out-File $errorLogFile -Append
    return
}


try
{
    #Connect to AzureAD
    Write-Output "Connecting to Azure AD..." | Out-File $logFile -Append
    Connect-AzureAD -Credential $Cred
    Write-Output "Connected successfully to Azure AD" | Out-File $logFile -Append
}
catch
{
    Write-Output "An error has occurred while connecting to Azure AD. Aborting..."  | Out-File $errorLogFile
    Write-Output ($_ | Out-String) | Out-File $errorLogFile -Append
    return
}

try
{
    #Get All Users from AzureAD
    Write-Output "Getting all users from Azure AD..." | Out-File $logFile -Append
    
++

= Get-AzureADUser -Filter "userType eq 'Member' and AccountEnabled eq true" -All:$True | select UserPrincipalName

    #$AllUsers = Get-AzureADUser -ObjectId ahmed.amin.mansour@Hk.com
    Write-Output "Got successfully all tenant enabled users." | Out-File $logFile -Append
}
catch
{
    Write-Output "An error has occurred while retriving all tenant users. Aborting..."  | Out-File $errorLogFile
    Write-Output ($_ | Out-String)  | Out-File $errorLogFile -Append | Out-File $errorLogFile -Append
    return
}

#Used to retrieve(return) the matched item from the SharePoint Hk Intranet properties lists
function GetUPSProperty-ListItem($currentPropName,$currentPropValue,$currentListTitle)
{
    $spQuery = ""

    if($currentPropValue.split('\').count -gt 1)
    {
        $currentSPPropItem =$null
        $currentPropValue = $currentPropValue.split('\')
        $currentMainPropValue = $currentPropValue[0]
        $currentSubPropValue = $currentPropValue[1]
        $spQuery = "<View Scope='RecursiveAll'><Query><Where>
                    <And>
                        <Eq><FieldRef Name='$spList_PropNameColumn'/><Value Type='Text'>$currentPropName</Value></Eq>
                        <And>
                            <Eq><FieldRef Name='$spList_PropValueColumn'/><Value Type='Text'>$currentMainPropValue</Value></Eq>
                            <Eq><FieldRef Name='$spList_SubPropValueColumn'/><Value Type='Text'>$currentSubPropValue</Value></Eq>
                        </And>
                    </And>
                </Where></Query></View>"
    }
    else
    {
        $spQuery = "<View Scope='RecursiveAll'><Query><Where>
                            <And>
                                <Eq><FieldRef Name='$spList_PropNameColumn'/><Value Type='Text'>$currentPropName</Value></Eq>
                                <And>
                                    <Eq><FieldRef Name='$spList_PropValueColumn'/><Value Type='Text'>$currentPropValue</Value></Eq>
                                    <IsNull><FieldRef Name='$spList_SubPropValueColumn'/></IsNull>
                                </And>
                            </And>
                        </Where></Query></View>"
    }


    $trialsCount = 0
    do
    {
        $trialsCount++
        $errorOccured = $false
        try
        {
            #This try catch blocks are needed because occassionally the get-pnplistitem and set-pnplistitem are failing
            $currentSPPropItem = Get-PnPListItem -List $currentListTitle -Query $spQuery -Connection $currentConnection -ErrorAction Stop
        }
        catch
        {
            $errorOccured = $true
            Write-Output "An error has occurred while getting the SharePoint Items. Hence, retrying...($trialsCount)" | Out-File $errorLogFile -Append
            Write-Output $_ | Out-File $errorLogFile -Append
            Start-Sleep -Seconds 120
            if($trialsCount -eq 6)
            {
                #Write-Host "ERROR! Failed to get the SharePoint Items for $currentPropName - $currentPropValue" -ForegroundColor Red
                Write-Output "ERROR! Failed to get the SharePoint Items for $currentPropName - $currentPropValue" | Out-File $errorLogFile -Append
            }
        }
    }while($errorOccured -and $trialsCount -lt 6)

    return $currentSPPropItem
}

#Used to return the list title based on the current UPS property index. Whereas, the index of the property is the same as the list title index
function GetUPSPropertyListTitle($currentProcessedPropValue)
{
    $uPSFields | %{if($_.DisplayName -eq $currentProcessedPropValue){$_.OutputListTitle}}
}

#Used to get the term GUID based on the provided values: Term Set Name, and Term Name
function Get-TermId($TermSetName, $TermName)
{
    $term = $null
    if($TermSetName -eq "Sender")
    {
        $term = $senderTerms | ?{$_.Name -eq $TermName}
    }
    elseif($TermSetName -eq "Topic")
    {
        $term = $topicTerms| ?{$_.Name -eq $TermName}
    }
    else
    {
       #TermSet is neither Sender nor Topic! 
    }
    
    if($term)
    {
        Write-Output "Term '$TermName' found in TermSet: $TermSetName"  | Out-File $logFile -Append
        return $term[0].Id
    }
    else
    {
        Write-Output "Term '$TermName' NOT found in TermSet: $TermSetName"  | Out-File $logFile -Append
        return ""
    }
}

#Used to update $uPSPropertiesValueCollections based on the provided property value
function ProcessPropertyValue($currentUPSPropertiesValueCollections,$uPSFieldDisplayName,$userRegion,$userCountry,$fieldValue)
{
    if($currentUPSPropertiesValueCollections -eq $null -or $currentUPSPropertiesValueCollections[$uPSFieldDisplayName] -eq $null)
    {
        #This property doesn't exist in the collection
        $currentUPSPropertiesValueCollections.Add($uPSFieldDisplayName,@{$fieldValue=1})
        if(-not [string]::IsNullOrEmpty($userRegion))
        {
            $currentUPSPropertiesValueCollections[$uPSFieldDisplayName].Add("$fieldValue\$userRegion",1)
        }

        if(-not [string]::IsNullOrEmpty($userCountry))
        {
            $currentUPSPropertiesValueCollections[$uPSFieldDisplayName].Add("$fieldValue\$userCountry",1)   
        }
    }
    elseif($currentUPSPropertiesValueCollections["$uPSFieldDisplayName"][$fieldValue] -eq $null)
    {
        #The collection contains the property (Sender, Topic, etc.), but not the value (Hk News, Substainability, etc.)
        $currentUPSPropertiesValueCollections[$uPSFieldDisplayName].Add($fieldValue,1)
        if(-not [string]::IsNullOrEmpty($userRegion))
        {
            $currentUPSPropertiesValueCollections[$uPSFieldDisplayName].Add("$fieldValue\$userRegion",1)
        }
        if(-not [string]::IsNullOrEmpty($userCountry))
        {
            $currentUPSPropertiesValueCollections[$uPSFieldDisplayName].Add("$fieldValue\$userCountry",1)
        }

    }
    elseif($currentUPSPropertiesValueCollections["$uPSFieldDisplayName"][$fieldValue])
    {
        
        #The collection contains the property (Sender, Topic, etc.), and the value (Hk News, Substainability, etc.). Hence, incrementing 1
        $currentUPSPropertiesValueCollections[$uPSFieldDisplayName][$fieldValue]++

        if(-not [string]::IsNullOrEmpty($userRegion))
        {
            if($currentUPSPropertiesValueCollections["$uPSFieldDisplayName"]["$fieldValue\$userRegion"] -eq $null)
            {
                $currentUPSPropertiesValueCollections[$uPSFieldDisplayName].Add("$fieldValue\$userRegion",1)
            }
            else
            {
                $currentUPSPropertiesValueCollections[$uPSFieldDisplayName]["$fieldValue\$userRegion"]++
            }
        }
        if(-not [string]::IsNullOrEmpty($userCountry))
        {
            if($currentUPSPropertiesValueCollections["$uPSFieldDisplayName"]["$fieldValue\$userCountry"] -eq $null)
            {
                $currentUPSPropertiesValueCollections["$uPSFieldDisplayName"].Add("$fieldValue\$userCountry",1)
            }
            else
            {
                $currentUPSPropertiesValueCollections[$uPSFieldDisplayName]["$fieldValue\$userCountry"]++
            }
        }
    }

    return $currentUPSPropertiesValueCollections
}

#Used to update the SharePoint list item with the completed wizard users count, Custom Bookmarking users count, or Apps Bookmarking Users count
function UpdatePropertyUsersCount($spPropName, $spPropValue, $updatedSPUsersCount)
{
    $desiredSPListItem = GetUPSProperty-ListItem -currentListTitle $completedWizardSPListTitle -currentPropName $spPropName -currentPropValue $spPropValue

    if($desiredSPListItem)
    {
        $item = Set-PnPListItem -Identity $desiredSPListItem.Id -List $completedWizardSPListTitle -Values @{"$spList_UsersCountColumn"="$updatedSPUsersCount"} -Connection $currentConnection
    }
    else
    {
        #No completed wizard item is created in the SharePoint list yet
        $item = Add-PnPListItem -List $completedWizardSPListTitle -Folder $spList_PrivateFolderName -Values @{"$spList_PropNameColumn" = $spPropName;"$spList_PropValueColumn" = $spPropValue;"$spList_UsersCountColumn" = "$updatedSPUsersCount"} -Connection $currentConnection 
    }
}

Write-Output $AllUsers | out-file $usersFilePath
Write-Output "Property Name,Property Value,Users Count" | out-file $outputFilePath
Write-Output "********* ERROR MESSAGES *********" | Out-File $errorLogFile

$allFieldsNames = $uPSFields | %{$_.Name}
$allFieldsNames = $allFieldsNames -join ';'
Write-Output "User Email;$allFieldsNames;Custom Bookmarks;App Bookmarks" | Out-File $CSVPath

$allUsersCount = $AllUsers.Count

Write-Output "Total Number of Enabled Users Found:"$allUsersCount  | Out-File $logFile -Append
 
$Counter = 0

Write-Output "Geting Sender terms"  | Out-File $logFile -Append

#$senderTerms = Get-PnPTerm -TermStore $termStore -TermGroup $termGroup -TermSet "Sender" -Recursive -IncludeChildTerms | %{if($_.TermsCount -gt 0){$_;$_.Terms}else{$_}}
try
{
    $senderTerms = Get-PnPTerm -TermGroup $termGroup -TermSet "Sender" -Recursive -IncludeChildTerms -ErrorAction Stop | %{if($_.TermsCount -gt 0){$_;$_.Terms}else{$_}}
}
catch
{
    Write-Output "An error has occurred while getting Sender Terms"  | Out-File $errorLogFile
    Write-Output ($_ | Out-String)  | Out-File $errorLogFile -Append
}

Write-Output "Geting Topic terms"  | Out-File $logFile -Append

#$topicTerms = Get-PnPTerm -TermStore $termStore -TermGroup $termGroup -TermSet "Topic" -Recursive -IncludeChildTerms | %{if($_.TermsCount -gt 0){$_;$_.Terms}else{$_}}
try
{
    $topicTerms = Get-PnPTerm -TermGroup $termGroup -TermSet "Topic" -Recursive -IncludeChildTerms -ErrorAction Stop | %{if($_.TermsCount -gt 0){$_;$_.Terms}else{$_}}
}
catch
{
    Write-Output "An error has occurred while getting Topic Terms"  | Out-File $errorLogFile
    Write-Output ($_ | Out-String)  | Out-File $errorLogFile -Append
}

ForEach($User in $AllUsers) #Users loop
{
    $Counter++
    Write-Progress -Activity "Extracting User Profile Data..." -Status "Getting User Profile $Counter of $allUsersCount" -PercentComplete (($Counter / $allUsersCount)  * 100)

    #Get User Profile

    $accountName = "i:0#.f|membership|$($User.UserPrincipalName)"
    $userProfilePropertiesData = $null

    $user = Get-PnPUserProfileProperty $accountName

    if($user.AccountName -eq $null)
    {
        Write-Output "SharePoint User profile doesn't exist. Checking the next user"  | Out-File $logFile -Append
        continue
    }

    do
    {
        try
        {
            $errorOccured = $null
            $userProfilePropertiesData = $upService.GetUserProfileByName($accountName)
        }
        catch
        {
            $errorOccured = $true
            Write-Output "An error Occured While getting the user profile for user'$accountName'" | Out-File $errorLogFile -Append
            Write-Output $_ | Out-File $errorLogFile -Append
            Write-Output "Sleeping for 120 seconds, as an error has occured while getting the user profile for user'$accountName'"  | Out-File $logFile -Append
            Start-Sleep 120
        }
    }while($errorOccured)
    
    $currentUserRegion = $null
    $currentUserCountry = $null

    if($userProfilePropertiesData -ne $null)
    {
        $fieldsValues=@{}
        $isWizardCompleted = $null
        $customBookmarks = $null
        $appsBookmarks = $null

        foreach($userProfilePropertyData in $userProfilePropertiesData)
        {
            if($userProfilePropertyData.Name -eq $uPSField_WelcomeWizardCompleted)
            {
                $isWizardCompleted = $userProfilePropertyData.Values | select -ExpandProperty Value
                if($isWizardCompleted -eq $false -or $isWizardCompleted -eq $null)
                {
                    #This user hasn't configured the welcome wizard. Hence, aborting
                    Write-Output "This user hasn't configured the welcome wizard. Hence, aborting..."  | Out-File $logFile -Append
                    break
                }
                else
                {
                    $completedWizardUsersCount++
                    Write-Output "Number of completed wizard users: $completedWizardUsersCount" | Out-File $logFile -Append
                }   
            }

            if($userProfilePropertyData.Name -eq $uPSField_CustomBookmarks)
            {
                $customBookmarks = $userProfilePropertyData.Values | select -ExpandProperty Value
                if([string]::IsNullOrEmpty($customBookmarks) -or $customBookmarks.Length -lt 4)
                {
                    #This user isn't using the custom bookmarking feature
                    Write-Output "This user is not using the Custom Bookmarking feature"  | Out-File $logFile -Append
                }
                else
                {
                    $customBookmarkingUsersCount++
                    Write-Output "Number of users using Custom Bookmarking feature: $customBookmarkingUsersCount" | Out-File $logFile -Append
                }   
            }
            
            if($userProfilePropertyData.Name -eq $uPSField_AppsBookmarks)
            {
                $appsBookmarks = $userProfilePropertyData.Values | select -ExpandProperty Value
                if([string]::IsNullOrEmpty($appsBookmarks) -or $appsBookmarks.Length -lt 4)
                {
                    #This user isn't using the App bookmarking feature
                    Write-Output "This user is not using the Apps Bookmarking feature"  | Out-File $logFile -Append
                }
                else
                {
                    $appsBookmarkingUsersCount++
                    Write-Output "Number of users using Apps Bookmarking feature: $appsBookmarkingUsersCount" | Out-File $logFile -Append
                }   
            }

            if($uPSFields.Where({$_.Name -eq $userProfilePropertyData.Name}))
            {
                if($userProfilePropertyData.Name -eq $region.Name)
                {
                    #This is the region user property
                    $currentUserRegion = $userProfilePropertyData.Values | select -ExpandProperty Value
                }
                elseif($userProfilePropertyData.Name -eq $country.Name)
                {
                    #This is the country user property
                    $currentUserCountry = $userProfilePropertyData.Values | select -ExpandProperty Value
                }
                else
                {
                    #This is one of the Hk Intranet user properties
                    $fieldsValues[$userProfilePropertyData.Name] = $userProfilePropertyData.Values | select -ExpandProperty Value
                }
            }
        }
        
        if($isWizardCompleted -eq $false -or $isWizardCompleted -eq $null){
            continue
        }

        for($i=0; $i -lt $uPSFields.Count; $i++)  #Hk Intranet Properties loop (per user)
        {
            $currentUPSField = $uPSFields[$i].Name
            $currentUPSField_DisplayName = $uPSFields[$i].DisplayName
            $currentUPSField_Level = $uPSFields[$i].Level
            $currentFieldValues = $fieldsValues[$currentUPSField]


            if($currentUPSField_Level -eq "Child")
            {
                #This is a child property; so, it will be recorded in the SharePoint list by its parent property
                continue
            }

            if($currentFieldValues -ne $null -and $currentFieldValues -ne "" -and $currentFieldValues.count -ne 0)
            {
                #This user already has data in Hk Intranet property (Sender, Topic, etc.)
                foreach($currentFieldValue in $currentFieldValues) #Hk Intranet single Property values loop (per user)
                {
                    #Iterating through the current ups property values
                    $uPSPropertiesValueCollections = ProcessPropertyValue $uPSPropertiesValueCollections $currentUPSField_DisplayName $currentUserRegion $currentUserCountry $currentFieldValue
                }
            }
            $currentFieldValues = $currentFieldValues -join "|"  #To be written in the output CSV file
            
            if($i -eq 0)
            {
                Write-Output "$($User.Email); $currentFieldValues" | Out-File $CSVPath -NoNewline -Append
            }
            else
            {
                Write-Output ";$currentFieldValues" | Out-File $CSVPath -NoNewline -Append
            }
        }
        
        Write-Output ";$currentUserCountry;$currentUserRegion;$customBookmarks;$appsBookmarks" | Out-File $CSVPath -Append
    }
    else
    {
        #User Profile Cannot be found

        Write-Output "this user profile cannot be found!" | Out-File $logFile -Append
    }

}

Write-Output "Completed iterating through all the user profiles" | Out-File $logFile -Append
Write-Output "Updating the SharePoint lists with the current statistics..." | Out-File $logFile -Append

Write-Output "Number of completed wizard users --- $completedWizardUsersCount ---"  | Out-File $logFile -Append
Write-Output "Number of users using Custom Bookmarks --- $customBookmarkingUsersCount ---"  | Out-File $logFile -Append
Write-Output "Number of users using App Bookmarks --- $appsBookmarkingUsersCount ---"  | Out-File $logFile -Append

UpdatePropertyUsersCount -spPropName "Completed Wizard" -spPropValue "All" -updatedSPUsersCount $completedWizardUsersCount
UpdatePropertyUsersCount -spPropName "Custom Bookmarks" -spPropValue "All" -updatedSPUsersCount $customBookmarkingUsersCount
UpdatePropertyUsersCount -spPropName "App Bookmarks" -spPropValue "All" -updatedSPUsersCount $appsBookmarkingUsersCount
#UpdateCompletedWizardUsersCount

foreach($uPSPropertyValueCol in $uPSPropertiesValueCollections.GetEnumerator())
{
    Write-Output "Iterating through Hk Intranet UPS properties - $($uPSPropertyValueCol.Name)" | Out-File $logFile -Append

    foreach($upsPropertyValue in $uPSPropertyValueCol.Value.GetEnumerator())
    {
        $upsMainPropertyValue = $null
        $upsSubPropertyValue = $null

        if($upsPropertyValue.Name.split('\').count -gt 1)
        {
            #this is a property with sub property in the form property\subproperty, e.g. Hk News\Germany
            $upsPropertyValueName = $upsPropertyValue.Name.split('\')
            $upsMainPropertyValue = $upsPropertyValueName[0]
            $upsSubPropertyValue = $upsPropertyValueName[1]
            Write-Output "Property with subproperty found: Property ($upsMainPropertyValue) \ Subproperty ($upsSubPropertyValue)" | Out-File $logFile -Append
        }
        Write-Output "Iterating through '$($uPSPropertyValueCol.Name)' property values - $($upsPropertyValue.Name)" | Out-File $logFile -Append
        Write-Output "$($uPSPropertyValueCol.Name),$($upsPropertyValue.Name),$($upsPropertyValue.Value)" | out-file $outputFilePath -Append

        $currentListTitle = GetUPSPropertyListTitle($uPSPropertyValueCol.Name)
        $currentSPPropItem = GetUPSProperty-ListItem -currentListTitle $currentListTitle -currentPropName $uPSPropertyValueCol.Name -currentPropValue $upsPropertyValue.Name

        if($currentSPPropItem -eq $null)
        {
            if($upsMainPropertyValue -and $upsSubPropertyValue)
            {
                Write-Output "This is a new '$upsMainPropertyValue' Property (not found in SharePoint list '$($currentListTitle)') with sub-property '$upsSubPropertyValue'" | Out-File $logFile -Append
                $trialsCount = 0
                do
                {
                    $trialsCount++
                    $errorOccured = $false
                    try
                    {
                        #This try catch blocks are needed because occassionally the get-pnplistitem and set-pnplistitem are failing
                        $item = Add-PnPListItem -List $currentListTitle -Folder $spList_PrivateFolderName -Values @{"$spList_PropNameColumn" = "$($uPSPropertyValueCol.Name)"; "$spList_PropValueColumn"="$upsMainPropertyValue";"$spList_SubPropValueColumn"="$upsSubPropertyValue";"$spList_UsersCountColumn"="$($upsPropertyValue.Value)"} -Connection $currentConnection -ErrorAction Stop
                        Write-Output "SharePoint item for '$upsMainPropertyValue' Property and with sub-property '$upsSubPropertyValue' has been created successfully" | Out-File $logFile -Append
                    }
                    catch
                    {
                        $errorOccured = $true
                        Write-Output "An error has occurred while creating the SharePoint Item for $($uPSPropertyValueCol.Name) - $upsMainPropertyValue\$upsSubPropertyValue. Hence, retrying...($trialsCount)" | Out-File $errorLogFile -Append
                        Write-Output $_ | Out-File $errorLogFile -Append
                        Start-Sleep -Seconds 120
                        
                        if($trialsCount -eq 6)
                        {
                            #Write-Host "ERROR! Failed to create the SharePoint Item for $($uPSPropertyValueCol.Name) - $upsMainPropertyValue\$upsSubPropertyValue" -ForegroundColor Red
                            Write-Output "ERROR! Failed to create the SharePoint Item for $($uPSPropertyValueCol.Name) - $upsMainPropertyValue\$upsSubPropertyValue" | Out-File $errorLogFile -Append
                        }
                    }
                }while($errorOccured -and $trialsCount -lt 6)
            }
            else
            {
                Write-Output "This is a new '$($uPSPropertyValueCol.Name)' Property (not found in SharePoint list '$($currentListTitle)')" | Out-File $logFile -Append
                $trialsCount = 0
                do
                {
                    $trialsCount++
                    $errorOccured = $false
                    try
                    {
                        #This try catch blocks are needed because occassionally the get-pnplistitem and set-pnplistitem are failing

                        $item = Add-PnPListItem -List $currentListTitle -Folder $spList_PublicFolderName -Values @{"$spList_PropNameColumn" = "$($uPSPropertyValueCol.Name)"; "$spList_PropValueColumn"="$($upsPropertyValue.Name)";"$spList_UsersCountColumn"="$($upsPropertyValue.Value)";"$spList_TermGUID"="$(Get-TermId -TermSetName $uPSPropertyValueCol.Name -TermName $upsPropertyValue.Name)"} -Connection $currentConnection -ErrorAction Stop
                        Write-Output "SharePoint item for '$($uPSPropertyValueCol.Name)' Property has been created successfully" | Out-File $logFile -Append
                    }
                    catch
                    {
                        $errorOccured = $true
                        Write-Output "An error has occurred while creating the SharePoint Item for $($uPSPropertyValueCol.Name) - $($upsPropertyValue.Name). Hence, retrying...($trialsCount)" | Out-File $errorLogFile -Append
                        Write-Output $_ | Out-File $errorLogFile -Append
                        Start-Sleep -Seconds 120
                        
                        if($trialsCount -eq 6)
                        {
                            Write-Output "ERROR! Failed to create the SharePoint Item for $($uPSPropertyValueCol.Name) - $($upsPropertyValue.Name)" | Out-File $errorLogFile -Append
                        }
                    }
                }while($errorOccured -and $trialsCount -lt 6)
                
            }
        }
        else
        {
            if($currentSPPropItem.Count -gt 1)
            {
                #There are more than one item with the same property, subproperty values. Hence, taking the first item
                $currentSPPropItem = $currentSPPropItem[0]   
            }
            if($upsMainPropertyValue -and $upsSubPropertyValue)
            {
                Write-Output "Existent '$upsMainPropertyValue' Property value with sub-property '$upsSubPropertyValue'. Updating its Users Count" | Out-File $logFile -Append
                $trialsCount = 0
                do
                {
                    $trialsCount++
                    $errorOccured = $false
                    try
                    {
                        #This try catch blocks are needed because occassionally the get-pnplistitem and set-pnplistitem are failing
                        $item = Set-PnPListItem -Identity $currentSPPropItem.Id -List $currentListTitle -Values @{"$spList_UsersCountColumn"="$($upsPropertyValue.Value)"} -Connection $currentConnection -ErrorAction Stop
                        Write-Output "Existent '$upsMainPropertyValue' Property value with sub-property '$upsSubPropertyValue' Users Count has been set successfully" | Out-File $logFile -Append
                    }
                    catch
                    {
                        $errorOccured = $true
                        Write-Output "An error has occurred while updating the SharePoint Item for $upsMainPropertyValue - $upsSubPropertyValue. Hence, retrying...($trialsCount)" | Out-File $errorLogFile -Append
                        Write-Output $_ | Out-File $errorLogFile -Append
                        Start-Sleep -Seconds 120
                        if($trialsCount -eq 6)
                        {
                            Write-Output "ERROR! Failed to update the SharePoint Item for $upsMainPropertyValue - $upsSubPropertyValue" | Out-File $errorLogFile -Append
                        }
                    }
                }while($errorOccured -and $trialsCount -lt 6)
            }
            else
            {
                Write-Output "Existent '$($uPSPropertyValueCol.Name)' Property value '$upsMainPropertyValue'. Updating its Users Count" | Out-File $logFile -Append
                $trialsCount = 0
                do
                {
                    $trialsCount++
                    $errorOccured = $false
                    try
                    {
                        #This try catch blocks are needed because occassionally the get-pnplistitem and set-pnplistitem are failing
                        #$item = Set-PnPListItem -Identity $currentSPPropItem.Id -List $currentListTitle -Values @{"$spList_UsersCountColumn"="$($upsPropertyValue.Value)"} -Connection $currentConnection -ErrorAction Stop

                        $item = Set-PnPListItem -Identity $currentSPPropItem.Id -List $currentListTitle -Values @{"$spList_UsersCountColumn"="$($upsPropertyValue.Value)";"$spList_TermGUID"="$(Get-TermId -TermSetName $uPSPropertyValueCol.Name -TermName $upsPropertyValue.Name)"} -Connection $currentConnection -ErrorAction Stop
                        Write-Output "Existent '$($uPSPropertyValueCol.Name)' Property value '$($upsPropertyValue.Name)' Users Count has been updated successfully" | Out-File $logFile -Append
                    }
                    catch
                    {
                        $errorOccured = $true
                        Write-Output "An error has occurred while updating the SharePoint Item for $($uPSPropertyValueCol.Name). Hence, retrying...($trialsCount)" | Out-File $errorLogFile -Append
                        Write-Output $_ | Out-File $errorLogFile -Append
                        Start-Sleep -Seconds 120
                        if($trialsCount -eq 6)
                        {
                            Write-Output "ERROR! Failed to update the SharePoint Item for $($uPSPropertyValueCol.Name)" | Out-File $errorLogFile -Append
                        }
                    }
                }while($errorOccured -and $trialsCount -lt 6)
            }
        }
    }
    Write-Output "Completed iterating through all '$($uPSPropertyValueCol.Name)' values" | Out-File $logFile -Append
}

Write-Output "Completed iterating through all Hk Intranet UPS properties values" | Out-File $logFile -Append
