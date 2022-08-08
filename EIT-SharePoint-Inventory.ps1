<##################################################################
 Name: EIT-Sharepoint-Inventory 

 .SYNOPSIS
 Get inventory of Sharepoint Site Collections, Sites, SubSites and the document items on Sharepoint Site Collections

 .DESCRIPTION
 The scripts catalogs all the document items in site collections as well as all the site collections and SubSites the account has access
 and generate output csv file

 The script can use AAD App Service principle to connect
 When prompted 
    Run Inventory:  Allows you proceed directly to starting the inventory process. Here you can connect using User Credentials with SharePoint Admin Privileges
                    or you can use an Existing Azure AD App Registration

    Register Azure AD App: This will allow you to create an Azure AD App Registration to use for the connection before starting the inventory process. This process will generate a
                            config file the script can use for Certificate Authentication. The account used for this process must be able to create App Registrations in Azure and
                            would need to be able to grant Admin Consent to the App Service permissions.

 ON PREMISE
    The script needs to be run on the sharepoint server and the account used to
    run the scripts needs to have db_owner permissions on the Sharepoint_Config database
    Add-SPShellAdmin -UserName CONTOSO\User1

    The following prerequistises
    1. Powershell version 5.1 or above
    2. SharePointPnPPowerShell2013/2016/2019 module needs to be installed depending Sharepoint version
        e.g. Install-Module PnP.Powershell

 SHAREPOINT ONLINE
    The account used in script needs to tenant admin

  The following prerequistises
    1. Powershell version 5.1 or above
    2. SharePointPnPPowerShellOnline module needs to be installed
        Install-Module PnP.Powershell

  OUTPUT
    The script outputs 2 csv files:
        1. EnvisionIT Document Inventory.csv
        2. EnvisionIT Site Collection Inventory.csv

##################################################################>

#General Use Function - used for timestamp outputs
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss.fff}] `t" -f (Get-Date)
}

function IIf($If, $Then, $Else) {
    If ($If -IsNot "Boolean") {$_ = $If}
    If ($If) {If ($Then -is "ScriptBlock") {&$Then} Else {$Then}}
    Else {If ($Else -is "ScriptBlock") {&$Else} Else {$Else}}
}

#Connect to sharepoint 
function Connect-PnPOnlineHelper {
    Param
    (
        [Parameter(Mandatory = $true)][string] $URL
    )

    if ($ClientId -and $Thumbprint -and $Tenant) {
        $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint -InformationAction Ignore
    }
    elseif ($IsSharePointOnline) {
        $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -Interactive
    }
    else {
        $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -CurrentCredentials
    }

    return $newConn
}

#Document Inventory Function
function Inventory-Site() {
    Param
    (
        [Parameter(Mandatory = $true)][string] $SiteUrl
    )

    try {
        $connSite = Connect-PnPOnlineHelper -Url $SiteUrl

        $listCounter = 0
        #Target multiple lists 
        $allLists = Get-PnPList -ErrorAction Stop -Connection $connSite | Where-Object {$_.BaseTemplate -eq 101 -or $_.BaseTemplate -eq 700}
        foreach ($rowList in $allLists) {

            $listCounter++
            Write-Progress -PercentComplete ($listCounter / ($allLists.Count) * 100) -Activity "Processing Lists $listCounter of $($allLists.Count)" -Status "Processing inventory from List '$($rowList.Title)' in $($SiteUrl)" -Id 3 -ParentId 2

            Write-Host "$(Get-TimeStamp) InventorySite: Processing List: $($rowList.Title) `t List Items: $($rowList.ItemCount)"
            $allItems = Get-PnPListItem -List $rowList.Title -PageSize 5000 -Connection $connSite

            $listItemCounter = 0    
            foreach ($item in $allItems) {
                $listItemCounter++
                Write-Progress -PercentComplete ($listItemCounter / ($rowList.ItemCount) * 100) -Activity "Processing Items $listItemCounter of $($rowList.ItemCount)" -Status "Processing list items of '$($rowList.Title)'" -Id 4 -ParentId 3

                if (($item.FileSystemObjectType) -eq "File") {
                    $rowItems = ''
                    $rowItems = '"'+$RootWebUrl+$item["FileRef"]+'","'+$SiteUrl+'","'+$item["File_x0020_Size"]+'","","'+$item["Created_x0020_Date"]+'","'+$item["Author"].LookupValue+'","'+$item["Author"].Email+'","","'+$item["Last_x0020_Modified"]+'","'+$item["Editor"].LookupValue+'","'+$item["Editor"].Email+'"'
                    $rowItems | Out-File $DocumentInventoryReportFile -Encoding utf8 -append
                }
            } # end foreach items
            Write-Progress -Id 4 -Activity "List Items Processing Done" -Completed
        } # end foreach Lists
        Write-Progress -Id 3 -Activity "List Processing Done" -Completed
        Write-Host -f Green "$(Get-TimeStamp) Inventory-Site: Processing Done for Site: $($SiteUrl)`n"
    }
    catch {
        Write-Host -f Red "$(Get-TimeStamp) Inventory-Site: Error Occurred while processing the Site: $($SiteUrl)" 
        Write-Host -f Red $_.Exception.Message 
    }
}

#Document Inventory Function
function Process-Sites() {
    Param
    (
        [parameter(Mandatory = $true)][string] $SiteCollUrl
    )

    try {

        Write-Host -f Yellow "$(Get-TimeStamp) InventorySite: Processing Site: $SiteCollUrl"
        #Connect to SharePoint
        $connSite = Connect-PnPOnlineHelper -Url $SiteCollUrl

        $spWeb = Get-PnPWeb -Includes Webs -Connection $connSite -ErrorAction Stop 

        if ($spWeb.ServerRelativeUrl -eq "/") {
            $RootWebUrl = $spWeb.Url.TrimEnd("/")
        }
        else {
            $RootWebUrl = ($spWeb.Url -replace $spWeb.ServerRelativeUrl).TrimEnd("/")
        }

        #Process the root site
        Inventory-Site -SiteUrl $spWeb.Url

        [int] $Total = $spWeb.Webs.Count
        [int] $i = 0

        if ($Total -gt 0) {
            $spSubWebs = Get-PnPSubWebs -Identity $spWeb -Recurse -Connection $connSite -ErrorAction Stop
            $Total = $spSubWebs.Count

            foreach ($spSubWeb in $spSubWebs) {
                $i++
                Write-Progress -PercentComplete ($i / ($Total) * 100) -Activity "Processing site $i of $($Total)" -Status "Processing Subsite $($spSubWeb.URL)'" -Id 2 -ParentId 1

                if ($spSubWeb.ServerRelativeUrl -ne $spWeb.ServerRelativeUrl) {
                    #Process the site
                    Inventory-Site -SiteUrl $spSubWeb.Url
                }
            }
            #Write-Progress -Id 2 -Activity "Site Processing Done" -Completed
        }
    }
    catch {
        Write-Host -f Red "$(Get-TimeStamp) Process-Sites: Error Occurred while processing the Site: $($SiteCollUrl)" 
        Write-Host -f Red $_.Exception.Message 
    }
}

#Site Collection Inventory Function
function Process-SitesOnline() {
    Param
    (
        [Parameter(Mandatory = $true)][string] $SiteURL
    )

    Try {

        $ConnSite = Connect-PnPOnlineHelper -Url $SiteURL 

        $CurrentWeb = Get-PnPWeb -Includes Webs -ErrorAction Stop

        [int] $Total = $CurrentWeb.Webs.Count
        [int] $i = 0

        if ($Total -gt 0) {
            $SubWebs = Get-PnPSubWeb -Identity $CurrentWeb -Recurse -Includes "WebTemplate"
            $Total = $SubWebs.Count

            #Loop throuh all Sub Sites
            foreach($SubWeb in $SubWebs) {
                $i++
                Write-Progress -PercentComplete ($i / ($Total) * 100) -Activity "Processing Subsites $i of $($Total)" -Status "Processing Subsite $($SubWeb.URL)'" -Id 2 -ParentId 1

                # Do Not Root Site again
                if ($SubWeb.ServerRelativeUrl -ne $CurrentWeb.ServerRelativeUrl) {
                    Write-Host "$(Get-TimeStamp) `tProcessing Subsite: $($SubWeb.Url)"

                    $SiteTemplate = ($SiteTemplates | Where-Object { $_.Name -eq $SubWeb.WebTemplate } | Select-Object -First 1).Title
                    If (-not $SiteTemplate) {$SiteTemplate = $SubWeb.WebTemplate}

                    $SiteDetails = New-Object PSObject

                    $SiteDetails | Add-Member NoteProperty 'Site name'($SubWeb.Title)
                    $SiteDetails | Add-Member NoteProperty 'URL'($SubWeb.URL)
                    $SiteDetails | Add-Member NoteProperty 'Site Type'("Subsite")
                    $SiteDetails | Add-Member NoteProperty 'Teams'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage used (GB)'("")
                    $SiteDetails | Add-Member NoteProperty 'Primary admin'($SubWeb.Author.LoginName)
                    $SiteDetails | Add-Member NoteProperty 'Hub'("")
                    $SiteDetails | Add-Member NoteProperty 'Template'($SiteTemplate)
                    $SiteDetails | Add-Member NoteProperty 'Last activity (UTC)'("")
                    $SiteDetails | Add-Member NoteProperty 'Date created'("")
                    $SiteDetails | Add-Member NoteProperty 'Created by'($SubWeb.Author.LoginName)
                    $SiteDetails | Add-Member NoteProperty 'Storage limit (GB)'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage used (%)'("")
                    $SiteDetails | Add-Member NoteProperty 'Microsoft 365 group'("")
                    $SiteDetails | Add-Member NoteProperty 'Files viewed or edited'("")
                    $SiteDetails | Add-Member NoteProperty 'Page views'("")
                    $SiteDetails | Add-Member NoteProperty 'Page visits'("")
                    $SiteDetails | Add-Member NoteProperty 'Files'("")
                    $SiteDetails | Add-Member NoteProperty 'Sensitivity'("")
                    $SiteDetails | Add-Member NoteProperty 'External sharing'("")
                    $SiteDetails | Add-Member NoteProperty 'Sharing Domain Restriction Mode'("")
                    $SiteDetails | Add-Member NoteProperty 'Sharing Allowed Domain List'("")
                    $SiteDetails | Add-Member NoteProperty 'Sharing Blocked Domain List'("")
                    $SiteDetails | Add-Member NoteProperty 'Default Link Permission'("")
                    $SiteDetails | Add-Member NoteProperty 'Default Sharing Link Type'("")
                    $SiteDetails | Add-Member NoteProperty 'External User Expiration In Days'("")
                    $SiteDetails | Add-Member NoteProperty 'Override Tenant Anonymous Link Expiration Policy'("")
                    $SiteDetails | Add-Member NoteProperty 'Override Tenant External User Expiration Policy'("")
                    $SiteDetails | Add-Member NoteProperty 'Allow Downloading Non Web Viewable Files'("")
                    $SiteDetails | Add-Member NoteProperty 'Allow Editing'("")
                    $SiteDetails | Add-Member NoteProperty 'Allow Self Service Upgrade'("")
                    $SiteDetails | Add-Member NoteProperty 'Anonymous Link Expiration In Days'("")
                    $SiteDetails | Add-Member NoteProperty 'Block Download Links File Type'("")
                    $SiteDetails | Add-Member NoteProperty 'Comments On Site Pages Disabled'("")
                    $SiteDetails | Add-Member NoteProperty 'Compatibility Level'("")
                    $SiteDetails | Add-Member NoteProperty 'Conditional Access Policy'("")
                    $SiteDetails | Add-Member NoteProperty 'Default Link To Existing Access'("")
                    $SiteDetails | Add-Member NoteProperty 'Deny Add And Customize Pages'("")
                    $SiteDetails | Add-Member NoteProperty 'Description'("")
                    $SiteDetails | Add-Member NoteProperty 'Disable App Views'("")
                    $SiteDetails | Add-Member NoteProperty 'Disable Company Wide Sharing Links'("")
                    $SiteDetails | Add-Member NoteProperty 'Disable Flows'("")
                    $SiteDetails | Add-Member NoteProperty 'Disable Sharing For Non Owners Status'("")
                    $SiteDetails | Add-Member NoteProperty 'Group Id'("")
                    $SiteDetails | Add-Member NoteProperty 'Hub Site Id'("")
                    $SiteDetails | Add-Member NoteProperty 'Information Segment'("")
                    $SiteDetails | Add-Member NoteProperty 'Limited Access File Type'("")
                    $SiteDetails | Add-Member NoteProperty 'Locale Id'("")
                    $SiteDetails | Add-Member NoteProperty 'Lock Issue'("")
                    $SiteDetails | Add-Member NoteProperty 'Lock State'("")
                    $SiteDetails | Add-Member NoteProperty 'Owner'("")
                    $SiteDetails | Add-Member NoteProperty 'Owner Login Name'("")
                    $SiteDetails | Add-Member NoteProperty 'Owner Name'("")
                    $SiteDetails | Add-Member NoteProperty 'Protection Level Name'("")
                    $SiteDetails | Add-Member NoteProperty 'PWA Enabled'("")
                    $SiteDetails | Add-Member NoteProperty 'Related Group Id'("")
                    $SiteDetails | Add-Member NoteProperty 'Resource Quota'("")
                    $SiteDetails | Add-Member NoteProperty 'Resource Quota Warning Level'("")
                    $SiteDetails | Add-Member NoteProperty 'Resource Usage Average'("")
                    $SiteDetails | Add-Member NoteProperty 'Resource Usage Current'("")
                    $SiteDetails | Add-Member NoteProperty 'Restricted To Geo'("")
                    $SiteDetails | Add-Member NoteProperty 'Sandboxed Code Activation Capability'("")
                    $SiteDetails | Add-Member NoteProperty 'Show People Picker Suggestions For Guest Users'("")
                    $SiteDetails | Add-Member NoteProperty 'Site Defined Sharing Capability'("")
                    $SiteDetails | Add-Member NoteProperty 'Social Bar On Site Pages Disabled'("")
                    $SiteDetails | Add-Member NoteProperty 'Status'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage Quota Type'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage Quota Warning Level'("")
                    $SiteDetails | Add-Member NoteProperty 'Storage Usage Current'("")
                    $SiteDetails | Add-Member NoteProperty 'Webs Count'("")
                    #Export details to CSV File
                    $SiteDetails | Export-CSV $SiteCollectionInventoryReportFile -Encoding UTF8 -NoTypeInformation -Append
                }
            }
        }
    }
    Catch {
        write-host -f Red "$(Get-TimeStamp) Error Processing Site: $SiteURL"
        Write-Host -f Red $_.Exception.Message 
   }
}

#Site Collection Inventory Function
function Process-SiteCollectionOnline() {
    
    $global:SiteCollectionInventoryReportFile = "$($PSScriptRoot)\EnvisionIT Site Collection Inventory.csv"

    # check the log file already exist
    if (test-path $SiteCollectionInventoryReportFile) {
        remove-item $SiteCollectionInventoryReportFile
    }
   
    # connect to sharepoint
    $Adminconn = Connect-PnPOnlineHelper -Url $SPSiteURL

    # Get Tenat Defaults
    $TenantInfo = Get-PnPTenant
    $DefaultSharingCapability = $TenantInfo.SharingCapability
    $DefaultRequireAnonymousLinksExpireInDays = $TenantInfo.RequireAnonymousLinksExpireInDays
    $DefaultSharingDomainRestrictionMode = $TenantInfo.SharingDomainRestrictionMode
    $DefaultSharingAllowedDomainList = $TenantInfo.SharingAllowedDomainList
    $DefaultSharingBlockedDomainList = $TenantInfo.SharingBlockedDomainList
    $DefaultSharingLinkType = $TenantInfo.DefaultSharingLinkType
    $DefaultLinkPermission = $TenantInfo.DefaultLinkPermission

    # Get all the site collection
    $SiteCollections = Get-PnPTenantSite -Connection $Adminconn

    # Get Templates on the Farm
    $SiteTemplates = Get-PnPWebTemplates | SELECT ID, Name, Title -Unique

    [int] $Total = $SiteCollections.Count
    [int] $i = 0
    Write-Host "$(Get-TimeStamp) `tTotal Site Collections found: $($Total)"

    foreach ($SiteCollection in $SiteCollections) {
        $i++
        Write-Progress -PercentComplete ($i / ($Total) * 100) -Activity "Processing Site Collections $i of $($Total)" -Status "Processing Site $($SiteCollection.URL)'" -Id 1

        Write-Host "$(Get-TimeStamp) `tProcessing Site Collection: $($SiteCollection.Url)"
        
        $conn = Connect-PnPOnlineHelper -Url $SiteCollection.Url

        $CurrentWeb = Get-PnPWeb -Includes Created
        $CurrentSiteCollection = Get-PnPTenantSite -Identity $SiteCollection.Url -Detailed -DisableSharingForNonOwnersStatus -Connection $conn

        if ($CurrentSiteCollection) {
            $SiteTemplate = ($SiteTemplates | Where-Object { $_.Name -eq $SiteCollection.Template } | Select-Object -First 1).Title
            If (-not $SiteTemplate) {$SiteTemplate = $CurrentSiteCollection.Template}

            $SiteDetails = New-Object PSObject

            $SiteDetails | Add-Member NoteProperty 'Site name'($CurrentSiteCollection.Title)
            $SiteDetails | Add-Member NoteProperty 'URL'($CurrentSiteCollection.Url)
            $SiteDetails | Add-Member NoteProperty 'Site Type'("Site Collection")
            $SiteDetails | Add-Member NoteProperty 'Teams'("")
            $SiteDetails | Add-Member NoteProperty 'Storage used (GB)'($CurrentSiteCollection.StorageUsageCurrent -f {0:N2})
            $SiteDetails | Add-Member NoteProperty 'Primary admin'($CurrentSiteCollection.OwnerEmail)
            $SiteDetails | Add-Member NoteProperty 'Hub'($CurrentSiteCollection.IsHubSite)
            $SiteDetails | Add-Member NoteProperty 'Template'($SiteTemplate)
            $SiteDetails | Add-Member NoteProperty 'Last activity (UTC)'($CurrentSiteCollection.LastContentModifiedDate.ToShortDateString())
            $SiteDetails | Add-Member NoteProperty 'Date created'($CurrentWeb.Created.ToShortDateString())
            $SiteDetails | Add-Member NoteProperty 'Created by'($CurrentWeb.Author.LoginName)
            $SiteDetails | Add-Member NoteProperty 'Storage limit (GB)'($CurrentSiteCollection.StorageQuota -f {0:N2})
            $SiteDetails | Add-Member NoteProperty 'Sensitivity'($CurrentSiteCollection.SensitivityLabel)
        
            $SiteDetails | Add-Member NoteProperty 'External sharing'((IIf (-not $CurrentSiteCollection.SharingCapability) $CurrentSiteCollection.SharingCapability $DefaultSharingCapability))
            $SiteDetails | Add-Member NoteProperty 'Sharing Domain Restriction Mode'((IIf (-not $CurrentSiteCollection.SharingDomainRestrictionMode) $CurrentSiteCollection.SharingDomainRestrictionMode $DefaultSharingDomainRestrictionMode))
            $SiteDetails | Add-Member NoteProperty 'Sharing Allowed Domain List'((IIf (-not $CurrentSiteCollection.SharingAllowedDomainList) $CurrentSiteCollection.SharingAllowedDomainList $DefaultSharingAllowedDomainList))
            $SiteDetails | Add-Member NoteProperty 'Sharing Blocked Domain List'((IIf (-not $CurrentSiteCollection.SharingBlockedDomainList) $CurrentSiteCollection.SharingBlockedDomainList $DefaultSharingBlockedDomainList))
            $SiteDetails | Add-Member NoteProperty 'Default Link Permission'((IIf (-not $CurrentSiteCollection.DefaultLinkPermission) $CurrentSiteCollection.DefaultLinkPermission $DefaultLinkPermission))
            $SiteDetails | Add-Member NoteProperty 'Default Sharing Link Type'((IIf (-not $CurrentSiteCollection.DefaultSharingLinkType) $CurrentSiteCollection.DefaultSharingLinkType $DefaultSharingLinkType))
            $SiteDetails | Add-Member NoteProperty 'External User Expiration In Days'((IIf ($CurrentSiteCollection.ExternalUserExpirationInDays -gt 0) $CurrentSiteCollection.ExternalUserExpirationInDays $DefaultRequireAnonymousLinksExpireInDays))

            $SiteDetails | Add-Member NoteProperty 'Override Tenant Anonymous Link Expiration Policy'($CurrentSiteCollection.OverrideTenantAnonymousLinkExpirationPolicy)
            $SiteDetails | Add-Member NoteProperty 'Override Tenant External User Expiration Policy'($CurrentSiteCollection.OverrideTenantExternalUserExpirationPolicy)
            $SiteDetails | Add-Member NoteProperty 'Allow Downloading Non Web Viewable Files'($CurrentSiteCollection.AllowDownloadingNonWebViewableFiles)
            $SiteDetails | Add-Member NoteProperty 'Allow Editing'($CurrentSiteCollection.AllowEditing)
            $SiteDetails | Add-Member NoteProperty 'Allow Self Service Upgrade'($CurrentSiteCollection.AllowSelfServiceUpgrade)
            $SiteDetails | Add-Member NoteProperty 'Anonymous Link Expiration In Days'($CurrentSiteCollection.AnonymousLinkExpirationInDays)
            $SiteDetails | Add-Member NoteProperty 'Block Download Links File Type'($CurrentSiteCollection.BlockDownloadLinksFileType)
            $SiteDetails | Add-Member NoteProperty 'Comments On Site Pages Disabled'($CurrentSiteCollection.CommentsOnSitePagesDisabled)
            $SiteDetails | Add-Member NoteProperty 'Compatibility Level'($CurrentSiteCollection.CompatibilityLevel)
            $SiteDetails | Add-Member NoteProperty 'Conditional Access Policy'($CurrentSiteCollection.ConditionalAccessPolicy)
            $SiteDetails | Add-Member NoteProperty 'Default Link To Existing Access'($CurrentSiteCollection.DefaultLinkToExistingAccess)
            $SiteDetails | Add-Member NoteProperty 'Deny Add And Customize Pages'($CurrentSiteCollection.DenyAddAndCustomizePages)
            $SiteDetails | Add-Member NoteProperty 'Description'($CurrentSiteCollection.Description)
            $SiteDetails | Add-Member NoteProperty 'Disable App Views'($CurrentSiteCollection.DisableAppViews)
            $SiteDetails | Add-Member NoteProperty 'Disable Company Wide Sharing Links'($CurrentSiteCollection.DisableCompanyWideSharingLinks)
            $SiteDetails | Add-Member NoteProperty 'Disable Flows'($CurrentSiteCollection.DisableFlows)
            $SiteDetails | Add-Member NoteProperty 'Disable Sharing For Non Owners Status'($CurrentSiteCollection.DisableSharingForNonOwnersStatus)
            $SiteDetails | Add-Member NoteProperty 'Group Id'($CurrentSiteCollection.GroupId)
            $SiteDetails | Add-Member NoteProperty 'Hub Site Id'($CurrentSiteCollection.HubSiteId)
            $SiteDetails | Add-Member NoteProperty 'Information Segment'($CurrentSiteCollection.InformationSegment)
            $SiteDetails | Add-Member NoteProperty 'Limited Access File Type'($CurrentSiteCollection.LimitedAccessFileType)
            $SiteDetails | Add-Member NoteProperty 'Locale Id'($CurrentSiteCollection.LocaleId)
            $SiteDetails | Add-Member NoteProperty 'Lock Issue'($CurrentSiteCollection.LockIssue)
            $SiteDetails | Add-Member NoteProperty 'Lock State'($CurrentSiteCollection.LockState)
            $SiteDetails | Add-Member NoteProperty 'Owner'($CurrentSiteCollection.Owner)
            $SiteDetails | Add-Member NoteProperty 'Owner Login Name'($CurrentSiteCollection.OwnerLoginName)
            $SiteDetails | Add-Member NoteProperty 'Owner Name'($CurrentSiteCollection.OwnerName)
            $SiteDetails | Add-Member NoteProperty 'Protection Level Name'($CurrentSiteCollection.ProtectionLevelName)
            $SiteDetails | Add-Member NoteProperty 'PWA Enabled'($CurrentSiteCollection.PWAEnabled)
            $SiteDetails | Add-Member NoteProperty 'Related Group Id'($CurrentSiteCollection.RelatedGroupId)
            $SiteDetails | Add-Member NoteProperty 'Resource Quota'($CurrentSiteCollection.ResourceQuota)
            $SiteDetails | Add-Member NoteProperty 'Resource Quota Warning Level'($CurrentSiteCollection.ResourceQuotaWarningLevel)
            $SiteDetails | Add-Member NoteProperty 'Resource Usage Average'($CurrentSiteCollection.ResourceUsageAverage)
            $SiteDetails | Add-Member NoteProperty 'Resource Usage Current'($CurrentSiteCollection.ResourceUsageCurrent)
            $SiteDetails | Add-Member NoteProperty 'Restricted To Geo'($CurrentSiteCollection.RestrictedToGeo)
            $SiteDetails | Add-Member NoteProperty 'Sandboxed Code Activation Capability'($CurrentSiteCollection.SandboxedCodeActivationCapability)
            $SiteDetails | Add-Member NoteProperty 'Show People Picker Suggestions For Guest Users'($CurrentSiteCollection.ShowPeoplePickerSuggestionsForGuestUsers)
            $SiteDetails | Add-Member NoteProperty 'Site Defined Sharing Capability'($CurrentSiteCollection.SiteDefinedSharingCapability)
            $SiteDetails | Add-Member NoteProperty 'Social Bar On Site Pages Disabled'($CurrentSiteCollection.SocialBarOnSitePagesDisabled)
            $SiteDetails | Add-Member NoteProperty 'Status'($CurrentSiteCollection.Status)
            $SiteDetails | Add-Member NoteProperty 'Storage Quota Type'($CurrentSiteCollection.StorageQuotaType)
            $SiteDetails | Add-Member NoteProperty 'Storage Quota Warning Level'($CurrentSiteCollection.StorageQuotaWarningLevel)
            $SiteDetails | Add-Member NoteProperty 'Storage Usage Current'($CurrentSiteCollection.StorageUsageCurrent)
            $SiteDetails | Add-Member NoteProperty 'Webs Count'($CurrentSiteCollection.WebsCount)

            #Export details to CSV File
            $SiteDetails | Export-CSV $SiteCollectionInventoryReportFile -Encoding UTF8 -NoTypeInformation -Append
    
            #Loop throuh all Sub Sites
            Process-SitesOnline -SiteURL $SiteCollection.Url
        }
    
    }
}

function Read-Host-With-Default {
    Param
    (
        [Parameter(Position=0)] [string] $Prompt,
        [Parameter(Position=1)] [string] $CurrentValue
    )

    $Entry = Read-Host "$Prompt (Default: $CurrentValue)"
    if ($Entry -eq "") {
        return $CurrentValue
    }
    else {
        return $Entry
    }
}

#Password Generation Function used for creating Certs with a Password
function GeneratePassword ([Parameter(Mandatory=$true)][int]$PasswordLength){
    Add-Type -AssemblyName System.Web
    $PassComplexCheck = $false

    do {
        $newPassword=[System.Web.Security.Membership]::GeneratePassword($PasswordLength,2)
        If ( ($newPassword -cmatch "[A-Z\p{Lu}\s]") `
            -and ($newPassword -cmatch "[a-z\p{Ll}\s]") `
            -and ($newPassword -match "[\d]") `
            -and ($newPassword -match "[^\w]")
            )
        {
            $PassComplexCheck=$True
        }
    } While ($PassComplexCheck -eq $false)

    return $newPassword
}

#Register Azure AD App
function RegisterAzureAdApp(){
    
    #Prompt for the Azure Tenant we'll need it later
    Write-Output "When prompted to login to Azure be sure to use an account with privileges to create App Registrations and to Grant Admin Consent for requested API Permissions"
    $tenantName = Read-Host "Enter the tenant name that you will be installing into (tenantName.onmicrosoft.com)"
    
    #Prompt for the App Name
    $appRegistrationName = Read-Host "Enter a name for the App Registration (Default: Envision IT SharePoint Inventory Scripts)"
    if($appRegistrationName -eq $null -or $appRegistrationName -eq ""){
        $appRegistrationName = "Envision IT SharePoint Inventory Scripts"
    }

    # Generate a password
    $UnsecurePassword = GeneratePassword (25)
    $certPassword = ConvertTo-SecureString $UnsecurePassword -AsPlainText -Force

    # Creates App Registration, generates certificate, uploads the cert to the app registration and cert store on local machine, prompts user to grant admin consent for permissions - all in 1 command
    $app = Register-PnPAzureADApp -ApplicationName $appRegistrationName -Tenant $tenantName `
                -CertificatePassword $certPassword `
                -SharePointApplicationPermissions "Sites.ReadWrite.All" `
                -Store CurrentUser `
                -Interactive
    
    #Outputs key data that can be reused at a later point to a local config file
    $json = @"
{
    "Tenant": "$($tenantName)",
    "App Name": "$($appRegistrationName)",
    "Client Id": "$($app.'AzureAppId/ClientId')",
    "Certificate Name": "CN=$($appRegistrationName)",
    "Certificate Thumbprint": "$($app.'Certificate Thumbprint')"
}
"@
    $outputConfig = $json | out-file -filepath "$PSScriptRoot\Envision IT Inventory Script Config.json"

    Write-Host -ForegroundColor Green "Azure AD App Registered! Config created at: $PSScriptRoot\Envision IT Inventory Script Config.json"
    Write-Host "Please check App '$($appRegistrationName)' with ClientID: $($app.'AzureAppId/ClientId') has been granted Admin Consent."
    #Pause and wait for confirmation to continue
    Read-Host “Press ENTER to continue once Admin Consent has been granted...”
}

#Main Inventory Function
function RunInventory(){    
    
    Write-Host -ForegroundColor Yellow "Inventory Process Started!"

    $global:SPSiteURL = Read-Host-With-Default "Enter the root site collection URL for SharePoint" $SPSiteURL
    $global:DocumentInventoryReportFile = "$($PSScriptRoot)\EnvisionIT Document Inventory.csv"
    [Boolean]$AllSiteCollections = $true
    
    [boolean]$Global:IsSharePointOnline = $SPSiteURL.ToLower() -like "*.sharepoint.com*"
    [string] $Global:RootWebUrl = $null
    
    if ($IsSharePointOnline) {
    
        Write-Host "How do you want to connect to SharePoint?"
        Write-Host "1. Using Azure AD App"
        Write-Host "2. Interactive/Current User"
        Do {
            $global:AppTypeId = Read-Host-With-Default "Enter the ID from the above list" $AppTypeId
        }
        Until (($AppTypeId -gt 0) -and ($AppTypeId -le 2))
    
        if ($AppTypeId -eq 1) {
            
            $config = Get-Content "$PSScriptRoot\Envision IT Inventory Script Config.json" | Out-String | ConvertFrom-Json
            
            if ($config.Tenant -eq $null -or $config.Tenant -eq ""){
                $Tenant = Read-Host-With-Default "Enter Tenant (tenantName.onmicrosoft.com)" $Tenant
            }
            else{
                Write-Host -ForegroundColor Yellow "Using Tenant: $($config.Tenant)"
                $Tenant = $config.Tenant
            }

            if ($config.'Client Id' -eq $null -or $config.'Client Id' -eq ""){
                $ClientId = Read-Host-With-Default "Enter ClientID" $ClientId
            }
            else{
                Write-Host -ForegroundColor Yellow "Using App $($config.'App Name') with ClientID: $($config.'Client Id')"
                $ClientId = $config.'Client Id'
            }

            if ($config.'Certificate Thumbprint' -eq $null -or $config.'Certificate Thumbprint' -eq ""){
                $Thumbprint = Read-Host-With-Default "Enter Certifcate Thumbprint" $Thumbprint
            }
            else{
                Write-Host -ForegroundColor Yellow "Using Thumbprint: $($config.'Certificate Thumbprint')"
                $Thumbprint = $config.'Certificate Thumbprint'
            }
        }
        elseif ($AppTypeId -eq 2) {
            $ClientId = $null
            $Tenant = $null
            $Thumbprint = $null
        }
    
        # connect to sharepoint
        $global:Adminconn = Connect-PnPOnlineHelper -Url $SPSiteURL
    
        #Get list of site collections for the tenant
    
        if ($AllSiteCollections -eq $true) {
            #Get list of site collections for the tenant
            $SiteCollections = Get-PnPTenantSite -Connection $Adminconn
        }
    }
    else {
        #On-Prem Connections
        $snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'}
        if ($snapin -eq $null) 
        {
            Write-Host "Loading SharePoint Powershell Snapin"
            Add-PSSnapin "Microsoft.SharePoint.Powershell"
        }
    
        #Get list of site collections for the tenant
        $SiteCollections = Get-SPSite -Limit All
    }

    Write-Host -ForegroundColor Yellow "Document Inventory Process has started!"
    
    # check the log file already exist
    if (test-path $DocumentInventoryReportFile) {
        remove-item $DocumentInventoryReportFile
    }
    
    $row = '"FullName","Site","FileSize","Attributes","Created","CreatedBy","CreatedByEmail","Accessed","Modified","ModifiedBy","ModifiedByEmail"'
    $row | Out-File $DocumentInventoryReportFile -Encoding utf8
    
    [int] $Total = $SiteCollections.Count
    [int] $j = 0
    
    if ($AllSiteCollections -eq $true) {
        $SiteCollections | Where-Object {$_.Url -cnotlike '*my.sharepoint.com/*' -and $_.Url -cnotlike '*/personal/*'}| ForEach {
            Write-Host -f Yellow "`n$(Get-TimeStamp) Processing Site Collection: $($_.Url)`n"
            $j++
            Write-Progress -PercentComplete ($j / ($Total) * 100) -Activity "Processing site collection $j of $($Total)" -Status "Processing Site Collection $($_.URL)'" -Id 1
    
            #if ($_.Url -clike '*/sites/2018-19WebsiteSet-up*') {
                Process-Sites -siteCollURL $_.Url
            #}
        }
    }
    else {
        Write-Host -f Yellow "`n$(Get-TimeStamp) Processing Site Collection: $SPSiteURL `n"
        Process-Sites -siteCollURL $SPSiteURL
    }

    #Run Site Collection Inventory
    Write-Host -ForegroundColor Yellow "Site Collection Inventory Process has started!"
    Process-SiteCollectionOnline

    Write-Host -f Green "$(Get-TimeStamp) All Done!"
}

##########################################
#  Main Scripts
##########################################

Write-Host "What would you like to do?"
Write-Host "1. Run Inventory"
Write-Host "2. Register Azure AD App"

Do {
    [int]$ChoiceId = Read-Host-With-Default "Enter the ID from the above list" $ChoiceId
}
Until (($ChoiceId -gt 0) -and ($ChoiceId -le 2))

if ($ChoiceId -eq 1) {
    RunInventory
}
elseif ($ChoiceId -eq 2){
    RegisterAzureAdApp
}
