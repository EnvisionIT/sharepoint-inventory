<##################################################################
 Name: EIT-Sharepoint-Inventory 

 .SYNOPSIS
 Get inventory of the document items on Sharepoint Site Collections

 .DESCRIPTION
 The scripts catalogs all the document items in site collections the account has access
 and generate output csv file

 The script can use AAD App Service principle to connect
 AAD App Registration
        $app = Register-PnPAzureADApp -Interactive -ApplicationName "EIT SharePoint Inventory Scripts" -Tenant <sitename>.sharepoint.com -OutPath "C:\Temp\EIT\Sharepoint Permissions Report" -CertificatePassword (ConvertTo-SecureString -String "PutYourPassw0rdHere" -AsPlainText -Force) -SharePointApplicationPermissions "Sites.FullControl.All"  -Store CurrentUser
        $app

 OR
 SharePoint Service principle(Legacy) to connect
    
 Follow the article for Sharepoint app registration
    https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs

    Shareoint App registration url
    https://contoso.sharepoint.com/_layouts/15/appregnew.aspx

    Granting the permission to Sharepoint app url
     https://contoso-admin.sharepoint.com/_layouts/15/appinv.aspx

    Permission XML for tentant wide permissions
     <AppPermissionRequests AllowAppOnlyPolicy="true">
        <AppPermissionRequest Scope="http://sharepoint/content/tenant" Right="FullControl" />
     </AppPermissionRequests>


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

.PARATMETER SPSiteURL
Specifiy the root Sharepoint site collection

.PARAMETER ReportFile
Output file full path

##################################################################>
#requires -version 5.1
##requires -module PnP.Powershell

Param
(
    [Parameter (Mandatory = $true)][string]$SPSiteURL = "https://envisionitdev-admin.sharepoint.com",
    [Parameter (Mandatory = $true)][string]$ReportFile = "C:\Temp\SharePoint_Inventory.csv"
)


function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss.fff}] `t" -f (Get-Date)
}

# Connect to sharepoint 
function Connect-PnPOnlineHelper {
    Param
    (
        [Parameter(Mandatory = $true)][string] $URL
    )

    if ($ClientId -and $ClientSecret -and $CertPath -and $Tenant) {
        $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -ClientId $ClientId -CertificatePath $CertPath -CertificatePassword $ClientSecret -Tenant $Tenant
    }
    elseif ($ClientId -and $Thumbprint -and $Tenant) {
        $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint -InformationAction Ignore
    }
    elseif ($ClientId -and $ClientSecret) {
        $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -ClientId $ClientId -ClientSecret $ClientSecret -WarningAction Ignore
    }
    elseif ($IsSharePointOnline) {
        $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -Interactive
    }
    else {
        $newConn = Connect-PnPOnline -ReturnConnection -Url $URL -CurrentCredentials
    }

    return $newConn
}


function Inventory-Site() {
    Param
    (
        [Parameter(Mandatory = $true)][string] $SiteUrl
    )
    
    try {
        $connSite = Connect-PnPOnlineHelper -Url $SiteUrl

        $listCounter = 0
        #Target multiple lists 
        $allLists = Get-PnPList -ErrorAction Stop | Where-Object {$_.BaseTemplate -eq 101} 
        foreach ($rowList in $allLists) {
            
            $listCounter++
            Write-Progress -PercentComplete ($listCounter / ($allLists.Count) * 100) -Activity "Processing Lists $listCounter of $($allLists.Count)" -Status "Processing inventory from List '$($rowList.Title)' in $($SiteUrl)" -Id 3 -ParentId 2

            Write-Host "$(Get-TimeStamp) InventorySite: Processing List: $($rowList.Title) `t List Items: $($rowList.ItemCount)"
            $allItems = Get-PnPListItem -List $rowList.Title -PageSize 5000
            
            $listItemCounter = 0    
            foreach ($item in $allItems) {
                $listItemCounter++
                Write-Progress -PercentComplete ($listItemCounter / ($rowList.ItemCount) * 100) -Activity "Processing Items $listItemCounter of $($rowList.ItemCount)" -Status "Processing list items of '$($rowList.Title)'" -Id 4 -ParentId 3
                
                if (($item.FileSystemObjectType) -eq "File") {
                    $rowItems = ''
                    $rowItems = '"'+$RootWebUrl+$item["FileRef"]+'","'+$SiteUrl+'","'+$item["File_x0020_Size"]+'","","'+$item["Created_x0020_Date"]+'","'+$item["Author"].LookupValue+'","'+$item["Author"].Email+'","","'+$item["Last_x0020_Modified"]+'","'+$item["Editor"].LookupValue+'","'+$item["Editor"].Email+'"'
                    $rowItems | Out-File $ReportFile -Encoding utf8 -append
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


##########################################
#  Main Scripts
##########################################
[boolean]$Global:IsSharePointOnline = $SPSiteURL.ToLower() -like "*.sharepoint.com*"
[string] $Global:ClientId = $null
[string] $Global:ClientSecret = $null
[string] $Global:RootWebUrl = $null

if ($IsSharePointOnline) {

    Write-Host "How do you want to connect to SharePoint?"
    Write-Host "1. Using Azure AD App"
    Write-Host "2. Using SharePoint App"
    Write-Host "3. Using Current User Credentials"
    Do {
        [int]$AppTypeId = Read-Host "Enter the ID from the above list"
    }
    Until (($AppTypeId -gt 0) -and ($AppTypeId -le 2))

    if ($AppTypeId -eq 1) {
        $ClientId = Read-Host "Application Client ID"
        #$ClientSecret = (ConvertTo-SecureString -AsPlainText 'myprivatekeypassword' -Force)
        $Tenant = Read-Host "Tenant"
        $Thumbprint = Read-Host "Certificate Thumbprint"
        $ClientSecret = $null
    }
    elseif ($AppTypeId -eq 2) {
        # Get the Client Id & Secret
        $ClientId = Read-Host "Application Client ID"
        $ClientSecretSecureString = Read-Host "Application Client Secret" -AsSecureString

        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecretSecureString)
        $ClientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

        $Tenant = $null
        $Thumbprint = $null
    }

    # connect to sharepoint
    $conn = Connect-PnPOnlineHelper -Url $SPSiteURL
    
    #Get list of site collections for the tenant
    $SiteCollections = Get-PnPTenantSite
}
else {
    
    $snapin = Get-PSSnapin | Where-Object {$_.Name -eq 'Microsoft.SharePoint.Powershell'}
    if ($snapin -eq $null) 
    {
        Write-Host "Loading SharePoint Powershell Snapin"
        Add-PSSnapin "Microsoft.SharePoint.Powershell"
    }

    #Get list of site collections for the tenant
    $SiteCollections = Get-SPSite -Limit All
}


# check the log file already exist
if (test-path $ReportFile) {
    remove-item $ReportFile
}

$row = '"FullName","Site","FileSize","Attributes","Created","CreatedBy","CreatedByEmail","Accessed","Modified","ModifiedBy","ModifiedByEmail"'
$row | Out-File $ReportFile -Encoding utf8

[int] $Total = $SiteCollections.Count
[int] $j = 0
# exclude my sites
$SiteCollections | Where-Object {$_.Url -cnotlike '*my.sharepoint.com/*' -and $_.Url -cnotlike '*/personal/*'}| ForEach {
    Write-Host -f Yellow "`n$(Get-TimeStamp) Processing Site Collection: $($_.Url)`n"
    $j++
    Write-Progress -PercentComplete ($j / ($Total) * 100) -Activity "Processing site collection $j of $($Total)" -Status "Processing Site Collection $($_.URL)'" -Id 1

    #if ($_.Url -clike '*/sites/2018-19WebsiteSet-up*') {
        Process-Sites -siteCollURL $_.Url
    #}
}
Write-Host -f Green "$(Get-TimeStamp) All Done!"