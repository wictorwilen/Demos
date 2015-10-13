# Remote Provisioning of Office 365 Artifacts

# WICTOR WILÉN
# SharePoint MVP
# Group Manager at Avanade Sweden
# @wictor
# http://www.wictorwilen.se

# This script: http://askwictor.com/uconnect-remote-prov 
#                     (not yet, after this session!!!)













###############################
# SCENARIO 1
# Use the Tenant Admin
# Requirements: A web browser

# Just use the web-browser: https://wictor-admin.sharepoint.com/_layouts/15/online/SiteCollections.aspx










###############################
# SCENARIO 2
# Use PowerShell 
# Requirements: SharePoint Online Management Shell
# Download at: http://askwictor.com/spo-powershell

$creds = Get-Credential
Connect-SPOService -Url https://wictor-admin.sharepoint.com/ -Credential $creds

New-SPOSite -Url https://wictor.sharepoint.com/sites/ManuallyProvisionedS `
    -Owner wictor@wictorwilen.se `
    -Title "A Site" `
    -Template "STS#0" `
    -LocaleId 1033 `
    -StorageQuota 10

Disconnect-SPOService








###############################
# SCENARIO 3
# Use PowerShell and PnP cmdlets
# Requirements: SPO and PnP PowerShell cmdlets
# Download at: http://askwictor.com/spo-powershell-pnp
# Schema: https://github.com/OfficeDev/PnP-Provisioning-Schema

$creds = Get-Credential
Connect-SPOService -Url https://wictor-admin.sharepoint.com/ -Credential $creds

New-SPOSite -Url https://wictor.sharepoint.com/sites/PnPProvisionedS `
    -Owner wictor@wictorwilen.se `
    -Title "A Site" `
    -Template "STS#0" `
    -LocaleId 1033 `
    -StorageQuota 10

Connect-SPOnline -Url https://wictor.sharepoint.com/sites/template -Credential $creds
Get-SPOProvisioningTemplate -Out c:\temp\PnPTemplate.xml -Schema LATEST -IncludeSiteCollectionTermGroup
Disconnect-SPOnline

Connect-SPOnline -Url https://wictor.sharepoint.com/sites/PnPProvisionedS -Credentials $creds
Apply-SPOProvisioningTemplate -Path c:\temp\PnPTemplate.xml
Disconnect-SPOnline

Disconnect-SPOService





###############################
# SCENARIO 4
# Use PowerShell and PnP cmdlets, with a supported list
# Requirements: SPO and PnP PowerShell cmdlets

$creds = Get-Credential
Connect-SPOService -Url https://wictor-admin.sharepoint.com/ -Credential $creds
Connect-SPOnline -Url https://wictor.sharepoint.com/sites/SiteCreator -Credentials $creds

$web = Get-SPOWeb 
$list = Get-SPOList -Identity "Site Requests" -Web $web
Get-SPOListItem -List $list | ForEach-Object {
    $item = $_
    if($item["Status"] -eq "Requested") {

        $template = "STS#0"
        if($item["Template"] -eq "Project") {
            $template = "PROJECTSITE#0"
        }
        New-SPOSite -Url "https://wictor.sharepoint.com/sites/$($item["Title"])" `
            -Owner wictor@wictorwilen.se `
            -Title $item["Title"] `
            -Template $template `
            -LocaleId 1033 `
            -StorageQuota 10 

        $item["Status"] = "Created"
        $item.Update()
    }
}

Disconnect-SPOnline
Disconnect-SPOService





###############################
# SCENARIO 5
# Use the Provisioning UX
# Requirements: SharePoint App + Azure Web Site + Azure Web Job
# Download at: http://askwictor.com/pnp-prov-ux

# https://wictor.sharepoint.com/sites/dp4/

####
# Clean up
$creds = Get-Credential
Connect-SPOService -Url https://wictor-admin.sharepoint.com/ -Credential $creds
Remove-SPOSite https://wictor.sharepoint.com/sites/ManuallyProvisioned
Remove-SPOSite https://wictor.sharepoint.com/sites/PnPProvisioned

###############################
# Groups
# PowerShell
# Requirements: None!
# More info at: http://www.wictorwilen.se/office-365-groups-for-admins-creating-groups

# Connect to EXO
$creds = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
    -Credential $creds -Authentication Basic -AllowRedirection
Import-PSSession $Session
 
New-UnifiedGroup `
    -DisplayName "Group: Project X" `
    -Alias "unifiedgroup-project-x" `
    -EmailAddresses "group-project-x@askwictor.com" `
    -AccessType Private 
 
Remove-PSSession $session


