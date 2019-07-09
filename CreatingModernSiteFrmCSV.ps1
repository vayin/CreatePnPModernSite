#Write in professional manner
#Regions, variables, functions, Try catch, logging,
Import-Module Microsoft.Online.SharePoint.Powershell -ErrorAction SilentlyContinue

Add-PSSnapIn Microsoft.SharePoint.PowerShell  -ErrorAction SilentlyContinue

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"

 
# Authenticate with the SharePoint site. 

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$ScriptDirCSV = $ScriptDir+"\SiteURLS.csv"

$SiteURLS = Import-Csv $ScriptDirCSV


<#$username = Read-Host -Prompt "Enter admin email:" 
$password = Read-Host -Prompt "Enter admin password:" -AsSecureString
$adminUrl = Read-Host -Prompt "Enter the Tenant Admin URL" 
#Connect-PnPOnline -Url $adminUrl -Credentials($username, $password);
#Other method
[System.Management.Automation.PSCredentials]$PSCredentials=New-Object System.Management.Automation.PSCredential($username,$password)#>
$AdminUrl=Read-Host -Prompt "Input the Site URL to connect via PNP"
Connect-PnPOnline -Url $AdminUrl -UseWebLogin

ForEach($item in $SiteURLS)
{
 $Title=$item.("Title");
 $SiteUrl=$item.("SiteURL");
 $Template=$item.("Template");
 $Description=$item.("Description");
 $owner=$item.("Owner");

New-PnPTenantSite `
  -Title $Title `
  -Url $SiteUrl `
  -Description $Description `
  -Owner $owner `
  -Lcid 1033 `
  -Template $Template `
  -TimeZone 11 `
  -StorageQuota 51200 `
  -StorageQuotaWarningLevel 45000 `
  -Wait


Write-Host("Site Provision Completed $Title") -BackgroundColor Green; 
}


