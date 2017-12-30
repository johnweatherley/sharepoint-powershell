# requires https://www.microsoft.com/en-us/download/details.aspx?id=42038


Param(
	[Parameter(Mandatory=$true)]
	[String]
	$tenant
)

############################################################################################

function Report-SPOSites(){
  $sites = Get-SPOSite
  $sites | Select-Object Title, Url, Owner, Template, StorageQuota, StorageUsageCurrent, ResourceQuota, ResourceUsageCurrent, Status, SharingCapability | Export-Csv -NoTypeInformation -Path SiteCollections.csv

  $webs = New-Object System.Collections.ArrayList

  "`"Title`",`"Url`",`"Created`",`"WebTemplate`",`"LastItemModifiedDate`"" | Add-Content Webs.csv

  foreach ($site in $sites)
  {
	$sitewebs = Get-SPOWebs $site.url $cred
    $webs.Add($sitewebs) | Out-Null
  }
}

############################################################################################

function Get-SPOWebs(){
param(
   $Url = $(throw "Please provide a Site Collection Url"),
   $Credential = $(throw "Please provide a Credentials")
)  
  $context = New-Object Microsoft.SharePoint.Client.ClientContext($Url)  
  $context.Credentials = $Credential 
  $web = $context.Web
  $context.Load($web)    
  $context.Load($web.Webs)  
  try
	{
		$context.ExecuteQuery()
	}
	catch
	{
		Write-Host "Not able to authenticate to retreive sites from $url.  User probably does not have the correct permissions." -foregroundcolor black -backgroundcolor Red
		return
	}  
  
  "`"$($web.Title)`",`"$($web.Url)`",`"$($web.Created)`",`"$($web.WebTemplate)`",`"$($web.LastItemModifiedDate)`"" | Add-Content Webs.csv

  foreach($subweb in $web.Webs)
  {
    Get-SPOWebs -Url $subweb.Url -Credential $Credential
  }
}

############################################################################################

function Get-SPOGroups(){
param(
   $url = $(throw "Please provide a Site Url")
)
  $groups = Get-SPOSiteGroup -site $url
  
  foreach($group in $groups)
  {
     Write-Host $group.Title -ForegroundColor "Yellow"
     Get-SPOSiteGroup -Site $siteURL -Group $group.Title | Select-Object -ExpandProperty Users
     Write-Host
  }
}

############################################################################################

Write-Host "Load CSOM libraries" -foregroundcolor black -backgroundcolor yellow
Set-Location $PSScriptRoot
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")

Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 

#Credentials to connect to office 365 site collection url 
$url ="https://$tenant-admin.sharepoint.com/"
$getcred = Get-Credential -credential $null
#$username="jweatherley@wolforg.onmicrosoft.com"
#$password="!4yourbuck"
#$Password = $password |ConvertTo-SecureString -AsPlainText -force
$cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($getcred.UserName, $getcred.Password)

Write-Host "Trying to authenticate to SharePoint Online Tenant site $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow  
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$Context.Credentials = $cred
$context.RequestTimeOut = 5000 * 60 * 10;
$web = $context.Web
$site = $context.Site 
$context.Load($web)
$context.Load($site)
try
{
    $context.ExecuteQuery()
    Write-Host "Authenticated to SharePoint Online Tenant site $url and get ClientContext object successfully" -foregroundcolor black -backgroundcolor Green
}
catch
{
    Write-Host "Not able to authenticate to SharePoint Online $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    return
}

############################################################################################

Connect-SPOService -Url $url -credential $getcred

Report-SPOSites