<#
.SYNOPSIS
  Connects the session to SharePoint Online.
.DESCRIPTION
  This script connects to SharePoint Online to provide you access to the commandlets and a CSOM context.
.PARAMETER Tenant
  The tenant name.  If your SharePoint Online URL is "https://domain.sharepoint.com" then your tenant name is "domain".
.PARAMETER URL
  Optional Site Collection URL if you wish to connect to the context of a Site Collection other than root.
.NOTES
  Version:        1.0
  Author:         John Weatherley
  Creation Date:  March 16, 2017
  Purpose/Change: Initial script development
.LINK
  http://www.johnweatherley.com
  https://github.com/johnweatherley/sharepoint-powershell
.EXAMPLE
  Connect-SPO -tenant domain
  Connect-SPO -tenant domain -url https://domain.sharepoint.com/sites/sitecollection
#>

Param(
	[Parameter(Mandatory=$true)]
	[String]
	$tenant,
    [string]
    $url
)

Write-Host "Load CSOM libraries" -foregroundcolor black -backgroundcolor yellow
Set-Location $PSScriptRoot
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Publishing.dll")

Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 

#Setup###########################################################################################

#Set Urls
if (!$url)
{
    $url = "https://$tenant.sharepoint.com/"
}

$urladmin ="https://$tenant-admin.sharepoint.com/"

#Hard code credentials here if you need to
$username = ""
$password = ""

if ([string]::IsNullOrEmpty($username))
{
    $getcred = Get-Credential -credential $null
    $username = $getcred.UserName
    $securepassword = $getcred.Password
}
else
{
    $securepassword = $password | ConvertTo-SecureString -AsPlainText -force
    $getcred = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securepassword
}

$cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securepassword)

#Connect###########################################################################################

Write-Host "Trying to connect to SharePoint Online site $urladmin" -foregroundcolor black -backgroundcolor yellow  

try 
{
    Connect-SPOService -Url $urladmin -credential $getcred
    Write-Host "Done!" -foregroundcolor black -backgroundcolor Green 
}
catch
{
    Write-Host "Error: $_.Exception.Message" -ForegroundColor Red
}

Write-Host "Trying to authenticate to SharePoint Online Tenant site $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow  
$global:Context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$global:Context.Credentials = $cred
$global:Context.RequestTimeOut = 5000 * 60 * 10;
$web = $global:Context.Web
$site = $global:Context.Site 
$global:Context.Load($web)
$global:Context.Load($site)
try
{
    $global:Context.ExecuteQuery()
    Write-Host "Done! Use `$Context" -foregroundcolor black -backgroundcolor Green 
}
catch
{
    Write-Host "Not able to authenticate to SharePoint Online $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    return
}