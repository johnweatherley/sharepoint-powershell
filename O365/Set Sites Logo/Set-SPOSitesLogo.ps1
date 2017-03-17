<#
.SYNOPSIS
  Applies a site logo at the Site Collection level from a CSV template.
.DESCRIPTION
  This script takes a CSV template and applies a site logo to a Site Collection as described in the template.
.PARAMETER CSV
  CSV file describing the Site Collections and logos.
  
  SiteURL,LogoUrl
  https://domain.sharepoint.com/,https://domain.sharepoint.com/SiteAssets/logo.png
  https://domain.sharepoint.com/sites/site1,https://domain.sharepoint.com/SiteAssets/logo.png
.NOTES
  Version:        1.0
  Author:         John Weatherley
  Creation Date:  March 17, 2017
  Purpose/Change: Initial script development
.LINK
  http://www.johnweatherley.com
  https://github.com/johnweatherley/sharepoint-powershell
.EXAMPLE
  Set-SPOSitesLogo -csv input.csv
#>

Param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_})]
	[String]
	$csv
)

#Setup###########################################################################################

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

$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securepassword)

#Load CSV###########################################################################################

$sites = Import-CSV $csv

#Loop CSV###########################################################################################

foreach ($row in $sites)
{
    Write-Host "$($row.SiteURL)" -ForegroundColor Green
    
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext("$($row.SiteURL)") 
    $ctx.Credentials = $credentials
 
    $web = $ctx.get_web()
    
    $ctx.Load($web)
    $ctx.ExecuteQuery()

    $web.SiteLogoUrl = "$($row.LogoURL)"
    $web.Update()
    $ctx.ExecuteQuery()    
}