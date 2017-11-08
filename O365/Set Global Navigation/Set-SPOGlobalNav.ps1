<#
.SYNOPSIS
  Creates navigation from an XML template.
.DESCRIPTION
  This script takes an XML template and builds out the navigation described in the template.
.PARAMETER URL
  The URL for the SharePoint Online tenant.
.PARAMETER XML
  XML file describing the navigation.
  
  <!-- Site Collections -->
  <sites>
	<!-- Site Collection
		 Attributes:
		 Url - Url for the site collection -->
	<site Url="https://domain.sharepoint.com/" >
		<!-- Global Nav -->
		<globalnav>
			<!-- Item on the Global Nav Bar.
				 Attributes:
				 Title - Text for the Item.
				 Url - Url for the Item.  Optional or use empty string for no link.
			-->
			<header Title="Link 1" Url="https://domain.sharepoint.com/link1">
				<!-- Sub Item
					 Attributes:
					 Title - Text for the Item.
					 Url - Url for the Item.  Optional or use empty string for no link.
				-->
				<nav Title="Link 1a" Url="https://domain.sharepoint.com/link1a" />
				<nav Title="Link 1b" Url="https://domain.sharepoint.com/link1b" />
			</header>
			<header Title="Link 2" Url="https://domain.sharepoint.com/link2">
				<!-- Sub Item
					 Attributes:
					 Title - Text for the Item.
					 Url - Url for the Item.  Optional or use empty string for no link.
				-->
				<nav Title="Link 2a" Url="https://domain.sharepoint.com/link2a" />
				<nav Title="Link 2b" Url="https://domain.sharepoint.com/link2b" />
			</header>
			<!-- Item on the Global Nav Bar without Sub Items. -->
			<header Title="Home" Url="http://domain.sharepoint.com/"/>
	  </globalnav>     
	</site>
  </sites>
.NOTES
  Version:        1.0
  Author:         John Weatherley
  Creation Date:  November 8, 2017
  Purpose/Change: Initial script development
.LINK
  http://www.johnweatherley.com
  https://github.com/johnweatherley/sharepoint-powershell
.EXAMPLE
  Set-SPOGlobalNav -url https://domain.sharepoint.com -xml input.xml
#>

Param(
	[Parameter(Mandatory=$true)]
	[String]
	$url,
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_})]
	[String]
	$xml
)

############################################################################################

Write-Host "Load CSOM libraries" -foregroundcolor black -backgroundcolor yellow
Set-Location $PSScriptRoot
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Publishing.dll")

Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 

#Function - Add Navigation###########################################################################################

function Add-MenuNavigation($xmlNavs)
{
	$Nodes = $Context.Web.Navigation.TopNavigationBar		
	$NavigationNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
	$NavigationNode.Title = $xmlNavs.Title
    
    if ($xmlNavs.Url)
    {
	    $NavigationNode.Url = $xmlNavs.Url
    }
	$NavigationNode.AsLastNode = $true
	$headerNode = $Nodes.Add($NavigationNode)
	$Context.Load($headerNode)
	
    try
    {
		$Context.ExecuteQuery()	
		Write-Host "Adding" $xmlNavs.Title "to Global Navigation Completed" -foregroundcolor black -backgroundcolor green	
    }
    catch
    {
		Write-Host Error Adding $xmlNavs.Title to Global Navigation $_.Exception.Message -foregroundcolor black -backgroundcolor Red
    }
	
	if ($xmlNavs.nav.Count -eq 0) { return }
	
	$xmlNavs.nav|
	ForEach-Object {
		$Node = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
		$Node.Title = $_.Title
		$Node.Url = $_.Url
		$Node.AsLastNode = $true			
		$Context.Load($headerNode.Children.Add($Node))
		
		try
		{
			$Context.ExecuteQuery()	
			Write-Host "Adding" $Node.Title "to Global Navigation Completed" -foregroundcolor black -backgroundcolor green	
		}
		catch
		{
			Write-Host Error Adding $Node.Title to Global Navigation $_.Exception.Message -foregroundcolor black -backgroundcolor Red
		}
	}	
}

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

#Test Connect###########################################################################################

Write-Host "Trying to authenticate to SharePoint Online Tenant site $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow  
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$Context.Credentials = $credentials
$Context.RequestTimeOut = 5000 * 60 * 10;
$web = $Context.Web
$site = $Context.Site
$Context.Load($web)
$Context.Load($site)
try
{
    $Context.ExecuteQuery()
    Write-Host "Done!" -foregroundcolor black -backgroundcolor Green 
}
catch
{
    Write-Host "Not able to authenticate to SharePoint Online $_.Exception.Message" -foregroundcolor black -backgroundcolor Red
    return
}

#Loop through the sites in the XML file###########################################################################################

#Read Data from XML file
Write-Host "Loading XML file" -foregroundcolor black -backgroundcolor yellow

[xml]$xmlContent = (Get-Content $xml)
if (-not $xmlContent)
{
    Write-Host "XML was not loaded successfully." -foregroundcolor black -BackgroundColor Red	
    return
}
	
Write-Host "Done!" -foregroundcolor black -backgroundcolor Green

#Start the loop
$xmlContent.sites.site |
ForEach-Object {
	Write-Host "Connecting to " $_.Url -foregroundcolor black -backgroundcolor yellow	
	$Context = New-Object Microsoft.SharePoint.Client.ClientContext($_.Url)
	$Context.Credentials = $credentials
	$context.RequestTimeOut = 5000 * 60 * 10;
	$web = $context.Web
	$site = $context.Site
	$context.Load($web)
	$context.Load($site)
	try
	{
		$context.ExecuteQuery()
		Write-Host "Done!" -foregroundcolor black -backgroundcolor Green 
	}
	catch
	{
		Write-Host "Not able to connect to " $_.Url " $_.Exception.Message" -foregroundcolor black -backgroundcolor Red	
		return
	}
	
	#Add Menu to Global Navigation
	$_.globalnav.header |
	ForEach-Object {
		Write-Host "Adding" $_.Title "to Global Navigation Starting..." -foregroundcolor black -backgroundcolor yellow
		Add-MenuNavigation ($_)	
	}
}