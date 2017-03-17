<#
.SYNOPSIS
  Applies folder permissions from a CSV template.
.DESCRIPTION
  This script takes a CSV template and applies folder permissions to a Document Library described in the template.
.PARAMETER URL
  The URL for the SharePoint Online Site Collection containing the Document Library.
.PARAMETER CSV
  CSV file describing the folder permissions.
  
  Path,Group,Permission,BreakInheritance,CopyFromParent,ApplyToChildren
  Shared Documents/Folder1/,Folder1Group,Edit,TRUE,TRUE,TRUE
  Shared Documents/Folder2/,Folder1Group,Contributor,TRUE,TRUE,TRUE
  Shared Documents/Folder2/,Folder2Group,Contributor,TRUE,TRUE,TRUE
.NOTES
  Version:        1.0
  Author:         John Weatherley
  Creation Date:  March 17, 2017
  Purpose/Change: Initial script development
.LINK
  http://www.johnweatherley.com
  https://github.com/johnweatherley/sharepoint-powershell
.EXAMPLE
  Set-SPOSetFolderPermissions -url https://domain.sharepoint.com -csv input.csv
#>

Param(
	[Parameter(Mandatory=$true)]
	[String]
	$url,
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_})]
	[String]
	$csv
)

############################################################################################

Write-Host "Load CSOM libraries" -foregroundcolor black -backgroundcolor yellow
Set-Location $PSScriptRoot
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll")
Add-Type -Path (Resolve-Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Publishing.dll")

Write-Host "CSOM libraries loaded successfully" -foregroundcolor black -backgroundcolor Green 

#GetRole###########################################################################################

Function GetRole
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, Position = 1)]
        [Microsoft.SharePoint.Client.RoleType]$rType
    )

    $web = $Context.Web
    if ($web -ne $null)
    {
        $roleDefs = $web.RoleDefinitions
        $Context.Load($roleDefs)
        $Context.ExecuteQuery()
        $roleDef = $roleDefs | Where-Object { $_.RoleTypeKind -eq $rType }
        return $roleDef
    }
    return $null
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

$cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securepassword)

#Connect###########################################################################################

Write-Host "Trying to authenticate to SharePoint Online Tenant site $url and get ClientContext object" -foregroundcolor black -backgroundcolor yellow  
$Context = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$Context.Credentials = $cred
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

#Load CSV###########################################################################################

$folderpermissions = Import-Csv $csv

#Loop CSV###########################################################################################

foreach ($folderperm in $folderpermissions)
{
    $folder = $web.GetFolderByServerRelativeUrl($web.ServerRelativeUrl + $folderperm.path)
    $context.Load($folder)
    $context.ExecuteQuery()
    
    #get role
    $roleTypeObject = [Microsoft.SharePoint.Client.RoleType]$folderperm.Permission
    $roleObj = GetRole $roleTypeObject
    $usrRDBC = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($Context)
    $usrRDBC.Add($roleObj)

    #get group
    $group = $web.SiteGroups.GetByName($folderperm.Group)

    # https://msdn.microsoft.com/en-us/library/microsoft.sharepoint.client.securableobject.breakroleinheritance(v=office.15).aspx
    if ([System.Convert]::ToBoolean($folderperm.BreakInheritance))
    {
        $folder.ListItemAllFields.BreakRoleInheritance([System.Convert]::ToBoolean($folderperm.CopyFromParent), [System.Convert]::ToBoolean($folderperm.ApplyToChildren))
    }

    # apply roles
    $context.Load($folder.ListItemAllFields.RoleAssignments.Add($group, $usrRDBC))
    $folder.Update()

    $context.ExecuteQuery()
}