<#
.SYNOPSIS
	Generates the STSSyncURLs for a Sharepoint Online Site using CSOM. Should be run from a SharePoint server.
	TODO: identify which DLL contains the microsoft.sharepoint.utilities.spencode method to allow running without a SharePoint install.
.DESCRIPTION
	This script is useful for generating a long list of STSSYNC urls without tediously connecting dozens of lists to Outlook, sharing them, and extracting
	the URL in that fashion. The output is a PSOBject which can be piped to a CSV for further manipulation, or to a separate script that creates appropriate GPOs
	or .REG files for further automation.
	Requires load-csomproperties.ps1 from: https://gist.github.com/glapointe/cc75574a1d4a225f401b
.EXAMPLE
	.\get-spostssyncurls | export-csv "urls.csv"
.NOTES
	Author	: Quinten Steenhuis
.LINK
	https://nonprofittechy.blogspot.com
.LINK
	https://msdn.microsoft.com/en-us/library/dd957390(v=office.12).aspx
#>

$cwd = split-path $myinvocation.mycommand.path
. (join-path $cwd "load-csomproperties.ps1")
<#
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll"
#>

$stsbase = "stssync://sts/?ver=1.1"
$rootURL = "https://MYCOMPANY.sharepoint.com"
$baseUrl = "https://MYCOMPANY.sharepoint.com/Units"
$adminURL = "https://MYCOMPANY-admin.sharepoint.com"

$username = "user@MYCOMPANY.com"
$password = Read-Host -Prompt "Enter password" -AsSecureString 


$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($baseUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password) 
$clientContext.Credentials = $credentials

if (!$clientContext.ServerObjectIsNull.Value) 
{ 
    Write-Host "Connected to SharePoint Online site: '$baseURL'" -ForegroundColor Green 
} 

$rootWeb = $clientContext.Web
$childWebs = $rootWeb.Webs
$clientContext.Load($rootWeb)
$clientContext.Load($childWebs)
$clientContext.ExecuteQuery()

function get-stssyncurl ($list) {
	
	$encoded = $stsbase + "&type="+$list.type+"&cmd=add-folder"	
	$ht = @{
		"base-url" = $list.baseURL;
		"list-url" = $list.listurl;
		"guid" = "{" + $list.guid + "}";
		"site-name" = $list.sitename;
		"list-name" = $list.listname
	}
	
	foreach ($h in $ht.getEnumerator()) {
		$encoded += "&" + $h.key + "=" + [microsoft.sharepoint.utilities.spencode]::urlencode($h.value)
	}
	
	foreach ($rep in $repl) {
		$encoded = $encoded -replace $rep.key,$rep.value	
	}
	
	# these characters appear to need /additional/ escaping, after urlencoding
	# and the hex value should be preceded by a vertical bar |
	# & \ [ ] | 
	$replPostRegex = "(%26)|(%5C)|(%5B)|(%5D)|(%7C)"	
	$encoded = $encoded -replace $replPostRegex,'|$1'
	return $encoded
}

function processWeb($web)
{
	# Skip some lists we'll never want to connect to Outlook
	$exclude = @(
		"Site Assets",
		"Site Pages",
		"MicroFeed",
		"Master Page Gallery",
		"Workflow History",
		"Workflow Tasks",
		"Workflows",
		"wfsvc",
		"Images",
		"Pages",
		"Composed Looks",
		"Master Pages",
		"Long Running Operation Status",
		"Notification List"
		"Project Policy Item List",
		"Quick Deploy Items",
		"Relationships List",
		"Reusable Content",
		"Site Collection Documents",
		"Site Collection Images",
		"Solution Gallery",
		"Style Library",
		"Suggested Content Browser Locations",
		"TaxonomyHiddenList",
		"Theme Gallery",
		"Translation Packages",
		"Translation Status",
		"User Information List",
		"Variation Labels",
		"Web Part Gallery",
		"wfpub",
		"appdata",
		"Form Templates",
		"Cache Profiles",
		"Access Requests",
		"Content and Structure Reports",
		"Content type publishing error log",
		"Converted Forms",
		"Device Channels",
		"fpdatasources",
		"List Template Gallery"
	)

    $lists = $web.Lists
    $clientContext.Load($web)
    $clientContext.Load($lists)

    $clientContext.ExecuteQuery()
	$lists = $lists | where {-not ($exclude -contains $_.Title)}

	$siteName = $web.Title

    foreach ($list in $Lists)
    {
		# using Gary Lapointe's implementation of lambda function to simplify loading Sharepoint Online CSMO properties -- Views property not loaded by default
		Load-CSOMProperties -object $list -propertyNames @("Views","rootFolder")
		$clientContext.ExecuteQuery()
		$views = @()
		foreach ($view in $list.Views) {
			$views += $view.Title
		}

		# if the baseType is "genericlist" we need to try searching the view types to determine which type of list this is
		# for Outlook equivalent
		
		$type = switch ($list.basetype) {
			"DocumentLibrary" {"documents"; break}
			"GenericList" {
				switch ($views) {
					"Threaded" {"discussions"; break}
					"Calendar" {"calendar"; break}
					"All Tasks" {"tasks"; break}
					"All Contacts" {"contacts"; break}
				};
				break
			}
		}
		
		$listURL = ($rootURL + $list.rootFolder.ServerRelativeURL).substring($web.URL.length)
		
		$retval = @{
			"type" = $type;
			"views" = $list.Views;
			"baseURL" = $web.URL;
			"listurl" = $listURL;
			"GUID" = $list.ID;
			"listName" = $list.Title;
			"siteName" = $siteName;
		}
		new-object psobject -property $retval
    }
}

$lists = @()
$lists += processWeb($rootweb)
foreach ($childWeb in $childWebs)
{
		$lists += processWeb($childWeb)
}

foreach ($list in $lists) {
	new-object psobject -property @{
		"siteName" = $list.siteName;
		"listName" = $list.listName;
		"stssyncurl" = (get-stssyncurl -list $list);
	}
}