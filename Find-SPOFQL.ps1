param(
    [Parameter(Mandatory=$true)]
    [string]$UserName,
	[Parameter(Mandatory=$false, ParameterSetName='All')]
    [switch]$All,
	[Parameter(Mandatory=$true, ParameterSetName='All')]
    [string]$TenantAdminSite,
	[Parameter(Mandatory=$false, ParameterSetName='Site')]
    [string]$Site,
	[Parameter(Mandatory=$true, ParameterSetName='Site')]
    [ValidateSet('SiteCollection','Site')]
    [string]$Scope
)

<#
.SYNOPSIS 
	Scans a single site and scope or ALL sites for FQL used in Result Sources and Query Rules.
	
.DESCRIPTION 
    Use of FQL has been deprecated in SharePoint Online.  This script will scan a single site 
    and scope or ALL sites for FQL used in Result Sources and Query Rules.  
    The script examines the query template for Result Sources and Query Rules for FQL keywords.
    To run, specificy a user name with Admin privileges.  Then choose to scan "All" sites or
    a single site.  When scanning all sites, add the -All switch and specify the tenant admin 
    site (typically has the format https://<my_tenant>-admin.sharepoint.com).  To scan a single 
    site, specify the site Url (e.g. https://<my_tenant>.sharepoint.com) and the SiteCollection
    or Site scope.
	
.NOTES
	=========================================
	File Name 	: Find-SPOFQL.ps1
    Author		: Eric Dixon, Brad Schlintz (he did something, really!)

	Requires	: 
		PowerShell Version 4.0 or greater, SharePointPnPPowerShellOnline
	
	========================================================================================
	This Sample Code is provided for the purpose of illustration only and is not intended to 
	be used in a production environment.  
	
		THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY
		OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
		WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.

	We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to 
	reproduce and distribute the object code form of the Sample Code, provided that You agree:
		(i) to not use Our name, logo, or trademarks to market Your software product in 
			which the Sample Code is embedded; 
		(ii) to include a valid copyright notice on Your software product in which the 
			 Sample Code is embedded; 
		and 
		(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against
			  any claims or lawsuits, including attorneys' fees, that arise or result from 
			  the use or distribution of the Sample Code.

	========================================================================================
	
.INPUTS
	UserName - SPO Admin User
    Parameter Set "All"
        TenantAdminSite - Url for tenant admin site
        All - Switch to scan all sites in tenant
    Parameter Set "Site"
        Site - Specific site to scan
        Scope - Scope to scan, SiteCollection or Site

.OUTPUTS
	Query Rules and/or Result Sources that may have FQL

.EXAMPLE
	.\Find-SPOFQL.ps1 -UserName admin@MyTenant.onmicrosoft.com -TenantAdminSite https://MyTenant-admin.sharepoint.com/ -All
    .\Find-SPOFQL.ps1 -UserName admin@m365x342170.onmicrosoft.com -Site https://MyTenant.sharepoint.com -Scope SiteCollection

#>


$fql_reserved_keywords = @("and", "or", "any", "andnot", "count", "decimal", "rank", "near", "onear", "int", "in32", "int64", "float", "double", "datetime", "max", "min", "range", "phrase", "scope", "filter", "not", "string", "starts-with", "ends-with", "equals", "words", "xrank")


function TemplateHasFQL {
param($template)
    if([string]::IsNullOrEmpty($template)){
        return $false
    }
    foreach($keyword in $fql_reserved_keywords){
        if($template.Contains("$keyword(")){
            return $true
        }
    }
    return $false
}

function ParseQueryRules {
param($xml,$site)

    $query_rules = $xml.SearchConfigurationSettings.SearchQueryConfigurationSettings.SearchQueryConfigurationSettings.QueryRules.QueryRule
    foreach($rule in $query_rules) {
        $query_template = $rule.CreateResultBlockActions.OrderedItems.CreateResultBlockAction._QueryTransform._QueryTemplate
        if(TemplateHasFQL -template $query_template){
            Write-Host "   * Found potential FQL in query template for Query Rule '$($rule._DisplayName)'" -ForegroundColor Yellow
            Write-Host "     $query_template" -ForegroundColor Yellow
        }
    }
}

function ParseResultSources {
param($xml)

    $sources = $xml.SearchConfigurationSettings.SearchQueryConfigurationSettings.SearchQueryConfigurationSettings.Sources.Source
    foreach($source in $sources){
        $query_template = $source.QueryTransform._QueryTemplate
        if(TemplateHasFQL -template $query_template){
            Write-Host "   * Found potential FQL in query template for Result Source '$($source.Name)'" -ForegroundColor Yellow
            Write-Host "     $query_template" -ForegroundColor Yellow
        }
        
    }
}


function Main {
#    $creds = (Get-Credential -Message "Please enter admin credientials for SPO" -UserName $Username)

#    if(-not $creds) {
#        Write-Host "Please supply valid credentials" -ForegroundColor Red
#        exit
#    }

    if($All){
        if(-not $TenantAdminSite.StartsWith("https://")) {
            $TenantAdminSite = "https://" + $TenantAdminSite
        }
#        Connect-SPOService $TenantAdminSite -Credential $creds
        Connect-PnPOnline $TenantAdminSite -UseWebLogin

        $sites = @($TenantAdminSite)
#        $sites += (Get-SPOSite).Url
        $sites += (Get-PnPTenantSite).Url
        $scopes = @("SiteCollection","Site")

    } else {
        if([string]::IsNullOrEmpty($Site) -or [string]::IsNullOrEmpty($Scope)) {
            Write-Host "Must enter a Site name and scope"
            exit
        }
        if(-not $Site.StartsWith("https://")) {
            $Site = "https://" + $Site
        }

        $sites = @($Site)
        $scopes = @($Scope)
    }

    Write-Host "Scanning search configuration for FQL in query templates..."

    foreach($s in $sites) {

#        Connect-PnPOnline $s -Credentials $creds
        Connect-PnPOnline $s -UseWebLogin

        if($s -eq $TenantAdminSite){
            Write-Host " * Checking site '$s' at scope 'Subscription'..."
            try {
                [xml]$xml = Get-PnPSearchConfiguration -Scope Subscription
            } catch {
                Write-Host " * Caught exception while trying to get configuration from site '$s' and scope '$scope'." -ForegroundColor Red
                Write-Host "$_" -ForegroundColor Red
                continue
            }

            if(-not $xml) {
                Write-Host " * Search Configuration for site '$s' and scope '$Scope' not found." -ForegroundColor Red
                continue
            }

            ParseQueryRules -xml $xml
            ParseResultSources -xml $xml

        } else {
            foreach($sc in $scopes){

                Write-Host " * Checking site '$s' at scope '$sc'..."

                try {
                    if($sc -eq "SiteCollection") {
                        [xml]$xml = Get-PnPSearchConfiguration -Scope Site
                    } else {
                        # site
                        [xml]$xml = Get-PnPSearchConfiguration -Scope Web
                    }
                } catch {
                    Write-Host " * Caught exception while trying to get configuration from site '$s' and scope '$scope'." -ForegroundColor Red
                    Write-Host "$_" -ForegroundColor Red
                    continue
                }

                if(-not $xml) {
                    Write-Host " * Search Configuration for site '$s' and scope '$Scope' not found." -ForegroundColor Red
                    continue
                }

                ParseQueryRules -xml $xml
                ParseResultSources -xml $xml
            }
        }

        Write-Host "   Done."
    }

    $creds = $null
    Write-Host "Done." -ForegroundColor Green
}

Main
