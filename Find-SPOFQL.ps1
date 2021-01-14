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
    Use of several FQL operators have been deprecated in SharePoint Online (as of Feb 2021).
    https://techcommunity.microsoft.com/t5/microsoft-search-blog/we-re-making-changes-to-search-in-sharepoint-online/ba-p/1971119
      
    The script examines the query template for Result Sources and Query Rules for FQL keywords.
    To run, specificy a user name with Admin privileges.  Then choose to scan "All" sites or
    a single site.  When scanning all sites, add the -All switch and specify the tenant admin 
    site (typically has the format https://<my_tenant>-admin.sharepoint.com).  To scan a single 
    site, specify the site Url (e.g. https://<my_tenant>.sharepoint.com) and the SiteCollection
    or Site scope.
	
    FQL Operators
    FAST Query Language (FQL) operators are keywords that specify Boolean operations or other 
    constraints to operands. The FQL operator syntax is as follows:

    [property-spec:]operator(operand [,operand]* [, parameter="value"]*)

    This script looks for the pattern "<operator>(" using the following reserved FQL keywords:
    "and", "or", "any", "andnot", "count", "decimal", "rank", "near", "onear", "int", "in32", 
    "int64", "float", "double", "datetime", "max", "min", "range", "phrase", "scope", "filter", 
    "not", "string", "starts-with", "ends-with", "equals", "words", "xrank"

    Also the following deprecated operators:
    "count","filter"

    Lastly, the following deprecated parameters for the specific operator:
    "string": ["linguistics","wildcard"]}

.NOTES
	=========================================
	File Name 	: Find-SPOFQL.ps1
    Author		: Eric Dixon, Brad Schlintz (he did something, really!)

	Requires	: 
		PowerShell Version 5.0 or greater, SharePointPnPPowerShellOnline
	
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
    .\Find-SPOFQL.ps1 -UserName admin@MyTenant.onmicrosoft.com -Site https://MyTenant.sharepoint.com -Scope SiteCollection

#>

if($host.Version.Major -lt 5){

Add-Type -TypeDefinition @"
    public enum ResourceType { 
        QueryRule;
        ResultSource;
    }
"@

} else {

enum ResourceType { 
    QueryRule;
    ResultSource;
}

}

$fql_reserved_keywords = @("and", "or", "any", "andnot", "count", "decimal", "rank", "near", "onear", "int", "in32", "int64", "float", "double", "datetime", "max", "min", "range", "phrase", "scope", "filter", "not", "string", "starts-with", "ends-with", "equals", "words", "xrank")
$fql_deprecated_operators = @("count","filter","any")
$fql_deprecated_paramters = ConvertFrom-Json '{"string":["linguistics","wildcard"]}'

function TemplateHasFQLOperator {
param($template, $keywords)
    if([string]::IsNullOrEmpty($template)){
        return $null
    }
    foreach($k in $keywords){
        if($template.Contains("$k(")){
            return $k
        }
    }
    return $null
}


function TemplateHasFQLParameter {
param($template, $parameters)
    if([string]::IsNullOrEmpty($template)){
        return $null, $null
    }
    foreach($k in $parameters.PSObject.Properties.Name){
        if($template.Contains("$k(")){
            foreach($p in $parameters.$k) {
                if($template.Contains("$p=")) {
                    return $k,$p
                }
            }
        }
    }
    return $null, $null
}


function TestForFQL{
param($template,$type,$displayName)

    switch($type) {
        [ResourceType]::QueryRule {$resourceType = "Query Rule"}
        [ResourceType]::ResultSource {$resourceType = "Result Source"}
    }

    $fql_op = TemplateHasFQLOperator -template $template -keywords $fql_reserved_keywords
    if($null -ne $fql_op){
        $dep_op = TemplateHasFQLOperator -template $template -keywords $fql_deprecated_operators
        if($null -ne $dep_op){
            Write-Host "   * Found deprecated FQL operator '$dep_op' in query template for $resourceType '$displayName'" -ForegroundColor Red
            Write-Host "     $template" -ForegroundColor Red
        } else {
            $param_op, $dep_param = TemplateHasFQLParameter -template $template -parameters $fql_deprecated_paramters
            if ($null -ne $param_op -and $null -ne $dep_param) {
                Write-Host "   * Found deprecated FQL parameter '$dep_param' for operator '$param_op' in query template for $resourceType '$displayName'" -ForegroundColor Red
                Write-Host "     $template" -ForegroundColor Red
            } else {
                Write-Host "   * Found FQL operator '$fql_op' in query template for $resourceType '$displayName'" -ForegroundColor Yellow
                Write-Host "     $template" -ForegroundColor Yellow
            }
        }
    }
    
}


function ParseQueryRules {
param($xml)

    $query_rules = $xml.SearchConfigurationSettings.SearchQueryConfigurationSettings.SearchQueryConfigurationSettings.QueryRules.QueryRule
    foreach($rule in $query_rules) {
        $query_template = $rule.CreateResultBlockActions.OrderedItems.CreateResultBlockAction._QueryTransform._QueryTemplate

        TestForFQL -template $query_template -type [ResourceType]::QueryRule -displayName $rule._DisplayName
    }
}


function ParseResultSources {
param($xml)

    $sources = $xml.SearchConfigurationSettings.SearchQueryConfigurationSettings.SearchQueryConfigurationSettings.Sources.Source
    foreach($source in $sources) {
        $query_template = $source.QueryTransform._QueryTemplate

        TestForFQL -template $query_template -type [ResourceType]::ResultSource -displayName $source.Name
    }
}


function Main {
    
    if($All){
        if(-not $TenantAdminSite.StartsWith("https://")) {
            $TenantAdminSite = "https://" + $TenantAdminSite
        }
        Connect-PnPOnline $TenantAdminSite -UseWebLogin

        $sites = @($TenantAdminSite)
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
    Write-Host
    Write-Host "Microsoft recommends, where applicable, using the default SharePoint query language, KQL where your business requirements can be similarly met."
    Write-Host "https://techcommunity.microsoft.com/t5/microsoft-search-blog/we-re-making-changes-to-search-in-sharepoint-online/ba-p/1971119"
}

Main
