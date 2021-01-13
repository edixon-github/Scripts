param(
    [Parameter(Mandatory=$true)]
    [string]$UserName,
	[Parameter(Mandatory=$False, ParameterSetName='All')]
    [Parameter(Mandatory=$false)]
    [switch]$All,
	[Parameter(Mandatory=$true, ParameterSetName='All')]
    [string]$TenantAdminSite,
	[Parameter(Mandatory=$False, ParameterSetName='Site')]
    [string]$Site,
	[Parameter(Mandatory=$true, ParameterSetName='Site')]
    [ValidateSet('SiteCollection','Site')]
    [string]$Scope
)

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
            Write-Host "   * Found potential FQL in query template for Query Rule $($rule._DisplayName)" -ForegroundColor Yellow
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
            Write-Host "   * Found potential FQL in query template for Result Source $($source.Name)" -ForegroundColor Yellow
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
