# Scripts
    Use of several FQL operators and parameters have been deprecated in SharePoint Online 
    (as of Feb 2021).
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

    This script looks for: 
        1.  The pattern "<operator>(" using the following reserved FQL keywords:
            "and", "or", "any", "andnot", "count", "decimal", "rank", "near", "onear", "int", 
            "in32", "int64", "float", "double", "datetime", "max", "min", "range", "phrase", 
            "scope", "filter", "not", "string", "starts-with", "ends-with", "equals", "words", 
            "xrank"

        2.  The following deprecated operators:
            "count","filter","any"

        3.  The following deprecated parameters for the specific operator:
            "string": ["linguistics","wildcard"]

    EXAMPLES
	# All Sites
	.\Find-SPOFQL.ps1 -UserName admin@MyTenant.onmicrosoft.com -TenantAdminSite https://MyTenant-admin.sharepoint.com/ -All
	# Specific site and scope
	.\Find-SPOFQL.ps1 -UserName admin@MyTenant.onmicrosoft.com -Site https://MyTenant.sharepoint.com -Scope SiteCollection

    EXAMPLE OUTPUT
	   * Checking site 'https://MyTenant-admin.sharepoint.com' at scope 'Subscription'...
	   * Found FQL operator 'string' in query template for Query Rule 'FQL Query Rule - Result Block'
	     title:string({subjectTerms})
	   * Found deprecated FQL operator 'filter' in query template for Result Source 'FQL Deprecated Opearator'
	     filter(and({searchTerms}, "fql"))
	   * Found deprecated FQL parameter 'linguistics' for operator 'string' in query template for Result Source 'FQL Deprecated Parameter'
	     string("{searchTerms}", linguistics="off")  
	   * Found FQL operator 'string' in query template for Result Source 'FQL Test Source'
	     title:string({searchTerms})
