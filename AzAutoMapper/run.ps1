using namespace System.Net
param($Request, $TriggerMetadata)
#
#
$TenantID = $Request.Query.TenantID
$User = $Request.Query.Username
 
$Body = @{
    'resource'      = 'https://graph.microsoft.com'
    'client_id'     = $ENV:ApplicationId
    'client_secret' = $ENV:ApplicationSecret
    'grant_type'    = "client_credentials"
    'scope'         = "openid"
}
$ClientToken = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$($TenantID)/oauth2/token" -Body $Body -ErrorAction Stop
$Headers = @{ "Authorization" = "Bearer $($ClientToken.access_token)" }
$UserID = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($User)" -Headers $Headers -Method GET -ContentType "application/json").email
 
$AllSitesURI = "https://graph.microsoft.com/v1.0/sites?search=$($UserID)"
$Sites = (Invoke-RestMethod -Uri $AllSitesURI -Headers $Headers -Method GET -ContentType "application/json").value
$MemberOf = foreach ($Site in $Sites) {
    $SiteWebID = $Site.id -split ','
    $SiteDrivesUri = "https://graph.microsoft.com/v1.0/sites/$($SiteWebID[1])/lists"
    $SitesDrivesReq = (Invoke-RestMethod -Uri $SiteDrivesUri -Headers $Headers -Method GET -ContentType "application/json").value | where-object { $_.Name -eq "Shared Documents" }
    if ($Site.description -like "*no-auto-map*") { continue }
    if ($null -eq [System.Web.HttpUtility]::UrlEncode($SitesDrivesReq.id)) { continue }
    [pscustomobject] @{
        SiteID    = [System.Web.HttpUtility]::UrlEncode("{$($SiteWebID[1])}")
        WebID     = [System.Web.HttpUtility]::UrlEncode("{$($SiteWebID[2])}")
        ListID    = [System.Web.HttpUtility]::UrlEncode("{$($SitesDrivesReq.id)}")
        WebURL    = [System.Web.HttpUtility]::UrlEncode($Site.webUrl)
        WebTitle  = [System.Web.HttpUtility]::UrlEncode($($Site.Name)).Replace("+", "%20")
        ListTitle = [System.Web.HttpUtility]::UrlEncode($SitesDrivesReq.name)
    }
 
}
 
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = $MemberOf
    })
