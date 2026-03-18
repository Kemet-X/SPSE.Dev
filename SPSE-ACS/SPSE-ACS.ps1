Add-Type -Path "C:\path to\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\path to\Microsoft.SharePoint.Client.Runtime.dll"


$SiteUrl = "https://your-sp-site"
$ClientId = "your-app-id"
$ClientSecret = "your-app-secret"

function Get-Realm($url) {
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("Authorization", "Bearer ")
        $null = $wc.DownloadString($url)
    }
    catch {
        $authHeader = $_.Exception.Response.Headers["WWW-Authenticate"]
        return ($authHeader.Split("=", 2)[1].Trim('"'))
    }
}

function Get-AppToken($url, $clientId, $clientSecret) {
    $realm = Get-Realm $url
    $client = $clientId  
   $resource = "$clientId@$realm"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $client
        client_secret = $clientSecret
        resource      = $resource
    }

    $tokenEndpoint = "https://accounts.accesscontrol.windows.net/$realm/tokens/OAuth/2"
     
    $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ErrorAction SilentlyContinue
    return $response.access_token 
}

$AccessToken = Get-AppToken  $SiteUrl $ClientId $ClientSecret 

$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$ctx.ExecutingWebRequest += {
    param($source, $args)
    $args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer $AccessToken"
}

$web = $ctx.Web
$ctx.Load($web)  
$ctx.ExecuteQuery()

Write-Host "Connected. Site title: $($web.Title)" -ErrorAction SilentlyContinue 
