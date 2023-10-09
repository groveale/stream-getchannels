## Before you start go to this URL and login https://web.microsoftstream.com/?NoSignUpCheck=1
## Then go to this URL https://ukwe-3.api.microsoftstream.com/api/videos?NoSignUpCheck=1&$top=100&$orderby=publishedDate%20desc&$expand=creator,events&$filter=published%20and%20(state%20eq%20%27Completed%27%20or%20contentSource%20eq%20%27livestream%27)&adminmode=true&api-version=1.4-private&$skip=0
## Open the developer tools and go to Application, copy the Cookie values into cookies.csv
[string]$region = "ukwe-3"
$cookiesCSVFilePath = ".\cookies.csv"
$outputCSVFilePath = ".\outputVideos.csv"
[string]$StreamPortalVideoViewRoot = "https://web.microsoftstream.com/video/"

$WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$cookies = Import-CSV $cookiesCSVFilePath

foreach ($cookie in $cookies) {
    $newCookie = New-Object System.Net.Cookie
    $newCookie.Name = $cookie.Name
    $newCookie.Value = $cookie.Value
    $newCookie.Domain = $cookie.Domain
    $WebSession.Cookies.Add($newCookie)
}

$headerParams = @{
    "ContentType" = "application/json"
}

$i = 0
$continue = $true
while ($continue -eq $true) {
    $i++
    $skip = ($i - 1) * 100
    Write-Output "Getting videos $skip - $($skip + 100)"
    $queryOptions = @{
        Method     = "GET"
        URI        = "https://$region.api.microsoftstream.com/api/videos?NoSignUpCheck=1&`$top=100&`$orderby=publishedDate%20desc&`$expand=creator,events&`$filter=published%20and%20(state%20eq%20%27Completed%27%20or%20contentSource%20eq%20%27livestream%27)&adminmode=true&api-version=1.4-private&`$skip=$skip"
        Headers    = $headerParams
        WebSession = $WebSession
    } 

    try {
        $videoResult = Invoke-RestMethod @queryOptions -ErrorAction Stop
    }
    catch {
        Write-Output $_.exception.message
    }

    if ($videoResult.value.length -eq 0) {
        $continue = $false
    }
    else {
        foreach ($video in $videoResult.value) {
            $videoInfo = [PSCustomObject]@{
                VideoID               = $video.id
                VideoName             = $video.name
                VideoURL              = $StreamPortalVideoViewRoot + $video.id
                VideoCreatorName      = $video.creator.name
                VideoCreatorEmail     = $video.creator.mail
                VideoCreationDate     = $video.created
                VideoModificationDate = $video.modified
                VideoLikes            = $video.metrics.likes
                VideoViews            = $video.metrics.views
                VideoComments         = $video.metrics.comments
                Videodescription      = $video.description
                VideoDuration         = $video.media.duration
                VideoHeight           = $video.media.height
                VideoWidth            = $video.media.width
                VideoIsAudioOnly      = $video.media.isAudioOnly
                VideoContentType      = $video.contentType
            }
        
            [array]$videoList += $videoInfo
        
        }
    }
}

$videoList | Export-CSV -NoTypeInformation $outputCSVFilePath