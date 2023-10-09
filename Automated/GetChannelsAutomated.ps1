## Before you start go to this URL and login https://web.microsoftstream.com/?NoSignUpCheck=1
## Then go to this URL https://ukwe-3.api.microsoftstream.com/api/channels?NoSignUpCheck=1&$top=100&$orderby=metrics%2Fvideos%20desc&$expand=creator,group&api-version=1.3-private&$skip=0
## Open the developer tools and go to Application, copy the Cookie values into cookies.csv
[string]$region = "ukwe-3"
$datestring = (get-date).ToString("yyyyMMdd-hhmm")
$cookiesCSVFilePath = ".\cookies-$region.csv"
$outputCSVFilePath = ".\outputChannels-$datestring.csv"
[string]$StreamPortalChannelRoot = "https://web.microsoftstream.com/channel/"

$WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$cookies = Import-CSV $cookiesCSVFilePath

$resfreshTimeOutMins = 0 

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

# Start the stopwatch
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Convert the pause duration to milliseconds
$pauseDurationMilliseconds = $resfreshTimeOutMins * 60 * 1000

$i = 0
$continue = $true
while ($continue -eq $true) {
    $i++
    $skip = ($i - 1) * 100
    Write-Output "Getting channels $skip - $($skip + 100)"
    $queryOptions = @{
        Method     = "GET"
        URI        = "https://$region.api.microsoftstream.com/api/channels?NoSignUpCheck=1&`$top=100&`$orderby=metrics%2Fvideos%20desc&`$expand=creator,group&api-version=1.3-private&`$skip=$skip"
        Headers    = $headerParams
        WebSession = $WebSession
    } 

    try {
        $channelResult = Invoke-RestMethod @queryOptions -ErrorAction Stop
    }
    catch {
        Write-Output $_.exception.message
    }

    if ($channelResult.value.length -eq 0) {
        $continue = $false
    }
    else {
        foreach ($channel in $channelResult.value) {
            $channelInfo = [PSCustomObject]@{
                ChannelID               = $channel.id
                ChannelName             = $channel.name
                ChannelURL              = $StreamPortalChannelRoot + $channel.id
                ChannelDefault          = $channel.isDefault
                ChannelFollowers        = $channel.metrics.follows
                ChannelVideos           = $channel.metrics.videos
                ChannelCreatorName      = $channel.creator.name
                ChannelCreatorEmail     = $channel.creator.mail
                ChannelCreationDate     = $channel.created
                ChannelModificationDate = $channel.modified
                O365GroupId             = $channel.group.id
                O365GroupName           = $channel.group.name
                O365GroupEmail          = $channel.group.aadGroup.mail
                O365GroupTotalChannels  = $channel.group.metrics.channels
                O365GroupTotalVideos    = $channel.group.metrics.videos
            }
        
            [array]$channelList += $channelInfo
        
        }
    }

    ## Do we need to pause and upate the cookies?
    if ($stopwatch.ElapsedMilliseconds -gt $pauseDurationMilliseconds)
    {
        ## Pause the execution
        Write-host " -----------------------------------------" -ForegroundColor DarkGreen
        Write-Host "Pausing as $resfreshTimeOutMins minutes have elapsed - Time to refresh the cookies"
        Write-Host "Update the cookies.csv file with new cookies"
        Read-Host -Prompt "Press Enter to continue ...."

        ## Read the cookies again
        $cookies = Import-CSV $cookiesCSVFilePath
        
        foreach ($cookie in $cookies) {
            $oldCookie = $WebSession.Cookies.GetAllCookies() | Where-Object { $_.Name -eq $cookie.Name }
            $oldCookie.Value = $cookie.Value
        }

        Write-Host "Cookie values uppdated - Resuming execution" -ForegroundColor DarkGreen

        ## Reset the stopwatch
        $stopwatch.Reset()
    }
}

$channelList | Export-CSV -NoTypeInformation $outputCSVFilePath