
[string]$PowerShellScriptFolder = pwd
[string]$streamJSONfolder = Join-Path -Path $PowerShellScriptFolder -ChildPath "ChannelsJSON"
Remove-Item -path $streamJSONfolder\* -include *.json -Force -Recurse

[string]$region = "euno-1"

[string]$StreamPortal = "https://web.microsoftstream.com/?NoSignUpCheck=1"
[string]$StreamPortalChannelRoot = "https://web.microsoftstream.com/channel/"
[string]$StreamAPIChannels100 = "https://$region.api.microsoftstream.com/api/channels?NoSignUpCheck=1&`$top=100&`$orderby=metrics%2Fvideos%20desc&`$expand=creator,group&api-version=1.3-private&`$skip="

[int]$Loopnumber = 2

Write-host " -----------------------------------------" -ForegroundColor Green
Write-Host "  =====>>>> PortalURL:", $StreamPortal
Start-Process -FilePath 'iexplore.exe' -ArgumentList $StreamPortal
Write-Host "      Enter your credentials to load O365 Stream portal" -ForegroundColor Magenta
Read-Host -Prompt "Press Enter to continue ...."

for($i=0;$i -lt $Loopnumber; $i++)
{
	Write-host " -----------------------------------------" -ForegroundColor Green
	$StreamAPIChannels100 = $StreamAPIChannels100 + $($i*100)
	Write-Host "  =====>>>> First 100 channels (from", $($i*100), "to", $(($i+1)*100), "):", $StreamAPIChannels100
	Start-Process -FilePath 'iexplore.exe' -ArgumentList $StreamAPIChannels100
	Write-Host "      Save the 100 channels (from", $($i*100), "to", $(($i+1)*100), ") into the folder $streamJSONfolder respecting the name channels100.json" -ForegroundColor Magenta
	Read-Host -Prompt "Press Enter to continue ...."
}

Write-host " -----------------------------------------" -ForegroundColor Green
$ChannelJSONFiles = Get-ChildItem -Path $streamJSONfolder -Recurse -Include *.json
[int]$channelscounter = 0
$ChanneljsonAggregateddata=@()
$data=@()

foreach($channelsjson in $ChannelJSONFiles)
{
	Write-host " -----------------------------------------" -ForegroundColor Green
	Write-Host "     =====>>>> JSON File:", $channelsjson, "- Path:", $channelsjson.FullName -ForegroundColor Yellow
	$Channeljsondata = Get-Content -Raw -Path $channelsjson.FullName | ConvertFrom-Json
	$ChanneljsonAggregateddata += $Channeljsondata
	Write-host " -----------------------------------------" -ForegroundColor Green
	#Write-Host "     =====>>>> Channel JSON Raw data:", $Channeljsondata -ForegroundColor green
	#Read-Host -Prompt "Press Enter to continue ...."
}

foreach($myChannel in $ChanneljsonAggregateddata.value)
{
	if($myChannel.metrics.videos -gt -1)
	{
		$channelscounter += 1
		$datum = New-Object -TypeName PSObject
		Write-host "        -----------------------------------------" -ForegroundColor Green
		Write-Host "        =====>>>> Channel (N°", $channelscounter ,") ID:", $myChannel.id, "- isDefault Channel:", $myChannel.isDefault -ForegroundColor green
		Write-Host "        ---- Channel Name:", $myChannel.name,"- Channel Portal URL:", $($StreamPortalChannelRoot + $myChannel.id)
		Write-Host "        ---- Channel CreationDate:", $myChannel.created,"- Channel ModificationDate:", $myChannel.modified
		Write-Host "        =====>>>> Channel Metrics Followers:", $myChannel.metrics.follows, "- Video Total:", $myChannel.metrics.videos -ForegroundColor Magenta
		Write-Host "        =====>>>> O365 Channel Creator Name: ", $myChannel.creator.name , " - Email:", $myChannel.creator.mail -ForegroundColor Magenta

		Write-Host " O365 GROUP Name:", $myChannel.group.name, "- ID:", $myChannel.group.id -ForegroundColor Yellow
		Write-Host "        =====>>>> O365 Group ID: ", $myChannel.group.id , " - Group Email:", $myChannel.group.aadGroup.mail -ForegroundColor Magenta
		Write-Host "        =====>>>> O365 Group Metrics Channel total:", $myChannel.group.metrics.channels, "- Video Total:", $myChannel.group.metrics.videos -ForegroundColor Magenta

		$datum | Add-Member -MemberType NoteProperty -Name ChannelID -Value $myChannel.id
		$datum | Add-Member -MemberType NoteProperty -Name ChannelName -Value $myChannel.name
		$datum | Add-Member -MemberType NoteProperty -Name ChannelURL -Value $($StreamPortalChannelRoot + $myChannel.id)
		$datum | Add-Member -MemberType NoteProperty -Name ChannelDefault -Value $myChannel.isDefault
		$datum | Add-Member -MemberType NoteProperty -Name ChannelFollowers -Value $myChannel.metrics.follows
		$datum | Add-Member -MemberType NoteProperty -Name ChannelVideos -Value $myChannel.metrics.videos
		$datum | Add-Member -MemberType NoteProperty -Name ChannelCreatorName -Value $myChannel.creator.name
		$datum | Add-Member -MemberType NoteProperty -Name ChannelCreatorEmail -Value $myChannel.creator.mail
		$datum | Add-Member -MemberType NoteProperty -Name ChannelCreationDate -Value $myChannel.created
		$datum | Add-Member -MemberType NoteProperty -Name ChannelModificationDate -Value $myChannel.modified

		$datum | Add-Member -MemberType NoteProperty -Name O365GroupId -Value $myChannel.group.id
		$datum | Add-Member -MemberType NoteProperty -Name O365GroupName -Value $myChannel.group.name
		$datum | Add-Member -MemberType NoteProperty -Name O365GroupEmail -Value $myChannel.group.aadGroup.mail
		$datum | Add-Member -MemberType NoteProperty -Name O365GroupTotalChannels -Value $myChannel.group.metrics.channels
		$datum | Add-Member -MemberType NoteProperty -Name O365GroupTotalVideos -Value $myChannel.group.metrics.videos

		$data += $datum
	}
}

$datestring = (get-date).ToString("yyyyMMdd-hhmm")
$fileName = ($PowerShellScriptFolder + "\O365StreamDetails_" + $datestring + ".csv")
	
Write-host " -----------------------------------------" -ForegroundColor Green
Write-Host (" >>> writing to file {0}" -f $fileName) -ForegroundColor Green
$data | Export-csv $fileName -NoTypeInformation
Write-host " -----------------------------------------" -ForegroundColor Green