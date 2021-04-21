$uri = "https://hooks.slack.com/services/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"
$filename = "APACTechSupRealTimeDashboard.xlsx"
$foldername = "C:\Users\jchoi\Tableau Software Inc\APAC Tech Support - Documents\"
$filepath = "C:\Users\jchoi\Tableau Software Inc\APAC Tech Support - Documents\$filename"

function MessageTo-Slack {
    param($ChannelUri, $Message)

    # $text = ":alert:Good day folks! Case list is here!:alert:"
    $body = ConvertTo-Json @{
        text="$text $Message"
    }
    
    $body = ConvertTo-Json @{
        # text=":alert: <!here> Restarting RealTime Angel"
        text=":alert: Restarting RealTime Angel"
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-RestMethod -Method POST -ContentType "application/json" -uri $ChannelUri -Body $body | Out-Null
}

function Slack-Mrkdwn {
    [CmdletBinding()]
    param ($Text)
    return (ConvertTo-Json -Depth 10 @{blocks=@(@{type="section";text=@{type="mrkdwn";text="$Text"}})})
} 

MessageTo-Slack -ChannelUri $uri -Message "Test"

Write-Host "Terminating RealTime Angel"
if ( (Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "RealTime Angel" }) ) {
    Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "RealTime Angel" } | Stop-Process -Force
    Start-Sleep -Seconds 2
}

Write-Host "Terminating Weekend Angel"
if ( (Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "Weekend Angel" }) ) {
    Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "Weekend Angel" } | Stop-Process -Force
    Start-Sleep -Seconds 2
}

Write-Host "Terminating Angel Monitor"
if ( (Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "Angel Monitor" }) ) {
    Get-Process powershell –ea 0 | Where-Object { $_.MainWindowTitle –like "Angel Monitor" } | Stop-Process -Force
    Start-Sleep -Seconds 2
}


Write-Host "Terminating Excel Process"
Get-Process | Where-Object { $_.ProcessName -match "Excel" } | Stop-Process
get-process excel | select MainWindowTitle, Id, StartTime
Write-Host "Terminating Excel Process X2"
Get-Process | Where-Object { $_.ProcessName -match "Excel" } | Stop-Process


$i = 1
Write-Host "Restarting Angel"

while(!(Test-Path $filepath)) {
    Start-Sleep -Seconds 1
    Write-Host $i
    $i++
    if ($i -eq 120) {
        Copy-Item C:\Downloads\APACTechSupRealTimeDashboard.xlsx -Destination $foldername -Recurse -Force
    }
}
Write-Host "Go!"

start powershell.exe C:\MyProjects\ps\Angel\RealTime-Cases3.ps1