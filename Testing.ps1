$configFilePath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"
$config = Get-Content -Path $configFilePath | ConvertFrom-Json

$DownloadPath = $config.downloadPath -replace '\$username', $env:UserName
$updateExeURL = $config.updateExeURL
$updateExeName = Join-Path $DownloadPath $config.updateExeName


function Get-Update {
    param($URL, $Output)

    try {
        $wc = New-Object System.Net.WebClient

        # Hook up DownloadProgressChanged event
        Register-ObjectEvent $wc DownloadProgressChanged -Action {
            Write-Progress -Activity "Downloading Update" -Status "$($EventArgs.ProgressPercentage)% Complete:" -PercentComplete $EventArgs.ProgressPercentage
        }

        # Record start time
        $startTime = Get-Date

        # Download the file synchronously
        $wc.DownloadFile($URL, $Output)

        # Record end time
        $endTime = Get-Date

        # Calculate download time and speed
        $downloadTime = $endTime - $startTime
        $fileSize = (Get-Item $Output).Length / 1MB
        $downloadSpeed = $fileSize / $downloadTime.TotalSeconds

        Write-Host "The file was downloaded to: $Output"
        Write-Host "Download time: $downloadTime"
        Write-Host "File size: $fileSize MB"
        Write-Host "Download speed: $downloadSpeed MB/sec"
    } catch {
        # Provide more details when an error occurs
        Write-Host "An error occurred while trying to download the file:"
        Write-Host "Error message: $($_.Exception.Message)"
        Write-Host "Error details: $($_.Exception.InnerException)"
        throw
    } finally {
        $wc.Dispose()
    }
}

Get-Update -URL $updateExeURL -Output "$updateExeName"