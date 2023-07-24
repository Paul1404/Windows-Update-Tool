$StartBanner = @"
:::       ::: :::    ::: :::::::::::       ::::::::   ::::::::  :::::::::  ::::::::::: ::::::::: :::::::::::      :::     :::      :::        :::::::  
:+:       :+: :+:    :+:     :+:          :+:    :+: :+:    :+: :+:    :+:     :+:     :+:    :+:    :+:          :+:     :+:    :+:+:       :+:   :+: 
+:+       +:+ +:+    +:+     +:+          +:+        +:+        +:+    +:+     +:+     +:+    +:+    +:+          +:+     +:+      +:+       +:+  :+:+ 
+#+  +:+  +#+ +#+    +:+     +#+          +#++:++#++ +#+        +#++:++#:      +#+     +#++:++#+     +#+          +#+     +:+      +#+       +#+ + +:+ 
+#+ +#+#+ +#+ +#+    +#+     +#+                 +#+ +#+        +#+    +#+     +#+     +#+           +#+           +#+   +#+       +#+       +#+#  +#+ 
 #+#+# #+#+#  #+#    #+#     #+#          #+#    #+# #+#    #+# #+#    #+#     #+#     #+#           #+#            #+#+#+#  #+#   #+#   #+# #+#   #+# 
  ###   ###    ########      ###           ########   ########  ###    ### ########### ###           ###              ###    ### ####### ###  #######  
"@

$EndBanner = @"
 ::::::::   ::::::::  ::::    ::::  :::::::::  :::        :::::::::: ::::::::::: :::::::::: :::::::::           :::    :::     :::     :::     ::: ::::::::::          :::          ::::    ::: ::::::::::: ::::::::  ::::::::::      :::::::::      :::   :::   ::: ::: 
:+:    :+: :+:    :+: +:+:+: :+:+:+ :+:    :+: :+:        :+:            :+:     :+:        :+:    :+:          :+:    :+:   :+: :+:   :+:     :+: :+:               :+: :+:        :+:+:   :+:     :+:    :+:    :+: :+:             :+:    :+:   :+: :+: :+:   :+: :+: 
+:+        +:+    +:+ +:+ +:+:+ +:+ +:+    +:+ +:+        +:+            +:+     +:+        +:+    +:+          +:+    +:+  +:+   +:+  +:+     +:+ +:+              +:+   +:+       :+:+:+  +:+     +:+    +:+        +:+             +:+    +:+  +:+   +:+ +:+ +:+  +:+ 
+#+        +#+    +:+ +#+  +:+  +#+ +#++:++#+  +#+        +#++:++#       +#+     +#++:++#   +#+    +:+          +#++:++#++ +#++:++#++: +#+     +:+ +#++:++#        +#++:++#++:      +#+ +:+ +#+     +#+    +#+        +#++:++#        +#+    +:+ +#++:++#++: +#++:   +#+ 
+#+        +#+    +#+ +#+       +#+ +#+        +#+        +#+            +#+     +#+        +#+    +#+          +#+    +#+ +#+     +#+  +#+   +#+  +#+             +#+     +#+      +#+  +#+#+#     +#+    +#+        +#+             +#+    +#+ +#+     +#+  +#+    +#+ 
#+#    #+# #+#    #+# #+#       #+# #+#        #+#        #+#            #+#     #+#        #+#    #+# #+#      #+#    #+# #+#     #+#   #+#+#+#   #+#             #+#     #+#      #+#   #+#+#     #+#    #+#    #+# #+#             #+#    #+# #+#     #+#  #+#        
 ########   ########  ###       ### ###        ########## ##########     ###     ########## #########  ##       ###    ### ###     ###     ###     ##########      ###     ###      ###    #### ########### ########  ##########      #########  ###     ###  ###    ### 
"@

# Define global variables
# Define a variable to hold an empty line (used for formatting)
$empty_line = ""

# Set the path to the configuration file
$configFilePath = Join-Path -Path $PSScriptRoot -ChildPath "config.json"

# Load the configuration data from the JSON file
$config = Get-Content -Path $configFilePath | ConvertFrom-Json

# Replace the "$username" placeholder in the download path with the actual username
$DownloadPath = $config.downloadPath -replace '\$username', $env:UserName

# Construct the path to the Windows Update Tool (WUT) in the download folder
$WUTPath = Join-Path $DownloadPath $config.updateExeName

# Extract the URL for downloading the Windows Update Tool
$updateExeURL = $config.updateExeURL

# Construct the full path to save the downloaded Windows Update Tool
$updateExeName = Join-Path $DownloadPath $config.updateExeName


# Creating Functions

<#
.SYNOPSIS
Displays an error message to the user.

.DESCRIPTION
This function is used to display an error message to the user when an error occurs during script execution.
It takes a single parameter, `$ErrorMessage`, which contains the error message to be shown.

.PARAMETER ErrorMessage
The error message to be displayed to the user.

.EXAMPLE
try {
    # Some code that may throw an error
    # ...
    throw "An error occurred while processing data."
}
catch {
    Trace-Error -ErrorMessage $_
}
# Displays an error message with the specific error message caught by the catch block.
#>

function Trace-Error {
    param (
        [Parameter(Mandatory=$true)]
        [string] $ErrorMessage
    )

    # Display the error message to the console
    Write-Host "An error occurred: $ErrorMessage"

    # Additional error handling code can be added here, such as logging the error or performing cleanup actions.
    # For example, you can log the error to a file or send an email notification to administrators.
    # If required, you can also prompt the user for further actions or provide additional feedback.

    # The function does not return any value, as it is meant to display the error message and handle the error scenario.
}


<#
.SYNOPSIS
Starts the Windows Update process using the Windows Update Tool.

.DESCRIPTION
This function changes the current working directory to the download folder specified in the global variable `$DownloadPath`.
It then attempts to start the Windows Update Tool by executing the `Update.exe` file in the current working directory.
If the `Update.exe` file is not found or the path is incorrect, an error message is displayed, and the function exits with an error code 1.

.NOTES
- The function relies on the `$DownloadPath` global variable, which should be defined before calling the function.
- The `Update.exe` file should be present in the download folder for the update process to succeed.

.EXAMPLE
$DownloadPath = "C:\Downloads"
Start-Update
# Changes to the specified download folder and starts the Windows Update Tool (Update.exe) from that location.
#>

function Start-Update {
    process {
        try {
            # Change to the download folder specified in the global variable $DownloadPath
            Write-Host "Changing to the download folder..."
            Set-Location $DownloadPath

            # Start the Windows Update Tool (Update.exe)
            Write-Host "Starting the Windows Update Tool..."
            .\Update.exe
        }
        catch {
            # If Update.exe is not found or path is incorrect, display an error message and exit with error code 1
            Write-Host "Error: Update.exe file not found or path is incorrect. Make sure Update.exe is present in the download folder."
            Write-Host "Current download folder path: $DownloadPath - Is this correct?"
            Exit 1
        }
    }
}


<#
.SYNOPSIS
Downloads the Windows Update Tool from the Microsoft website.

.DESCRIPTION
This function downloads the Windows Update Tool from the provided URL and saves it to the specified output location.
It uses the `System.Net.WebClient` class to perform the download asynchronously, allowing it to display download progress.
The function requires two parameters:
- `$URL`: The URL from which to download the Windows Update Tool.
- `$Output`: The full path to save the downloaded file.

.EXAMPLE
$downloadURL = "https://example.com/update.exe"
$downloadPath = "C:\Downloads\update.exe"
Get-Update -URL $downloadURL -Output $downloadPath
# Downloads the Windows Update Tool from the provided URL and saves it to the specified output location.
#>

function Get-Update {
    param($URL, $Output)
    
    try {
        # Create a new WebClient object
        $wc = New-Object System.Net.WebClient

        # Download the file asynchronously from the provided URL to the specified output path
        $wc.DownloadFileAsync([uri]$URL, $Output)
        
        # Create an event handler to show the download progress
        Register-ObjectEvent -InputObject $wc -EventName DownloadProgressChanged -Action {
            Write-Progress -Activity "Downloading Update" -Status "$($EventArgs.ProgressPercentage)% Complete:" -PercentComplete $EventArgs.ProgressPercentage
        }
        
        # Wait until the download is finished
        while ($wc.IsBusy) {
            Start-Sleep -Seconds 1
        }
    } catch {
        # If the download fails, throw an error
        throw "Failed to download file."
    } finally {
        # Dispose of the WebClient object to free resources
        $wc.Dispose()
    }
}


<#
.SYNOPSIS
Synchronizes the system time with the user's confirmation.

.DESCRIPTION
This function retrieves the current system time using the `Get-Date` cmdlet and displays it to the console.
It then prompts the user to confirm whether the system time is correct by entering "Y" for yes or "N" for no.
If the user confirms that the system time is incorrect ("N"), the function calls `Sync-WindowsTime` to synchronize the system time.

.NOTES
- The function relies on the `Sync-WindowsTime` function to perform the time synchronization.
- The user's response is case-sensitive (i.e., "Y" or "N" must be entered in uppercase).

.EXAMPLE
Sync-Time
# Prompts the user to confirm if the system time is correct and syncs the time if the user responds "N".
#>

function Sync-Time {
    process {
        # Get the current system time and display it to the console
        $CurrentTime = (Get-Date).ToString('T')
        Write-Host "Current system time: $CurrentTime"

        # Ask the user for confirmation if the system time is correct
        [string]$TimeCorrect = Read-Host "Is your system time correct? (Y/N)"

        # Synchronize the system time if the user confirms it's incorrect
        if ($TimeCorrect -eq "N") {
            # Call the Sync-WindowsTime function to synchronize the system time
            Sync-WindowsTime
        } else {
            # If the user responds "Y" or anything other than "N", exit the function without making any changes
            Exit
        }
    }
}


<#
.SYNOPSIS
Checks if the script is running with administrator rights.

.DESCRIPTION
This function determines whether the script is running with elevated privileges (administrator rights).
It uses the `[Security.Principal.WindowsPrincipal]` class to check the current PowerShell session's role.
If the script is running with administrator rights, the function returns `True`.
If the script is not running with administrator rights, an error is displayed using the `Trace-Error` function, and the function returns `False`.

.EXAMPLE
if (Test-AdminRights) {
    Write-Host "The script is running with administrator rights."
} else {
    Write-Host "The script needs to be run as an administrator."
}
#>

function Test-AdminRights {
    if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator") -eq $false) {
        # If not running with admin rights, display an error and return false
        Trace-Error "This script needs to be run in an elevated (administrator-level) PowerShell window."
        Return $false
    } else {
        # If running with admin rights, return true
        Write-Output "Script is running with admin rights."
        Return $true
    }
}


<#
.SYNOPSIS
Checks if a file exists at the specified path.

.DESCRIPTION
This function checks if a file exists at the given file path. It uses the `Test-Path` cmdlet to determine file existence.
The function takes a single parameter, `$FilePath`, which is the full path to the file being checked.
If the file exists, the function returns `True`; otherwise, it returns `False`.

.PARAMETER FilePath
The full path to the file to be checked for existence.

.EXAMPLE
$filePath = "C:\Downloads\update.exe"
if (Test-FileExists -FilePath $filePath) {
    Write-Host "The file exists."
} else {
    Write-Host "The file does not exist."
}
#>

function Test-FileExists {
    param (
        [Parameter(Mandatory=$true)]
        [string] $FilePath
    )

    Write-Output "Checking if the update file exists in the download folder..."

    if (Test-Path -Path $FilePath -PathType Leaf) {
        # If the file exists, return true
        Write-Host "The file exists!"
        Return $true
    } else {
        # If the file does not exist, return false
        Write-Host "The file does not exist..."
        Return $false
    }
}


<#
.SYNOPSIS
Synchronizes the system time if the script is running with administrator rights.

.DESCRIPTION
This function checks if the PowerShell session has elevated privileges (administrator rights).
If it is running with administrator rights, the function proceeds with synchronizing the system time using the `w32tm.exe` utility.
If the script is not run with administrator rights, a warning is displayed, and the function exits without making any changes to the system.

.EXAMPLE
if (Sync-WindowsTime) {
    Write-Host "System time successfully synchronized."
}
#>

function Sync-WindowsTime {
    # Requires administrator rights
    # Check if the PowerShell session is elevated (has been run as an administrator)
    If (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator") -eq $false) {
        # Display a warning message if not running with admin rights
        $empty_line | Out-String
        Write-Warning "It seems that this script is run in a 'normal' PowerShell window."
        $empty_line | Out-String
        Write-Verbose "Please consider running this script in an elevated (administrator-level) PowerShell window." -Verbose
        $empty_line | Out-String
        $admin_text = "For performing system-altering procedures, such as removing scheduled tasks, disabling services, or writing to the registry, elevated rights are mandatory. An elevated PowerShell session can be initiated by starting PowerShell with the 'run as an administrator' option."
        Write-Output $admin_text
        $empty_line | Out-String
        "Exiting without making any changes to the system." | Out-String
        Return $false
    } Else {
        # Return true if running with admin rights
        Return $true
    }
}


<#
.SYNOPSIS
Retrieves the current Windows username.

.DESCRIPTION
This function fetches the username of the currently logged-in Windows user by accessing the `System.Environment.UserName` property.
It does not require any parameters and will return the current Windows username as a string.

.EXAMPLE
$user = Get-CurrentWindowsUser
Write-Host "The current Windows user is: $user"
#>

function Get-CurrentWindowsUser {
    [CmdletBinding()]
    param()

    # Get the current Windows username
    $username = [System.Environment]::UserName

    # Output the username
    Write-Output "Current Windows user: $username"
}


<#
.SYNOPSIS
Checks for an active internet connection.

.DESCRIPTION
This function checks if the system has an active internet connection by attempting to create a web request using the `[Activator]` class.
It determines internet connectivity based on the response from the request. If a connection is available, the function returns True. Otherwise, it returns False.

.EXAMPLE
$internetStatus = Get-InternetConnection
if ($internetStatus -eq $true) {
    Write-Host "Internet connection is available."
} else {
    Write-Host "No internet connection detected."
}
#>

function Get-InternetConnection {
    # Check for internet connectivity using a web request
    If (([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet) -eq $false) {
        # If no internet connection, throw an error and return False
        Trace-Error "The Internet connection doesn't seem to be working. Exiting without syncing the time."
        Return $false
    } Else {
        # If there is an internet connection, return True
        Write-Output "Internet connection is working."
        Return $true
    }
}


# Main script starts here
# Check if there is an internet connection
if (Get-InternetConnection) {
    # Check if the script is running with admin rights
    if (Test-AdminRights) {
        # Check if the Windows Update Tool is already downloaded
        if ($false) {
            Start-Update
        } else {
            # If not downloaded, get the Windows Update Tool from the URL and start the update process
            Get-Update -URL $updateExeURL -Output "$updateExeName"
            Start-Update
        }
    }
}

# Print ending message
Write-Output $empty_line
Write-Output $empty_line
Write-Host "$EndBanner"
