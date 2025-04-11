<#
.SYNOPSIS
    Installs the Powershell WinGet client module.
.DESCRIPTION
    Checks to see if the Powershell WinGet client module is installed. If it isn't, attempts to install it.
    If the install fails, exits the program.
.NOTES
    The module can be found here: https://www.powershellgallery.com/packages/Microsoft.WinGet.Client/1.10.320.
#>
function Install-WinGetClientModule {
    # Check if official Powershell WinGet client is installed.
    if (Get-Module -ListAvailable -Name Microsoft.WinGet.Client) {
        Write-Host "Powershell Module for the Windows Package Manager Client is installed."

    } else {
        # Install NuGet package provider, which is a requirement to install the module.
        Write-Host "Installing NuGet..."
        Register-PackageSource -Name "NuGet" -Location "https://api.nuget.org/v3/index.json" -ProviderName NuGet

        # Check if NuGet installed successfully.
        if ($?) {
            Write-Host "Successfully installed." -ForegroundColor Green
        } else {
            # If install could not be completed, exit script.
            Write-Host "Could not install NuGet." -ForegroundColor Red
            exit
        }

        # Set PSGallery as a trusted repository to install from.
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

        # Install the module.
        Write-Host "Installing the Powershell Module for the Windows Package Manager Client..."
        Install-Module -Name Microsoft.WinGet.Client

        # Check if the module installed successfully.
        if ($?) {
            Write-Host "Successfully installed." -ForegroundColor Green
        } else {
            # If install could not be completed, exit script.
            Write-Host "Could not install the Powershell Module for the Windows Package Manager Client." -ForegroundColor Red
            exit
        }
    }
}


<#
.SYNOPSIS
    Get WinGet applications that require updates.
#>
function Get-ApplicationsWithUpdates {
    # Store the applications with updates available through WinGet
    $applicationsWithUpdates = Get-WinGetPackage | Where-Object IsUpdateAvailable

    return $applicationsWithUpdates
}


<#
.SYNOPSIS
    Update applications though WinGet.
#>
function Update-Applications {
    param (
        [Parameter(Mandatory = $true)]
        [Array]$applicationsToUpdate
    )
    
    # Update each application provided.
    foreach ($app in $applicationsToUpdate) {
        $appName = $app.Name
        $appId = $app.Id
        Write-Host "Updating $appName..."
        
        # Update the application
        Update-WinGetPackage -Id $appId -Scope Any -Mode Silent -Force
        
        # Check if the application was successfully updated.
        if ($?) {
            Write-Host "$appName updated to the latest version." -ForegroundColor Green
        } else {
            # If the update failed, attempt to reinstall.
            Write-Host "Failed to update $appName. Attempting to reinstall..."
            Uninstall-WinGetPackage -Id $appId  -Scope Any -Mode Silent -Force
            Install-WinGetPackage -Id $appId  -Scope Any -Mode Silent -Force

            # Check if the application reinstalled successfully.
            if ($?) {
                Write-Host "$appName updated to the latest version." -ForegroundColor Green
            } else {
                Write-Host "Failed to update $appName." -ForegroundColor Red
            }
        }
    }
}


function Main {
    # Start logging.
    $logDate = Get-Date -Format FileDateTime
    Start-Transcript -Path C:\Users\$env:USERNAME\Documents\Update-WinGetApplicationsLogs\log-$logDate.txt

    # Get Powershell WinGet client.
    Install-WinGetClientModule

    # Check for application updates through WinGet.
    $applicationsToUpdate = Get-ApplicationsWithUpdates
    Write-Host $applicationsToUpdate

    # Check that the number of applications to update isn't 0.
    if ($applicationsToUpdate.Count -gt 0) {
        # Update applications.
        Update-Applications -ApplicationsToUpdate $applicationsToUpdate
    } else {
        Write-Host "All applications are up to date."
    } 

    # Stop logging.
    Stop-Transcript
}


# Run the script.
Main
