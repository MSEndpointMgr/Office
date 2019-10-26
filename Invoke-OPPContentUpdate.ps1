<#
.SYNOPSIS
    Update the content of an Office 365 ProPlus application created in ConfigMgr.

.DESCRIPTION
    This script will ensure that the latest content version is updated when this script is either manually triggered or scheduled to run by doing the following:
    - Download the latest Office Deployment Tool executable and replace existing setup.exe in application content source path
    - Update the Office 365 ProPlus application content to the latest version
    - Update the detection method of the application in ConfigMgr with the latest versioning details

.PARAMETER OfficePackagePath
    Specify the full path to the Office application content source.

.PARAMETER OfficeApplicationName
    Specify the Office application display name.

.PARAMETER OfficeConfigurationFile
    Specify the Office application configuration file name, e.g. 'configuration.xml'.

.PARAMETER SkipDetectionMethodUpdate
    When True, update the detection method for the specified Office 365 ProPlus application.

.EXAMPLE
    # Update the content of an Office 365 ProPlus application named 'Office 365 ProPlus 64-bit' to the latest version:
    .\Invoke-OPPContentUpdate.ps1 -OfficePackagePath "C:\Source\Apps\Office365\x64" -OfficeApplicationName "Office 365 ProPlus 64-bit" -OfficeConfigurationFile "configuration.xml" -SkipDetectionMethodUpdate $false -Verbose

    # Update the content of an Office 365 ProPlus application named 'Office 365 ProPlus 64-bit' to the latest version, but don't update the application detection method:
    .\Invoke-OPPContentUpdate.ps1 -OfficePackagePath "C:\Source\Apps\Office365\x64" -OfficeApplicationName "Office 365 ProPlus 64-bit" -OfficeConfigurationFile "configuration.xml" -SkipDetectionMethodUpdate $true -Verbose

.NOTES
    FileName:    Invoke-OPPContentUpdate.ps1
    Author:      Nickolaj Andersen
    Contact:     @NickolajA
    Created:     2019-10-22
    Updated:     2019-10-26

    Version history:
    1.0.0 - (2019-10-22) Script created
    1.0.1 - (2019-10-25) Added the SkipDetectionMethodUpdate parameter to provide functionality that will not update the detection method
    1.0.2 - (2019-10-26) Added so that Distribution Points will automatically be updated
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [parameter(Mandatory = $false, HelpMessage = "Specify the full path to the Office application content source.")]
    [ValidateNotNullOrEmpty()]
    [string]$OfficePackagePath = "E:\CMsource\Apps\Microsoft\Office 365 ProPlus\x64",

    [parameter(Mandatory = $false, HelpMessage = "Specify the Office application display name.")]
    [ValidateNotNullOrEmpty()]
    [string]$OfficeApplicationName = "Office 365 ProPlus 64-bit (Semi-Annual)",

    [parameter(Mandatory = $false, HelpMessage = "Specify the Office application configuration file name, e.g. 'configuration.xml'.")]
    [ValidateNotNullOrEmpty()]
    [string]$OfficeConfigurationFile = "configuration.xml",

    [parameter(Mandatory = $false, HelpMessage = "When True, update the detection method for the specified Office 365 ProPlus application.")]
    [ValidateNotNullOrEmpty()]
    [bool]$SkipDetectionMethodUpdate = $true
)
Begin {
    # Import ConfigMgr module, required to update the detection method of the Office application
    try {
        Write-Verbose -Message "Attempting to import the Configuration Manager module"
        Import-Module -Name "$(Split-Path -Path $env:SMS_ADMIN_UI_PATH -Parent)\ConfigurationManager.psd1" -ErrorAction Stop -Verbose:$false
        $SiteCode = Get-PSDrive -PSProvider CMSite -Verbose:$false | Select-Object -ExpandProperty Name
    }
    catch [System.Exception] {
        Write-Warning -Message "$($_.Exception.Message). Line: $($_.InvocationInfo.ScriptLineNumber)"; break
    }    
}
Process {
    # Functions
    function Start-DownloadFile {
        param(
            [parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$URL,
    
            [parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$Path,
    
            [parameter(Mandatory=$true)]
            [ValidateNotNullOrEmpty()]
            [string]$Name
        )
        Begin {
            # Construct WebClient object
            $WebClient = New-Object -TypeName System.Net.WebClient
        }
        Process {
            # Create path if it doesn't exist
            if (-not(Test-Path -Path $Path)) {
                New-Item -Path $Path -ItemType Directory -Force | Out-Null
            }
    
            # Start download of file
            $WebClient.DownloadFile($URL, (Join-Path -Path $Path -ChildPath $Name))
        }
        End {
            # Dispose of the WebClient object
            $WebClient.Dispose()
        }
    }

    Write-Verbose -Message "Initiating Office application content update process"

    try {
        # Download latest Office Deployment Tool
        $ODTDownloadURL = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
        $WebResponseURL = ((Invoke-WebRequest -Uri $ODTDownloadURL -UseBasicParsing -ErrorAction Stop -Verbose:$false).links | Where-Object { $_.outerHTML -like "*click here to download manually*" }).href
        $ODTFileName = Split-Path -Path $WebResponseURL -Leaf
        $ODTFilePath = (Join-Path -Path $env:windir -ChildPath "Temp")
        Write-Verbose -Message "Attempting to download latest Office Deployment Toolkit executable"
        Start-DownloadFile -URL $WebResponseURL -Path $ODTFilePath -Name $ODTFileName

        try {
            # Extract latest ODT file
            $ODTExecutable = (Join-Path -Path $ODTFilePath -ChildPath $ODTFileName)
            $ODTExtractionPath = (Join-Path -Path $ODTFilePath -ChildPath (Get-ChildItem -Path $ODTExecutable).VersionInfo.ProductVersion)
            $ODTExtractionArguments = "/quiet /extract:$($ODTExtractionPath)"

            # Extract ODT files
            Write-Verbose -Message "Attempting to extract the setup.exe executable from Office Deployment Toolkit"
            Start-Process -FilePath $ODTExecutable -ArgumentList $ODTExtractionArguments -Wait -ErrorAction Stop

            try {
                # Determine if ODT needs to be updated in Office package folder
                $ODTCurrentVersion = (Get-ChildItem -Path (Join-Path -Path $OfficePackagePath -ChildPath "setup.exe") -ErrorAction Stop).VersionInfo.ProductVersion
                Write-Verbose -Message "Determined current Office Deployment Toolkit version as: $($ODTCurrentVersion)"
                $ODTLatestVersion = (Get-ChildItem -Path (Join-Path -Path $ODTExtractionPath -ChildPath "setup.exe") -ErrorAction Stop).VersionInfo.ProductVersion
                Write-Verbose -Message "Determined latest Office Deployment Toolkit version as: $($ODTLatestVersion)"

                try {
                    if ([System.Version]$ODTLatestVersion -gt [System.Version]$ODTCurrentVersion) {
                        # Replace existing setup.exe in Office package path with extracted
                        Write-Verbose -Message "Current Office Deployment Toolkit version needs to be updated to latest version, attempting to copy latest setup.exe"
                        Copy-Item -Path (Join-Path -Path $ODTExtractionPath -ChildPath "setup.exe") -Destination (Join-Path -Path $OfficePackagePath -ChildPath "setup.exe") -Force -ErrorAction Stop
                    }

                    try {
                        # Cleanup downloaded ODT content and executable
                        Write-Verbose -Message "Attempting to remove downloaded Office Deployment Toolkit temporary content files"
                        Remove-Item -Path $ODTExtractionPath -Recurse -Force -ErrorAction Stop
                        Remove-Item -Path $ODTExecutable -Force -ErrorAction Stop

                        try {
                            # Determine existing Office package version in \office\data folder
                            Write-Verbose -Message "Attempting to detect currect version information for existing Office 365 ProPlus content"
                            $OfficeDataFolderRoot = (Join-Path -Path $OfficePackagePath -ChildPath "office\data")
                            $OfficeDataFolderCurrent = Get-ChildItem -Path $OfficeDataFolderRoot -Directory -ErrorAction Stop
                            $OfficeDataFileCurrent = Get-ChildItem -Path $OfficeDataFolderRoot -Filter "v*_*.cab" -ErrorAction Stop

                            try {
                                # Construct arguments for setup.exe and call the executable and let it complete before we continue
                                $OfficeArguments = "/download $($OfficeConfigurationFile)"
                                Write-Verbose -Message "Attempting to update the Office 365 ProPlus application content based on configuration file"
                                Start-Process -FilePath "setup.exe" -ArgumentList $OfficeArguments -WorkingDirectory $OfficePackagePath -Wait -ErrorAction Stop

                                # Cleanup older Office data folder versions
                                Write-Verbose -Message "Checking to see if previous Office 365 ProPlus application content version should be removed"
                                if ((Get-ChildItem -Path $OfficeDataFolderRoot -Directory | Measure-Object).Count -ge 2) {
                                    Write-Verbose -Message "Previous Office 365 ProPlus application content should be cleaned up"

                                    # Remove old data folder
                                    Write-Verbose -Message "Attempting to remove Office 365 ProPlus application content directory: $($OfficeDataFolderCurrent.Name)"
                                    Remove-Item -Path $OfficeDataFolderCurrent.FullName -Recurse -Force

                                    # Remove old data cab file
                                    Write-Verbose -Message "Attempting to remove Office 365 ProPlus application content cabinet file: $($OfficeDataFileCurrent.Name)"
                                    Remove-Item -Path $OfficeDataFileCurrent.FullName -Force
                                }

                                try {
                                    # Get latest Office data version
                                    $OfficeDataLatestVersion = (Get-ChildItem -Path $OfficeDataFolderRoot -Directory -ErrorAction Stop).Name
                                    Write-Verbose -Message "Office 365 ProPlus application content is now determined at version: $($OfficeDataLatestVersion)"

                                    try {
                                        # Set location to ConfigMgr drive
                                        Write-Verbose -Message "Changing location to ConfigMgr drive: $($SiteCode):"
                                        Set-Location -Path ($SiteCode + ":") -ErrorAction Stop -Verbose:$false

                                        try {
                                            # Get Office application deployment type object
                                            Write-Verbose -Message "Attempting to retrieve Office 365 ProPlus application deployment type for application: $($OfficeApplicationName)"
                                            $OfficeDeploymentType = Get-CMDeploymentType -ApplicationName $OfficeApplicationName -ErrorAction Stop -Verbose:$false

                                            if ($SkipDetectionMethodUpdate -eq $false) {
                                                try {
                                                    # Create a new registry detection method
                                                    Write-Verbose -Message "Attempting to create new registry detection clause object for Office 365 ProPlus application deployment type: $($OfficeDeploymentType.LocalizedDisplayName)"
                                                    $DetectionClauseArgs = @{
                                                        ExpressionOperator = "GreaterEquals"
                                                        Hive = "LocalMachine"
                                                        KeyName = "Software\Microsoft\Office\ClickToRun\Configuration"
                                                        PropertyType = "Version"
                                                        ValueName = "VersionToReport"
                                                        ExpectedValue = $OfficeDataLatestVersion
                                                        Value = $true
                                                        ErrorAction = "Stop"
                                                        Verbose = $false
                                                    }
                                                    $DetectionClauseRegistryKeyValue = New-CMDetectionClauseRegistryKeyValue @DetectionClauseArgs
    
                                                    try {
                                                        # Construct string array with logical name of enhanced detection method registry name
                                                        [string[]]$OfficeApplicationDetectionMethodLogicalName = ([xml]$OfficeDeploymentType.SDMPackageXML).AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.SimpleSetting.LogicalName
                                                        Write-Verbose -Message "Enhanced detection method logical name for existing registry detection clause was determined as: $($OfficeApplicationDetectionMethodLogicalName)"
    
                                                        # Remove existing detection method and add new with updated version info
                                                        Write-Verbose -Message "Attempting to replace existing detection clause with new containing latest Office 365 ProPlus application content version"
                                                        Set-CMScriptDeploymentType -InputObject $OfficeDeploymentType -RemoveDetectionClause $OfficeApplicationDetectionMethodLogicalName -AddDetectionClause $DetectionClauseRegistryKeyValue -ErrorAction Stop  -Verbose:$false
                                                    }
                                                    catch [System.Exception] {
                                                        Write-Warning -Message "Failed to update registry detection clause for Office 365 ProPlus application deployment type. Error message: $($_.Exception.Message)"
                                                    }
                                                }
                                                catch [System.Exception] {
                                                    Write-Warning -Message "Failed to create new registry detection clause object. Error message: $($_.Exception.Message)"
                                                }
                                            }
                                            
                                            try {
                                                # Update Distribution Points
                                                Write-Verbose -Message "Attempting to update Distribution Points for application: $($OfficeApplicationName)"
                                                Update-CMDistributionPoint -ApplicationName $OfficeApplicationName -DeploymentTypeName $OfficeDeploymentType.LocalizedDisplayName -ErrorAction Stop -Verbose:$false

                                                Write-Verbose -Message "Successfully completed Office 365 ProPlus application content update process"
                                            }
                                            catch [System.Exception] {
                                                Write-Warning -Message "Failed to update Distribution Points for application. Error message: $($_.Exception.Message)"
                                            }
                                        }
                                        catch [System.Exception] {
                                            Write-Warning -Message "Failed to retrieve deployment type object for Office 365 ProPlus application. Error message: $($_.Exception.Message)"
                                        }                                            
                                    }
                                    catch [System.Exception] {
                                        Write-Warning -Message "Failed to change current location to ConfigMgr drive. Error message: $($_.Exception.Message)"
                                    }
                                }
                                catch [System.Exception] {
                                    Write-Warning -Message "Failed to determine the latest Office 365 application content version. Error message: $($_.Exception.Message)"
                                }
                            }
                            catch [System.Exception] {
                                Write-Warning -Message "Failed to update Office 365 ProPlus application content. Error message: $($_.Exception.Message)"
                            }
                        }
                        catch [System.Exception] {
                            Write-Warning -Message "Failed to detect currect version information for existing Office 365 ProPlus content. Error message: $($_.Exception.Message)"
                        }
                    }
                    catch [System.Exception] {
                        Write-Warning -Message "Failed to cleanup downloaded Office Deployment Toolkit content from temporary location. Error message: $($_.Exception.Message)"
                    }
                }
                catch [System.Exception] {
                    Write-Warning -Message "Failed to copy new setup.exe to Office application content source. Error message: $($_.Exception.Message)"
                }
            }
            catch [System.Exception] {
                Write-Warning -Message "Failed to determine version numbers for Office Deployment Toolkit for comparison. Error message: $($_.Exception.Message)"
            }
        }
        catch [System.Exception] {
            Write-Warning -Message "Failed to extract Office Deployment Toolkit. Error message: $($_.Exception.Message)"
        }
    }
    catch [System.Exception] {
        Write-Warning -Message "Failed to download the latest Office Deployment Toolkit. Error message: $($_.Exception.Message)"
    }
}
End {
    # Set location back to filesystem drive
    Set-Location -Path $env:SystemDrive
}