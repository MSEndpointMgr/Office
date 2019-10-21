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

# Edit the following variables
$OfficePackagePath = "E:\CMsource\Apps\Microsoft\Office 365 ProPlus\x64"
$OfficeApplicationName = "Office 365 ProPlus 64-bit (Semi-Annual)"

Import-Module -Name "$(Split-Path -Path $env:SMS_ADMIN_UI_PATH -Parent)\ConfigurationManager.psd1" -ErrorAction Stop -Verbose:$false
$SiteCode = Get-PSDrive -PSProvider CMSite -Verbose:$false | Select-Object -ExpandProperty Name

# Ensure package path exist before proceeding
if (Test-Path -Path $OfficePackagePath) {
    # Download latest Office Deployment Tool
    $ODTDownloadURL = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
    $WebResponseURL = ((Invoke-WebRequest -Uri $ODTDownloadURL -UseBasicParsing).links | Where-Object { $_.outerHTML -like "*click here to download manually*" }).href
    $ODTFileName = Split-Path -Path $WebResponseURL -Leaf
    $ODTFilePath = (Join-Path -Path $env:windir -ChildPath "Temp")
    Start-DownloadFile -URL $WebResponseURL -Path $ODTFilePath -Name $ODTFileName

    # Extract latest ODT file
    $ODTExecutable = (Join-Path -Path $ODTFilePath -ChildPath $ODTFileName)
    $ODTExtractionPath = (Join-Path -Path $ODTFilePath -ChildPath (Get-ChildItem -Path $ODTExecutable).VersionInfo.ProductVersion)
    $ODTExtractionArguments = "/quiet /extract:$($ODTExtractionPath)"

    # Extract ODT files
    Start-Process -FilePath $ODTExecutable -ArgumentList $ODTExtractionArguments -Wait

    # Determine if ODT needs to be updated in Office package folder
    $ODTCurrentVersion = (Get-ChildItem -Path (Join-Path -Path $OfficePackagePath -ChildPath "setup.exe")).VersionInfo.ProductVersion
    $ODTLatestVersion = (Get-ChildItem -Path (Join-Path -Path $ODTExtractionPath -ChildPath "setup.exe")).VersionInfo.ProductVersion

    if ([System.Version]$ODTLatestVersion -gt [System.Version]$ODTCurrentVersion) {
        # Replace existing setup.exe in Office package path with extracted
        Copy-Item -Path (Join-Path -Path $ODTExtractionPath -ChildPath "setup.exe") -Destination (Join-Path -Path $OfficePackagePath -ChildPath "setup.exe") -Force
    }

    # Cleanup downloaded ODT content
    Remove-Item -Path $ODTExtractionPath -Recurse -Force

    # Remove downloaded ODT executable
    Remove-Item -Path $ODTExecutable -Force

    # Determine existing Office package version in \office\data folder
    $OfficeDataFolderRoot = (Join-Path -Path $OfficePackagePath -ChildPath "office\data")
    $OfficeDataFolderCurrent = Get-ChildItem -Path $OfficeDataFolderRoot -Directory
    $OfficeDataFileCurrent = Get-ChildItem -Path $OfficeDataFolderRoot -Filter "v*_*.cab"

    # Construct arguments for setup.exe and call the executable and let it complete before we continue
    $OfficeArguments = "/download configuration.xml"
    Start-Process -FilePath "setup.exe" -ArgumentList $OfficeArguments -WorkingDirectory $OfficePackagePath -Wait

    # Cleanup older Office data folder versions
    if ((Get-ChildItem -Path $OfficeDataFolderRoot -Directory | Measure-Object).Count -ge 2) {
        # Remove old data folder
        Remove-Item -Path $OfficeDataFolderCurrent.FullName -Recurse -Force

        # Remove old data cab file
        Remove-Item -Path $OfficeDataFileCurrent.FullName -Force
    }

    # Get latest Office data version
    $OfficeDataLatestVersion = (Get-ChildItem -Path $OfficeDataFolderRoot -Directory).Name

    # Set location to ConfigMgr drive
    Set-Location -Path ($SiteCode + ":")

    # Get Office application deployment type object
    $OfficeDeploymentType = Get-CMDeploymentType -ApplicationName $OfficeApplicationName

    # Create a new registry detection method
    $DetectionClauseRegistryKeyValue = New-CMDetectionClauseRegistryKeyValue -ExpressionOperator GreaterEquals -Hive LocalMachine -KeyName "Software\Microsoft\Office\ClickToRun\Configuration" -PropertyType Version -ValueName "VersionToReport" -ExpectedValue $OfficeDataLatestVersion -Value
    
    # Construct string array with logical name of enhanced detection method registry name
    [string[]]$OfficeApplicationDetectionMethodLogicalName = ([xml]$OfficeDeploymentType.SDMPackageXML).AppMgmtDigest.DeploymentType.Installer.CustomData.EnhancedDetectionMethod.Settings.SimpleSetting.LogicalName
    
    # Remove existing detection method and add new with updated version info
    Set-CMScriptDeploymentType -InputObject $OfficeDeploymentType -RemoveDetectionClause $OfficeApplicationDetectionMethodLogicalName -AddDetectionClause $DetectionClauseRegistryKeyValue

    # Set location back to filesystem drive
    Set-Location -Path $env:SystemDrive
}