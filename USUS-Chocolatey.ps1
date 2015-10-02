<#
.SYNOPSIS
	Creates Chocolatey Packages from USUS Software Master
.NOTES
	File Name		: USUS-Chocolatey.ps1
	Author			: Jason Lorsung (jason@ususcript.com)
	Last Update		: 2015-10-01
	Version			: 1.0
.USAGE
	USUS-Chocolatey.ps1 -SoftwareMasterFile "SoftwareMaster.xml" -IncludesDir "IncludesDir" -ChocolateyRepo "ChocolateyRepo"
.FLAGS
	-SoftwareMasterFile		Use this to specify your SoftwareMasterFile file for the script to use
	-IncludesDir			Use this to specify where the script should get reqired items from (NuGet.exe)
	-ChocolateyRepo			Use this to specify where the script should store the Chocolatey Packages
	-DebugEnable			Use this to enable Debug output
#>

param([Parameter(Mandatory=$True)][string]$SoftwareMasterFile,[Parameter(Mandatory=$True)][string]$IncludesDir,[Parameter(Mandatory=$True)][string]$ChocolateyRepo)

[string]$Timestamp = $(get-date -f "yyyy-MM-dd HH:mm")

#Define Functions

Function ChocolateyPackage
{
	Param($Version32,$Version64,$PackageName,$HumanReadableName,$Timestamp,$IsMSI,$BitCount,$ChocolateyRepo,$SoftwareMaster,$SoftwareMasterFile)
	
	IF ($BitCount -eq 32 -Or $BitCount -eq 96)
	{
		IF ($Version32.Chocolatey)
		{
			IF (([datetime]$Version32.Chocolatey.Updated) -ge ([datetime]$Version32.Updated))
			{
				Write-Debug "Chocolatey Package already up to date, skipping..."
				Return
			}
		}
		
		$SoftwareVersion = $Version32.name
		$Location32 = $Version32.Location
		IF ($BitCount -eq 96)
		{
			$Location64 = $Version64.Location
		}
	}
	
	IF ($BitCount -eq 64)
	{
		IF ($Version64.Chocolatey)
		{
			IF (([datetime]$Version64.Chocolatey.Updated) -ge ([datetime]$Version64.Updated))
			{
				Write-Debug "Chocolatey Package already up to date, skipping..."
				Return
			}
		}
		
		$SoftwareVersion = $Version64.name
		$Location64 = $Version64.Location
	}
	
	$ChocolateyNuget = $ChocolateyRepo + "\" + $PackageName
	$NugetLocation = $env:TEMP + "\ChocolateyNuget"

	IF (Test-Path $NugetLocation)
	{
		Remove-Item $NugetLocation -Recurse -Force -ErrorAction SilentlyContinue
		Try
		{
			New-Item $NugetLocation -Type Directory -ErrorAction Stop | Out-Null
		} Catch {
			Write-Debug "Could not create program directory of $NugetLocation.
	Please ensure that the user running this script has Write permissions to this location, and try again.`r`n"
		} 	
	} ELSE {
		Try
		{
			New-Item $NugetLocation -Type Directory -ErrorAction Stop | Out-Null
		} Catch {
			Write-Debug "Could not create program directory of $NugetLocation.
	Please ensure that the user running this script has Write permissions to this location, and try again.`r`n"
		} 
	}
	
	#NuSpec Creation
	
	$NuSpecCommand = "& '" + $IncludesDir + "\nuget.exe' spec " + '"' + $NugetLocation + "\" + $PackageName + '"'
	Invoke-Expression $NuSpecCommand | Out-Null
	$ChocolateySpecLocation = $NugetLocation + "\" + $PackageName + ".nuspec"
	
	#NuSpec Configuration
	
	[xml]$ChocolateySpec = Get-Content $ChocolateySpecLocation
	$ChocolateySpec.package.metadata.id = $PackageName
	
	$ChocolateySpec.package.metadata.version = [string]$SoftwareVersion
	
	IF ($BitCount -eq 32 -Or $BitCount -eq 96)
	{
		$Extras = $Version32.Extras32
	}
	
	IF ($BitCount -eq 64)
	{
		$Extras = $Version64.Extras64
	}
	
	IF ($Extras.Author)
	{
		$ChocolateySpec.package.metadata.authors = [string]$Extras.Author
	} ELSE {
		$ChocolateySpec.package.metadata.authors = "USUS to Chocolatey"
	}
	
	IF ($Extras.Owner)
	{
		$ChocolateySpec.package.metadata.owners = [string]$Extras.Owner
	} ELSE {
		$ChocolateySpec.package.metadata.owners = "USUS to Chocolatey"
	}
	
	IF (!($Extras.Tags))
	{
		$ChocolateySpec.package.metadata.RemoveChild($ChocolateySpec.package.metadata.SelectSingleNode("tags")) | Out-Null
	} ELSE {
		$ChocolateySpec.package.metadata.tags = $Extras.Tags
	}
	
	IF ($Extras.Description)
	{
		$ChocolateySpec.package.metadata.description = [string]$Extras.Description
	} ELSE {
		$ChocolateySpec.package.metadata.description = "Installs the latest $HumanReadableName - Version ($SoftwareVersion)
Updated by USUS - $Timestamp"
	}
	
	IF ($Extras.Summary)
	{
		$ChocolateySpec.package.metadata.AppendChild($ChocolateySpec.CreateElement("summary")) | Out-Null
		$ChocolateySpec.package.metadata.summary = [string]$Extras.Summary
	}
	
	$ChocolateySpec.package.metadata.AppendChild($ChocolateySpec.CreateElement("title")) | Out-Null
	$ChocolateySpec.package.metadata.title = $HumanReadableName
	$ChocolateySpec.package.metadata.RemoveChild($ChocolateySpec.package.metadata.SelectSingleNode("licenseUrl")) | Out-Null
	$ChocolateySpec.package.metadata.RemoveChild($ChocolateySpec.package.metadata.SelectSingleNode("projectUrl")) | Out-Null
	
	IF ($Extras.IconURL)
	{
		$ChocolateySpec.package.metadata.iconUrl = [string]$Extras.IconURL
	} ELSE {
		$ChocolateySpec.package.metadata.RemoveChild($ChocolateySpec.package.metadata.SelectSingleNode("iconUrl")) | Out-Null
	}
	
	$ChocolateySpec.package.metadata.RemoveChild($ChocolateySpec.package.metadata.SelectSingleNode("releaseNotes")) | Out-Null
	$ChocolateySpec.package.metadata.RemoveChild($ChocolateySpec.package.metadata.SelectSingleNode("copyright")) | Out-Null
	$ChocolateySpec.package.metadata.dependencies.RemoveAll()
	$ChocolateySpec.Save($ChocolateySpecLocation)
	
	#Chocolatey Install Script Creation
		
	$NugetTools = $NugetLocation + "\Tools"
	
	Try
	{
		New-Item $NugetTools -Type Directory -ErrorAction Stop | Out-Null
	} Catch {
		Write-Debug "Could not create program directory of $NugetTools.
Please ensure that the user running this script has Write permissions to this location, and try again.`r`n"
	}
	
	$PackageArguments = ""
	IF ($IsMSI)
	{		
		IF ($Extras.SilentInstall)
		{
			$PackageArguments = $PackageArguments + " /qn"
		}
		IF ($Extras.NoReboot)
		{
			$PackageArguments = $PackageArguments + " /norestart"
		}
	}
	$PackageArguments = $PackageArguments + " " + $Extras.CustomOptions
	$PackageArguments = $PackageArguments.Trim()
	$ChocolateyInstallScriptLocation = $NugetTools + "\chocolateyInstall.ps1"
	
	$ChocolateyInstallScript = '$packageName = ' + "'$PackageName'
" + '$version = ' + "'" + $SoftwareVersion + "'" + "
" + '$fileType = '

	IF ($IsMSI)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + "'msi'"
	} ELSE {
		$ChocolateyInstallScript = $ChocolateyInstallScript + "'exe'"
	}
	
	$ChocolateyInstallScript = $ChocolateyInstallScript + "
" + '$installArgs = ' + "'$PackageArguments'
"

	IF ($BitCount -eq 32)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + '$url = ' + "'$Location32'"
	}
	
	IF ($BitCount -eq 64)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + '$url64 = ' + "'$Location64'"
	}
	
	IF ($BitCount -eq 96)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + '$url = ' + "'$Location32'" + '
$url64 = ' + "'$Location64'"
	}
	
	$ChocolateyInstallScript = $ChocolateyInstallScript + "
" + '$majorVersion = ([version] $version).Major

'

	IF ($Extras.WMIPackageName)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + '$alreadyInstalled = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "' + $Extras.WMIPackageName + '" -And $_.Version -eq $version}

IF ($alreadyInstalled)
{
	Write-Output $(' + "'$HumanReadableName '" + ' + $version + ' + "' is already installed.')
} ELSE {"
	}
	$ChocolateyInstallScript = $ChocolateyInstallScript + "Install-ChocolateyPackage " + '$packageName $fileType $installArgs '
	
	IF ($BitCount -eq 32)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + '$url'
	} 
	IF ($BitCount -eq 64)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + '$url64'
	}
	IF ($BitCount -eq 96)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + '$url $url64'
	}
	
	IF ($WMIPackageName)
	{
		$ChocolateyInstallScript = $ChocolateyInstallScript + "
}"
	}
	
	$ChocolateyInstallScript | Out-File $ChocolateyInstallScriptLocation
	
	#Generate Nuget
	
	$NugetCommand = "& '" + $IncludesDir + "\nuget.exe' pack " + '"' + $ChocolateySpecLocation + '"' + " -OutputDirectory " + '"' + $ChocolateyRepo + '"'
	
	Invoke-Expression $NugetCommand | Out-Null
	
	Remove-Item $NugetLocation -Recurse -Force -ErrorAction SilentlyContinue
	
	IF ($BitCount -eq 32 -Or $BitCount -eq 96)
	{
		IF (!($Version32.Chocolatey))
		{
			$ChocolateyVersion = $SoftwareMaster.CreateElement("Chocolatey")
			$Version32.AppendChild($ChocolateyVersion) | Out-Null
			$Updated = $SoftwareMaster.CreateElement("Updated")
			$ChocolateyVersion.AppendChild($Updated) | Out-Null
			$Updated.InnerText = $Timestamp
		} ELSE {
			$Version32.Chocolatey.Updated = $Timestamp
		}
	}
	
	IF ($BitCount -eq 64 -Or $BitCount -eq 96)
	{
		IF (!($Version64.Chocolatey))
		{
			$ChocolateyVersion = $SoftwareMaster.CreateElement("Chocolatey")
			$Version64.AppendChild($ChocolateyVersion) | Out-Null
			$Updated = $SoftwareMaster.CreateElement("Updated")
			$ChocolateyVersion.AppendChild($Updated) | Out-Null
			$Updated.InnerText = $Timestamp
		} ELSE {
			$Version64.Chocolatey.Updated = $Timestamp
		}
	}
	
	$SoftwareMaster.Save($SoftwareMasterFile)
	
}

IF ($DebugEnable)
{
	$DebugPreference = "Continue"
}

IF (!(Test-Path $SoftwareMasterFile))
{
	Write-Output "Cannot Find $SoftwareMaster. Make sure the account running this script has access to this
	file and try again."
	Exit
}

IF (!(Test-Path $IncludesDir))
{
	Write-Output "Cannot Find Includes Directory $IncludesDir. Cannot create Chocolatey Packages without it."
	Exit
}

$NugetPath = $IncludesDir + "\" + "nuget.exe"

IF (!(Test-Path $NugetPath))
{
	$NugetUrl = "https://nuget.org/nuget.exe"
	$header = "USUS-Chocolatey"
	$WebClient = New-Object System.Net.WebClient
	$WebClient.Headers.Add("user-agent", $header)
	IF (!(Test-Path $NugetPath))
	{
		TRY
		{
			$WebClient.DownloadFile($NugetUrl,$NugetPath)
		} CATCH [System.Net.WebException] {
			Start-Sleep 30
			TRY
			{
				$WebClient.DownloadFile($NugetUrl,$NugetPath)
			} CATCH [System.Net.WebException] {
				Write-Output "Could not download installer from $NugetUrl.
Please check that the web server is reachable. The error was:"
				Write-Output $_.Exception.ToString()
				Write-Output "`r`n"
			}
		}
	}
}

IF (!(Test-Path $NugetPath))
{
	Write-Output "Cannot download Nuget from $NugetUrl. Packaging is unable to complete without this file."
	Exit
}

IF (!(Test-Path $ChocolateyRepo))
{
	Write-Output "Cannot find ChocolateyRepo $ChocolateyRepo. Make sure the account running this script has access to this
	directory and try again."
}

[xml]$SoftwareMaster = Get-Content $SoftwareMasterFile

ForEach ($Software in $SoftwareMaster.SoftwarePackages.software)
{
	$PackageName = $Software.Name
	$HumanReadableName = $Software.HumanReadableName
	
	IF ($Software.IsMSI)
	{
		$IsMSI = $True
	} ELSE {
		$IsMSI = $False
	}
	
	IF ($Software.Versions32)
	{
		$Has32Bit = $True
	} ELSE {
		$Has32Bit = $False
	}
	
	IF ($Software.Versions64)
	{
		$Has64Bit = $True
	} ELSE {
		$Has64Bit = $False
	}
	
	IF ($Has32Bit)
	{
		ForEach ($Version32 in $Software.Versions32.version)
		{
			IF ($Has64Bit)
			{
				$Software32Version = $Version32.name

				IF (($Software.Versions64.version | Where-Object { $_.name -eq $Version32.name }).Count -ne 0)
				{
					$Version64 = $Software.Versions64.version | Where-Object { $_.name -eq $Version32.name } | Select-Object
				} ELSE {
					IF ($Version64)
					{
						Remove-Variable Version64
					}
				} 
			} ELSE {
				IF ($Version64)
				{
					Remove-Variable Version64
				}
			}
			
			IF ($Version64)
			{
				ChocolateyPackage -Version32 $Version32 -Version64 $Version64 -PackageName $PackageName -HumanReadableName $HumanReadableName -Timestamp $Timestamp -IsMSI $IsMSI -BitCount 96 -ChocolateyRepo $ChocolateyRepo -SoftwareMaster $SoftwareMaster -SoftwareMasterFile $SoftwareMasterFile
			} ELSE {
				ChocolateyPackage -Version32 $Version32 -PackageName $PackageName -HumanReadableName $HumanReadableName -Timestamp $Timestamp -IsMSI $IsMSI -BitCount 32 -ChocolateyRepo $ChocolateyRepo -SoftwareMaster $SoftwareMaster -SoftwareMasterFile $SoftwareMasterFile
			}
		}		

	}
	
	IF ($Has64Bit)
	{
		ForEach ($Version64 in $Software.Versions64.version)
		{
			ChocolateyPackage -Version64 $Version64 -PackageName $PackageName -HumanReadableName $HumanReadableName -Timestamp $Timestamp -IsMSI $IsMSI -BitCount 64 -ChocolateyRepo $ChocolateyRepo -SoftwareMaster $SoftwareMaster -SoftwareMasterFile $SoftwareMasterFile
		}
	}	
}