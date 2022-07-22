# Powershell Script to download SaRACmd, extract, and destroy the world.
# This script does not care which version it is, it will remove Office;
# No questions asked. No prisoners. Do not pass 'Go'.
# Script v2.0
# THINGS TO DO STILL:
# [ ]Close open office applications despite versions
# [ ]Run setup.exe silently locally
# [ ]
#

# Office 365 & 2019 configuration parameters
[CmdletBinding(DefaultParameterSetName = 'XMLFile')]
param(
  [Parameter(ParameterSetName = 'XMLFile')][String]$ConfigurationXMLFile,
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$AcceptEULA = 'TRUE',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('Broad', 'Targeted', 'Monthly')]$Channel = 'Broad',
  [Parameter(ParameterSetName = 'NoXML')][Switch]$DisplayInstall = $False,
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('Groove', 'Outlook', 'OneNote', 'Access', 'OneDrive', 'Publisher', 'Word', 'Excel', 'PowerPoint', 'Teams', 'Lync')][Array]$ExcludeApps,
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('64', '32')]$OfficeArch = '64',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('O365ProPlusRetail', 'O365BusinessRetail')]$OfficeEdition = 'O365ProPlusRetail',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet(0, 1)]$SharedComputerLicensing = '0',
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$EnableUpdates = 'TRUE',
  [Parameter(ParameterSetName = 'NoXML')][String]$LoggingPath,
  [Parameter(ParameterSetName = 'NoXML')][String]$SourcePath,
  [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$PinItemsToTaskbar = 'TRUE',
  [Parameter(ParameterSetName = 'NoXML')][Switch]$KeepMSI = $False,
  [String]$OfficeInstallDownloadPath = 'C:\saratemp\office',
  [Switch]$CleanUpInstallFiles = $False
)

# SaRACMD URL:
$url = "https://aka.ms/SaRA_CommandLineVersionFiles"

# Create working directory & manage the files:
New-Item -Path 'C:\saratemp\' -ItemType Directory
New-Item -Path 'C:\saratemp\office\' -ItemType Directory
Invoke-WebRequest -Uri $url -OutFile "C:\saratemp\download.zip"
Expand-Archive -LiteralPath 'C:\saratemp\download.zip' -DestinationPath C:\saratemp\expanded\

# Close open Office applications
#

#   Remove Office versions
# cmd.exe /c "C:\saratemp\expanded\SaRACmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion All"
Start-Process "C:\saratemp\expanded\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All" -Wait -PassThru

# Creates an XML configuration file for the Office Deployment Tool
# Sets a variable to download the Office Deployment Tool
function Set-XMLFile {

  if ($ExcludeApps) {
    $ExcludeApps | ForEach-Object {
      $ExcludeAppsString += "<ExcludeApp ID =`"$_`" />"
    }
  }

  if ($OfficeArch) {
    $OfficeArchString = "`"$OfficeArch`""
  }

  if ($KeepMSI) {
    $RemoveMSIString = $Null
  }
  else {
    $RemoveMSIString = '<RemoveMSI />'
  }

  if ($Channel) {
    $ChannelString = "Channel=`"$Channel`""
  }
  else {
    $ChannelString = $Null
  }

  if ($SourcePath) {
    $SourcePathString = "SourcePath=`"$SourcePath`""
  }
  else {
    $SourcePathString = $Null
  }

  if ($DisplayInstall) {
    $SilentInstallString = 'Full'
  }
  else {
    $SilentInstallString = 'None'
  }

  if ($LoggingPath) {
    $LoggingString = "<Logging Level=`"Standard`" Path=`"$LoggingPath`" />"
  }
  else {
    $LoggingString = $Null
  }

  $OfficeXML = [XML]@"
  <Configuration>
    <Add OfficeClientEdition=$OfficeArchString $ChannelString $SourcePathString  >
      <Product ID="$OfficeEdition">
        <Language ID="MatchOS" />
        $ExcludeAppsString
      </Product>
    </Add>
    <Property Name="PinIconsToTaskbar" Value="$PinItemsToTaskbar" />
    <Property Name="SharedComputerLicensing" Value="$SharedComputerlicensing" />
    <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
    <Updates Enabled="$EnableUpdates" />
    $RemoveMSIString
    $LoggingString
  </Configuration>
"@

  $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")

}
function Get-ODTURL {

  [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'

  $MSWebPage | ForEach-Object {
    if ($_ -match 'url=(https://.*officedeploymenttool.*\.exe)') {
      $matches[1]
    }
  }

}

if (!($ConfigurationXMLFile)) {
  Set-XMLFile
}

$ConfigurationXMLFile = "$OfficeInstallDownloadPath\OfficeInstall.xml"
$ODTInstallLink = Get-ODTURL

# Download the Office Deployment Tool
Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"

# Run the Office Deployment Tool setup
Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait

# Run the O365 install
$Silent = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure $ConfigurationXMLFile" -Wait -PassThru

# Cleanup after yourself:
Set-Location \
Remove-Item -LiteralPath "C:\saratemp\" -Force -Recurse
