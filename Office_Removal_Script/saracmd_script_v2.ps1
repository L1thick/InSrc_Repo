# Powershell Script to download SaRACmd & ODT, extract, run, and cleanup.
# This script does not care which version it is, it will remove Office prior;
# Change the "Microsoft Office XML configuration" values as needed.
# Remove everything below line 54 to the 'Cleanup' section.
# Script v2.2
#

# Microsoft Office XML configuration:
[CmdletBinding(DefaultParameterSetName = 'XMLFile')]
param(
  [Parameter(ParameterSetName = 'XMLFile')][String]$XMLFile,
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
  [Switch]$CleanUpInstallFiles = $False
)

# Gather Microsoft Application URLs:
function Get-ODTURL {

  [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'

  $MSWebPage | ForEach-Object {
    if ($_ -match 'url=(https://.*officedeploymenttool.*\.exe)') {
      $matches[1]
    }
  }

}
$ODTLink = Get-ODTURL
$SaRACMDLink = "https://aka.ms/SaRA_CommandLineVersionFiles"

# Create working directory:
New-Item -Path 'C:\saratemp\' -ItemType Directory
New-Item -Path 'C:\saratemp\office\' -ItemType Directory

# Close open Office applications:
#

# Download & Run SaRACMD:
Invoke-WebRequest -Uri $SaRACMDLink -OutFile "C:\saratemp\download.zip"
Expand-Archive -LiteralPath 'C:\saratemp\download.zip' -DestinationPath C:\saratemp\expanded\
cmd.exe /c "C:\saratemp\expanded\SaRACmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion All"
# Start-Process "C:\saratemp\expanded\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All" -Wait -PassThru

# REMOVE ALL BELOW HERE TO "# Cleanup:" IF YOU DO NOT WANT TO REINSTALL OFFICE WITH THIS SCRIPT!

# Create the Office Customization XML file:
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

  $OfficeXML.Save("C:\saratemp\office\OfficeInstall.xml")

}
if (!($XMLFile)) {
  Set-XMLFile
}
$XMLFile = "C:\saratemp\office\OfficeInstall.xml"

# Download & Run Office Deployment Tool:
Invoke-WebRequest -Uri $ODTLink -OutFile "C:\saratemp\office\ODTSetup.exe"
#
# TEST THE CMD COMMANDS
#
# cmd.exe /c "C:\saratemp\office\ODTSetup.exe /quiet /extract:C:\saratemp\office"
# cmd.exe /c "C:\saratemp\office\Setup.exe /configure C:\saratemp\office\OfficeInstall.xml"
Start-Process "C:\saratemp\office\ODTSetup.exe" -ArgumentList "/quiet /extract:C:\saratemp\office" -Wait
Start-Process "C:\saratemp\office\Setup.exe" -ArgumentList "/configure $XMLFile" -Wait -PassThru

# Cleanup:
Remove-Item -LiteralPath "C:\saratemp\" -Force -Recurse
