# Powershell Script to download SaRACmd, extract, and destroy the world.
# This script does not care which version it is, it will remove Office;
# No questions asked. No prisoners. Do not pass 'Go'.

# Intune 'Install Command'
# %windir%\sysnative\windowspowershell\v1.0\powershell.exe -ExecutionPolicy Bypass -file "saracmd_script.ps1"

# SaRACMD URL. Change if needed:
$url = "https://aka.ms/SaRA_CommandLineVersionFiles"


# Create working directory & manage the file:
New-Item -Path 'C:\saratemp\' -ItemType Directory
wget -Uri $url -OutFile "C:\saratemp\download.zip"
Set-Location \saratemp\
Expand-Archive -LiteralPath 'C:\saratemp\download.zip' -DestinationPath C:\saratemp\expanded\
Set-Location \saratemp\expanded\

# Office application management:
#   Close open Office applications
cmd.exe /c "taskkill /f /im winword.exe"
cmd.exe /c "taskkill /f /im excel.exe"
cmd.exe /c "taskkill /f /im outlook.exe"
cmd.exe /c "taskkill /f /im mspub.exe"
cmd.exe /c "taskkill /f /im powerpnt.exe"
cmd.exe /c "taskkill /f /im onenote.exe"
#   Remove Office versions
cmd.exe /c "C:\saratemp\expanded\SaRACmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion All"

# Cleanup after yourself:
Set-Location \
Remove-Item -LiteralPath "C:\saratemp\" -Force -Recurse
exit
