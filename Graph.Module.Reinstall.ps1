#Required running with elevated right.
if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
   Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
   break
}

#Uninstall Microsoft.Graph modules except Microsoft.Graph.Authentication
$moduleNumber = 0
$Modules = Get-Module Microsoft.Graph* -ListAvailable | 
Where-Object {$_.Name -ne "Microsoft.Graph.Authentication"} | Select-Object Name -Unique

Foreach ($Module in $Modules){
  $ModuleName = $Module.Name
  $Versions = Get-Module $ModuleName -ListAvailable
  Foreach ($Version in $Versions){
    $ModuleVersion = $Version.Version
    Write-Progress -Activity "Removing the module $ModuleName..." -Status "Processing $($moduleNumber + 1) / $($Modules.count)"
    Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion -ErrorAction SilentlyContinue
    $moduleNumber++
  }
}

#Fix installed modules
$installedModuleNumber = 0
$InstalledModules = Get-InstalledModule Microsoft.Graph* | 
Where-Object {$_.Name -ne "Microsoft.Graph.Authentication"} | Select-Object Name -Unique

Foreach ($InstalledModule in $InstalledModules){
  $InstalledModuleName = $InstalledModule.Name
  $InstalledVersions = Get-Module $InstalledModuleName -ListAvailable
  Foreach ($InstalledVersion in $InstalledVersions){
    $InstalledModuleVersion = $InstalledVersion.Version
    Write-Progress -Activity "Removing the module $ModuleName..." -Status "Processing $($installedModuleNumber + 1) / $($Modules.count)"
    Uninstall-Module $InstalledModuleName -RequiredVersion $InstalledModuleVersion -ErrorAction SilentlyContinue
    $installedModuleNumber++
  }
}

#Uninstall Microsoft.Graph.Authentication
$i = 0
$ModuleName = "Microsoft.Graph.Authentication"
$Versions = Get-Module $ModuleName -ListAvailable

Foreach ($Version in $Versions){
  $ModuleVersion = $Version.Version
  Write-Progress -Activity "Removing the module $ModuleName..." -Status "Processing $($i + 1) / $($Modules.count)"
  Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion
  $i++
}

#Install-Module Microsoft.Graph
#Install-Module Microsoft.Graph.Beta

Write-Host "`nInstalling the Microsoft Graph PowerShell module..."
Install-Module Microsoft.Graph -Force
Install-Module Microsoft.Graph.Beta -Force
Write-Host "Done."
