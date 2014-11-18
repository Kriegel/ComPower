# test modulcontext for PowerShell version 2.0
$ModuleRootPath = $MyInvocation.MyCommand.Module.ModuleBase

# if Modulecontext is empty
# test Scriptcontext for PowerShell version 2.0
If(-not $ModuleRootPath) {
  $ModuleRootPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# if we are running PowerShell 3.0 or higher we are using the $PSScriptRoot variable
If(([int]$PSVersionTable.PSVersion.Major) -gt 2) {
    $ModuleRootPath = $PSScriptRoot
}

# The Registry entrys are changing only if a new COM Component ist registered or unregistered
# This happens not very often so we use a Variable to hold the Registered COM components from Registry
# as Array which is used as a static cache
$Script:ComRegisteredCache = [System.Collections.ArrayList]@()

# load Functions and include it into the Module scope 

. (Join-Path -Path $ModuleRootPath -ChildPath 'Get-ComRegistered.ps1')
. (Join-Path -Path $ModuleRootPath -ChildPath 'Get-ComRot.ps1')
. (Join-Path -Path $ModuleRootPath -ChildPath 'Get-ComRotRegistered.ps1')
. (Join-Path -Path $ModuleRootPath -ChildPath 'Remove-ComObject.ps1')
. (Join-Path -Path $ModuleRootPath -ChildPath 'New-Object.ps1')


# init the Registry cache
Write-Verbose 'Loading informations from Registry, this takes a while!' -Verbose
Get-ComRegistered
Write-Verbose 'Finish to load informations from Registry. Module ComPower is ready to use!' -Verbose