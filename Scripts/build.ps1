set-location $env:USERPROFILE\github\worddoc

if ($NuGetApiKey -eq $null) { 
    Write-Error -Message "No PowerShell GalleryAPI" 
    Break
    }
else {
    Write-host "Using $powershellgalleryAPI"
    }

$psd = Import-PowerShellDataFile -path ".\WordDoc\WordDoc.psd1"

Update-ModuleManifest -Path ".\WordDoc\WordDoc.psd1" -ModuleVersion "1.1.$($(get-date).DayofYear)"
Test-ModuleManifest -path ".\WordDoc\WordDoc.psd1"

$psd.RootModule
$psd.ModuleVersion
$psd.Copyright

Write-host "Check date before continuing -> BUILD ((GET-DATE).DayOfYear)"
pause
Import-Module .\module\WordDoc.psd1
. .\Scripts\example-1-simple.ps1
. .\Scripts\example-1-detailed.ps1
Write-host "Only continue if all OK"
pause
Publish-Module -path ".\Module\" -NuGetApiKey $NuGetApiKey -WhatIf