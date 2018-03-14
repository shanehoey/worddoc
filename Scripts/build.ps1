exit

set-location $env:USERPROFILE\github\worddoc

$NuGetApiKey = $NuGetApiKey
$NuGetApiKey

$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert
$cert | format-table subject,issuer

$version = "1.2.1"

Update-ModuleManifest -Path ".\WordDoc\WordDoc.psd1" -ModuleVersion $version

Set-AuthenticodeSignature -filepath ".\WordDoc\worddoc.psd1" -Certificate $cert
(Get-AuthenticodeSignature -FilePath ".\WordDoc\worddoc.psd1").Status

Set-AuthenticodeSignature -filepath ".\WordDoc\worddoc.psm1" -Certificate $cert
(Get-AuthenticodeSignature -FilePath ".\WordDoc\worddoc.psm1").Status

Test-ModuleManifest -path ".\WordDoc\WordDoc.psd1"

Remove-Module WordDoc -ErrorAction SilentlyContinue
Import-Module .\WordDoc\WordDoc.psd1 

New-WordInstance
New-WordDocument

$fa_github  = [char]0xf09b
$fontawesometext = "Font Awesome 5 Brands Regular"
add-wordtext  -text $fa_github -Font $fontawesometext -Size 45 -NoParagraph
add-wordtext "https://shanehoey.github.io/worddoc/"

Close-WordDocument -SaveOptions wdDoNotSaveChanges
Close-WordDocument
Close-WordInstance

get-command -Module WordDoc | select name,version

#Manually run these 
. .\Scripts\example-1-simple.ps1
. .\Scripts\example-2-detailed.ps1
. .\Scripts\example-3-template.ps1

### MANUAL GitHUB Commit to master

### IMPORTANT ONLY RUN AFTER ALL ABOVE IS COMPLETED
pause
Publish-Module -path .\WordDoc -NuGetApiKey $NuGetApiKey