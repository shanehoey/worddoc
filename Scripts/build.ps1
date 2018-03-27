exit

set-location $env:USERPROFILE\github\worddoc

$NuGetApiKey = $NuGetApiKey
$NuGetApiKey

$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert
$cert | format-table subject,issuer

$version = "1.2.2"

Update-ModuleManifest -Path ".\WordDoc\WordDoc.psd1" -ModuleVersion $version

Set-AuthenticodeSignature -filepath ".\WordDoc\worddoc.psd1" -Certificate $cert
(Get-AuthenticodeSignature -FilePath ".\WordDoc\worddoc.psd1").Status

Set-AuthenticodeSignature -filepath ".\WordDoc\worddoc.psm1" -Certificate $cert
(Get-AuthenticodeSignature -FilePath ".\WordDoc\worddoc.psm1").Status

Test-ModuleManifest -path ".\WordDoc\WordDoc.psd1" | select -expand ExportedCommands | Fl

Remove-Module WordDoc -ErrorAction SilentlyContinue
Import-Module .\WordDoc\WordDoc.psm1 

get-command -Module WordDoc | select name,version

New-WordInstance
New-WordDocument

$fa_github  = [char]0xf09b
$fontawesometext = "Font Awesome 5 Brands Regular"
add-wordtext  -text $fa_github -Font $fontawesometext -Size 45 -NoParagraph -TextColor wdColorAqua
add-wordtext "https://shanehoey.github.io/worddoc/" -TextColor wdColorAqua

$worddocumment=Get-WordDocument
[Microsoft.Office.Core.MsoAutoShapeType]$shape = "msoShapeRectangle"


$pagewidth = (get-worddocument).pagesetup.pagewidth
$pageheight = (get-worddocument).pagesetup.pageheight

$newshape = $worddocument.shapes.AddShape($shape,0,0,$pagewidth,$pageheight)
$newshape.Line.Weight =10
$newshape.line.Visible = 0




add-wordshape -shape msoShapeRectangle -left 0 -top 0 -Width $pagewidth -Height ($pageheight/2) -zorder msoSendBehindText -UserPicture "http://source.unsplash.com/random" -PictureEffect msoEffectCement 
add-wordshape -shape msoShapeRectangle -left 0 -top ($pageheight/2) -Width $pagewidth -Height ($pageheight/2) -zorder msoSendBehindText -themecolor msoThemeColorDark1 



Close-WordDocument -SaveOptions wdDoNotSaveChanges
Close-WordDocument
Close-WordInstance



#Manually run these 
. .\Scripts\example-1-simple.ps1
. .\Scripts\example-2-detailed.ps1
. .\Scripts\example-3-template.ps1

### MANUAL GitHUB Commit to master

### IMPORTANT ONLY RUN AFTER ALL ABOVE IS COMPLETED
pause
Publish-Module -path .\WordDoc -NuGetApiKey $NuGetApiKey