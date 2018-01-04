#requires -version 4.0
#requires -module WordDoc

$Text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse et urna eget lectus rhoncus molestie rutrum in diam. Etiam quis convallis risus. Phasellus quis viverra nulla. Etiam non eleifend enim. Fusce dictum euismod mauris, sit amet pulvinar tellus iaculis vitae. Suspendisse finibus lobortis consequat. Morbi a lobortis libero. Etiam est orci, facilisis ac lacus et, maximus aliquam elit. Nulla sed nisl a elit convallis pulvinar ultrices eget urna."
$Service = Get-Service | Where-Object status -eq stopped | Where-Object name -like p*
$Process = Get-Process | Where-Object name -like p* | Select-Object -Property name,id

$wi1 = New-WordInstance -returnobject
$wi1.Name

$wi2 = New-WordInstance -returnobject
$wi2.Name

$wd1 = New-WordDocument -returnobject -WordInstance $wi1
$wd1.Name

$wd2 = New-WordDocument -returnobject -WordInstance $wi2
$wd2.Name

Add-WordCoverPage -CoverPage Facet -WordInstance $wi1 -WordDocument $wd1
Add-WordCoverPage -CoverPage Banded -WordInstance $wi2 -WordDocument $wd2

Add-WordBreak -breaktype NewPage -WordInstance $wi1 -WordDocument $wd1
Add-WordBreak -breaktype NewPage -WordInstance $wi2 -WordDocument $wd2

Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle -WordDocument $wd1 
Add-WordTOC  -Tableader 0 -IncludePageNumbers $true  -WordInstance $wi1 -wordDocument $wd1

Add-WordBreak -breaktype NewPage  -WordInstance $wi1 -WordDocument $wd1

Add-WordText -text "Heading1" -WDBuiltinStyle wdStyleHeading1 -WordDocument $wd1
Add-WordText -text $text -WDBuiltinStyle wdStyleDefaultParagraphFont -WordDocument $wd1

Add-WordText -text "Heading1" -WDBuiltinStyle wdStyleHeading2 -WordDocument $wd2
Add-WordText -text $text -WDBuiltinStyle wdStyleDefaultParagraphFont -WordDocument $wd2


Add-WordTable -Object $service -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -WordDocument $wd1
Add-WordTable -Object $process -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -WordDocument $wd2

Set-WordOrientation -Orientation Landscape -WordInstance $wi1

Update-WordTOC -WordDocument $wd1

Set-WordBuiltInProperty "My Title" -WdBuiltInProperty wdPropertyTitle -WordDocument $wd1
Set-WordBuiltInProperty "Shane Hoey" -WdBuiltInProperty wdPropertyAuthor -WordDocument $wd1
Set-WordBuiltInProperty "My SubTitle" -WdBuiltInProperty wdPropertySubject -WordDocument $wd1

Set-WordBuiltInProperty "My Title" -WdBuiltInProperty wdPropertyTitle -WordDocument $wd1
Set-WordBuiltInProperty "Shane Hoey" -WdBuiltInProperty wdPropertyAuthor -WordDocument $wd1
Set-WordBuiltInProperty "My SubTitle" -WdBuiltInProperty wdPropertySubject -WordDocument $wd1


Save-WordDocument -WordDocument $wd1
Save-WordDocument -filename worddoc.doc -WordSaveFormat wdFormatDocument -WordDocument $wd2

Close-WordDocument -WordDocument $wd1  -WordInstance $wi1 
Close-WordInstance -WordInstance $wi1

Close-WordInstance -WordInstance $wi2



