#requires -version 4.0
#requires -module WordDoc

$object   = Get-Process -name po* | Select-Object name,id | Select-Object -First 3
$Text     =  "Lorem ipsum dolor sit amet, consectetur adipiscing elit."

New-WordInstance
New-WordDocument

Add-WordCoverPage -CoverPage Facet 
Add-WordBreak -breaktype NewPage

Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle
Add-WordTOC  -Tableader 0 -IncludePageNumbers $true
Add-WordBreak -breaktype NewPage 

Add-WordText -text "Heading1" -WDBuiltinStyle wdStyleHeading1
Add-WordText -text $text -WDBuiltinStyle wdStyleDefaultParagraphFont

Add-WordText -text "Heading2" -WDBuiltinStyle wdStyleHeading2
Add-WordText -text $text -WDBuiltinStyle wdStyleDefaultParagraphFont

Add-WordText -text "Heading3" -WDBuiltinStyle wdStyleHeading3
Add-WordText -text $text -WDBuiltinStyle wdStyleDefaultParagraphFont

Add-WordText -text "Heading4" -WDBuiltinStyle wdStyleHeading4 -WdColor wdColorGreen
Add-WordText -text $text -WDBuiltinStyle wdStyleDefaultParagraphFont 

Add-WordBreak -breaktype Section
Add-WordBreak -breaktype NewPage
Set-WordOrientation -Orientation Landscape

Add-WordText -text "Heading1" -WDBuiltinStyle  wdStyleHeading1
Add-WordTable -Object $object -WDTableFormat wdTableFormatGrid1
Add-WordTable -Object $object -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5'
Add-WordTable -Object $object -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -FirstColumn $false -BandedRow $false
Add-WordTable -Object $object -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -FirstColumn $false -BandedRow $false -VerticleTable
Add-WordTable -Object $object -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -FirstColumn $false -BandedRow $false -VerticleTable -WdAutoFitBehavior wdAutoFitWindow 

Add-WordBreak -breaktype Section
Set-WordOrientation -Orientation Portrait

Update-WordTOC 

Set-WordBuiltInProperty "My Title" -WdBuiltInProperty wdPropertyTitle
Set-WordBuiltInProperty "Shane Hoey" -WdBuiltInProperty wdPropertyAuthor
Set-WordBuiltInProperty "My SubTitle" -WdBuiltInProperty wdPropertySubject

Save-WordDocument
Save-WordDocument -filename worddoc.pdf -WordSaveFormat wdFormatPDF
Save-WordDocument -filename worddoc.docx -WordSaveFormat wdFormatDocumentdefault
Save-WordDocument -filename worddoc.doc -WordSaveFormat wdFormatDocument
Save-WordDocument -filename worddoc.html -WordSaveFormat wdFormatHTML

Close-WordDocument
Close-WordInstance
