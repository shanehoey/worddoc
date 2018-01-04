#requires -version 4.0
#requires -module WordDoc

Import-Module -Name Worddoc 

$Text1 = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse et urna eget lectus rhoncus molestie rutrum in diam. Etiam quis convallis risus. Phasellus quis viverra nulla. Etiam non eleifend enim. Fusce dictum euismod mauris, sit amet pulvinar tellus iaculis vitae. Suspendisse finibus lobortis consequat. Morbi a lobortis libero. Etiam est orci, facilisis ac lacus et, maximus aliquam elit. Nulla sed nisl a elit convallis pulvinar ultrices eget urna."
$Text2 = "Mauris quis mattis lorem. Curabitur interdum commodo velit non interdum. Morbi auctor purus vel enim consectetur tempor. Nunc non nisl in felis blandit porta. Donec pellentesque felis id diam semper, ac egestas lectus ullamcorper. Aliquam feugiat purus eget quam elementum, ac viverra tortor elementum. Fusce tincidunt et purus quis sollicitudin. Aliquam gravida vel leo et posuere. Aenean rhoncus ante nec sapien semper, at tempus tellus dictum. Pellentesque risus risus, facilisis sit amet metus rutrum, semper lobortis orci. Quisque viverra, tellus nec pulvinar rhoncus, tortor massa faucibus lorem, ut ullamcorper mi mi nec dui. Aliquam id nulla eget nunc aliquet mattis vitae ut risus."

$Service = Get-Service | Where-Object status -eq stopped | Where-Object name -like p*
$Process = Get-Process | Where-Object name -like p* | Select-Object -Property name,id

$wi = New-WordInstance
$wd = New-WordDocument

get-WordInstance -WordInstance $w1
get-WordDocument -WordDocument $wd 

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

Save-WordDocument -filename worddoc.html -WordSaveFormat wdFormatHTML
Save-WordDocument -filename worddoc.pdf -WordSaveFormat wdFormatPDF
Save-WordDocument -filename worddoc.docx -WordSaveFormat wdFormatDocumentDefault
