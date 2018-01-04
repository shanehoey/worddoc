#requires -version 4.0
#requires -module WordDoc

$object   = Get-Process -name po* | Select-Object name,id | Select-Object -First 3
$Text     =  "Lorem ipsum dolor sit amet, consectetur adipiscing elit."

New-WordInstance
New-WordDocument

Add-WordTemplate 

Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle
Add-WordTOC  -Tableader 0 -IncludePageNumbers $true
Add-WordBreak -breaktype NewPage 

Add-WordText -text "Heading 1" -WDBuiltinStyle wdStyleHeading1
Add-WordText -text $text -WDBuiltinStyle wdStyleDefaultParagraphFont

Add-WordText -text "Heading 2" -WDBuiltinStyle  wdStyleHeading1
Add-WordTable -Object $object -WDTableFormat wdTableFormatColorful1

Close-WordDocument
Close-WordInstance
