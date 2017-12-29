#requires -version 3.0
#requires -module WordDoc

Import-Module Worddoc

$Text     =  "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
$Word     =  New-WordInstance 
$WordDoc  =  New-WordDocument -Word $word

Add-WordCoverPage -CoverPage Banded -Word $Word -WordDoc $WordDoc 
Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle -WordDoc $WordDoc 
Add-WordTOC -word $word -WordDoc $WordDoc 
Add-WordBreak -breaktype NewPage -word $Word -WordDoc $WordDoc 
Add-WordText -text 'Heading1' -WDBuiltinStyle wdStyleHeading1 -WordDoc $WordDoc
Add-WordText -text $Text -WDBuiltinStyle wdStyleNormal -WordDoc $WordDoc 
