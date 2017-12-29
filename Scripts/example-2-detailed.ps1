#requires -version 3.0
#requires -module WordDoc

Import-Module -Name Worddoc -Force

$Text1 = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Suspendisse et urna eget lectus rhoncus molestie rutrum in diam. Etiam quis convallis risus. Phasellus quis viverra nulla. Etiam non eleifend enim. Fusce dictum euismod mauris, sit amet pulvinar tellus iaculis vitae. Suspendisse finibus lobortis consequat. Morbi a lobortis libero. Etiam est orci, facilisis ac lacus et, maximus aliquam elit. Nulla sed nisl a elit convallis pulvinar ultrices eget urna."
$Text2 = "Mauris quis mattis lorem. Curabitur interdum commodo velit non interdum. Morbi auctor purus vel enim consectetur tempor. Nunc non nisl in felis blandit porta. Donec pellentesque felis id diam semper, ac egestas lectus ullamcorper. Aliquam feugiat purus eget quam elementum, ac viverra tortor elementum. Fusce tincidunt et purus quis sollicitudin. Aliquam gravida vel leo et posuere. Aenean rhoncus ante nec sapien semper, at tempus tellus dictum. Pellentesque risus risus, facilisis sit amet metus rutrum, semper lobortis orci. Quisque viverra, tellus nec pulvinar rhoncus, tortor massa faucibus lorem, ut ullamcorper mi mi nec dui. Aliquam id nulla eget nunc aliquet mattis vitae ut risus."

$Service = Get-Service | Where-Object status -eq stopped | Where-Object name -like p*
$Process = Get-Process | Where-Object name -like p* | Select-Object -Property name,id

$Word  = New-WordInstance -Visable $True -Verbose 
$WordDoc = New-WordDocument -WordInstance $Word  

Add-WordCoverPage -CoverPage Banded -WordInstance $Word  -WordDoc $WordDoc
Add-WordBreak -breaktype NewPage -WordInstance $Word  -WordDoc $WordDoc
Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle -WordDoc $WordDoc
Add-WordTOC -Word $Word  -WordDoc $WordDoc
Add-WordBreak -breaktype NewPage -WordInstance $Word  -WordDoc $WordDoc
Add-WordText -text 'Heading1' -WDBuiltinStyle wdStyleHeading1 -WordDoc $WordDoc
Add-WordText -text $text1 -WDBuiltinStyle wdStyleNormal -WordDoc $WordDoc
Add-WordText -text 'Heading2' -WDBuiltinStyle wdStyleHeading2 -WordDoc $WordDoc
Add-WordText -text $text2 -WDBuiltinStyle wdStyleNormal -WordDoc $WordDoc
Add-WordText -text 'Heading3' -WDBuiltinStyle wdStyleHeading3 -WordDoc $WordDoc
Add-WordText -text $text1 -WDBuiltinStyle wdStyleNormal -WordDoc $WordDoc
Add-WordText -text 'Heading4' -WDBuiltinStyle wdStyleHeading4 -WordDoc $WordDoc
Add-WordText -text $text2 -WDBuiltinStyle wdStyleNormal -WordDoc $WordDoc
Add-WordText -text 'Heading5' -WDBuiltinStyle wdStyleHeading5 -WordDoc $WordDoc 
Add-WordText -text $text1 -WDBuiltinStyle wdStyleNormal -WordDoc $WordDoc
Add-WordText -text 'Heading6' -WDBuiltinStyle wdStyleHeading6 -WordDoc $WordDoc
Add-WordText -text $text2 -WDBuiltinStyle wdStyleNormal -WordDoc $WordDoc
Add-WordText -text 'Bullet' -WDBuiltinStyle wdStyleListBullet -WordDoc $WordDoc
Add-WordText -text 'Bullet' -WDBuiltinStyle wdStyleListBullet -WordDoc $WordDoc

Add-WordTable -Object $Process -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -FirstColumn $false -WordDoc $WordDoc
Add-WordTable -Object $Process -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 5 Dark' -GridAccent 'Accent 5' -VerticleTable -HeaderRow $false -WordDoc $WordDoc
Add-WordBreak -breaktype NewPage -WordInstance $Word  -WordDoc $WordDoc
Add-WordBreak -breaktype Section -WordInstance $Word  -WordDoc $WordDoc
Add-WordText -text title -WDBuiltinStyle wdStyleHeading1 -WordDoc $WordDoc
Set-WordOrientation -Orientation Landscape -WordInstance $Word 

Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -GridTable 'Grid Table 4' -GridAccent 'Accent 5' -WordDoc $WordDoc
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormat3DEffects1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormat3DEffects2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormat3DEffects3
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatClassic1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatClassic2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatClassic3
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatClassic4
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatColorful1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatColorful2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatColorful3
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatColumns1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatColumns2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatColumns3
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatColumns4
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatColumns5
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatContemporary
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatElegant
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatGrid1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatGrid2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatGrid3
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatGrid4
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatGrid5
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatGrid6
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatGrid7
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatGrid8
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatList1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatList2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatList3
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatList4
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatList5
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatList6
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatList7
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatList8
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatNone
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatProfessional
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatSimple1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatSimple2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatSimple3
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatSubtle1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatSubtle2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatWeb1
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatWeb2
Add-WordTable -Object ($Service | Select-Object -Property Displayname,ServiceName,Status) -WordDoc $WordDoc -WdAutoFitBehavior wdAutoFitWindow -WdDefaultTableBehavior wdWord9TableBehavior  -WDTableFormat wdTableFormatWeb3

Add-WordBreak -breaktype Section -WordInstance $Word  -WordDoc $WordDoc
Add-WordBreak -breaktype NewPage -WordInstance $Word  -WordDoc $WordDoc

Set-WordOrientation -Orientation Portrait -WordInstance $Word 
Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyTitle   -Text "Title"      -WordDoc $WordDoc
Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyCompany -Text "Company"    -WordDoc $WordDoc
Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyAuthor  -Text "Shane Hoey" -WordDoc $WordDoc
Update-WordTOC -WordDoc $WordDoc 

#Save-WordDocument -WordSaveFormat wdFormatHTML -filename scripthtml -folder $home\documents
#Save-WordDocument -WordSaveFormat wdFormatPDF -filename scriptpdf -folder $home\documents
Save-WordDocument -WordSaveFormat wdFormatDocument -filename scriptdoc -folder $home\documents
Close-WordDocument -Word $Word -WordDoc $WordDoc
