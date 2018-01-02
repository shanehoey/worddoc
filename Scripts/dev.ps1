
break

## DO NOT USE 

### This file just a scrath file for testing 

Set-Location .\worddoc

Import-Module .\WordDoc\WordDoc.psd1 
Import-Module .\WordDoc\WordDoc.psd1 -Force

Import-Module .\WordDoc\WordDoc.psm1
Import-Module .\WordDoc\WordDoc.psm1 -Force 

remove-Module WordDoc

New-WordInstance
get-WordInstance

New-WordDocument
New-WordDocument


[Enum]::GetNames([Microsoft.Office.Interop.Word.WdPageColor]) 

[Enum]::GetNames([Microsoft.Office.Interop.Word.WdSaveFormat ]) 
[Enum]::GetNames([Microsoft.Office.Interop.Word.WdSaveOptions ]) 


$document.Styles | select NameLocal


[enum]::GetNames([microsoft.office.interop.word])

[Microsoft.Office.Interop.Word.WdAlertLevel
[Microsoft.Office.Interop.Word.WdAlignmentTabAlignment
[Microsoft.Office.Interop.Word.WdAlignmentTabRelative
[Microsoft.Office.Interop.Word.WdAnimation
[enum]::GetNames([Microsoft.Office.Interop.Word.WdApplyQuickStyleSets])

[enum]::GetNames([Microsoft.Office.Interop.Word.WdTableFormat])

[Microsoft.Office.Interop.Word.WdArabicNumeral
[Microsoft.Office.Interop.Word.WdAraSpeller
[Microsoft.Office.Interop.Word.WdArrangeStyle
[Microsoft.Office.Interop.Word.WdAutoFitBehavior
[Microsoft.Office.Interop.Word.WdAutoMacros
[Microsoft.Office.Interop.Word.WdAutoVersions
[Microsoft.Office.Interop.Word.WdBaselineAlignment
[Microsoft.Office.Interop.Word.WdBookmarkSortBy
[Microsoft.Office.Interop.Word.WdBorderDistanceFrom
[Microsoft.Office.Interop.Word.WdBorderType
[Microsoft.Office.Interop.Word.WdBorderTypeHID
[Microsoft.Office.Interop.Word.WdBreakType
[Microsoft.Office.Interop.Word.WdBrowserLevel
[Microsoft.Office.Interop.Word.WdBrowseTarget
[Microsoft.Office.Interop.Word.WdBuildingBlockTypes
[enum]::GetNames([Microsoft.Office.Interop.Word.WdBuiltInProperty])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdBuiltinStyle])
[Microsoft.Office.Interop.Word.WdCalendarType
[Microsoft.Office.Interop.Word.WdCalendarTypeBi
[Microsoft.Office.Interop.Word.WdCaptionLabelID
[Microsoft.Office.Interop.Word.WdCaptionNumberStyle
[Microsoft.Office.Interop.Word.WdCaptionNumberStyleHID
[Microsoft.Office.Interop.Word.WdCaptionPosition
[enum]::GetNames([Microsoft.Office.Interop.Word.WdCellColor])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdCellVerticalAlignment])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdCharacterCase])

[enum]::GetNames([Microsoft.Office.Interop.Word.WdDefaultListBehavior])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdDefaultTableBehavior])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdPageColor])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdPageFit])

[enum]::GetNames([Microsoft.Office.Interop.Word.WdThemeColorIndex])
$s = [Microsoft.Office.Interop.Word.WdThemeColorIndex]"wdThemeColorAccent5"

[enum]::GetNames([Microsoft.Office.Interop.Word.WdTableDirection])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdTableFieldSeparator])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdTableFormat])
$table.style = [Microsoft.Office.Interop.Word.WdTableFormat]"wdTableFormatList5"
$table.linkstyle = [Microsoft.Office.Interop.Word.WdTableFormat]"wdTableFormatList5"
[enum]::GetNames([Microsoft.Office.Interop.Word.WdTableFormatApply])
[enum]::GetNames([Microsoft.Office.Interop.Word.WdTablePosition])
