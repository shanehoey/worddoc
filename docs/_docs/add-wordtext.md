---
title: "add-wordtext"
excerpt: "Adds text to MS Word Document."
category: "help"
---

# Add-WordText
**Module** WordDoc

## SYNOPSIS
Adds text to MS Word Document.

## DESCRIPTION
Adds text to MS Word Document.

## SYNTAX

```
Add-WordText [-text] <String> [-WdColor {wdColorBlack | wdColorDarkRed | wdColorRed | wdColorDarkGreen | wdColorOliveGreen | wdColorBrown | wdColorOrange | wdColorGreen | wdColorDarkYellow | wdColorLightOrange | wdColorLime | 
wdColorGold | wdColorBrightGreen | wdColorYellow | wdColorGray95 | wdColorGray90 | wdColorGray875 | wdColorGray85 | wdColorGray80 | wdColorGray75 | wdColorGray70 | wdColorGray65 | wdColorGray625 | wdColorDarkTeal | wdColorPlum 
| wdColorGray60 | wdColorSeaGreen | wdColorGray55 | wdColorDarkBlue | wdColorViolet | wdColorTeal | wdColorGray50 | wdColorGray45 | wdColorIndigo | wdColorBlueGray | wdColorGray40 | wdColorTan | wdColorLightYellow | 
wdColorGray375 | wdColorGray35 | wdColorGray30 | wdColorGray25 | wdColorRose | wdColorAqua | wdColorGray20 | wdColorLightGreen | wdColorGray15 | wdColorGray125 | wdColorGray10 | wdColorGray05 | wdColorBlue | wdColorPink | 
wdColorLightBlue | wdColorLavender | wdColorSkyBlue | wdColorPaleBlue | wdColorTurquoise | wdColorLightTurquoise | wdColorWhite | wdColorAutomatic}] [-WDBuiltinStyle {wdStyleTocHeading | wdStyleBibliography | wdStyleBookTitle 
| wdStyleIntenseReference | wdStyleSubtleReference | wdStyleIntenseEmphasis | wdStyleSubtleEmphasis | wdStyleIntenseQuote | wdStyleQuote | wdStyleListParagraph | wdStyleTableMediumList1Accent1 | 
wdStyleTableMediumShading2Accent1 | wdStyleTableMediumShading1Accent1 | wdStyleTableLightGridAccent1 | wdStyleTableLightListAccent1 | wdStyleTableLightShadingAccent1 | wdStyleTableColorfulGrid | wdStyleTableColorfulList | 
wdStyleTableColorfulShading | wdStyleTableDarkList | wdStyleTableMediumGrid3 | wdStyleTableMediumGrid2 | wdStyleTableMediumGrid1 | wdStyleTableMediumList2 | wdStyleTableMediumList1 | wdStyleTableMediumShading2 | 
wdStyleTableMediumShading1 | wdStyleTableLightGrid | wdStyleTableLightList | wdStyleTableLightShading | wdStyleNormalObject | wdStyleNormalTable | wdStyleHtmlVar | wdStyleHtmlTt | wdStyleHtmlSamp | wdStyleHtmlPre | 
wdStyleHtmlKbd | wdStyleHtmlDfn | wdStyleHtmlCode | wdStyleHtmlCite | wdStyleHtmlAddress | wdStyleHtmlAcronym | wdStyleHtmlNormal | wdStylePlainText | wdStyleNavPane | wdStyleEmphasis | wdStyleStrong | wdStyleHyperlinkFollowed 
| wdStyleHyperlink | wdStyleBlockQuotation | wdStyleBodyTextIndent3 | wdStyleBodyTextIndent2 | wdStyleBodyText3 | wdStyleBodyText2 | wdStyleNoteHeading | wdStyleBodyTextFirstIndent2 | wdStyleBodyTextFirstIndent | wdStyleDate | 
wdStyleSalutation | wdStyleSubtitle | wdStyleMessageHeader | wdStyleListContinue5 | wdStyleListContinue4 | wdStyleListContinue3 | wdStyleListContinue2 | wdStyleListContinue | wdStyleBodyTextIndent | wdStyleBodyText | 
wdStyleDefaultParagraphFont | wdStyleSignature | wdStyleClosing | wdStyleTitle | wdStyleListNumber5 | wdStyleListNumber4 | wdStyleListNumber3 | wdStyleListNumber2 | wdStyleListBullet5 | wdStyleListBullet4 | wdStyleListBullet3 
| wdStyleListBullet2 | wdStyleList5 | wdStyleList4 | wdStyleList3 | wdStyleList2 | wdStyleListNumber | wdStyleListBullet | wdStyleList | wdStyleTOAHeading | wdStyleMacroText | wdStyleTableOfAuthorities | wdStyleEndnoteText | 
wdStyleEndnoteReference | wdStylePageNumber | wdStyleLineNumber | wdStyleCommentReference | wdStyleFootnoteReference | wdStyleEnvelopeReturn | wdStyleEnvelopeAddress | wdStyleTableOfFigures | wdStyleCaption | 
wdStyleIndexHeading | wdStyleFooter | wdStyleHeader | wdStyleCommentText | wdStyleFootnoteText | wdStyleNormalIndent | wdStyleTOC9 | wdStyleTOC8 | wdStyleTOC7 | wdStyleTOC6 | wdStyleTOC5 | wdStyleTOC4 | wdStyleTOC3 | 
wdStyleTOC2 | wdStyleTOC1 | wdStyleIndex9 | wdStyleIndex8 | wdStyleIndex7 | wdStyleIndex6 | wdStyleIndex5 | wdStyleIndex4 | wdStyleIndex3 | wdStyleIndex2 | wdStyleIndex1 | wdStyleHeading9 | wdStyleHeading8 | wdStyleHeading7 | 
wdStyleHeading6 | wdStyleHeading5 | wdStyleHeading4 | wdStyleHeading3 | wdStyleHeading2 | wdStyleHeading1 | wdStyleNormal}] [-WordDocument <Document>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>Add-WordText -text "Heading 1" -WdColor Value -WDBuiltinStyle Value -WordDocument Value
```

Adds text to document

### -------------------------- EXAMPLE 2 --------------------------


```
PS C:\>Add-WordText -text "Heading 1" -WdColor Value -WDBuiltinStyle Value -WordDocument Value
```

Adds text to document


## PARAMETERS

### text

Text to add to word Document

```
Type String
Parameter Sets: 
Aliases: 
Required: true
Position: 1
Default Value:
Accept pipeline input: false
```
### WdColor

Color of Text

```
Type WdColor
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:wdColorAutomatic
Accept pipeline input: false
```
### WDBuiltinStyle

Builtin Stype to use

```
Type WdBuiltinStyle
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:wdStyleDefaultParagraphFont
Accept pipeline input: false
```
### WordDocument

WordDocument Object

```
Type Document
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:$Script:WordDocument
Accept pipeline input: false
```
### CommonParameters

This function only supports -verbose

## RELATED LINKS


https://shanehoey.github.io/worddoc/docs/add-wordtext
