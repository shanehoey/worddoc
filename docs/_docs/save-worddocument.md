---
title: "save-worddocument"
excerpt: "Describe purpose of "Save-WordDocument" in 1-2 sentences."
category: "help"
---

# Save-WordDocument
**Module** WordDoc

## SYNOPSIS
Describe purpose of "Save-WordDocument" in 1-2 sentences.

## DESCRIPTION
Add a more complete description of what the function does.

## SYNTAX

```
Save-WordDocument [-WordDocument <Document>] [[-WordSaveFormat] {wdFormatDocument | wdFormatDocument97 | wdFormatTemplate | wdFormatTemplate97 | wdFormatText | wdFormatTextLineBreaks | wdFormatDOSText | 
wdFormatDOSTextLineBreaks | wdFormatRTF | wdFormatUnicodeText | wdFormatEncodedText | wdFormatHTML | wdFormatWebArchive | wdFormatFilteredHTML | wdFormatXML | wdFormatXMLDocument | wdFormatXMLDocumentMacroEnabled | 
wdFormatXMLTemplate | wdFormatXMLTemplateMacroEnabled | wdFormatDocumentDefault | wdFormatPDF | wdFormatXPS | wdFormatFlatXML | wdFormatFlatXMLMacroEnabled | wdFormatFlatXMLTemplate | wdFormatFlatXMLTemplateMacroEnabled | 
wdFormatOpenDocumentText | wdFormatStrictOpenXMLDocument}] [[-filename] <String>] [-folder <String>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>Save-WordDocument -WordDocument Value -WordSaveFormat Value -filename Value -folder Value
```

Describe what this call does


## PARAMETERS

### WordDocument



```
Type Document
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:$Script:WordDocument
Accept pipeline input: false
```
### WordSaveFormat

Describe parameter -WordSaveFormat.

```
Type WdSaveFormat
Parameter Sets: 
Aliases: 
Required: false
Position: 1
Default Value:wdFormatDocumentDefault
Accept pipeline input: false
```
### filename

Describe parameter -filename.

```
Type String
Parameter Sets: 
Aliases: 
Required: false
Position: 2
Default Value:document.docx
Accept pipeline input: false
```
### folder

Describe parameter -folder.

```
Type String
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:[Environment]::GetFolderPath('MyDocuments')
Accept pipeline input: false
```
### CommonParameters

This function only supports -verbose

## RELATED LINKS


https://shanehoey.github.io/worddoc/docs/save-worddocument
