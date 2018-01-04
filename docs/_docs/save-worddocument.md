---
title: "save-worddocument"
excerpt: "Save a Word Document (also .pdf and .html)"
category: "help"
---

# Save-WordDocument
**Module** WordDoc

## SYNOPSIS
Save a Word Document (also .pdf and .html)

## DESCRIPTION
Save a Word Document (also .pdf and .html)

## SYNTAX

```
Save-WordDocument [-WordDocument <Document>] -filename <String> -WordSaveFormat {wdFormatDocument | wdFormatDocument97 | wdFormatTemplate | wdFormatTemplate97 | wdFormatText | wdFormatTextLineBreaks | wdFormatDOSText | 
wdFormatDOSTextLineBreaks | wdFormatRTF | wdFormatUnicodeText | wdFormatEncodedText | wdFormatHTML | wdFormatWebArchive | wdFormatFilteredHTML | wdFormatXML | wdFormatXMLDocument | wdFormatXMLDocumentMacroEnabled | 
wdFormatXMLTemplate | wdFormatXMLTemplateMacroEnabled | wdFormatDocumentDefault | wdFormatPDF | wdFormatXPS | wdFormatFlatXML | wdFormatFlatXMLMacroEnabled | wdFormatFlatXMLTemplate | wdFormatFlatXMLTemplateMacroEnabled | 
wdFormatOpenDocumentText | wdFormatStrictOpenXMLDocument} [-folder <String>] [<CommonParameters>]

Save-WordDocument [-WordDocument <Document>] [-folder <String>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>Save-WordDocument -WordSaveFormat wdFormatDocument -filename worddoc.docx -folder c:\users\shane\documents\
```

Saves document as a standard Word Document in c:\users\shane\documents\worddoc.docx

### -------------------------- EXAMPLE 2 --------------------------


```
PS C:\>Save-WordDocument
```

Opens a save-as GUI, allowing you to save as a docx, html, or pdf file.


## PARAMETERS

### WordDocument

Word Document to save

```
Type Document
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:$Script:WordDocument
Accept pipeline input: false
```
### filename

Filename to save document as

```
Type String
Parameter Sets: 
Aliases: 
Required: true
Position: named
Default Value:
Accept pipeline input: false
```
### WordSaveFormat

Format to save document as.

```
Type WdSaveFormat
Parameter Sets: 
Aliases: 
Required: true
Position: named
Default Value:
Accept pipeline input: false
```
### folder

Folder to save document in

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
