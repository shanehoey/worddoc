---
title: "set-wordbuiltinproperty"
excerpt: "Describe purpose of "Set-WordBuiltInProperty" in 1-2 sentences."
category: "help"
---

# Set-WordBuiltInProperty
**Module** WordDoc

## SYNOPSIS
Describe purpose of "Set-WordBuiltInProperty" in 1-2 sentences.

## DESCRIPTION
Add a more complete description of what the function does.

## SYNTAX

```
Set-WordBuiltInProperty [-WdBuiltInProperty] {wdPropertyTitle | wdPropertySubject | wdPropertyAuthor | wdPropertyKeywords | wdPropertyComments | wdPropertyTemplate | wdPropertyLastAuthor | wdPropertyRevision | 
wdPropertyAppName | wdPropertyTimeLastPrinted | wdPropertyTimeCreated | wdPropertyTimeLastSaved | wdPropertyVBATotalEdit | wdPropertyPages | wdPropertyWords | wdPropertyCharacters | wdPropertySecurity | wdPropertyCategory | 
wdPropertyFormat | wdPropertyManager | wdPropertyCompany | wdPropertyBytes | wdPropertyLines | wdPropertyParas | wdPropertySlides | wdPropertyNotes | wdPropertyHiddenSlides | wdPropertyMMClips | wdPropertyHyperlinkBase | 
wdPropertyCharsWSpaces} [-text] <String> [-WordDocument <Document>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>Set-WordBuiltInProperty -WdBuiltInProperty Value -text Value -WordDocument Value
```

Describe what this call does


## PARAMETERS

### WdBuiltInProperty

Describe parameter -WdBuiltInProperty.

```
Type WdBuiltInProperty
Parameter Sets: 
Aliases: 
Required: true
Position: 1
Default Value:
Accept pipeline input: false
```
### text

Describe parameter -text.

```
Type String
Parameter Sets: 
Aliases: 
Required: true
Position: 2
Default Value:
Accept pipeline input: false
```
### WordDocument

Describe parameter -WordDocument.

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


https://shanehoey.github.io/worddoc/docs/set-wordbuiltinproperty
