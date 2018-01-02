---
title: "Set-WordBuiltInProperty"
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
wdPropertyCharsWSpaces} [-text] <String> [-WordDoc <Document>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
Set-WordBuiltInProperty -WdBuiltInProperty Value -text Value -WordDoc Value
```
PS C:\>
Describe what this call does

## PARAMETERS

### -WdBuiltInProperty

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
### -text

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
### -WordDoc

Describe parameter -WordDoc.

```
Type Document
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:$Script:WordDoc
Accept pipeline input: false
```
### CommonParameters

This function only supports -verbose

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
# Set-WordBuiltInProperty
