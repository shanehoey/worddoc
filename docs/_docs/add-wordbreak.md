---
title: "add-wordbreak"
excerpt: "Create a new Break (newpage,section or paragraph)"
category: "help"
---

# Add-WordBreak
**Module** WordDoc

## SYNOPSIS
Create a new Break (newpage,section or paragraph)

## DESCRIPTION
Create a new Break (newpage,section or paragraph)

## SYNTAX

```
Add-WordBreak [-breaktype <String>] [-WordInstance <Application>] [-WordDocument 
<Document>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>Add-WordBreak -breaktype NewPage
```

Creates a NewPage Break


## PARAMETERS

### breaktype

Type of break (newpage,section or paragraph)

```
Type String
Parameter Sets: 
Aliases: 
Required: false
Position: 1
Default Value:NewPage
Accept pipeline input: false
```
### WordInstance

Describe parameter -WordInstance.

```
Type Application
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:$Script:WordInstance
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


https://shanehoey.github.io/worddoc/docs/add-wordinstance
