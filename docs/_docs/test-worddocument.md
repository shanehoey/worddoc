---
title: "test-worddocument"
excerpt: "Returns True or False if parsed object is a MS-Word Document."
category: "help"
---

# test-WordDocument
**Module** WordDoc

## SYNOPSIS
Returns True or False if parsed object is a MS-Word Document.

## DESCRIPTION
Returns True or False if parsed object is a MS-Word Document.

## SYNTAX

```
test-WordDocument [[-WordDocument] <Object>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>test-WordDocument -WordDocument $wd
```

tests is $wd is a MS Word Document Object


## PARAMETERS

### WordDocument

Object that you want to check if it is a MS Word Document

```
Type Object
Parameter Sets: 
Aliases: 
Required: false
Position: 1
Default Value:$Script:WordDocument
Accept pipeline input: false
```
### CommonParameters

This function only supports -verbose

## RELATED LINKS


https://shanehoey.github.io/worddoc/docs/test-worddoc
