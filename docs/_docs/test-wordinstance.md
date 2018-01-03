---
title: "test-wordinstance"
excerpt: "Returns True or False if parsed object is a MS-Word Application."
category: "help"
---

# test-WordInstance
**Module** WordDoc

## SYNOPSIS
Returns True or False if parsed object is a MS-Word Application.

## DESCRIPTION
Returns True or False if parsed object is a MS-Word Application.

## SYNTAX

```
test-WordInstance [[-WordInstance] <Object>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>test-WordInstance -WordInstance $wi
```

Tests is $wi is a MS Word Application object


## PARAMETERS

### WordInstance

Object that you want to check if it is a MS Word Application

```
Type Object
Parameter Sets: 
Aliases: 
Required: false
Position: 1
Default Value:$Script:WordInstance
Accept pipeline input: false
```
### CommonParameters

This function only supports -verbose

## RELATED LINKS


https://shanehoey.github.io/worddoc/docs/test-wordinstance
