---
title: "new-wordinstance"
excerpt: "The New-WordInstance function starts a new instance of MS Word."
category: "help"
---

# New-WordInstance
**Module** WordDoc

## SYNOPSIS
The New-WordInstance function starts a new instance of MS Word.

## DESCRIPTION
The New-WordInstance function starts a new instance of MS Word.

## SYNTAX

```
New-WordInstance [[-WindowState] {wdWindowStateNormal | wdWindowStateMaximize | 
wdWindowStateMinimize}] [[-Visible] <Boolean>] [-ReturnObject] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>New-WordInstance -WindowState wdWindowStateMaximize -Visable True
```

Create a new Word Instance that is maximised and is visable.

### -------------------------- EXAMPLE 2 --------------------------


```
PS C:\>New-WordInstance -Visable False
```

Create a new Word Instance that is hidden.

### -------------------------- EXAMPLE 3 --------------------------


```
PS C:\>$wi = New-WordInstance -ReturnObject
```

Create a word instance that is stored in a local variable.


## PARAMETERS

### WindowState

Set the MS Word application wdWindowStateMaximize, wdWindowStateMinimize, 
wdWindowStateNormal

```
Type WdWindowState
Parameter Sets: 
Aliases: 
Required: false
Position: 1
Default Value:wdWindowStateMaximize
Accept pipeline input: false
```
### Visible

Makes MS Word application Visable or Hidden

```
Type Boolean
Parameter Sets: 
Aliases: 
Required: false
Position: 2
Default Value:True
Accept pipeline input: false
```
### ReturnObject

When used the function will return the Word Instance as an Object to be stored in a 
variable in the local shell. 
If using this method you must use worddocobject as well, and manually parse these 
objects to all functions.

```
Type SwitchParameter
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### CommonParameters

This function only supports -verbose

## RELATED LINKS

New-WordInstance

https://shanehoey.github.io/worddoc/docs/new-wordinstance

