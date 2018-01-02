---
title: "new-WordInstance"
excerpt: "The new-wordinstance function starts a new instance of MS Word."
category: "help"
---

# new-WordInstance
**Module** WordDoc

## SYNOPSIS
The new-wordinstance function starts a new instance of MS Word.

## DESCRIPTION
The new-wordinstance function starts a new instance of MS Word.

## SYNTAX

```
new-WordInstance [-WordInstanceObject] [[-Visable] <Boolean>] [<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
new-WordInstance -Visable True
```
PS C:\>

### -------------------------- EXAMPLE 2 --------------------------


```
new-WordInstance -Visable False
```
PS C:\>

### -------------------------- EXAMPLE 3 --------------------------


```
$wi = new-wordinstance -wordinstanceobject
```
PS C:\>
$wd = new-worddoc      -wordinstance $wi  -worddocobject

## PARAMETERS

### -WordInstanceObject

Returns an Word Instance Object.

```
Type SwitchParameter
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### -Visable

Makes the Word object Visable or Hidden

```
Type Boolean
Parameter Sets: 
Aliases: 
Required: false
Position: 1
Default Value:True
Accept pipeline input: false
```
### CommonParameters

This function only supports -verbose

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
# new-WordInstance
