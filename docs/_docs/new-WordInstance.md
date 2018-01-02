---
title: "new-wordinstance"
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
PS C:\>new-WordInstance -Visable True
```

Create a new Word Instance that is visable

### -------------------------- EXAMPLE 2 --------------------------


```
PS C:\>new-WordInstance -Visable False
```

Create a new Word Instance that is hidden

### -------------------------- EXAMPLE 3 --------------------------


```
PS C:\>$wi = new-wordinstance -wordinstanceobject
```

Create a word instance that is stored in a local variable


## PARAMETERS

### WordInstanceObject

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
### Visable

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

new-WordInstance
