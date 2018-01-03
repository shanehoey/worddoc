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

When used the function will return the Word Instance as an Object to be stored in a variable in the local shell. 
If using this method you must use worddocobject as well, and manually parse these objects to all functions.

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

Makes MS Word application Visable or Hidden

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

## RELATED LINKS

new-wordinstance

https://shanehoey.github.io/worddoc/docs/new-wordinstance

