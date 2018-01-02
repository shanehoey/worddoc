---
title: "add-wordtable"
excerpt: "Describe purpose of "Add-WordTable" in 1-2 sentences."
category: "help"
---

# Add-WordTable
**Module** WordDoc

## SYNOPSIS
Describe purpose of "Add-WordTable" in 1-2 sentences.

## DESCRIPTION
Add a more complete description of what the function does.

## SYNTAX

```
Add-WordTable [-Object] <PSObject> [-WdAutoFitBehavior {wdAutoFitFixed | wdAutoFitContent | wdAutoFitWindow}] [-WdDefaultTableBehavior {wdWord8TableBehavior | wdWord9TableBehavior}] [-HeaderRow <Boolean>] [-TotalRow <Boolean>] 
[-BandedRow <Boolean>] [-FirstColumn <Boolean>] [-LastColumn <Boolean>] [-BandedColumn <Boolean>] [-WDTableFormat {wdTableFormatNone | wdTableFormatSimple1 | wdTableFormatSimple2 | wdTableFormatSimple3 | wdTableFormatClassic1 
| wdTableFormatClassic2 | wdTableFormatClassic3 | wdTableFormatClassic4 | wdTableFormatColorful1 | wdTableFormatColorful2 | wdTableFormatColorful3 | wdTableFormatColumns1 | wdTableFormatColumns2 | wdTableFormatColumns3 | 
wdTableFormatColumns4 | wdTableFormatColumns5 | wdTableFormatGrid1 | wdTableFormatGrid2 | wdTableFormatGrid3 | wdTableFormatGrid4 | wdTableFormatGrid5 | wdTableFormatGrid6 | wdTableFormatGrid7 | wdTableFormatGrid8 | 
wdTableFormatList1 | wdTableFormatList2 | wdTableFormatList3 | wdTableFormatList4 | wdTableFormatList5 | wdTableFormatList6 | wdTableFormatList7 | wdTableFormatList8 | wdTableFormat3DEffects1 | wdTableFormat3DEffects2 | 
wdTableFormat3DEffects3 | wdTableFormatContemporary | wdTableFormatElegant | wdTableFormatProfessional | wdTableFormatSubtle1 | wdTableFormatSubtle2 | wdTableFormatWeb1 | wdTableFormatWeb2 | wdTableFormatWeb3}] 
[-RemoveProperties] [-VerticleTable] [-NoParagraph] [-WordDoc <Document>] [<CommonParameters>]

Add-WordTable [-Object] <PSObject> [-WdAutoFitBehavior {wdAutoFitFixed | wdAutoFitContent | wdAutoFitWindow}] [-WdDefaultTableBehavior {wdWord8TableBehavior | wdWord9TableBehavior}] [-HeaderRow <Boolean>] [-TotalRow <Boolean>] 
[-BandedRow <Boolean>] [-FirstColumn <Boolean>] [-LastColumn <Boolean>] [-BandedColumn <Boolean>] [-PlainTable <String>] [-RemoveProperties] [-VerticleTable] [-NoParagraph] [-WordDoc <Document>] [<CommonParameters>]

Add-WordTable [-Object] <PSObject> [-WdAutoFitBehavior {wdAutoFitFixed | wdAutoFitContent | wdAutoFitWindow}] [-WdDefaultTableBehavior {wdWord8TableBehavior | wdWord9TableBehavior}] [-HeaderRow <Boolean>] [-TotalRow <Boolean>] 
[-BandedRow <Boolean>] [-FirstColumn <Boolean>] [-LastColumn <Boolean>] [-BandedColumn <Boolean>] [-GridTable <String>] [-GridAccent <String>] [-RemoveProperties] [-VerticleTable] [-NoParagraph] [-WordDoc <Document>] 
[<CommonParameters>]

Add-WordTable [-Object] <PSObject> [-WdAutoFitBehavior {wdAutoFitFixed | wdAutoFitContent | wdAutoFitWindow}] [-WdDefaultTableBehavior {wdWord8TableBehavior | wdWord9TableBehavior}] [-HeaderRow <Boolean>] [-TotalRow <Boolean>] 
[-BandedRow <Boolean>] [-FirstColumn <Boolean>] [-LastColumn <Boolean>] [-BandedColumn <Boolean>] [-ListTable <String>] [-ListAccent <String>] [-RemoveProperties] [-VerticleTable] [-NoParagraph] [-WordDoc <Document>] 
[<CommonParameters>]
```


## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------


```
PS C:\>Add-WordTable -Object Value -WdAutoFitBehavior Value -WdDefaultTableBehavior Value -HeaderRow Value -TotalRow Value -BandedRow Value -FirstColumn Value -LastColumn Value -BandedColumn Value -RemoveProperties -VerticleTable -NoParagraph -WordDoc Value
```

Describe what this call does

### -------------------------- EXAMPLE 2 --------------------------


```
PS C:\>Add-WordTable -WDTableFormat Value
```

Describe what this call does

### -------------------------- EXAMPLE 3 --------------------------


```
PS C:\>Add-WordTable -PlainTable Value
```

Describe what this call does

### -------------------------- EXAMPLE 4 --------------------------


```
PS C:\>Add-WordTable -GridTable Value -GridAccent Value
```

Describe what this call does

### -------------------------- EXAMPLE 5 --------------------------


```
PS C:\>Add-WordTable -ListTable Value -ListAccent Value
```

Describe what this call does


## PARAMETERS

### Object

Describe parameter -Object.

```
Type PSObject
Parameter Sets: 
Aliases: 
Required: true
Position: 1
Default Value:
Accept pipeline input: true (ByValue)
```
### WdAutoFitBehavior

Describe parameter -WdAutoFitBehavior.

```
Type WdAutoFitBehavior
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:wdAutoFitContent
Accept pipeline input: false
```
### WdDefaultTableBehavior

Describe parameter -WdDefaultTableBehavior.

```
Type WdDefaultTableBehavior
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:wdWord9TableBehavior
Accept pipeline input: false
```
### HeaderRow

Describe parameter -HeaderRow.

```
Type Boolean
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:True
Accept pipeline input: false
```
### TotalRow

Describe parameter -TotalRow.

```
Type Boolean
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### BandedRow

Describe parameter -BandedRow.

```
Type Boolean
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:True
Accept pipeline input: false
```
### FirstColumn

Describe parameter -FirstColumn.

```
Type Boolean
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### LastColumn

Describe parameter -LastColumn.

```
Type Boolean
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### BandedColumn

Describe parameter -BandedColumn.

```
Type Boolean
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### WDTableFormat

Describe parameter -WDTableFormat.

```
Type WdTableFormat
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:wdTableFormatNone
Accept pipeline input: false
```
### PlainTable

Describe parameter -PlainTable.

```
Type String
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:Table Grid
Accept pipeline input: false
```
### GridTable

Describe parameter -GridTable.

```
Type String
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:Grid Table 1 Light
Accept pipeline input: false
```
### ListTable

Describe parameter -ListTable.

```
Type String
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:List Table 1 Light
Accept pipeline input: false
```
### ListAccent

Describe parameter -ListAccent.

```
Type String
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:Accent 1
Accept pipeline input: false
```
### GridAccent

Describe parameter -GridAccent.

```
Type String
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:Accent 1
Accept pipeline input: false
```
### RemoveProperties

Describe parameter -RemoveProperties.

```
Type SwitchParameter
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### VerticleTable

Describe parameter -VerticleTable.

```
Type SwitchParameter
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### NoParagraph

Describe parameter -NoParagraph.

```
Type SwitchParameter
Parameter Sets: 
Aliases: 
Required: false
Position: named
Default Value:False
Accept pipeline input: false
```
### WordDoc

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

Add-WordTable
