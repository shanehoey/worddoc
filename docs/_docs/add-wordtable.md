---
title: "add-wordtable"
excerpt: "add-wordtable"
category: "help"
---
## {{ page.title }}
{{ page.excerpt }}

### Syntax

`Add-WordTable [-Object] <psobject> [-WdAutoFitBehavior <WdAutoFitBehavior>] [-WdDefaultTableBehavior <WdDefaultTableBehavior>] [-HeaderRow <bool>] [-TotalRow <bool>] [-BandedRow <bool>] [-FirstColumn <bool>] [-LastColumn <bool>] [-BandedColumn <bool>] [-WDTableFormat <WdTableFormat>] [-RemoveProperties] [-VerticleTable] [-NoParagraph] [-WordDoc <Object>] [<CommonParameters>]`

`Add-WordTable [-Object] <psobject> [-WdAutoFitBehavior <WdAutoFitBehavior>] [-WdDefaultTableBehavior <WdDefaultTableBehavior>] [-HeaderRow <bool>] [-TotalRow <bool>] [-BandedRow <bool>] [-FirstColumn <bool>] [-LastColumn <bool>] [-BandedColumn <bool>] [-PlainTable <string>] [-RemoveProperties] [-VerticleTable] [-NoParagraph] [-WordDoc <Object>] [<CommonParameters>]`

`Add-WordTable [-Object] <psobject> [-WdAutoFitBehavior <WdAutoFitBehavior>] [-WdDefaultTableBehavior <WdDefaultTableBehavior>] [-HeaderRow <bool>] [-TotalRow <bool>] [-BandedRow <bool>] [-FirstColumn <bool>] [-LastColumn <bool>] [-BandedColumn <bool>] [-GridTable <string>] [-GridAccent <string>] [-RemoveProperties] [-VerticleTable] [-NoParagraph] [-WordDoc <Object>] [<CommonParameters>]`

`Add-WordTable [-Object] <psobject> [-WdAutoFitBehavior <WdAutoFitBehavior>] [-WdDefaultTableBehavior <WdDefaultTableBehavior>] [-HeaderRow <bool>] [-TotalRow <bool>] [-BandedRow <bool>] [-FirstColumn <bool>] [-LastColumn <bool>] [-BandedColumn <bool>] [-ListTable <string>] [-ListAccent <string>] [-RemoveProperties] [-VerticleTable] [-NoParagraph] [-WordDoc <Object>] [<CommonParameters>]`

### Parameters

<table class="table table-striped table-bordered table-condensed visible-on">
	<thead>
		<tr>
			<th>Name</th>
			<th class="visible-lg visible-md">Alias</th>
			<th>Description</th>
			<th class="visible-lg visible-md">Required?</th>
			<th class="visible-lg">Pipeline Input</th>
			<th class="visible-lg">Default Value</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td><nobr>BandedColumn</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>BandedRow</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>FirstColumn</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>GridAccent</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td></td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>GridTable</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td></td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>HeaderRow</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>LastColumn</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>ListAccent</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td></td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>ListTable</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td></td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>NoParagraph</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>Object</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>psobject to send to word</td>
			<td class="visible-lg visible-md">true</td>
			<td class="visible-lg">true \(ByValue\)</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>PlainTable</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>RemoveProperties</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td></td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>TotalRow</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>VerticleTable</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>WDTableFormat</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>WdAutoFitBehavior</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>WdDefaultTableBehavior</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
		<tr>
			<td><nobr>WordDoc</nobr></td>
			<td class="visible-lg visible-md">None</td>
			<td>Add help message for user</td>
			<td class="visible-lg visible-md">false</td>
			<td class="visible-lg">false</td>
			<td class="visible-lg"></td>
		</tr>
	</tbody>
</table>			
