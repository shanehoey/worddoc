> This project is no longer under development,  if you are wanting to create documentation from powershell, try using markdown instead, along with pandoc to convert to word documents.


# WordDoc - Create Word Documents directly from PowerShell

This project's documentation is hosted on Git Pages
https://worddoc.shanehoey.com/

The latest release is hosted on PowerShell Gallery 
https://www.powershellgallery.com/packages/WordDoc/

## Distributed under the MIT License
This project is distrubuted under the MIT License. The license can be viewed [here](https://github.com/shanehoey/worddoc/blob/master/LICENSE)

## Project Notes
This Project contains Powershell sample scripts that can be reused / adapted. Please do not just execute scripts without understanding what each and every line will do.


WordDoc PowerShell Module
Word Doc is a PowerShell Module that helps you create documents directly from powershell. This simple module enables your to quickly and effortlessly create Word Documents directly from Powershell.

**Please Note:** This project is hosted on [GitHib](https://github.com/shanehoey/worddoc), and full documentation is available on [GitPages](https://worddoc.shanehoey.com).

## Highlights include
 * Generate a Word Documents directly from PowerShell
 * Create Tables from PowerShell Objects
 * Update page title, author etc
 * Save as PDF, WordDoc, HTML and more

 ## Installation (windows 10)

 Installation instructions for other versions of windows available [here](https://worddoc.shanehoey.com/getting-started/)

```powershell
install-module -name worddoc -scope currentuser
```

## Example Usage
A simple example to show you how to create a word document, create a cover page, Table of Contents,  add some word text, and add some objects as a word table.

```powershell
Import-Module Worddoc 
New-WordInstance 
New-WordDocument
Add-WordCoverPage -CoverPage Banded 
Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle 
Add-WordTOC
Add-WordBreak -breaktype NewPage 
Add-WordText -text 'Heading1' -WDBuiltinStyle wdStyleHeading1 
$a = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.' 
Add-WordText -text $a -WDBuiltinStyle wdStyleNormal
$s = get-service | select name,status
Add-WordTable -object $s
Save-WordDocument 
Close-WordDocument
Close-WordInstance
```
