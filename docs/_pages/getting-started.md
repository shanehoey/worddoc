---
title: "WordDoc Getting Started"
excerpt: "Super Simple Installation via PowerShell Gallery"
header:
  cta_label: "<i class='fa fa-file'></i> View Documentation"
  cta_url: "/docs"
---

This PowerShell module enables you to **create Word Documents directly from Powershell**. The only prequisites are that you have PowerShell v5 and MS Word installed on your computer. The Installation of the module could not be simpler as the current release is hosted on [PowerShell Gallery](https://www.powershellgallery.com/packages/WordDoc/)

## Latest Release

The latest release of this project is hosted on [PowerShell Gallery](https://www.powershellgallery.com/packages/WordDoc/)

{{ site.btn_poshgal }}

This development version of this project is hosted on [GitHub](https://www.github.com/shanehoey/WordDoc/)

{{ site.btn_github }}
{{ site.btn_github_watch }}
{{ site.btn_github_fork }}
{{ site.btn_github_star }}

## Install latest release from PowerShell Gallery (Powershell v5)

Install latest released version directly from PowerShell Gallary

```powershell
install-module -modulename WordDoc -scope currentuser
```

## Install manually  (Powershell v3-v4)

1. Copy the worddoc.psm1 file to a WordDoc folder into one of the following folders
 * `%userprofile%\Documents\WindowsPowerShell\Modules\WordDoc`
 * `%windir%\System32\WindowsPowerShell\v1.0\Modules\WordDoc`

## Example Script

Once you have installed the module you can now Create Word Document's from powershell

```powershell
Import-Module Worddoc 
$a = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.' 
New-WordInstance 
New-WordDocument
Add-WordCoverPage -CoverPage Banded
Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle
Add-WordTOC
Add-WordBreak -breaktype NewPage
Add-WordText -text 'Heading1' -WDBuiltinStyle wdStyleHeading1
Add-WordText -text $text -WDBuiltinStyle wdStyleNormal
```

**ProTip:** Be sure to check out other [example scripts](/worddoc/scripts/) 
{: .notice--success}
