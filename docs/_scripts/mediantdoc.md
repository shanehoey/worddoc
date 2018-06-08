---
title: "MediantDoc"
excerpt: "MediantDoc will generate a as-built document for an AudioCodes Mediant Device, either via the ini file or irectly from the device"
category: "general"
---

## CCE Design Doc Script

Quickly and effortless create a Skype for Business Cloud Connector Eddition (CCE) Design Document or As Built Document using cloudconnector.ini and Powershell.

Hightlights include:
 * Generate a full Design or As Built document from the cloudconnectpr.ini file
 * Full List of Servers Created
 * Firewall Requirements
 * Certificate Requirements
 
[Download the Production version on Powershell Gallery](https://powershellgallery.com/mediantdoc/)
[Download the Development version  on GitHUB](https://github.com/shanehoey/mediantdoc/)
[Download the Prequiste WordDoc module on Powershell Gallery](https://powershellgallery.com/worddoc/)
[Download the Prequiste Mediant module on Powershell Gallery](https://powershellgallery.com/mediant/)


## Easy Installation via PowerShell Gallery
```powershell
install-module worddoc -scope currentuser
install-module mediant -scope currentuser
install-script mediantdoc -scope currentuser
```

## Easy Updates via PowerShell Gallery
```powershell
update-module worddoc -scope currentuser
update-module mediant -scope currentuser
update-script mediantdoc -scope currentuser
```

## Example 
 .\mediantdoc.ps1 


**ProTip:** Be sure to check out the other [example scripts](/worddoc/scripts/) 
{: .notice--info}
