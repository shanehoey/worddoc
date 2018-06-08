---
title: "cce Design Doc"
excerpt: "cce Design Doc will generate a design document for Skype for Business online Cloud Connector Edition from a cloudconnector.ini fle"
category: "general"
---

## CCE Design Doc Script

Quickly and effortless create a Skype for Business Cloud Connector Eddition (CCE) Design Document or As Built Document using cloudconnector.ini and Powershell.

Hightlights include:
 * Generate a full Design or As Built document from the cloudconnectpr.ini file
 * Full List of Servers Created
 * Firewall Requirements
 * Certificate Requirements
 
[Download the Production version on Powershell Gallery](https://powershellgallery.com/ccedesigndoc/)
[Download the latest on GitHUB](https://github.com/shanehoey/ccedesigndoc/)
[Download the Prequiste WordDoc module on Powershell Gallery](https://powershellgallery.com/worddoc/)

## Easy Installation via PowerShell Gallery
```powershell
install-module worddoc -scope currentuser
install-script ccedesigndoc -scope currentuser
```

## Easy Updates via PowerShell Gallery
```powershell
update-module worddoc -scope currentuser
update-script ccedesigndoc -scope currentuser
```

## Example 
 .\cceDesignDoc.ps1 -filepath .\cloudconnector.ini



**ProTip:** Be sure to check out the other [example scripts](/worddoc/scripts/) 
{: .notice--info}
