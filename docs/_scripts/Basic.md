---
title: "Basic Example"
excerpt: "Basic Examples"
category: "general"
---

## Example Script

Once you have installed the module you can now Create Word Document's from powershell

```powershell
Import-Module Worddoc 
$a = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit.' 
$word = New-WordInstance 
$worddoc = New-WordDocument -Word $word
Add-WordCoverPage -CoverPage Banded -Word $word -WordDoc $worddoc 
Add-WordText -text 'Table of Contents' -WDBuiltinStyle wdStyleTitle -WordDoc $worddoc 
Add-WordTOC -word $word -WordDoc $worddoc 
Add-WordBreak -breaktype NewPage -word $word -WordDoc $worddoc 
Add-WordText -text 'Heading1' -WDBuiltinStyle wdStyleHeading1 -WordDoc $worddoc
Add-WordText -text $text -WDBuiltinStyle wdStyleNormal -WordDoc $worddoc #WordDoc
```

**ProTip:** Be sure to check out the other [example scripts](/worddoc/scripts/) 
{: .notice--info}
