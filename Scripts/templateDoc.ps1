#Requires -Module WordDoc

<#

.AUTHOR Shane Hoey
.COPYRIGHT 2018 Shane Hoey
.LICENSEURI https://shanehoey.github.io/worddocdoc/license
.PROJECTURI https://shanehoey.github.io/worddoc

MIT License

Copyright (c) 2016-2018 Shane Hoey

Permission is hereby granted, free of charge, to any person obtaining a copy 
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#>

[cmdletbinding(DefaultParameterSetName = "Default")]
Param(  

    [ValidateScript( { (Test-Path $_) -and ((Get-Item $_).Extension -like ".do*x") })]  
    [Parameter(ValueFromPipeline = $false, Mandatory = $false, ParameterSetName = "Default")]
    [string]$WordTemplate,
    
    [Parameter(ValueFromPipeline = $false, Mandatory = $false, ParameterSetName = "Default")]
    [string]$DesignJsonURL,

    [Parameter(ValueFromPipeline = $false, Mandatory = $false, ParameterSetName = "Default")]
    [bool]$NotifyUpdates = $true,

    [Parameter(ValueFromPipeline = $false, Mandatory = $false, ParameterSetName = "Default")]
    [string]$DocumentTitle = "TemplateDoc",

    [Parameter(ValueFromPipeline = $false, Mandatory = $false, ParameterSetName = "Default")]
    [string]$DocumentCustomer = "Shane Hoey",

    [Parameter(ValueFromPipeline = $false, Mandatory = $false, ParameterSetName = "Default")]
    [string]$DocumentAuthor = "Shane Hoey"

)

    try { import-module -name WordDoc -ErrorAction Stop } catch { Write-Warning "WordDoc Module is required , to install ->  install-module -name worddoc -scope currentuser"; break }

    if(!($DesignJsonURL)) { $DesignJsonURL ='https://shanehoey.com/templatedoc.json' }    #Change this 
    $VersionScript = "523129c9-3f83-4efc-839f-973c2c205a28"                         #Change this (new-guid)
    $VersionURL = 'https://shanehoey.com/versions/templateDoc.json'                 #Change this 
    $useragent = 'templateDoc'                                                      #Change this
    
    #Used to quickly enable/disabled specific sections 
    $section = @{}
    $section["CoverPage"] = $true
    $section["Overview"] = $true
    $section["Examples"] = $true

    if ($PSBoundParameters.ContainsKey('WordTemplate')) 
    {
        $TemplateFile = (get-item -path $WordTemplate).fullname
    }    
    else 
    {
        write-host "Load Word Template ?" -foregroundcolor Yellow 
        switch (($host.ui.PromptForChoice("", "Do you want to use an existing word Template ??", [System.Management.Automation.Host.ChoiceDescription[]]((New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"), (New-Object System.Management.Automation.Host.ChoiceDescription "&No")), 1))) 
        {
            0 {  
                Write-warning -Message "Due to a bug the open file dialog box may be behind other windows"
                Add-Type -AssemblyName System.Windows.Forms
                $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $OpenFileDialog.initialDirectory = [Environment]::GetFolderPath('MyDocuments')
                $OpenFileDialog.filter = 'Word Document (*.docx)|*.docx|Word Template (*.dotx)|*.dotx'
                $OpenFileDialog.title = 'Select Word Template to import'
                $result = $OpenFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))
                if ($result -eq [Windows.Forms.DialogResult]::OK) 
                {
                    $TemplateFile = $OpenFileDialog.filename
                }
                else 
                {
                    Write-Verbose "No file selected" -VERBOSE
                } 
                Remove-Variable -Name OpenFileDialog  -ErrorAction SilentlyContinue
                Remove-Variable -name result  -ErrorAction SilentlyContinue
            }
        }
    }

    try 
    {
        Write-verbose -Message "Downloading design document from $designJsonURL" -Verbose
        $DesignText = (invoke-WebRequest -uri $DesignJsonURL -ContentType "text/plain").content -split '\n' | convertfrom-json
        Write-verbose -Message "Downloading $designJsonURL Complete" -Verbose
        $section["DesignText"] = $true
    } 
    catch 
    {
        $section["DesignText"] = $false
        write-warning "Unable to download $designJsonURL design template, defaulting to a blank document"
    }

    #region Version Control & 14 Day usage stats, please do not remove this section.
    # It is only used for version Control and unique users via github 
    # Collecting the stats gives me an indication how often this script is used to determine if I should continue developing it, or concentrate on other projects
    # If you want to silence the notice set notify to $false rather than deleting the section
    # thank in advance
    try 
    {
        $Version = (Invoke-WebRequest -Uri $VersionURL -UserAgent $useragent -Method Get -DisableKeepAlive -TimeoutSec 2 ).content | convertfrom-json 
        if (($thisversion -ne $version.release) -and ($thisversion -ne $version.dev)) 
        {
            if ($notifyupdates) 
            { 
                Write-Warning  "`n**********************`nScript has been Updated`n**********************`nMore details available at $($version.link)"
                start-sleep -Seconds 5 
            }
        }
    }
    catch 
    {
        Write-Warning "Unable to check for updates"
    }
    #endregion

    #region Create Word Document 
        New-WordInstance 
        New-WordDocument
        if ($TemplateFile) {
            Add-WordTemplate -filename $TemplateFile
        }
    #endregion

    #Turn of spelling to speed up creating doc
    #TODO - Need to work out how to turn spelling on/off on a per document basis. 
    (get-wordInstance).options.checkspellingasyoutype = $false

    #region Coverpage
    if($section.CoverPage) { 
        
        #region AddCoverpage
            for ($i = 0; $i -lt 3; $i++) 
            {
                Add-WordBreak -breaktype Paragraph 
            }
            Add-WordText -text $DocumentTitle -WDBuiltinStyle wdStyleTitle
            for ($i = 0; $i -lt 3; $i++) 
            {
                Add-WordBreak -breaktype Paragraph 
            }
            Add-WordText -text 'for' -WDBuiltinStyle wdStyleTitle
            for ($i = 0; $i -lt 3; $i++) 
            {
                Add-WordBreak -breaktype Paragraph 
            }
            Add-WordText -text $DocumentCustomer -WDBuiltinStyle wdStyleTitle
            Add-WordBreak -breaktype NewPage
        #endregion 

        #region License
            $license = "MIT License`nCopyright (c) 2016-2018 Shane Hoey`rPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the 'Software'), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:`nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.`nTHE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."
            Add-WordBreak -breaktype Paragraph
            Add-WordText -text 'This document has been created with wordDoc which has been distributed under the MIT license. For more information visit http://shanehoey.github.io/worddoc/' -WDBuiltinStyle wdStyleBookTitle
            Add-WordBreak -breaktype Paragraph
            #bug with bold/italic in worddoc module
            $selection = (Get-WordDocument).application.selection
            $selection.font.Bold = $False
            $selection.ParagraphFormat.Alignment = 3
            Add-WordText -text $license -WDBuiltinStyle wdStyleNormal
            Add-WordBreak -breaktype NewPage
        #endregion

        #region Shameless Plug
        for ($i = 0; $i -lt 3; $i++) 
        {
        Add-WordBreak -breaktype Paragraph 
        }
        Add-WordText -text 'Are you using this commercially? Show your appreciation and encourage more development of this script at https://paypal.me/shanehoey' -WDBuiltinStyle wdStyleIntenseQuote
        #endregion

        #region TOC
            Add-WordBreak -breaktype NewPage
            Add-WordText -text 'Contents' -WDBuiltinStyle wdStyleTOCHeading 
            Add-WordTOC 
            Add-WordBreak -breaktype NewPage
        #endregion 

        #region update document settings
            Set-WordBuiltInProperty -WdBuiltInProperty wdPropertytitle -text $DocumentTitle
            Set-WordBuiltInProperty -WdBuiltInProperty wdPropertySubject -text "$Documenttitle for $documentCustomer"
            Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyAuthor -text $([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('UwBoAGEAbgBlACAASABvAGUAeQA=')))
            Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyComments -text $([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('aAB0AHQAcABzADoALwAvAHMAaABhAG4AZQBoAG8AZQB5AC4AZwBpAHQAaAB1AGIALgBpAG8ALwB3AG8AcgBkAGQAbwBjAC8A')))
            Set-WordBuiltInProperty -WdBuiltInProperty wdPropertyManager -text $([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String('UwBoAGEAbgBlACAASABvAGUAeQA=')))
        #endregion

    }
    #endregion 
    if ($section.Overview) 
    {
        Add-WordText "Overview" -WDBuiltinStyle wdStyleHeading1 
        if($section.DesignText) { $designtext.textOverview | ForEach-Object  { Add-wordtext -text $_ -WDBuiltinStyle wdStyleNormal } }
        Add-WordBreak -breaktype NewPage
    }

    if ($section.Examples) 
    {
        Add-WordText "Example One" -WDBuiltinStyle wdStyleHeading1 
        if($section.DesignText) 
        {   
            $DesignText.textExample1 | ForEach-Object { Add-wordtext -text $_ -WDBuiltinStyle wdStyleNormal } 
            $DesignText.tableExample1 | foreach-object { Add-WordTable -Object $_  -GridTable 'Grid Table 4' -GridAccent 'Accent 3' -WdAutoFitBehavior wdAutoFitWindow } 
        }
        Add-WordText "Example Two" -WDBuiltinStyle wdStyleHeading1 
        if($section.DesignText) 
        { 
            $designtext.textExample2 | ForEach-Object  { Add-wordtext -text $_ -WDBuiltinStyle wdStyleNormal } 
            $DesignText.tableExample2 | foreach-object { Add-WordTable -Object $_ -GridTable 'Grid Table 4' -GridAccent 'Accent 3' -WdAutoFitBehavior wdAutoFitWindow } 
        }
    }
