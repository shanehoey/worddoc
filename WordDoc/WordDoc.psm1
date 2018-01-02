#requires -version 3.0
<#
    .SYNOPSIS
    WordDoc helps you quickly generate Word Documents from PowerShell quickly and effortless.

    .DESCRIPTION
    WordDoc helps you quickly generate Word Documents from PowerShell quickly and effortless.
      
    .EXAMPLE
    import-module -name WordDoc
    Imports the WordDoc module into the current Powershell Instance

    .LINK
    Project Site
    https://shanehoey.com/worddoc

    .LINK
    Credits
    https://shanehoey.github.io/worddoc/credits/

    .LINK
    License
    https://shanehoey.github.io/worddoc/license/
#>

function new-WordInstance {
  <#
    .SYNOPSIS
    The new-wordinstance function starts a new instance of MS Word.

    .DESCRIPTION
    The new-wordinstance function starts a new instance of MS Word.

    .PARAMETER WordInstanceObject
    Returns an Word Instance Object.

    .PARAMETER Visable
    Makes the Word object Visable or Hidden

    .EXAMPLE
    new-WordInstance -Visable True
    
    .EXAMPLE
    new-WordInstance -Visable False
    
    .EXAMPLE
    $wi = new-wordinstance -wordinstanceobject
    $wd = new-worddoc      -wordinstance $wi  -worddocobject

    .LINK
    https://shanehoey.github.io/worddoc/docs/new-wordinstance

  #>


    [CmdletBinding()]
    Param( 
        [Parameter(Mandatory = $false, HelpMessage = 'Returns object Rather than creating a Script Scope variable')]
        [switch]$WordInstanceObject,

        [Parameter(Mandatory = $false, HelpMessage = 'Display Word or keep it hidden')]
        [bool]$Visable = $true
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { Add-Type -AssemblyName Microsoft.Office.Interop.Word }
        catch { Write-Warning  -Message "$($MyInvocation.InvocationName) - Unable to add Word Assembly, Word must be installed for this module... exiting" ; break }
        if (!($WordInstanceObject)) { 
            try { if (test-path -Path variable:script:WordInstance) {throw 'WordInstance already exists'} }
            catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" ; break }
        }
    }
    Process { 
        try { $WordInstance = new-Object -ComObject Word.Application -Property @{Visible = $Visable}  
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - Unable to Invoke Word... exiting" ; break }
        try { if ($WordInstanceObject) { return $WordInstance } else { new-Variable -Name 'WordInstance' -Value $WordInstance -scope script -ErrorAction SilentlyContinue } }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - Unable to create variable... exiting" ; break }   
    }
    End { 
        Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" 
    }
}

function test-WordInstance {
  <#
    .SYNOPSIS
    Describe purpose of "test-WordInstance" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .EXAMPLE
    test-WordInstance -WordInstance $wi
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/test-wordinstance
    
  #>

 
    [CmdletBinding()]
    Param(  
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = 'Word Instance Object')] 
        $WordInstance = $Script:WordInstance
    )
    Begin {
        Add-Type -AssemblyName Microsoft.Office.Interop.Word
Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
    }
    Process { 
        if ($WordInstance -is [Microsoft.Office.Interop.Word.Application]) {
            write-verbose -Message 'Type is [Microsoft.Office.Interop.Word.Application]'
        }
        else { 
            write-verbose -Message 'Type is NOT [Microsoft.Office.Interop.Word.Application]'
            throw 'WordInstance was not of type [Microsoft.Office.Interop.Word.Application]' 
        }
    }
    End { 
        Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" 
    }
}

function test-WordDoc {
  <#
    .SYNOPSIS
    Describe purpose of "test-WordDoc" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    test-WordDoc -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/test-worddoc

  #>


    [CmdletBinding()]
    Param(  
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = 'Word Document Object')] 
        $WordDoc = $Script:WordDoc
    )
    Begin { 
        Add-Type -AssemblyName Microsoft.Office.Interop.Word
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
    }
    Process {
        if ($WordDoc -is [Microsoft.Office.Interop.Word.Document]) {
            write-verbose -Message 'Type is [Microsoft.Office.Interop.Word.Document]'
        }
        else { 
            write-verbose -Message 'Type is NOT [Microsoft.Office.Interop.Word.Document]'
            throw 'WordDoc was not of type [Microsoft.Office.Interop.Word.Document]' 
        }
    }
    End { 
        Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" 
    }
}

function get-WordInstance {
  <#
    .SYNOPSIS
    Describe purpose of "get-WordInstance" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .EXAMPLE
    get-WordInstance -WordInstance Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/get-wordinstance

  #>

 
    [CmdletBinding()]
    Param(  
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = 'Word Instance Object')] 
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { test-wordinstance -wordinstance $wordInstance }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { return $wordInstance }
    End { Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" }
}

function get-WordDoc {
  <#
    .SYNOPSIS
    Describe purpose of "get-WordDoc" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    get-WordDoc -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/get-worddoc

  #>


    [CmdletBinding()]
    Param(  
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = 'Word Document Object')] 
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:WordDoc
    )
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { return $worddoc}
    End { Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" }
}

function new-WordDocument {
  <#
    .SYNOPSIS
    Describe purpose of "new-WordDocument" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .PARAMETER WordDocObject
    Describe parameter -WordDocObject.

    .EXAMPLE
    new-WordDocument -WordInstance Value -WordDocObject
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/new-worddocument

  #>


    [CmdletBinding()]
    Param(  
        [Parameter(Mandatory = $false, Position = 0, HelpMessage = 'Word Instance Object')] 
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
 
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [switch]$WordDocObject
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try { test-wordinstance -WordInstance $wordinstance }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            $WordDoc = $WordInstance.Documents.Add()
            $WordDoc.Activate()
            if ($WordDocObject) {
                return $WordDoc  
            }
            else {
                new-Variable -Name WordDoc -Value $WordDoc -Scope script -ErrorAction SilentlyContinue
            }  
        }
        catch {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
        }
    }
    End { 
        Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" 
    }
}

function Save-WordDocument {
  <#
    .SYNOPSIS
    Describe purpose of "Save-WordDocument" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .PARAMETER WordSaveFormat
    Describe parameter -WordSaveFormat.

    .PARAMETER filename
    Describe parameter -filename.

    .PARAMETER folder
    Describe parameter -folder.

    .EXAMPLE
    Save-WordDocument -WordDoc Value -WordSaveFormat Value -filename Value -folder Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/save-worddocument

  #>


    [CmdletBinding()]
    Param( 
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:worddoc,

        [Parameter(Mandatory = $false, HelpMessage = 'Configure the save format', Position = 0)]
        [Microsoft.Office.Interop.Word.WdSaveFormat]$WordSaveFormat = 'wdFormatDocumentDefault',
     
        [Parameter(Mandatory = $false, HelpMessage = 'Name of Filename', Position = 1)]
        [string]$filename = 'document.docx',
    
        [Parameter(Mandatory = $false, HelpMessage = 'Name of Folder Path, defaults to MyDocuments')]
        [String]$folder = [Environment]::GetFolderPath('MyDocuments')
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            if (!($PSBoundParameters.ContainsKey('filename'))) { 
              Add-Type -AssemblyName System.windows.forms
              $SaveFileDialog = new-Object -TypeName System.Windows.Forms.saveFileDialog
              $SaveFileDialog.initialDirectory =  $folder
              $SaveFileDialog.filter = 'WordDocuments (*.docx)| *.docx | All Documents (*.*)| *.* '
               $null = $SaveFileDialog.ShowDialog() 
              $WordDoc.SaveAs([ref]($SaveFileDialog.filename) , $WordSaveFormat)
            }
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)"  }
}

function Close-WordDocument {
  <#
    .SYNOPSIS
    Describe purpose of "Close-WordDocument" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Close-WordDocument -WordInstance Value -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/close-worddocument

  #>


    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, HelpMessage = 'Word Instance Object')]
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
  
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:WordDoc
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"  
        try { 
            test-wordinstance -WordInstance $wordinstance
            test-worddoc -Worddoc $worddoc
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process {     
        try {
            $WordDoc.Close() 
            #$WordInstance.Quit()  
        }
        catch {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
        }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}

function Close-WordInstance {
  <#
    .SYNOPSIS
    Describe purpose of "Close-WordInstance" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .EXAMPLE
    Close-WordInstance -WordInstance Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/close-wordinstance

  #>


    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        #Todo cast type instead ie [Microsoft.Office.Interop.Word.Application]$WordInstance but does not work
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance
  
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { 
            test-wordinstance -WordInstance $wordinstance
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process {     
        try {
            $WordInstance.Quit()  
        }
        catch {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
        }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}

function Add-WordText {
  <#
    .SYNOPSIS
    Describe purpose of "Add-WordText" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER text
    Describe parameter -text.

    .PARAMETER WdColor
    Describe parameter -WdColor.

    .PARAMETER WDBuiltinStyle
    Describe parameter -WDBuiltinStyle.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Add-WordText -text Value -WdColor Value -WDBuiltinStyle Value -WordDoc Value
    Describe what this call does


    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordtext

  #>


    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, HelpMessage = 'Text to add to word document.' )] 
        [String]$text,
    
        [Parameter(Mandatory = $false, HelpMessage = 'Color of text.' )] 
        [Microsoft.Office.Interop.Word.WdColor]$WdColor = 'wdColorAutomatic',
    
        [Parameter(Mandatory = $false, HelpMessage = 'Built in style of text.' )] 
        [Microsoft.Office.Interop.Word.WdBuiltinStyle]$WDBuiltinStyle = 'wdStyleDefaultParagraphFont',
    
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $script:WordDoc
    )
    Begin {
        Add-Type -AssemblyName Microsoft.Office.Interop.Word
Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            if ($PSBoundParameters.ContainsKey('WDBuiltinStyle')) { Write-verbose -Message "$WDBuiltinStyle"; $WordDoc.application.selection.Style = $WDBuiltinStyle }
            if ($PSBoundParameters.ContainsKey('WdColor')) { Write-verbose -Message "$wdcolor"; $WordDoc.Application.Selection.font.Color = $WdColor.value__ }
            $WordDoc.Application.Selection.TypeText("$($text)")    
            $WordDoc.Application.Selection.TypeParagraph() 
            $WordDoc.application.selection.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]'wdStyleNormal'
        }
        catch {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
        }
    }
    End { 
        Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" 
    }
}

function Add-WordBreak {
  <#
    .SYNOPSIS
    Describe purpose of "Add-WordBreak" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER breaktype
    Describe parameter -breaktype.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Add-WordBreak -breaktype Value -WordInstance Value -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordbreak


  #>


    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)] 
        [Parameter(ParameterSetName = 'GridTable')]
        [ValidateSet('NewPage', 'Section', 'Paragraph')]
        [string]$breaktype = 'NewPage',
   
        #Todo cast type instead ie [Microsoft.Office.Interop.Word.Application]$WordInstance but does not work
        [Parameter(Mandatory = $false, HelpMessage = 'Word Instance Object')]
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
    
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:WordDoc
    )
 
    Begin {
        Add-Type -AssemblyName Microsoft.Office.Interop.Word
Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try { test-wordinstance -WordInstance $wordinstance }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {  
            switch ($breaktype) { 
                'NewPage' { $WordInstance.Selection.InsertNewPage() }
                'Section' { $WordInstance.Selection.Sections.Add() }
                'Paragraph' { $WordInstance.Selection.InsertParagraph() } 
            }
            [Void]$WordDoc.application.selection.goto([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc')
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { 
        Write-Verbose -Message "End    : $($Myinvocation.InvocationName)"
    }
}

function Set-WordBuiltInProperty {
  <#
    .SYNOPSIS
    Describe purpose of "Set-WordBuiltInProperty" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WdBuiltInProperty
    Describe parameter -WdBuiltInProperty.

    .PARAMETER text
    Describe parameter -text.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Set-WordBuiltInProperty -WdBuiltInProperty Value -text Value -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/set-wordbuiltinproperty

  #>


    [CmdletBinding()]
    param(
    
        [Parameter(Position = 0, HelpMessage = 'Add help message for user', Mandatory = $true)] 
        [Microsoft.Office.Interop.Word.WdBuiltInProperty]$WdBuiltInProperty,
    
        [Parameter(Position = 1, HelpMessage = 'Add help message for user', mandatory = $true)] 
        [String]$text,
    
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:WordDoc
    )
    Begin { 
        Write-Verbose -Message "Start  : $($Myinvocation.InvocationName)" 
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try { 
            Write-Verbose -Message $WdBuiltInProperty
            $WordDoc.BuiltInDocumentProperties.item($WdBuiltInProperty).value = $text
        }
        catch {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
        }
    }
    End { 
        Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" 
    }
}

function Add-WordCoverPage {
  <#
    .SYNOPSIS
    Describe purpose of "Add-WordCoverPage" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER CoverPage
    Describe parameter -CoverPage.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Add-WordCoverPage -CoverPage Value -WordInstance Value -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordcoverpage

  #>


    [CmdletBinding()]
    param(
        #Todo cast type instead ie [Microsoft.Office.Interop.Word]
        [Parameter(Position = 0)] 
        [ValidateSet('Austin', 'Banded', 'Facet', 'Filigree', 'Grid', 'Integral', 'Ion (Dark)', 'Ion (Light)', 'Motion', 'Retrospect', 'Semaphore', 'Sideline', 'Slice (Dark)', 'Slice (Light)', 'Viewmaster', 'Whisp')]  
        [string]$CoverPage = 'Facet',
  
    
        [Parameter(Mandatory = $false, HelpMessage = 'Word Instance Object')]
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
    
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:WordDoc
    )  
    Begin { 
        Add-Type -AssemblyName Microsoft.Office.Interop.Word
Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try { test-wordinstance -WordInstance $wordinstance }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            $Selection = $WordDoc.application.selectio
            $WordInstance.Templates.LoadBuildingBlocks()
            $bb = $WordInstance.templates | Where-Object -Property name -EQ -Value 'Built-In Building Blocks.dotx'
            $part = $bb.BuildingBlockEntries.item($CoverPage)
            $null = $part.Insert($WordInstance.Selection.range, $true) 
            [Void]$Selection.goto([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc')
        }
        catch {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
        }
    }
    End { 
        Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" 
    }
}

function Set-WordOrientation {
  <#
    .SYNOPSIS
    Describe purpose of "Set-WordOrientation" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER Orientation
    Describe parameter -Orientation.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .EXAMPLE
    Set-WordOrientation -Orientation Value -WordInstance Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/set-wordorientation

  #>


    [CmdletBinding()]
    param(
        [Parameter(Position = 0, HelpMessage = 'Orientation of page', Mandatory = $true)] 
        [ValidateSet('Portrait', 'Landscape')]  
        [string]$Orientation,
  
        [Parameter(Mandatory = $false, HelpMessage = 'Word Instance Object')]
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance
            
    )
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try { test-wordinstance -WordInstance $wordinstance }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            switch ($Orientation) {
                'Portrait' { $WordInstance.Selection.PageSetup.Orientation = 0 }
                'Landscape' { $WordInstance.Selection.PageSetup.Orientation = 1 }    
            }
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}

function Add-WordTOC {
  <#
    .SYNOPSIS
    Describe purpose of "Add-WordTOC" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Add-WordTOC -WordInstance Value -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordtoc

  #>


    [CmdletBinding()]  
    param (
        #Todo cast type instead ie [Microsoft.Office.Interop.Word.Application]$WordInstance but does not work
        [Parameter(Mandatory = $false, HelpMessage = 'Word Instance Object')]
        [ValidateScript( {test-wordinstance -WordInstance $_})]
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
  
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [ValidateScript( {test-worddoc -Worddoc $_})]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:WordDoc
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try { test-wordinstance -WordInstance $wordinstance }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process {  
        try {
            $toc = $WordDoc.TablesOfContents.Add($WordInstance.selection.Range)
            $toc.TabLeader = 0
            $toc.HeadingStyles 
            $WordDoc.Application.Selection.TypeParagraph()
        }
        catch {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
        }
    }
    End {
        Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" 
    }
}

function Update-WordTOC {
  <#
    .SYNOPSIS
    Describe purpose of "Update-WordTOC" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Update-WordTOC -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/update-wordtoc

  #>


    [CmdletBinding()]   
    param (
        [Parameter(Mandatory = $false)]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:WordDoc
    )
    Begin {
        Write-Verbose -Message "Start  : $($Myinvocation.InvocationName)" 
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try { $WordDoc.Fields | ForEach-Object -Process { $_.Update() } }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}

function Add-WordTable {
  <#
    .SYNOPSIS
    Describe purpose of "Add-WordTable" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER Object
    Describe parameter -Object.

    .PARAMETER WdAutoFitBehavior
    Describe parameter -WdAutoFitBehavior.

    .PARAMETER WdDefaultTableBehavior
    Describe parameter -WdDefaultTableBehavior.

    .PARAMETER HeaderRow
    Describe parameter -HeaderRow.

    .PARAMETER TotalRow
    Describe parameter -TotalRow.

    .PARAMETER BandedRow
    Describe parameter -BandedRow.

    .PARAMETER FirstColumn
    Describe parameter -FirstColumn.

    .PARAMETER LastColumn
    Describe parameter -LastColumn.

    .PARAMETER BandedColumn
    Describe parameter -BandedColumn.

    .PARAMETER WDTableFormat
    Describe parameter -WDTableFormat.

    .PARAMETER PlainTable
    Describe parameter -PlainTable.

    .PARAMETER GridTable
    Describe parameter -GridTable.

    .PARAMETER ListTable
    Describe parameter -ListTable.

    .PARAMETER ListAccent
    Describe parameter -ListAccent.

    .PARAMETER GridAccent
    Describe parameter -GridAccent.

    .PARAMETER RemoveProperties
    Describe parameter -RemoveProperties.

    .PARAMETER VerticleTable
    Describe parameter -VerticleTable.

    .PARAMETER NoParagraph
    Describe parameter -NoParagraph.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Add-WordTable -Object Value -WdAutoFitBehavior Value -WdDefaultTableBehavior Value -HeaderRow Value -TotalRow Value -BandedRow Value -FirstColumn Value -LastColumn Value -BandedColumn Value -RemoveProperties -VerticleTable -NoParagraph -WordDoc Value
    Describe what this call does

    .EXAMPLE
    Add-WordTable -WDTableFormat Value
    Describe what this call does

    .EXAMPLE
    Add-WordTable -PlainTable Value
    Describe what this call does

    .EXAMPLE
    Add-WordTable -GridTable Value -GridAccent Value
    Describe what this call does

    .EXAMPLE
    Add-WordTable -ListTable Value -ListAccent Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordtable

  #>


    [CmdletBinding()]
    param(
        [Parameter(Position = 0, HelpMessage = 'psobject to send to word', Mandatory = $true, ValuefromPipeline = $true)]    
        [psobject]$Object,
  
        [Parameter(HelpMessage = 'Add help message for user')] 
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]$WdAutoFitBehavior = 'wdAutoFitContent',

        [Parameter(HelpMessage = 'Add help message for user')] 
        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]$WdDefaultTableBehavior = 'wdWord9TableBehavior', 

        [Parameter(HelpMessage = 'Add help message for user')]
        [bool]$HeaderRow = $true,
    
        [Parameter(HelpMessage = 'Add help message for user')]
        [bool]$TotalRow = $false,
    
        [Parameter(HelpMessage = 'Add help message for user')]
        [bool]$BandedRow = $true,
    
        [Parameter(HelpMessage = 'Add help message for user')]
        [bool]$FirstColumn = $false,
    
        [Parameter(HelpMessage = 'Add help message for user')]
        [bool]$LastColumn = $false,
    
        [Parameter(HelpMessage = 'Add help message for user')]
        [bool]$BandedColumn = $false,

        [Parameter(Mandatory = $false, ParameterSetName = 'WDTableFormat', HelpMessage = 'Add help message for user')]
        [Microsoft.Office.Interop.Word.WdTableFormat]$WDTableFormat = 'wdTableFormatNone',
    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        [Parameter(Mandatory = $false, ParameterSetName = 'PlainTable', HelpMessage = 'Add help message for user')]
        [validateSet('Table Grid', 'Table Grid Light', 'Plain Table 1', 'Plain Table 2', 'Plain Table 3', 'Plain Table 4', 'Plain Table 5')]
        [String]$PlainTable = 'Table Grid',

    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        #Todo:  Investigate $table.ApplyStyleDirectFormatting("Grid Table 5 Dark")
        [Parameter( Mandatory = $false, ParameterSetName = 'GridTable')]
        [ValidateSet('Grid Table 1 Light', 'Grid Table 2', 'Grid Table 3', 'Grid Table 4', 'Grid Table 5 Dark', 'Grid Table 6 Colorful', 'Grid Table 7 Colorful')]
        [String]$GridTable = 'Grid Table 1 Light',
    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        [Parameter( Mandatory = $false, ParameterSetName = 'ListTable')]
        [ValidateSet('List Table 1 Light', 'List Table 2', 'List Table 3', 'List Table 4', 'List Table 5 Dark', 'List Table 6 Colorful', 'List Table 7 Colorful')]
        [String]$ListTable = 'List Table 1 Light',
    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        [Parameter( Mandatory = $false, ParameterSetName = 'ListTable')]
        [ValidateSet('Accent 1', 'Accent 2', 'Accent 3', 'Accent 4', 'Accent 5', 'Accent 6')]
        [String]$ListAccent = 'Accent 1',
    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        [Parameter( Mandatory = $false, ParameterSetName = 'GridTable')]
        [ValidateSet('Accent 1', 'Accent 2', 'Accent 3', 'Accent 4', 'Accent 5', 'Accent 6')]
        [string]$GridAccent = 'Accent 1',
    
        [Parameter( Mandatory = $false)]
        [switch]$RemoveProperties,
    
        [Parameter( Mandatory = $false, HelpMessage = 'Add help message for user')]
        [switch]$VerticleTable,
    
        [Parameter( Mandatory = $false, HelpMessage = 'Add help message for user')]
        [switch]$NoParagraph,
    
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object')]
        [Microsoft.Office.Interop.Word.Document][Object]$WordDoc = $Script:WordDoc
    )
   
    Begin {
        Add-Type -AssemblyName Microsoft.Office.Interop.Word
try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            $TableRange = $WordDoc.application.selection.range  
            if (!($VerticleTable)) {
                $Columns = @($Object | Get-Member -MemberType Property, NoteProperty).count
                if ($RemoveProperties) { $Rows = @($Object).count } 
                else {$Rows = @($Object).count + 1 }
            }
            if ($VerticleTable) {
                if ($RemoveProperties) { $Columns = @($Object).count } 
                else {$Columns = @($Object).count + 1 }
                $Rows = @($Object | Get-Member -MemberType Property, NoteProperty).count
            }
            $Table = $WordDoc.Tables.Add($TableRange, $Rows, $Columns, $WdDefaultTableBehavior, $WdAutoFitBehavior) 
            if ($PSBoundParameters.ContainsKey('WDTableFormat')) { $Table.autoformat([Microsoft.Office.Interop.Word.WdTableFormat]::$WDTableFormat) }  
            if ($PSBoundParameters.ContainsKey('PlainTable')) { $Table.style = $PlainTable } 
            if ($PSBoundParameters.ContainsKey('GridTable')) { 
                if ($PSBoundParameters.ContainsKey('GridAccent')) {
                    $Table.style = ($GridTable + ' - ' + $GridAccent) 
                }
                else { $Table.style = $GridTable } 
            } 
            if ($PSBoundParameters.ContainsKey('ListTable')) {
                if ($PSBoundParameters.ContainsKey('ListAccent')) { $Table.style = ($ListTable + ' - ' + $ListAccent) }
                else { $Table.style = $ListTable } 
            }  
            if ($PSBoundParameters.ContainsKey('HeaderRow')) {
                if ($HeaderRow) { $Table.ApplyStyleHeadingRows = $true }
                else { $Table.ApplyStyleHeadingRows = $false } 
            }
            if ($PSBoundParameters.ContainsKey('TotalRow')) {
                if ($TotalRow) { $Table.ApplyStyleLastRow = $true }
                else { $Table.ApplyStyleLastRow = $false } 
            }
            if ($PSBoundParameters.ContainsKey('BandedRow')) {
                if ($BandedRow) { $Table.ApplyStyleRowBands = $true }
                else { $Table.ApplyStyleRowBands = $false} 
            }
            if ($PSBoundParameters.ContainsKey('FirstColumn')) {
                if ($FirstColumn) { $Table.ApplyStyleFirstColumn = $true }
                else { $Table.ApplyStyleFirstColumn = $false } 
            }
            if ($PSBoundParameters.ContainsKey('LastColumn')) {
                if ($LastColumn) { $Table.ApplyStyleLastColumn = $true }
                else { $Table.ApplyStyleLastColumn = $false } 
            }
            if ($PSBoundParameters.ContainsKey('BandedColumn')) {
                if ($BandedColumn) { $Table.ApplyStyleColumnBands = $true }
                else { $Table.ApplyStyleColumnBands = $false } 
            }
            [int]$Row = 1
            [int]$Col = 1
            $PropertyNames = @()
            if ($Object -is [Array]) {[ARRAY]$HeaderNames = $Object[0].psobject.properties | ForEach-Object -Process { $_.Name }} 
            else { [ARRAY]$HeaderNames = $Object.psobject.properties | ForEach-Object -Process { $_.Name } }
   
            if ($RemoveProperties) { $Table.ApplyStyleHeadingRows = $false }
     
            if (!($VerticleTable)) {
                for ($i = 0; $i -le $Columns - 1; $i++) {
                    $PropertyNames += $HeaderNames[$i]
                    if (!$RemoveProperties) {
                        $Table.Cell($Row, $Col).Range.Text = $HeaderNames[$i]
                    }
                    $Col++
                }
                if (!$RemoveProperties)
                { $Row = 2 }
   
                $Object | 
                    ForEach-Object -Process {
                    $Col = 1
                    for ($i = 0; $i -le $Columns - 1; $i++) {      
                        $Table.Cell($Row, $Col).Range.Text = (($_."$($PropertyNames[$i])") -as [System.string])
                        $Col++
                    }    
                    $Row++
                }
            } 
            if ($VerticleTable) {
                for ($i = 0; $i -le $Rows - 1; $i++) {
                    $PropertyNames += $HeaderNames[$i]
                    if (!$RemoveProperties) {
                        $Table.Cell($Row, $Col).Range.Text = $HeaderNames[$i]
                    }
                    $Row++
                }    
                if (!$RemoveProperties) { 
                    $Col = 2 
                }
                $Object | 
                    ForEach-Object -Process {
                    $Row = 1
                    for ($i = 0; $i -le $Rows - 1; $i++) {      
                        $Table.Cell($Row, $Col).Range.Text = (($_."$($PropertyNames[$i])") -as [System.string])
                        $Row++
                    }    
                    $Col++
                }
            }
            $Selection = $WordDoc.application.selection
            [Void]$Selection.goto([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc')
            if (!($NoParagraph)) { $WordDoc.Application.Selection.TypeParagraph() }
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}

function Get-WordBuiltinStyle {
  <#
    .SYNOPSIS
    Describe purpose of "Get-WordBuiltinStyle" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .EXAMPLE
    Get-WordBuiltinStyle
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/get-wordbuiltinstyle

  #>


    [CmdletBinding()]
    param()
    Begin { Add-Type -AssemblyName Microsoft.Office.Interop.Word
Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" }
    Process { 
        try { [Enum]::GetNames([Microsoft.Office.Interop.Word.WdBuiltinStyle]) | ForEach-Object -Process {[pscustomobject]@{ Style = $_ } } 
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}

function Get-WordWdTableFormat {
  <#
    .SYNOPSIS
    Describe purpose of "Get-WordWdTableFormat" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .EXAMPLE
    Get-WordWdTableFormat
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/get-wordwdtableformat

  #>


    [CmdletBinding()]
    param()

    Begin { Add-Type -AssemblyName Microsoft.Office.Interop.Word
Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" }
    Process { 
        try { [Enum]::GetNames([Microsoft.Office.Interop.Word.WdTableFormat]) | ForEach-Object -Process {[pscustomobject]@{ Style = $_ } } 
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}

function Add-WordTemplate {
  <#
    .SYNOPSIS
    Describe purpose of "Add-WordTemplate" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER filename
    Describe parameter -filename.

    .PARAMETER WordDoc
    Describe parameter -WordDoc.

    .EXAMPLE
    Add-WordTemplate -filename Value -WordDoc Value
    Describe what this call does

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordtemplate

  #>


    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Add word document or template to import', Position = 0, ParameterSetName = 'Default')]
        [ValidateScript( { test-Path -Path $_ })] 
        [string]$filename,

        [Parameter(Mandatory = $false, HelpMessage = 'Word Document Object', ParameterSetName = 'Default')]
        [Microsoft.Office.Interop.Word.Document]$WordDoc = $Script:WordDoc  
    )   
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { test-worddoc -Worddoc $worddoc }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    
        if (!($PSBoundParameters.ContainsKey('filename'))) { 
            Add-Type -AssemblyName System.windows.forms 
            $OpenFileDialog = new-Object -TypeName System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory =  [Environment]::GetFolderPath('Desktop')
            $OpenFileDialog.filter = 'WordDocuments (*.docx)| *.docx | *.dotx'
            $null = $OpenFileDialog.ShowDialog()
            $filename = $OpenFileDialog.filename
        }

    }
    Process { 
        try { $WordDoc.Application.Selection.InsertFile([ref]($filename)) }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}