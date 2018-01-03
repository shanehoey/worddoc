#requires -version 4.0
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

try { Add-Type -AssemblyName Microsoft.Office.Interop.Word }
catch { Write-Warning  -Message "$($MyInvocation.InvocationName) - Unable to add Word Assembly, Word must be installed for this module to work" }

function new-WordInstance {
  <#
    .SYNOPSIS
    The new-wordinstance function starts a new instance of MS Word.

    .DESCRIPTION
    The new-wordinstance function starts a new instance of MS Word.

    .PARAMETER WordInstanceObject
    When used the function will return the Word Instance as an Object to be stored in a variable in the local shell. 
    If using this method you must use worddocobject as well, and manually parse these objects to all functions. 

    .PARAMETER Visable
    Makes MS Word application Visable or Hidden

    .EXAMPLE
    new-WordInstance -Visable True
    
    Create a new Word Instance that is visable
    
    .EXAMPLE
    new-WordInstance -Visable False

    Create a new Word Instance that is hidden
    
    .EXAMPLE
    $wi = new-wordinstance -wordinstanceobject

    Create a word instance that is stored in a local variable
    
    .INPUTS

    .OUTPUTS
     [Microsoft.Office.Interop.Word.Application]

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK

    new-wordinstance

    https://shanehoey.github.io/worddoc/docs/new-wordinstance

  #>

    [CmdletBinding()]
    Param( 
        [switch]$WordInstanceObject,

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

function get-WordInstance {
    <#
      .SYNOPSIS
      This function is used to return a Word Instance created automatically by Word Doc Module
  
      .DESCRIPTION
      This function is used to return a Word Instance created automatically by Word Doc Module
  
      .PARAMETER WordInstance
      Not required as this function will work without using WordInstance Parameter
  
      .EXAMPLE
      get-WordInstance -WordInstance Value
      Describe what this call does
  
      .NOTES
      for more examples visit https://shanehoey.github.io/worddoc/
  
      .LINK
      https://shanehoey.github.io/worddoc/docs/get-wordinstance
  
    #>
 
    [CmdletBinding()]
    Param(  
        [Parameter(Position = 0,DontShow)] 
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance
    )
    Begin { 
          Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
          try { $null = test-wordinstance -wordinstance $wordInstance }
          catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
      }
      Process { return $wordInstance }
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
  
      .PARAMETER WordDocumentObject
      Describe parameter -WordDocObject.
  
      .EXAMPLE
      new-WordDocument -WordInstance Value -WordDocObject
      Describe what this call does
  
      .NOTES
      for more examples visit https://shanehoey.github.io/worddoc/
  
      .LINK
      https://shanehoey.github.io/worddoc/docs/new-worddocument
  
    #>
    
      [CmdletBinding()]
      Param(  
          [Parameter(Position = 0)] 
          [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
   
          [switch]$WordDocObject
      )
      Begin { 
          Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
          try { $null = test-wordinstance -WordInstance $wordinstance }
          catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
      }
      Process { 
          try {
              $WordDocument = $WordInstance.Documents.Add()
              $WordDocument.Activate()
              try { if ($WordDocObject) { return $WordDocument } else { new-Variable -Name 'WordDocument' -Value $WordDocument -scope script -ErrorAction SilentlyContinue } }
              catch { Write-Warning -Message "$($MyInvocation.InvocationName) - Unable to create variable... exiting" ; break }   
            }
            catch {
                Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
            }

      }
      End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)"  }
  }



  function test-WordInstance {
  <#
    .SYNOPSIS
    Returns True or False if parsed object is a MS-Word Application.

    .DESCRIPTION
    Returns True or False if parsed object is a MS-Word Application.

    .PARAMETER WordInstance
    Object that you want to check if it is a MS Word Application

    .EXAMPLE

    test-WordInstance -WordInstance $wi
    
    Tests is $wi is a MS Word Application object

    .INPUTS

    .OUTPUTS
    [Boolean]

    This function returns a Boolean.

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/test-wordinstance
    
  #>

    [CmdletBinding()]
    Param(  
        [Parameter(Position = 0)] 
        $WordInstance = $Script:WordInstance
    )
    Begin { Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" }
    Process { 
        if ($WordInstance -is [Microsoft.Office.Interop.Word.Application]) {
            return $true
        }
        else { 
            throw "Object is not type [Microsoft.Office.Interop.Word.Document]"
            return $false
        }
    }
    End { Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" }
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
  
      .NOTES
      for more examples visit https://shanehoey.github.io/worddoc/
  
      .LINK
      https://shanehoey.github.io/worddoc/docs/close-wordinstance
  
    #>
  
    [CmdletBinding()]
    param(
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance
      )
      Begin { 
          Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
          try { 
              $null = test-wordinstance -WordInstance $wordinstance
          }
          catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
      }
      Process {     
          try {
              $WordInstance.Quit()  
              remove-variable WordInstance -Scope Script
          }
          catch {
              Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
          }
      }
      End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
  }
  

  function test-WordDocument {
  <#
    .SYNOPSIS
    Returns True or False if parsed object is a MS-Word Document.

    .DESCRIPTION
    Returns True or False if parsed object is a MS-Word Document.

    .PARAMETER WordDocument
    Object that you want to check if it is a MS Word Document

    .EXAMPLE
    test-WordDocument -WordDocument $wd

    tests is $wd is a MS Word Document Object

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/test-worddoc

  #>

    [CmdletBinding()]
    Param(  
        [Parameter(Position = 0)] 
        $WordDocument = $Script:WordDocument
    )
    Begin { Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" }
    Process {
        if ($WordDocument -is [Microsoft.Office.Interop.Word.Document]) {
            return $true
        }
        else { 
            throw "Object is not type [Microsoft.Office.Interop.Word.Document]"
            return $false
        }
    }
    End { Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" }
}

function get-WordDocument {
  <#
    .SYNOPSIS
    Describe purpose of "get-WordDoc" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    get-WordDocument -WordDocument Value
   

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/get-worddoc

  #>


    [CmdletBinding()]
    Param(  
        [Parameter(Position = 0,DontShow)] 
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
    )
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { $null = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { return $WordDocument}
    End { Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" }
}

function Save-WordDocument {
  <#
    .SYNOPSIS
    Describe purpose of "Save-WordDocument" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordDocumentument
    Describe parameter -WordDocument.

    .PARAMETER WordSaveFormat
    Describe parameter -WordSaveFormat.

    .PARAMETER filename
    Describe parameter -filename.

    .PARAMETER folder
    Describe parameter -folder.

    .EXAMPLE
    Save-WordDocument -WordDocument Value -WordSaveFormat Value -filename Value -folder Value
    Describe what this call does

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/save-worddocument

  #>

    [CmdletBinding()]
    Param( 
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument,

        [Parameter(Position = 0)]
        [Microsoft.Office.Interop.Word.WdSaveFormat]$WordSaveFormat = 'wdFormatDocumentDefault',
     
        [Parameter(Position = 1)]
        [string]$filename = 'document.docx',
    
        [String]$folder = [Environment]::GetFolderPath('MyDocuments')
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { test-WordDocument -WordDocument $WordDocument }
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
              $WordDocument.SaveAs([ref]($SaveFileDialog.filename) , $WordSaveFormat)
            }
            else { $WordDocument.SaveAs([ref]((Join-Path -path $folder -ChildPath $filename)) , $WordSaveFormat) }
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

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Close-WordDocument -WordInstance Value -WordDocument Value
    Describe what this call does

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/close-worddocument

  #>


    [CmdletBinding()]
    param(
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
  
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"  
        try { 
            $null = test-wordinstance -WordInstance $wordinstance
            $null = test-WordDocument -WordDocument $worddoc
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process {     
        try {
            $WordDocument.Close() 
            remove-variable WordDocument -Scope Script

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
    Adds text to MS Word Document.

    .DESCRIPTION
    Adds text to MS Word Document.

    .PARAMETER text
    Text to add to word Document

    .PARAMETER WdColor
    Color of Text

    .PARAMETER WDBuiltinStyle
    Builtin Stype to use 

    .PARAMETER WordDocument
    WordDocument Object 

    .EXAMPLE

    Add-WordText -text "Heading 1" -WdColor Value -WDBuiltinStyle Value -WordDocument Value
    
    Adds text to document 

    .EXAMPLE

    Add-WordText -text "Heading 1" -WdColor Value -WDBuiltinStyle Value -WordDocument Value
    
    Adds text to document 

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordtext

  #>

    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true )] 
        [String]$text,
    
        [Microsoft.Office.Interop.Word.WdColor]$WdColor = 'wdColorAutomatic',
    
        [Microsoft.Office.Interop.Word.WdBuiltinStyle]$WDBuiltinStyle = 'wdStyleDefaultParagraphFont',
    
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
    )
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { $null  = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            if ($PSBoundParameters.ContainsKey('WDBuiltinStyle')) { Write-verbose -Message "$WDBuiltinStyle"; $WordDocument.application.selection.Style = $WDBuiltinStyle }
            if ($PSBoundParameters.ContainsKey('WdColor')) { Write-verbose -Message "$wdcolor"; $WordDocument.Application.Selection.font.Color = $WdColor.value__ }
            $WordDocument.Application.Selection.TypeText("$($text)")    
            $WordDocument.Application.Selection.TypeParagraph() 
            $WordDocument.application.selection.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]'wdStyleNormal'
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

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Add-WordBreak -breaktype Value -WordInstance Value -WordDocument Value
    Describe what this call does

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordbreak


  #>

    [CmdletBinding()]
    param (
        [Parameter(Position = 0)] 
        [Parameter(ParameterSetName = 'GridTable')]
        [ValidateSet('NewPage', 'Section', 'Paragraph')]
        [string]$breaktype = 'NewPage',
   
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
    
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
    )
 
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try {  $null = test-wordinstance -WordInstance $wordinstance 
               $null =  test-WordDocument -WordDocument $WordDocument 
            }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {  
            switch ($breaktype) { 
                'NewPage' { $null =  $WordInstance.Selection.InsertNewPage() }
                'Section' { $null = $WordInstance.Selection.Sections.Add() }
                'Paragraph' { $null =  $WordInstance.Selection.InsertParagraph() } 
            }
            $null = $WordDocument.application.selection.goto([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc')
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

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Set-WordBuiltInProperty -WdBuiltInProperty Value -text Value -WordDocument Value
    Describe what this call does

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/set-wordbuiltinproperty

  #>


    [CmdletBinding()]
    param(
    
        [Parameter(Position = 0, Mandatory = $true)] 
        [Microsoft.Office.Interop.Word.WdBuiltInProperty]$WdBuiltInProperty,
    
        [Parameter(Position = 1, mandatory = $true)] 
        [String]$text,
    
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
    )
    Begin { 
        Write-Verbose -Message "Start  : $($Myinvocation.InvocationName)" 
        try { $null = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try { 
            Write-Verbose -Message $WdBuiltInProperty
            $WordDocument.BuiltInDocumentProperties([Microsoft.Office.Interop.Word.WdBuiltInProperty]$WdBuiltInProperty) = $text
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

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Add-WordCoverPage -CoverPage Value -WordInstance Value -WordDocument Value
    Describe what this call does

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordcoverpage

  #>

    [CmdletBinding()]
    param(
        #Todo cast type instead ie [Microsoft.Office.Interop.Word]
        [Parameter(Position = 0)] 
        [ValidateSet('Austin', 'Banded', 'Facet', 'Filigree', 'Grid', 'Integral', 'Ion (Dark)', 'Ion (Light)', 'Motion', 'Retrospect', 'Semaphore', 'Sideline', 'Slice (Dark)', 'Slice (Light)', 'Viewmaster', 'Whisp')]  
        [string]$CoverPage = 'Facet',
    
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
    
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
    )  
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try {  $null = test-wordinstance -WordInstance $wordinstance 
               $null = test-WordDocument -WordDocument $WordDocument 
            }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            $Selection = $WordDocument.application.selection
            $WordInstance.Templates.LoadBuildingBlocks()
            $bb = $WordInstance.templates | Where-Object -Property name -EQ -Value 'Built-In Building Blocks.dotx'
            $part = $bb.BuildingBlockEntries.item($CoverPage)
            $null = $part.Insert($WordInstance.Selection.range, $true) 
            $null = $Selection.goto([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc')
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

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/set-wordorientation

  #>


    [CmdletBinding()]
    param(
        [Parameter(Position = 0, HelpMessage = 'Orientation of page', Mandatory = $true)] 
        [ValidateSet('Portrait', 'Landscape')]  
        [string]$Orientation,
  
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance
            
    )
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try {  $null = test-wordinstance -WordInstance $wordinstance }
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

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Add-WordTOC -WordInstance Value -WordDocument Value
    Describe what this call does

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordtoc

  #>


    [CmdletBinding()]  
    param (
        #Todo cast type instead ie [Microsoft.Office.Interop.Word.Application]$WordInstance but does not work
        [ValidateScript( {test-wordinstance -WordInstance $_})]
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
  
        [ValidateScript( {test-WordDocument -WordDocument $_})]
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument,
        
        [ValidateRange(0,5)]
        [Int]$Tableader = 0,

        [ValidateRange(0,5)]
        [Int]$IncludePageNumbers = $TRUE
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
        try {  $null = test-wordinstance -WordInstance $wordinstance 
               $null = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process {  
        try {
            $toc = $WordDocument.TablesOfContents.Add($WordInstance.selection.Range)
            $toc.Tableader = $Tableader
            $toc.IncludePageNumbers = $IncludePageNumbers
            $WordDocument.Application.Selection.TypeParagraph()
            $null = $WordDocument.Application.Selection.goto([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc')
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

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Update-WordTOC -WordDocument Value
    Describe what this call does

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/update-wordtoc

  #>


    [CmdletBinding()]   
    param (
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
    )
    Begin {
        Write-Verbose -Message "Start  : $($Myinvocation.InvocationName)" 
        try { $null = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try { $null =  $WordDocument.Fields | ForEach-Object -Process { $_.Update() } }
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

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Add-WordTable -Object Value -WdAutoFitBehavior Value -WdDefaultTableBehavior Value -HeaderRow Value -TotalRow Value -BandedRow Value -FirstColumn Value -LastColumn Value -BandedColumn Value -RemoveProperties -VerticleTable -NoParagraph -WordDocument Value
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

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordtable

  #>


    [CmdletBinding()]
    param(
        [Parameter(Position = 0, HelpMessage = 'psobject to send to word', Mandatory = $true, ValuefromPipeline = $true)]    
        [psobject]$Object,
  
        [Microsoft.Office.Interop.Word.WdAutoFitBehavior]$WdAutoFitBehavior = 'wdAutoFitContent',

        [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]$WdDefaultTableBehavior = 'wdWord9TableBehavior', 

        [bool]$HeaderRow = $true,
    
        [bool]$TotalRow = $false,
    
        [bool]$BandedRow = $true,
    
        [bool]$FirstColumn = $false,
    
        [bool]$LastColumn = $false,
    
        [bool]$BandedColumn = $false,

        [Parameter(ParameterSetName = 'WDTableFormat')]
        [Microsoft.Office.Interop.Word.WdTableFormat]$WDTableFormat = 'wdTableFormatNone',
    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        [Parameter(ParameterSetName = 'PlainTable')]
        [validateSet('Table Grid', 'Table Grid Light', 'Plain Table 1', 'Plain Table 2', 'Plain Table 3', 'Plain Table 4', 'Plain Table 5')]
        [String]$PlainTable = 'Table Grid',

    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        #Todo:  Investigate $table.ApplyStyleDirectFormatting("Grid Table 5 Dark")
        [Parameter( ParameterSetName = 'GridTable')]
        [ValidateSet('Grid Table 1 Light', 'Grid Table 2', 'Grid Table 3', 'Grid Table 4', 'Grid Table 5 Dark', 'Grid Table 6 Colorful', 'Grid Table 7 Colorful')]
        [String]$GridTable = 'Grid Table 1 Light',
    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        [Parameter( ParameterSetName = 'ListTable')]
        [ValidateSet('List Table 1 Light', 'List Table 2', 'List Table 3', 'List Table 4', 'List Table 5 Dark', 'List Table 6 Colorful', 'List Table 7 Colorful')]
        [String]$ListTable = 'List Table 1 Light',
    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        [Parameter( ParameterSetName = 'ListTable')]
        [ValidateSet('Accent 1', 'Accent 2', 'Accent 3', 'Accent 4', 'Accent 5', 'Accent 6')]
        [String]$ListAccent = 'Accent 1',
    
        #Todo:  Investigate how to do better thru [Microsoft.Office.Interop.Word.??????]
        [Parameter( ParameterSetName = 'GridTable')]
        [ValidateSet('Accent 1', 'Accent 2', 'Accent 3', 'Accent 4', 'Accent 5', 'Accent 6')]
        [string]$GridAccent = 'Accent 1',
    
        [switch]$RemoveProperties,
    
        [switch]$VerticleTable,
    
        [switch]$NoParagraph,
    
        [Microsoft.Office.Interop.Word.Document][Object]$WordDocument = $Script:WordDocument
    )
   
    Begin {
        try { $null = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
            $TableRange = $WordDocument.application.selection.range  
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
            $Table = $WordDocument.Tables.Add($TableRange, $Rows, $Columns, $WdDefaultTableBehavior, $WdAutoFitBehavior) 
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
            $Selection = $WordDocument.application.selection
            $null = $Selection.goto([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToBookmark, $null, $null, '\EndOfDoc')
            if (!($NoParagraph)) { $WordDocument.Application.Selection.TypeParagraph() }
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

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/get-wordbuiltinstyle

  #>

    [CmdletBinding()]
    param()
    Begin { Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" }
    Process { 
        try { [Enum]::GetNames([Microsoft.Office.Interop.Word.WdBuiltinStyle]) | ForEach-Object -Process {[pscustomobject]@{ Style = $_ } } }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" } }
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

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/get-wordwdtableformat

  #>


    [CmdletBinding()]
    param()

    Begin { Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" }
    Process { 
        try { [Enum]::GetNames([Microsoft.Office.Interop.Word.WdTableFormat]) | ForEach-Object -Process {[pscustomobject]@{ Style = $_ } } }
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

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Add-WordTemplate -filename Value -WordDocument Value
    Describe what this call does

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordtemplate

  #>


    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, HelpMessage = 'Add word document or template to import', Position = 0, ParameterSetName = 'Default')]
        [ValidateScript( { test-Path -Path $_ })] 
        [string]$filename,

        [Parameter(ParameterSetName = 'Default')]
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument  
    )   
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { test-WordDocument -WordDocument $WordDocument }
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
        try { $WordDocument.Application.Selection.InsertFile([ref]($filename)) }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
}
# SIG # Begin signature block
# MIINCgYJKoZIhvcNAQcCoIIM+zCCDPcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUP48FtRVmln5HwKwI4OdRy0ix
# 8oqgggpMMIIFFDCCA/ygAwIBAgIQDq/cAHxKXBt+xmIx8FoOkTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE4MDEwMzAwMDAwMFoXDTE5MDEw
# ODEyMDAwMFowUTELMAkGA1UEBhMCQVUxGDAWBgNVBAcTD1JvY2hlZGFsZSBTb3V0
# aDETMBEGA1UEChMKU2hhbmUgSG9leTETMBEGA1UEAxMKU2hhbmUgSG9leTCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANAI9q03Pl+EpWcVZ7PQ3AOJ17k6
# OoS9SCIbZprs7NhyRIg7mKzxdcHMnjKwUe/7NDlt5mYzXT2yY/0MeUkyspiEs1+t
# eiHJ6IIs9llWgPGOkV4Ro5fZzlutqeeaomEW/ulH7mVjihVCR6mP/O09YSNo0Dv4
# AltYmVXqhXTB64NdwupL2G8fmTmVUJsww9abtGxy3mhL/l2W3VBcozZbCZVw363p
# 9mjeR9WUz5AxZji042xldKB/97cNHd/2YyWuJ8eMlYfRqz1nVgmmpuU+SuApRult
# hy6wNEngVmJBVhH/a8AH29dEZNL9pzhJGRwGBFi+m/vIr5SFhQVFZYJy79kCAwEA
# AaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsqCqOl6nEDwGD5LfZldQ5YMB0GA1Ud
# DgQWBBROEIC6bKfPIk2DtUTZh7HSa5ajqDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0l
# BAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1oDOgMYYvaHR0cDovL2NybDMuZGln
# aWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1nMS5jcmwwNaAzoDGGL2h0dHA6Ly9j
# cmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3VyZWQtY3MtZzEuY3JsMEwGA1UdIARF
# MEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2lj
# ZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggrBgEFBQcBAQR4MHYwJAYIKwYBBQUH
# MAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBOBggrBgEFBQcwAoZCaHR0cDov
# L2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0U0hBMkFzc3VyZWRJRENvZGVT
# aWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggEBAIly
# KESC2V2sBAl6sIQiHRRgQ9oQdtQamES3fVBNHwmsXl76DdjDURDNi6ptwve3FALo
# ROZHkrjTU+5r6GaOIopKwE4IXkboVoPBP0wJ4jcVm7kcfKJqllSBGZfpnSUjlaRp
# EE5k1XdVAGEoz+m0GG+tmb9gGblHUiCAnGWLw9bmRoGbJ20a0IQ8jZsiEq+91Ft3
# 1vJSBO2RRBgqHTama5GD16OyE3Aps5ypaKYXuq0cnNZCaCasRtDJPolSP4KQ+NVg
# Z/W/rDiO8LNOTDwGcZ2bYScAT88A5KX42wiKnKldmyXnd4ffrwWk8fPngR5sVhus
# Arv6TbwR8dRMGwXwQqMwggUwMIIEGKADAgECAhAECRgbX9W7ZnVTQ7VvlVAIMA0G
# CSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0
# IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEwMjIxMjAwMDBaFw0yODEwMjIxMjAw
# MDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNV
# BAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNz
# dXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7RZmxOttE9X/lqJ3bMtdx6nadBS63
# j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p0WfTxvspJ8fTeyOU5JEjlpB3gvmh
# hCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj6YgsIJWuHEqHCN8M9eJNYBi+qsSy
# rnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grkV7tKtel05iv+bMt+dDk2DZDv5LVO
# pKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHyDxL0xY4PwaLoLFH3c7y9hbFig3NB
# ggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMBAAGjggHNMIIByTASBgNVHRMBAf8E
# CDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEFBQcDAzB5
# BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0
# cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNy
# bDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgGCmCGSAGG/WwAAgQwKjAoBggrBgEF
# BQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAKBghghkgBhv1sAzAd
# BgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1DlgwHwYDVR0jBBgwFoAUReuir/SS
# y4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQELBQADggEBAD7sDVoks/Mi0RXILHwl
# KXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q3yBVN7Dh9tGSdQ9RtG6ljlriXiSB
# ThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/kLEbBw6RFfu6r7VRwo0kriTGxycq
# oSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dcIFzZcbEMj7uo+MUSaJ/PQMtARKUT
# 8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6dGRrsutmQ9qzsIzV6Q3d9gEgzpkx
# Yz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT+hKUGIUukpHqaGxEMrJmoecYpJpk
# Ue8xggIoMIICJAIBATCBhjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdp
# Q2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBAhAOr9wAfEpcG37G
# YjHwWg6RMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkG
# CSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEE
# AYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQU2b/edQ4VDxMUAN0qZa1VhgpOLzANBgkq
# hkiG9w0BAQEFAASCAQCtnSKW+/IqmsLDK6QoL+jOS/SKQMlU9joiIADhNGSRl3V/
# aCWTZ4tIsfRu4cHZ6PuMNYlMrAZ2JtGs2oH5gcH2KcFp55uQKfqo44S0A/UF5f8C
# PP2FP3ztmH+LH7RkfeZ9UBqsE0faTBxLwJCGkoGu9n/4o2bASm0azmwv4GaaKCUI
# Jk5gckHKQhB++5tfva9krc+4mNeNJEVdxKY2bwlvWDbb3zKsUJp9PsKWjLdylS36
# 7t3Xsz/Dl2yYY/Cp+KuKfjA4IRuj6p203/2mZgFiiEfsAVPkPxDX99hEbOyv+pkM
# 9DaFiysbpguIIFoiUsHsEMjJiMDdYTMvKhY+WVkY
# SIG # End signature block
