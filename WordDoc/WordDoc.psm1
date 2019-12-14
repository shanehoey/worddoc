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
    https://shanehoey.github.io/worddoc/

    .LINK
    License
    https://shanehoey.github.io/worddoc/license/

#>

try { Add-Type -AssemblyName Microsoft.Office.Interop.Word }
catch { Write-Warning  -Message "$($MyInvocation.InvocationName) - Unable to add Word Assembly, Word must be installed for this module to work" }

function New-WordInstance {
  <#
    .SYNOPSIS
    The New-WordInstance function starts a new instance of MS Word.

    .DESCRIPTION
    The New-WordInstance function starts a new instance of MS Word.

    .PARAMETER WindowState
    Set the MS Word application wdWindowStateMaximize, wdWindowStateMinimize, wdWindowStateNormal

    .PARAMETER Visible
    Makes MS Word application Visable or Hidden

    .PARAMETER ReturnObject
    When used the function will return the Word Instance as an Object to be stored in a variable in the local shell. 
    If using this method you must use worddocobject as well, and manually parse these objects to all functions. 


    .EXAMPLE
    New-WordInstance -WindowState wdWindowStateMaximize -Visable True 
    
    Create a new Word Instance that is maximised and is visable.
    
    .EXAMPLE
    New-WordInstance -Visable False 

    Create a new Word Instance that is hidden.
    
    .EXAMPLE
    $wi = New-WordInstance -ReturnObject

    Create a word instance that is stored in a local variable.
    
    .INPUTS

    .OUTPUTS
     [Microsoft.Office.Interop.Word.Application]

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK

    New-WordInstance

    https://shanehoey.github.io/worddoc/docs/new-wordinstance

  #>

    [CmdletBinding()]
    Param( 
        [Microsoft.Office.Interop.Word.WdWindowState]$WindowState = "wdWindowStateMaximize",
        [alias("Visable")]
        [bool]$Visible = $true,
        [switch]$ReturnObject
        )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { Add-Type -AssemblyName Microsoft.Office.Interop.Word }
        catch { Write-Warning  -Message "$($MyInvocation.InvocationName) - Unable to add Word Assembly, Word must be installed for this module... exiting" ; break }
        if (!($returnobject)) { 
            try { if (test-path -Path variable:script:WordInstance) {throw 'WordInstance already exists'} }
            catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" ; break }
            }
        }
    Process { 
        try 
        { 
            #$WordInstance = New-Object -ComObject Word.Application -Property @{Visible = $Visible} 
            $WordInstance = New-Object -ComObject Word.Application
            if ($Visible) 
            {
                Write-Warning "*** MS Word may be behind this window! ***"
                $WordInstance.Visible = $Visible
                $WordInstance.Activate()
                $WordInstance.WindowState = $WindowState
            }
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - Unable to Invoke Word... exiting" ; break }
        try { if ($returnobject) { return $WordInstance } else { New-Variable -Name 'WordInstance' -Value $WordInstance -scope script -ErrorAction SilentlyContinue } }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - Unable to create variable... exiting" ; break }   
        }
    End { 
        Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***" 
    }
    }

function Get-WordInstance {
    <#
      .SYNOPSIS
      This function is used to return a Word Instance created automatically by Word Doc Module
  
      .DESCRIPTION
      This function is used to return a Word Instance created automatically by Word Doc Module
  
      .PARAMETER WordInstance
      Not required as this function will work without using WordInstance Parameter
  
      .EXAMPLE
      Get-WordInstance -WordInstance Value
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

function Test-WordInstance {
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
        [Microsoft.Office.Interop.Word.WdSaveOptions]$SaveOptions = "wdPromptToSaveChanges",
        [Microsoft.Office.Interop.Word.WdOriginalFormat]$OriginalFormat = "wdPromptUser",
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument,
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance

      )
    Begin { 
          Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
          try { $null = test-wordinstance -WordInstance $wordinstance }
          catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
        }
    Process {     
        try {

            if ($wordinstance.activedocument){
                Close-WordDocument -SaveOptions $SaveOptions -OriginalFormat $OriginalFormat
                if (test-path variable:script:Worddocument) { remove-variable WordDocument -Scope Script }
            }
            if ($WordInstance.ActiveWindow.Active) { throw "Unable to exit - Document Not Closed" }
            if (!($wordinstance.ActiveDocument)){ 
                $WordInstance.Quit()
                if (test-path variable:script:WordInstance) { remove-variable WordInstance -Scope Script }
            }

            }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
        }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
    }

function New-WordDocument {
    <#
      .SYNOPSIS
      Describe purpose of "New-WordDocument" in 1-2 sentences.
  
      .DESCRIPTION
      Add a more complete description of what the function does.
  
      .PARAMETER WordInstance
      Describe parameter -WordInstance.
  
      .PARAMETER returnobject
      Describe parameter -returnobject.
  
      .EXAMPLE
      New-WordDocument -WordInstance Value -WordDocObject
  
      .NOTES
      for more examples visit https://shanehoey.github.io/worddoc/
  
      .LINK
      https://shanehoey.github.io/worddoc/docs/new-worddocument
  
    #>
    
      [CmdletBinding()]
      Param(  
          [Parameter(Position = 0)] 
          [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,

          [switch]$returnobject
        )
      Begin { 
          Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"
          try { $null = test-wordinstance -WordInstance $wordinstance }
          catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
          if (!($returnobject)) { 
            try { if (test-path -Path variable:script:WordDocument) {throw 'WordDocument already exists'} }
            catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" ; break }
            }
        }
      Process { 
          try {

              $WordDocument = $WordInstance.Documents.Add()
              $WordDocument.Activate()
            
              try { if ($returnobject) { return $WordDocument } else { New-Variable -Name 'WordDocument' -Value $WordDocument -scope script -ErrorAction SilentlyContinue } }
              catch { Write-Warning -Message "$($MyInvocation.InvocationName) - Unable to create variable... exiting" ; break }   
            }
            catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }

        }
      End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)"  }
  }

function Get-WordDocument {
  <#
    .SYNOPSIS
    Describe purpose of "Get-WordDoc" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Get-WordDocument -WordDocument Value
   

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/get-worddocument

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

function Test-WordDocument {
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
    https://shanehoey.github.io/worddoc/docs/test-worddocument

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

function Save-WordDocument {
  <#
    .SYNOPSIS
    Save a Word Document (also .pdf and .html)

    .DESCRIPTION
    Save a Word Document (also .pdf and .html)

    .PARAMETER WordDocument
    Word Document to save

    .PARAMETER WordSaveFormat
    Format to save document as.

    .PARAMETER filename
    Filename to save document as

    .PARAMETER folder
    Folder to save document in 

    .EXAMPLE
    Save-WordDocument -WordSaveFormat wdFormatDocument -filename worddoc.docx -folder c:\users\shane\documents\
    
    Saves document as a standard Word Document in c:\users\shane\documents\worddoc.docx

      .EXAMPLE
    Save-WordDocument 

    Opens a save-as GUI, allowing you to save as a docx, html, or pdf file. 

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/save-worddocument

  #>
    [CmdletBinding()]
    Param( 
        
        [Parameter(Mandatory=$true,ParameterSetName = "SaveAs")]
        [string]$Filename,

        [Parameter(ParameterSetName = "SaveAs")]
        [String]$Folder = [Environment]::GetFolderPath('MyDocuments'),

        [Parameter(Mandatory=$false,ParameterSetName = "SaveAs")]
        [Microsoft.Office.Interop.Word.WdSaveFormat]$WordSaveFormat,
     
        [Parameter(ParameterSetName = "Default")]
        [Parameter(ParameterSetName = "SaveAs")]
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
        
    )
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { $null = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try { 
            if ($PSCmdlet.ParameterSetName -eq "SaveAs") { 
                $filepath = Join-Path -path $folder -ChildPath $filename
                $WordDocument.SaveAs([ref][system.object]$filepath , [ref]$WordSaveFormat) 
            }
            else
            {
                if (!($WordDocument.saved)) { $WordDocument.Save() }
            } 
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)" }
        }
    End { Write-Verbose -Message "[End] *** $($Myinvocation.InvocationName) ***"   }
    }

function Close-WordDocument {
  <#
    .SYNOPSIS
    Describe purpose of "Close-WordDocument" in 1-2 sentences.

    .DESCRIPTION
    Add a more complete description of what the function does.
    
    .PARAMETER SaveChanges
    Describe parameter -wdPromptToSaveChanges.

    .PARAMETER OriginalFormat
    Describe parameter -wdPromptUser.

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
        [Microsoft.Office.Interop.Word.WdSaveOptions]$SaveOptions = "wdPromptToSaveChanges",
        [Microsoft.Office.Interop.Word.WdOriginalFormat]$OriginalFormat = "wdPromptUser",
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument,
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance
    )
    Begin { 
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***"  
        try { 
            $null = test-WordDocument -WordDocument $worddocument
        }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process {     
        try 
        {
            if ($WordInstance.ActiveDocument) { 
                $WordDocument.close($SaveOptions,$OriginalFormat)
                if (!($wordinstance.ActiveDocument)) { if (test-path variable:script:Worddocument) { remove-variable WordDocument -Scope Script } }
            }
        }
        catch 
        {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message) - Document Not Saved"
        }
    }
    End { Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" }
    }

function Add-WordBreak {
  <#
    .SYNOPSIS
    Create a new Break (newpage,section or paragraph)

    .DESCRIPTION
     Create a new Break (newpage,section or paragraph)

    .PARAMETER breaktype
    Type of break (newpage,section or paragraph)

    .PARAMETER WordInstance
    Describe parameter -WordInstance.

    .PARAMETER WordDocument
    Describe parameter -WordDocument.

    .EXAMPLE
    Add-WordBreak -breaktype NewPage
    Creates a NewPage Break

    .NOTES
    for more examples visit https://shanehoey.github.io/worddoc/

    .LINK
    https://shanehoey.github.io/worddoc/docs/add-wordinstance


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
    
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
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
        [Parameter(Mandatory = $false, HelpMessage = 'Word Document or Template to import', Position = 0, ParameterSetName = 'Default')]
        [ValidateScript({ test-Path -Path $_ })] 
        [string]$filename,

        [Parameter(ParameterSetName = 'Default')]
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument  
        )   
    Begin {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { $null = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
        if ($PSBoundParameters.ContainsKey('filename')) {
            $filename = (get-item -path $filename).fullname
            } 
        else { 
            Add-Type -AssemblyName System.windows.forms 
            $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory =  [Environment]::GetFolderPath('Desktop')
            $OpenFileDialog.filter = 'Word Documents (*.docx)|*.docx|Word Templates (*.dotx)|*.dotx'
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

        [Microsoft.Office.Interop.Word.WdBuiltinStyle]$WDBuiltinStyle,
    
        [switch]$Allcaps,
        [switch]$Bold,
        [switch]$DoubleStrikeThrough,
        [switch]$Italic,
        [Int]$Size,
        [switch]$SmallCaps,
        [switch]$StrikeThrough,
        [switch]$SubScript,
        [switch]$SuperScript,
        [alias("WdColor")]
        [Microsoft.Office.Interop.Word.WdColor]$TextColor,
        [switch]$Underline,
        [string]$Font,
        [Microsoft.Office.Interop.Word.WdParagraphAlignment]$Align,

        [switch]$NoParagraph,

        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
        
    )
    Begin 
    {
        Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
        try { $null  = test-WordDocument -WordDocument $WordDocument }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try {
         

            if ($PSBoundParameters.ContainsKey('WDBuiltinStyle')) { $WordDocument.application.selection.Style = $WDBuiltinStyle }
            if ($PSBoundParameters.ContainsKey('Allcaps')) { $WordDocument.Application.Selection.Font.Allcaps = $true }
            if ($PSBoundParameters.ContainsKey('Bold')) { $WordDocument.Application.Selection.Font.Bold = $true }
            if ($PSBoundParameters.ContainsKey('DoubleStrikeThrough')) { $WordDocument.Application.Selection.Font.DoubleStrikeThrough = $true }
            if ($PSBoundParameters.ContainsKey('Italic')) { $WordDocument.Application.Selection.Font.Italic = $true }
            if ($PSBoundParameters.ContainsKey('Size')) { $WordDocument.Application.Selection.Font.Size = $Size }
            if ($PSBoundParameters.ContainsKey('SmallCaps')) { $WordDocument.Application.Selection.Font.smallcaps = $true }
            if ($PSBoundParameters.ContainsKey('StrikeThrough')) { $WordDocument.Application.Selection.Font.StrikeThrough = $true }
            if ($PSBoundParameters.ContainsKey('SubScript')) { $WordDocument.Application.Selection.Font.Subscript = $true }
            if ($PSBoundParameters.ContainsKey('SuperScript')) { $WordDocument.Application.Selection.Font.Superscript = $true }
            if ($PSBoundParameters.ContainsKey('Underline')) { $WordDocument.Application.Selection.Font.Underline = $true }
            if ($PSBoundParameters.ContainsKey('TextColor')) { $WordDocument.Application.Selection.font.Color = $TextColor.value__ }
            if ($PSBoundParameters.ContainsKey('Font')) { if ($worddocument.application.FontNames -contains $font) { $worddocument.Application.Selection.font.name = $Font } }
            if ($PSBoundParameters.ContainsKey('Align')) { $worddocument.Application.Selection.ParagraphFormat.Alignment = $align } 

            $WordDocument.Application.Selection.TypeText("$($text)")
            if (!($noparagraph)) { $WordDocument.Application.Selection.TypeParagraph() }
            
            $worddocument.Application.Selection.Font.reset()
            $worddocument.Application.Selection.select()

        }
        catch {
            Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
        }
    }
    End { 
        Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" 
    }
    }

function Add-WordShape {
  
      [CmdletBinding()]
      param(

        [Parameter(Mandatory=$True,ParameterSetName = "Default")]
        [Parameter(Mandatory=$True,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$True,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$True,ParameterSetName = "PresetTexture")]
        [int]$left=0,
        [Parameter(Mandatory=$True,ParameterSetName = "Default")]
        [Parameter(Mandatory=$True,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$True,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$True,ParameterSetName = "PresetTexture")]
        [int]$Top=0,
        [Parameter(Mandatory=$True,ParameterSetName = "Default")]
        [Parameter(Mandatory=$True,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$True,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$True,ParameterSetName = "PresetTexture")]
        [int]$Width=300, 
        [Parameter(Mandatory=$True,ParameterSetName = "Default")]
        [Parameter(Mandatory=$True,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$True,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$True,ParameterSetName = "PresetTexture")]
        [int]$Height=300,
        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [Microsoft.Office.Core.MsoAutoShapeType]$shape = "msoShapeRectangle",
        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [Microsoft.Office.Core.MsoZOrderCmd]$zorder = "msoSendBehindText",
        
        [Parameter(Mandatory=$True,ParameterSetName = "themecolor")]
        [Microsoft.Office.Core.MsoThemeColorIndex]$themecolor,
        
        [Parameter(Mandatory=$True,ParameterSetName = "UserPicture")]
        [string]$UserPicture, 

        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Microsoft.Office.Core.msoPictureEffectType]$PictureEffect = "msoEffectNone",

        [Parameter(Mandatory=$True,ParameterSetName = "PresetTexture")]
        [Microsoft.Office.Core.MsoPresetTexture]$PresetTexture,
        
        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [int]$Lineweight=0,

        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [int]$LineStyle=1,

        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [Switch]$LineVisible,

        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [switch]$RelativeVerticalPosition,

        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [switch]$RelativeHorizontalPosition,

        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [switch]$LockPosition,

        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance,
        
        [Parameter(Mandatory=$false,ParameterSetName = "Default")]
        [Parameter(Mandatory=$false,ParameterSetName = "themecolor")]
        [Parameter(Mandatory=$false,ParameterSetName = "UserPicture")]
        [Parameter(Mandatory=$false,ParameterSetName = "PresetTexture")]
        [Microsoft.Office.Interop.Word.Document]$WordDocument = $Script:WordDocument
          
      )
      Begin 
      {
          Write-Verbose -Message "[Start] *** $($Myinvocation.InvocationName) ***" 
          try { $null  = test-WordDocument -WordDocument $WordDocument }
          catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
      }
      Process { 
          try {
              
           $newshape = $worddocument.shapes.AddShape($shape,$Left,$top,$Width,$Height)
           
           if ($PSBoundParameters.ContainsKey('themecolor')) { $newshape.Fill.ForeColor.ObjectThemeColor = [int]($themecolor.value__) }
           if ($PSBoundParameters.ContainsKey('zorder')) { $newshape.zorder($zorder) }
           if ($PSBoundParameters.ContainsKey('UserPicture')) { $newshape.fill.UserPicture($UserPicture) }
           if ($PSBoundParameters.ContainsKey('PictureEffect')) { $newshape.fill.PictureEffects.Insert($PictureEffect) }
           if ($PSBoundParameters.ContainsKey('PresetTexture')) { $newshape.fill.UserTextured($PresetTexture) }
           if ($PSBoundParameters.ContainsKey('LineVisible')) { $newshape.Line.Visible = -1 } else { $newshape.Line.Visible = 0 }
           if ($PSBoundParameters.ContainsKey('Lineweight')) { $newshape.Line.Weight = $Lineweight } else { $newshape.Line.Weight = 0 }
           if ($PSBoundParameters.ContainsKey('LineStyle')) { $newshape.Line.Style = $LineStyle } else { $newshape.Line.Style = 1 }
           if ($PSBoundParameters.ContainsKey('RelativeVerticalPosition')) { $newshape.RelativeVerticalPosition = 1 }
           if ($PSBoundParameters.ContainsKey('RelativeHorizontalPosition')) { $newshape.RelativeHorizontalPosition = 1 }
           if ($PSBoundParameters.ContainsKey('LockPosition')) { $newshape.LockAnchor = 1 }
            

           #(Get-WordDocument).shapes(1).fill.forecolor.objectthemecolor =9
           #(Get-WordDocument).shapes(1).fill.backcolor.objectthemecolor =4
           #(Get-WordDocument).shapes(1).line.forecolor.objectthemecolor =9
           #(Get-WordDocument).shapes(1).line.fill.backcolor.objectthemecolor =4

          }
          catch {
              Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"
          }
      }
      End { 
          Write-Verbose -Message "End    : $($Myinvocation.InvocationName)" 
      }
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
            $WordDocument.Application.Selection.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]'wdStyleNormal'
            $WordDocument.Application.Selection.Font.reset()

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
        [ValidateScript( {test-wordinstance -WordInstance $_})]
        [Microsoft.Office.Interop.Word.Application]$WordInstance = $Script:WordInstance
    )
    Begin {
        Write-Verbose -Message "Start  : $($Myinvocation.InvocationName)" 
        try {  $null = test-wordinstance -WordInstance $wordinstance }
        catch { Write-Warning -Message "$($MyInvocation.InvocationName) - $($_.exception.message)"; break }
    }
    Process { 
        try { $WordInstance.ActiveDocument.TablesOfContents[1].Update() }
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

# SIG # Begin signature block
# MIINHwYJKoZIhvcNAQcCoIINEDCCDQwCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUnYAhi+K/Xi53q7Abher6ouBl
# skqgggphMIIFKTCCBBGgAwIBAgIQD8tApulPpYV/uEuZ3XX3/jANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE5MDIwOTAwMDAwMFoXDTE5MTAx
# NTEyMDAwMFowZjELMAkGA1UEBhMCQVUxEzARBgNVBAgTClF1ZWVuc2xhbmQxGDAW
# BgNVBAcTD1JvY2hlZGFsZSBTb3V0aDETMBEGA1UEChMKU2hhbmUgSG9leTETMBEG
# A1UEAxMKU2hhbmUgSG9leTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# ALwO4uf2IRVuz+vei74RR98B7LYaN0CFslmxAOISgihLCAHy6TpWNShnOFQBHz4B
# vKAX86W5532uyh8pr4pN4UistsyzggFaYrYl7x6KWLGzt/ku0nx4CYnoZaGNdeDc
# oJ7ukJvaEmD6CDBmIwMYOa7gDih07EAlq1ZCHXLZKTcvQ1YBHkn0sxIDyg3ilrQK
# mO8G5JHh17GGb+n6OzUWNwYRwCmktEXDMJYVtgmjSVwLbFU+SPgGld5lnzqELjgh
# NvuVXsdSotJXIXjBAjuZComoSYdEVukYVhNh228TgH/M45M2yLLBgLPnvd/L7gUy
# /cAEBd45hrjNuwXhXVrgzl0CAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsq
# CqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBRMl/fVAn1vK9RW7FdPr5dMUDNCMTAO
# BgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1
# oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1n
# MS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3Vy
# ZWQtY3MtZzEuY3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUH
# AgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggr
# BgEFBQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNv
# bTBOBggrBgEFBQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lD
# ZXJ0U0hBMkFzc3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAw
# DQYJKoZIhvcNAQELBQADggEBADq0MMofNx0tgG3mARjfSWbIE6fUWPDqJwFVfjWy
# vu+u7qQk6d0RP8EF25najMaEyg6X1Q/Cb6Lo6O9ILn56QKjqtELyFNvq+Ei0hBs7
# jk/+DAZqhuKFFtVle9hSbM0R41b5viZK+yBrh2SD6kGYSg81XVvzuaWYmNQESoW9
# bLOnO0QTcuz2Pe/0hYwqUnlCzm3yl9M485TBJdnB754YBgKcrYSLL57Kit4c2U7D
# rdP0YxAQdjMY9xQacd8Rc16sSyCmi2Q3b8xSkBXSCyqCnkEYMK9n3hlMGw0aM000
# 4rJaeT94x77x1nhpyKMMHgaK+XmDPMnuYPsKZxX4QE9GCtYwggUwMIIEGKADAgEC
# AhAECRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEw
# MjIxMjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNV
# BAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7
# RZmxOttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p
# 0WfTxvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj
# 6YgsIJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grk
# V7tKtel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHy
# DxL0xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMB
# AAGjggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjAT
# BgNVHSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCB
# gQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lD
# ZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgG
# CmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQu
# Y29tL0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1
# DlgwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEL
# BQADggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q
# 3yBVN7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/
# kLEbBw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dc
# IFzZcbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6
# dGRrsutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT
# +hKUGIUukpHqaGxEMrJmoecYpJpkUe8xggIoMIICJAIBATCBhjByMQswCQYDVQQG
# EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNl
# cnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBT
# aWduaW5nIENBAhAPy0Cm6U+lhX+4S5nddff+MAkGBSsOAwIaBQCgeDAYBgorBgEE
# AYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwG
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBT3bLYm
# 2XPp1APvDn4ucbbCqY8m0DANBgkqhkiG9w0BAQEFAASCAQBIEBUSo/AsXLQiL34n
# MYmOuK4vkN+2vqIkGXUSe6uJSKtHSaWT3PBXO0yN0YKgGauF2qeib9iEQUiGzdUw
# C8NQvWLE7lpb5J3cxj2q2nweAg195VV88cAsC945DYz6A+8EZW9lGcT7g4R1QePC
# jL2iOMa9tNFG3YaXpdbQbEl69S3C8MnNyJ2Y9a3aGvANH+fMtTmNSG5yAVHK7mSz
# 7VWk7xmAZywdq0hdRdD0atglRrgAhJ6dziWk6d1ttsHQ1LcFCyEXz/No9OjJMtwT
# o9z4hG3xGeyrFTNZ1jgzMpoFoP4dktJ60ZmjumgDXhF2c8nqSMfu4bsNymd/Tlqt
# VH26
# SIG # End signature block
