# Set up some variables
$onApp = $Null
$xmlSections=''
$strPages=''
$pageID=''
$xmlPage = New-Object System.Xml.XmlDocument
$xmlNewPage = New-Object System.Xml.XmlDocument
$xmlPageDoc = New-Object System.Xml.XmlDocument
$schema = $Null
Function Start-ONApp {
[CmdletBinding()]
  Param()
  if ( -not $script:onApp)  {
    try {
      Write-Verbose "onApp not found"
      $interOp = Get-Item $env:WinDir\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15*\*
      Write-Verbose -Message "Interop Assembly found at: $($interOp.FullName)"
      Add-Type -LiteralPath $interOp.FullName
      # $script:onApp = New-Object -ComObject OneNote.Application
      $script:onApp =  New-Object Microsoft.Office.Interop.OneNote.ApplicationClass
    }
    catch [System.Runtime.InteropServices.COMException] {
      Write-Error "Unable to create COM Object - is OneNote installed?"
      Break
    }

    $script:xmlNs = New-Object System.Xml.XmlNamespaceManager($xmlPageDoc.NameTable)
    $onProcess = Get-Process onenote
    $onVersion = $onProcess.ProductVersion.Split(".")
    Write-Verbose "OneNote version $($onVersion[0]) detected"
    #$onApp | Get-Member | Out-Host
    switch ($onVersion[0]) {
      "16" { $script:schema = "http://schemas.microsoft.com/office/onenote/2013/onenote" }
      "15" { $script:schema = "http://schemas.microsoft.com/office/onenote/2013/onenote" }
      "14" { $script:schema = "http://schemas.microsoft.com/office/onenote/2010/onenote" }
    }
    $xmlNs.AddNamespace("one",$schema)
  }
  else {
      Write-Verbose "onApp found"
      $message  = $onApp.GetType()
      Write-Verbose $message
  }
}

Function Get-ONApp {
  Start-ONApp
  $ONApp
}
Function Get-ONHierarchy {
  Start-ONApp
  $onApp.getHierarchy($null,[Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages,[ref]$strPages)
  $xmlPageDoc.LoadXML($strPages)
}
Function Stop-ONApp {
  [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($onApp)
  $script:onApp = $Null
  #Remove-Variable onApp
  [GC]::Collect()
}
Function Get-ONNoteBooks {
  Start-ONApp
  $xmlNoteBooks = $xmlPageDoc.SelectNodes("//one:Notebook",$xmlNs)
  $xmlNoteBooks
}

Function Get-ONPages {
  [CmdletBinding(DefaultParameterSetName='All')]
  Param(
    [Parameter(ParameterSetName='NotebookName',Mandatory)]
    [String]$NoteBookName,
    [Parameter(ParameterSetName='NotebookId',Mandatory)]
    [String]$NoteBookId,
    [Parameter(ParameterSetName='SectionName',Mandatory)]
    [String]$SectionName,
    [Parameter(ParameterSetName='SectionId',Mandatory)]
    [String]$SectionId
  )
  Start-ONApp
  Write-Verbose $PSCmdlet.ParameterSetName
  switch ($PSCmdlet.ParameterSetName) {
    "NotebookName"  { $xpath = "//one:Notebook[@name='$NotebookName']//one:Page"}
    "NotebookId"    { $xpath = "//one:Notebook[@ID='$NotebookId']//one:Page"}
    "SectionName"   { $xpath = "//one:Section[@name='$SectionName']/one:Page"}
    "SectionId"     { $xpath = "//one:Section[@ID='$SectionId']/one:Page"} 
    "All"           { $xpath = "//one:Page" }
  }
  Write-Verbose $xpath
  $xmlPages = $xmlPageDoc.SelectNodes($xpath,$xmlNS)
  $xmlPages
}

Function Get-ONSectionGroups {
  [CmdletBinding(DefaultParameterSetName='All')]
  Param(
    [Parameter(ParameterSetName='NotebookName',Mandatory)]
    [String]$NoteBookName,
    [Parameter(ParameterSetName='NotebookId',Mandatory)]
    [String]$NoteBookId
  )
  Start-ONApp
  Write-Verbose $PSCmdlet.ParameterSetName
  switch ($PSCmdlet.ParameterSetName) {
    "NotebookName"  { $xpath = "//one:Notebook[@name='$NotebookName']/one:SectionGroup[not(@isRecycleBin='true')]"}
    "NotebookId"    { $xpath = "//one:Notebook[@ID='$NotebookId']/one:SectionGroup[not(@isRecycleBin='true')]"}
    "All"           { $xpath = "//one:SectionGroup[not(@isRecycleBin='true')]" }
  }
  Write-Verbose $xpath
  $xmlSectionGroups = $xmlPageDoc.SelectNodes($xpath,$xmlNS)
  $xmlSectionGroups
}

Function Get-ONSections {
  [CmdletBinding(DefaultParameterSetName='All')]
  Param(
    [Parameter(ParameterSetName='NotebookName',Mandatory)]
    [String]$NoteBookName,
    [Parameter(ParameterSetName='NotebookId',Mandatory)]
    [String]$NoteBookId
  )
  Start-ONApp
  Write-Verbose $PSCmdlet.ParameterSetName
  switch ($PSCmdlet.ParameterSetName) {
    "NotebookName"  { $xpath = "//one:Notebook[@name='$NotebookName']/one:Section"}
    "NotebookId"    { $xpath = "//one:Notebook[@ID='$NotebookId']/one:Section"}
    "All"           { $xpath = "//one:Section" }
  }
  Write-Verbose $xpath
  $xmlSections = $xmlPageDoc.SelectNodes($xpath,$xmlNS)
  $xmlSections
}
Function Get-ONSection {
[CmdletBinding(DefaultParameterSetName='Name')]
  Param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Section Name?',ParameterSetName='Name')]
    [Alias('Name')]
    [string[]]$Section,
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Section Id?',ParameterSetName='Id')]
    [string[]]$Id
  )
  Start-ONApp
  $xpathList = @()
  Write-Verbose $PSCmdlet.ParameterSetName
  switch ($PSCmdlet.ParameterSetName) {
    'Name' { 
      foreach ($s in $Section ) { $xpathList += "//one:Section[@name='$s']" }
    }
    'Id'   { 
      foreach ($i in $Id) { $xpathList += "//one:Section[@ID='$i']" }
    }
  }
  foreach ($xpath in $xpathList) {
    Write-Verbose $PSCmdlet.ParameterSetName
    $xmlSection = $xmlPageDoc.SelectSingleNode("$xpath",$xmlNs)
    $xmlSection
  }
}
Function New-ONPage {
[CmdletBinding()]
  Param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Section ID?')]
    [Alias('id')]
    [string[]]$SectionID,
    [Alias('name')]
    [string]$Title
  )
  Begin {
    Start-ONApp
    $strPage = ''
  }
  Process {
    $onApp.createNewPage($SectionID,[ref]$pageID)
    $onApp.getPageContent($pageID,[ref]$strPage)
    $xmlNewPage.LoadXML($strPage)
    if ($Title) {
      $xmlNewPage.Page.Title.OE.T.InnerText = $Title
    }
    $xmlNewPage.Page
  }
}

Function Get-ONNoteBook {
[CmdletBinding(DefaultParameterSetName='Name')]
  Param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Notebook Name?',ParameterSetName='Name')]
    [Alias('Name')]
    [string[]]$NoteBook,
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Notebook ID?',ParameterSetName='Id')]
    [string[]]$Id
  )
  Start-ONApp
  switch ($PSCmdlet.ParameterSetName) {
      'Name' { $xpath = "//one:Notebook[@name='$Notebook']"}
      'Id'   { $xpath = "//one:Notebook[@ID='$Id']"}
  }
  $xmlNoteBook = $xmlPageDoc.SelectSingleNode("$xpath",$xmlNs)
  $xmlNoteBook
}
Function Add-ONElement {
[CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true,Position=1)]$Element,
    [Parameter(Mandatory=$true,Position=2)]$Parent
  )
  Start-ONApp
  $Parent.AppendChild($Element)
}
Function New-ONElement {
[CmdletBinding()]
  Param(
    [Parameter(Mandatory=$true,Position=1)]$Element,
    [Parameter(Mandatory=$true,Position=2)]$Document
  )
  Start-ONApp
  $Document.OwnerDocument.CreateNode([system.xml.xmlnodetype]::Element,"one:$Element",$schema)
}
Function New-ONCData {
  [CmdletBinding()]
    Param(
      [Parameter(Mandatory=$true,Position=1)]$CData,
      [Parameter(Mandatory=$true,Position=2)]$Document
    )
    Start-ONApp
    $Document.OwnerDocument.CreateNode([system.xml.xmlnodetype]::CDATA,$CData,$schema)
}
Function Update-ONPage {
[CmdletBinding()]
  Param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True)]
    $PageContent
  )
  Begin {
    Start-ONApp
  }
  Process {
    Write-Verbose $PageContent.GetType().Name
    if ($PageContent.GetType().Name -eq 'XmlElement' ) {
      $PageContent = $PageContent.OuterXml
      Write-Verbose $PageContent
    }
    $onApp.UpdatePageContent($PageContent)
    Get-ONHierarchy
  }
}
Function Get-ONPage {
[CmdletBinding(DefaultParameterSetName='Name')]
  Param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Page Name?',ParameterSetName='Name')]
    [Alias('Name')]
    [string[]]$Page,
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Page ID?',ParameterSetName='Id')]
    [string[]]$Id
  )
  Begin {
    Start-ONApp
  }
  Process {
    $xpathlist = @()
      $strPageContent=''
      Write-Verbose $PSCmdlet.ParameterSetName
      switch ($PSCmdlet.ParameterSetName) {
        'Name' { 
          foreach ($p in $Page ) { $xpathlist += "//one:Page[@name='$p'" }
        }
        'Id'   { 
          foreach ($i in $Id) { $xpathlist += "//one:Page[@ID='$i'" }
        }
      }
      foreach ($xpath in $xpathlist) {
        $onPage = $xmlPageDoc.SelectSingleNode("$xpath and (@isInRecycleBin!='true' or not (@isInRecycleBin))]",$xmlNs)
        # Write-Verbose $onPage.OuterXml
        if ($onPage) {
            $onApp.GetPageContent($onPage.id,[ref]$strPageContent)     
            $xmlPage.LoadXML($strPageContent)
            $xmlPage.Page
        }
      }
  }
}
Function Show-OnPage {
  [CmdletBinding(DefaultParameterSetName='Name')]
  Param (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Page Name?',
    Position=0,
    ParameterSetName='Name')]
    [string]$Name,
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Page Id?',
    ParameterSetName='Id')]
    [string]$Id,
    [Parameter(ParameterSetName='Name')]
    [Parameter(ParameterSetName='Id')]
    [switch]$NewWindow
  )
  switch ($PSCmdlet.ParameterSetName) {
    "Name" { $navPage = Get-ONPage -Page $Name} 
    "Id"   { $NavPage = Get-ONPage -Id $Id }
  }
  if ($NewWindow.IsPresent) {
    $WindowFlag = -1
  }
  else {
    $WindowFlag = 0
  }
  $onApp.NavigateTo($navPage.id,$Null,$WindowFlag)
}

Function Show-ONGUI {
[CmdletBinding()]
Param()
Start-ONApp
$onApp.NavigateTo($Null)
}

Function Publish-ONObject {
  [CmdletBinding()]
  Param (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='Please provide a OneNote object ID')]
    [Alias('Identity')]
    [string[]]$Id,
    [Parameter(Mandatory=$True,
    HelpMessage='Please provide a valid OneNote export type')]
    [ValidateSet("PDF","XPS","DOC","EMF","ONEPKG","MHT","HTML")]
    [Alias('Type')]
    [string[]]$Format,
    [Parameter(Mandatory=$True,
    HelpMessage='Please provide a file path')]
    [Alias('FilePath')]
    [string[]]$Path
  )
  switch ($Format.ToLower()) {
    "onepkg"  {$PublishFormat = 1;break}
    "mht"     {$PublishFormat = 2;break}
    "pdf"     {$PublishFormat = 3;break}
    "xps"     {$PublishFormat = 4;break}
    "doc"     {$PublishFormat = 5;break}
    "emf"     {$PublishFormat = 6;break}
    "html"    {$PublishFormat = 7;break}
    default   {$PublishFormat = -1;break}
  }
  Write-Verbose $PublishFormat
  if ($PublishFormat -ge 0) {
    $onApp.Publish($Id,$Path,$PublishFormat,"")
  }
}
Get-ONHierarchy
<#
#Get-ONNoteBooks |gm
$myNoteBook = Get-ONNoteBook -NoteBook "Stuart's Notebook"
$myNoteBook 
#Get-ONPages | Select Name
#Get-ONPage -Page "The Trouble with Tablets" | Select *
#Get-ONPages | Where-Object {$_.Name -like '*tablet*'} | Get-ONPage | Get-Member
$myPage = Get-ONPages | Where-Object {$_.Name -like '*tablet*'} | Get-ONPage 
$myPage
#>
