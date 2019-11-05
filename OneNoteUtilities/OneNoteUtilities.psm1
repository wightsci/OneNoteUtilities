
# Set up some variables
$onApp = $Null
$xmlSections=''
$xmlPages=''
$pageID=''
$xmlPage = New-Object System.Xml.XmlDocument
$xmlNewPage = New-Object System.Xml.XmlDocument
$xmlPageDoc = New-Object System.Xml.XmlDocument
$schema = $Null
Function Start-ONApp {
[CmdletBinding()]
param()
if ( -not $script:onApp)  {
  try {
    Write-Verbose "onApp not found"
    $script:onApp = New-Object -ComObject OneNote.Application
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
Function Get-ONHierarchy {
Start-ONApp
$onApp.getHierarchy($null,[Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages,[ref]$xmlPages)
$xmlPageDoc.LoadXML($xmlPages)
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
Start-ONApp
$xmlPages = $xmlPageDoc.SelectNodes("//one:Page",$xmlNS)
$xmlPages
}

Function Get-ONSections {
Start-ONApp
$xmlSections = $xmlPageDoc.SelectNodes("//one:Section",$xmlNS)
$xmlSections
}
Function Get-ONSection {
[CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='What Section?')]
    [Alias('Name')]
    [string[]]$Section
  )
  Start-ONApp
$xmlSection = $xmlPageDoc.SelectSingleNode("//one:Section[@name=`"$($Section)`"]",$xmlNs)
$xmlSection
}
Function New-ONPage {
[CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='What Section?')]
    [Alias('id')]
    [string[]]$SectionID
  )
Begin {
  Start-ONApp
}
Process {
$onApp.createNewPage($SectionID,[ref]$pageID)
$onApp.getPageContent($pageID,[ref]$xmlPage)
$xmlNewPage.LoadXML($xmlPage)
$xmlNewPage.Page
}
}

Function Get-ONNoteBook {
[CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='What Notebook?')]
    [Alias('Name')]
    [string[]]$NoteBook
  )
  Start-ONApp
$xmlNoteBook = $xmlPageDoc.SelectSingleNode("//one:Notebook[@name=`"$($NoteBook)`"]",$xmlNs)
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
Function Update-ONPage {
[CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True)]
    [string[]]$PageContent
  )
Begin {
  Start-ONApp
}
  Process {
    $onApp.UpdatePageContent($PageContent)
}
}
Function Get-ONPage {
[CmdletBinding()]
  param
  (
    [Parameter(Mandatory=$True,
    ValueFromPipeline=$True,
    ValueFromPipelineByPropertyName=$True,
    HelpMessage='What Page?')]
    [Alias('Name')]
    [string[]]$Page
  )
  Begin {
    Start-ONApp
  }
Process {
    $xmlPageContent=''
    $onPage = $xmlPageDoc.SelectSingleNode("//one:Page[@name=`"$Page`" and (@isInRecycleBin!=`"true`" or not (@isInRecycleBin))]",$xmlNs)
    Write-Verbose $xmlPageDoc.OuterXml
    if ($onPage) {
        $onApp.GetPageContent($onPage.id,[ref]$xmlPageContent)
        
        $xmlPage.LoadXML($xmlPageContent)
        $xmlPage.Page
        }
    }
}
Function Show-OnPage {
  [CmdletBinding()]
  param (
  [Parameter(Mandatory=$True,
  ValueFromPipeline=$True,
  ValueFromPipelineByPropertyName=$True,
  HelpMessage='What Page?')]
  [Alias('Name')]
  [string[]]$Page
  )
  $navPage = Get-OnPage -Page $Page
  $onApp.NavigateTo($navPage.id,$Null)
}

Function Publish-ONObject {
  [CmdletBinding()]
  param (
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
