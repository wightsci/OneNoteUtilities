#Some comment text
#More comment
try {
$onApp = New-Object -ComObject OneNote.Application
}
catch [System.Runtime.InteropServices.COMException] {
  Write-Error "Unable to create COM Object - is OneNote installed?"
  Break
}
# Set up some variables
$xmlSections=''
$xmlPages=''
$pageID=''
$xmlPage = New-Object System.Xml.XmlDocument
$xmlPageDoc = New-Object System.Xml.XmlDocument
$xmlNs = New-Object System.Xml.XmlNamespaceManager($xmlPageDoc.NameTable)
$onProcess = Get-Process onenote
$onVersion = $onProcess.ProductVersion.Split(".")
"OneNote version $($onVersion[0]) detected"
switch ($onVersion[0]) {
    "16" { $schema = "http://schemas.microsoft.com/office/onenote/2013/onenote" }
    "15" { $schema = "http://schemas.microsoft.com/office/onenote/2013/onenote" }
    "14" { $schema = "http://schemas.microsoft.com/office/onenote/2010/onenote" }
    }
$xmlNs.AddNamespace("one",$schema)

Function Get-ONHierarchy {
<#
.SYNOPSIS
Gets the current OneNote Hierarchy
.DESCRIPTION
Loads the current OneNote Hierarchy for use by other functions
.EXAMPLE
Get-ONHierarchy
#>
$onApp.getHierarchy($null,[Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages,[ref]$xmlPages)
$xmlPageDoc.LoadXML($xmlPages)
}

Function Get-ONNoteBooks {
<#
  .SYNOPSIS
  Gets OneNote Notebooks
  .DESCRIPTION
  Returns OneNote XML Schema based elements representing Notebooks
  .EXAMPLE
  Get-ONNoteBooks
#>
$xmlNoteBooks = $xmlPageDoc.SelectNodes("//one:Notebook",$xmlNs)
$xmlNoteBooks
}

Function Get-ONPages {
<#
  .SYNOPSIS
  Gets OneNote Pages
  .DESCRIPTION
  Returns OneNote XML Schema based elements representing Pages
  .EXAMPLE
  Get-ONPages
#>
$xmlPages = $xmlPageDoc.SelectNodes("//one:Page",$xmlNS)
$xmlPages
}

Function Get-ONSections {
<#
  .SYNOPSIS
  Gets OneNote Sections
  .DESCRIPTION
  Returns OneNote XML Schema based elements representing Sections
  .EXAMPLE
  Get-ONSections
#>
$xmlSections = $xmlPageDoc.SelectNodes("//one:Section",$xmlNS)
$xmlSections
}
Function Get-ONSection {
<#
.SYNOPSIS
Gets OneNote Section
.DESCRIPTION
Returns OneNote XML Schema based elements representing a Section
.PARAMETER Section
The Section name to query. Just one.
#>
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
$xmlSection = $xmlPageDoc.SelectSingleNode("//one:Section[@name=`"$($Section)`"]",$xmlNs)
$xmlSection
}
Function New-ONPage {
<#
.SYNOPSIS
Create a new OneNote Page.
.DESCRIPTION
Returns a OneNote XML Schema based element representing the new page.
.PARAMETER SectionID
The ID of the Section in which the Page is to be created.
.EXAMPLE
Get-ONSections | Where-Object { $_.name -like '*unfiled*' } | New-ONPage

xml           Page
---           ----
version="1.0" Page

This example uses the Get-ONSections command and standard PowerShell
filtering to pass objects to New-ONPage via the pipeline. New-ONPage
then returns a Page XmlElement object for each object received.

.INPUTS
Any object with an 'id' property
.OUTPUTS
System.Xml.XmlElement extended by the currently selected OneNote schema.
#>
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
Process {
$onApp.createNewPage($SectionID,[ref]$pageID)
$onApp.getPageContent($pageID,[ref]$xmlPage)
$xmlPageDoc.LoadXML($xmlPage)
$xmlPageDoc
}
}

Function Get-ONNoteBook {
<#
  .SYNOPSIS
  Gets a OneNote Notebook
  .DESCRIPTION
  Returns OneNote XML Schema based element representing a specific Notebook
  .EXAMPLE
  Get-ONNoteBook -NoteBook 'My NoteBook'
  .PARAMETER NoteBook
  The NoteBook name to query. Just one.
#>
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
$xmlNoteBook = $xmlPageDoc.SelectSingleNode("//one:Notebook[@name=`"$($NoteBook)`"]",$xmlNs)
$xmlNoteBook
}
Function New-ONElement {
<#
.SYNOPSIS
Creates a OneNote XML Schema based element
.DESCRIPTION
Creates an element of the specified type in 
the specified XML document's DOM using the 
currently in-use schema.
.EXAMPLE
New-ONElement -Element "T" -Document $XMLDoc
.PARAMETER
Element
.PARAMETER
Document
#>
[CmdletBinding()]
Param(
[Parameter(Mandatory=$true,Position=1)]$Element,
[Parameter(Mandatory=$true,Position=2)]$Document
)
$Document.CreateElement("one","$Element",$schema)
}
Function Get-ONPage {
<#
.SYNOPSIS
Gets a OneNote Page.
.DESCRIPTION
Returns OneNote XML Schema based element representing a specific Page.
Ignores pages in the recycle bin.
.EXAMPLE
Get-ONPage -Page "My Page"
This example returns a Page XmlElement object representing the page
with the exact name "My Page".
.EXAMPLE
Get-ONPages | Where-Object { $_.Name -like 'OneNote*' } | Get-ONPage
This example uses the Get-ONPages command and standard PowerShell
filtering to pass objects to Get-ONPage via the pipeline. Get-ONPage
then returns a Page XmlElement object for each object received.
.EXAMPLE
Get-Service | Where-Object { $_.Name -like '*winrm*' } | Get-ONPage

one              : http://schemas.microsoft.com/office/onenote/2013/onenote
ID               : {D7B35AD3-1559-0CBB-0F63-F10786864060}{1}{E19476877483600779377920100891604390372276781}
name             : WinRM
dateTime         : 2016-06-25T14:28:56.000Z
lastModifiedTime : 2016-06-25T14:30:28.000Z
pageLevel        : 1
lang             : en-GB
QuickStyleDef    : {PageTitle, p}
PageSettings     : PageSettings
Title            : Title
Outline          : Outline

This example returns a Page XmlElement that whose name matches that
of the object passed down the pipeline. 
.INPUTS
Any object with a 'Page' or 'Name' property.
.OUTPUTS
System.Xml.XmlElement extended by the currently selected OneNote schema.
This includes the full content of the page, unlike the objects returned
by the Get-ONPages command.
.NOTES
This function uses the XPath SelectSingleNode method 'under the hood'.
This means:
    In the event of multiple pages having the same name, only the first 
    will be returned.
    The page search is case-sensitive.
.LINK
Get-ONPages
#>
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
Process {
    $xmlPageContent=''
    $onPage = $xmlPageDoc.SelectSingleNode("//one:Page[@name=`"$Page`" and (@isInRecycleBin!=`"true`" or not (@isInRecycleBin))]",$xmlNs)
    if ($onPage) {
        $onApp.GetPageContent($onPage.id,[ref]$xmlPageContent)
        
        $xmlPage.LoadXML($xmlPageContent)
        $xmlPage.Page
        }
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
