---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Get-ONPage

## SYNOPSIS
Gets a OneNote Page.

## SYNTAX

### Name (Default)
```
Get-ONPage -Page <String[]> [<CommonParameters>]
```

### Id
```
Get-ONPage -Id <String[]> [<CommonParameters>]
```

## DESCRIPTION
Returns OneNote XML Schema based element representing a specific Page.
Ignores pages in the recycle bin.

## EXAMPLES

### EXAMPLE 1
```
Get-ONPage -Page "My Page"
```

This example returns a Page XmlElement object representing the page
with the exact name "My Page".

### EXAMPLE 2
```
Get-ONPages | Where-Object { $_.Name -like 'OneNote*' } | Get-ONPage
```

This example uses the Get-ONPages command and standard PowerShell
filtering to pass objects to Get-ONPage via the pipeline.
Get-ONPage
then returns a Page XmlElement object for each object received.

### EXAMPLE 3
```
Get-Service | Where-Object { $_.Name -like '*winrm*' } | Get-ONPage
```

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

### EXAMPLE 4
```
 Get-ONPage -id '{AB5DB915-FB77-0D89-1B94-8D316660CFCB}{1}{E1910021276453986493171911072997640903877411}'

one              : http://schemas.microsoft.com/office/onenote/2013/onenote
ID               : {AB5DB915-FB77-0D89-1B94-8D316660CFCB}{1}{E1910021276453986493171911072997640903877411}
name             : Article list
dateTime         : 2017-04-15T11:41:36.000Z
lastModifiedTime : 2017-08-24T13:00:13.000Z
pageLevel        : 1
lang             : en-GB
QuickStyleDef    : {PageTitle, p}
PageSettings     : PageSettings
Title            : Title
Outline          : Outline
```

This example returns a page based on its ID.

## PARAMETERS

### -Id
Page Id?

```yaml
Type: String[]
Parameter Sets: Id
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Page
What Page?

```yaml
Type: String[]
Parameter Sets: Name
Aliases: Name

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Any object with a 'Page' or 'Name' property.
## OUTPUTS

### System.Xml.XmlElement extended by the currently selected OneNote schema.
### This includes the full content of the page, unlike the objects returned
### by the Get-ONPages command.
## NOTES
This function uses the XPath SelectSingleNode method 'under the hood'.
This means:
    In the event of multiple pages having the same name, only the first 
    will be returned.
    The page search is case-sensitive.

## RELATED LINKS

[Get-ONPages]()

