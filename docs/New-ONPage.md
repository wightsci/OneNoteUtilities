---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# New-ONPage

## SYNOPSIS
Create a new OneNote Page.

## SYNTAX

```
New-ONPage [-SectionID] <String[]> [<CommonParameters>]
```

## DESCRIPTION
Returns a OneNote XML Schema based element representing the new page.

## EXAMPLES

### EXAMPLE 1
```
Get-ONSections | Where-Object { $_.name -like '*unfiled*' } | New-ONPage
```

xml           Page
---           ----
version="1.0" Page

This example uses the Get-ONSections command and standard PowerShell
filtering to pass objects to New-ONPage via the pipeline.
New-ONPage
then returns a Page XmlElement object for each object received.

## PARAMETERS

### -SectionID
The ID of the Section in which the Page is to be created.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: id

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### Any object with an 'id' property
## OUTPUTS

### System.Xml.XmlElement extended by the currently selected OneNote schema.
## NOTES

## RELATED LINKS
