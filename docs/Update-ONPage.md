---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Update-ONPage

## SYNOPSIS
Updates an existing OneNote page

## SYNTAX

```
Update-ONPage [-PageContent] <String[]> [<CommonParameters>]
```

## DESCRIPTION
Updates a OneNote page using the currently in-use schema.

## EXAMPLES

### EXAMPLE 1
```
Update-ONPage $myPage.OuterXML
```

## PARAMETERS

### -PageContent
An xml string containing the updated page content

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
