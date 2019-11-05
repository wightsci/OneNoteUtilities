---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Get-ONNoteBook

## SYNOPSIS
Gets a OneNote Notebook

## SYNTAX

```
Get-ONNoteBook [-NoteBook] <String[]> [<CommonParameters>]
```

## DESCRIPTION
Returns OneNote XML Schema based element representing a specific Notebook

## EXAMPLES

### EXAMPLE 1
```
Get-ONNoteBook -NoteBook 'My NoteBook'
```

## PARAMETERS

### -NoteBook
The NoteBook name to query.
Just one.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: Name

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
