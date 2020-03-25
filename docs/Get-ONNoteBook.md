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

### Name (Default)
```
Get-ONNoteBook -NoteBook <String[]> [<CommonParameters>]
```

### Id
```
Get-ONNoteBook -Id <String[]> [<CommonParameters>]
```

## DESCRIPTION
Returns OneNote XML Schema based element representing a specific Notebook

## EXAMPLES

### EXAMPLE 1
```
Get-ONNoteBook -NoteBook 'My NoteBook'
```

## PARAMETERS

### -Id
Notebook Id?

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

### -NoteBook
The NoteBook name to query.
Just one.

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

## OUTPUTS

## NOTES

## RELATED LINKS
