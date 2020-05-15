---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Get-ONNoteBook

## SYNOPSIS
Gets one or more OneNote Notebooks.

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
Returns one or more OneNote XML Schema based elements representing specific Notebooks.

## EXAMPLES

### EXAMPLE 1
```
Get-ONNoteBook -NoteBook 'My NoteBook'
```

## PARAMETERS

### -Id
The Notebook ID to query.

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
The Notebook name to query.

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
