---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Get-ONSections

## SYNOPSIS
Gets OneNote Sections

## SYNTAX

### All (Default)
```
Get-ONSections [<CommonParameters>]
```

### NotebookName
```
Get-ONSections -NoteBookName <String> [<CommonParameters>]
```

### NotebookId
```
Get-ONSections -NoteBookId <String> [<CommonParameters>]
```

## DESCRIPTION
Returns OneNote XML Schema based elements representing Sections.
By default, all Sections from all Notebooks are returned and can be
filtered using standard cmdlets like Where-Object. As an
alternative you can specify the names or IDs of the Notebook hosting the Sections.


## EXAMPLES

### EXAMPLE 1
```
Get-ONSections
```

## PARAMETERS

### -NoteBookId
The ID of the Notebook hosting the Sections.

```yaml
Type: String
Parameter Sets: NotebookId
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoteBookName
The name of the Notebook hosting the Sections.

```yaml
Type: String
Parameter Sets: NotebookName
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS
