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

This command returns all Sections in all Notebooks

### EXAMPLE 2
```
Get-ONSections -NoteBookName 'Real World Samples'

name             : Student pages
ID               : {C19F5E9B-C37B-0C25-0FC7-55FCE4E36F7B}{26}{B0}
path             : https://d.docs.live.net/816f7725bef00a5f/Documents/Real World Samples/Student pages.one
lastModifiedTime : 2019-07-16T15:03:46.000Z
color            : #8AA8E4
Page             : {Cincinnati Country Day School samples, Graphing/Color Coding , Forces, Gravity, and Newton's Laws of M6otion, La Religión...}

name             : Lesson Plans
ID               : {70863F49-36D4-0CB0-1CBD-AF8C35E05883}{29}{B0}
path             : https://d.docs.live.net/816f7725bef00a5f/Documents/Real World Samples/Lesson Plans.one
lastModifiedTime : 2020-03-25T19:31:49.000Z
color            : #D5A4BB
Page             : {3rd Grade Math, Egyptian numbers - Handout #2, Number System in China - Handout #8, Trace the Graph...}
...
```

This command returns all of the sections from the named Notebook.

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
