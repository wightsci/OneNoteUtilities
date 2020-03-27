---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Get-ONSectionGroups

## SYNOPSIS
Gets OneNote Section Groups.

## SYNTAX

### All (Default)
```
Get-ONSectionGroups [<CommonParameters>]
```

### NotebookName
```
Get-ONSectionGroups -NoteBookName <String> [<CommonParameters>]
```

### NotebookId
```
Get-ONSectionGroups -NoteBookId <String> [<CommonParameters>]
```

## DESCRIPTION
Returns OneNote XML Schema based elements representing Section Groups.

## EXAMPLES

### Example 1
```powershell
PS C:\> Get-SectionGroups
```

This command returns all Section Groups in all Notebooks.

## PARAMETERS

### -NoteBookId
The ID of the Notebook containing the Section Groups.

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
The name of the Notebook containg the Section Groups.

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

### None

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
