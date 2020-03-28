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

### Example 2
```powershell
PS C:\> Get-ONSectionGroups -NoteBookName 'WebNotes'

name             : First Section Group
ID               : {1F3C5AB9-BEDB-49AE-8FFE-C0EEB19817D5}{1}{B0}
path             : https://d.docs.live.net/816f7725bef99999/WebNotes/First Section Group/
lastModifiedTime : 2020-03-28T11:44:42.000Z
Section          : Section
```

This command returns all Section Groups in the specified Notebook

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
