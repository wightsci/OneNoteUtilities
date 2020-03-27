---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Get-ONPages

## SYNOPSIS
Gets OneNote Pages

## SYNTAX

### All (Default)
```
Get-ONPages [<CommonParameters>]
```

### NotebookName
```
Get-ONPages -NoteBookName <String> [<CommonParameters>]
```

### NotebookId
```
Get-ONPages -NoteBookId <String> [<CommonParameters>]
```

### SectionName
```
Get-ONPages -SectionName <String> [<CommonParameters>]
```

### SectionId
```
Get-ONPages -SectionId <String> [<CommonParameters>]
```

## DESCRIPTION
Returns OneNote XML Schema based elements representing Pages

## EXAMPLES

### EXAMPLE 1
```
Get-ONPages
```

## PARAMETERS

### -NoteBookId
{{ Fill NoteBookId Description }}

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
{{ Fill NoteBookName Description }}

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

### -SectionId
{{ Fill SectionId Description }}

```yaml
Type: String
Parameter Sets: SectionId
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SectionName
{{ Fill SectionName Description }}

```yaml
Type: String
Parameter Sets: SectionName
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
