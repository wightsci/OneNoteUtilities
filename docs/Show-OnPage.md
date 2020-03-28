---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Show-OnPage

## SYNOPSIS
Displays a page in the OneNote user interface.

## SYNTAX

### Name (Default)
```
Show-OnPage -Name <String> [-NewWindow] [<CommonParameters>]
```

### Id
```
Show-OnPage -Id <String> [-NewWindow] [<CommonParameters>]
```

## DESCRIPTION
Displays a page in the OneNote user interface.

## EXAMPLES

### Example 1
```powershell
PS C:\> Show-ONPage -Id '{C19F5E9B-C37B-0C25-0FC7-55FCE4E36F7B}{26}{E18372038253285132566191417462735909894706105}'
```

This command displays the specified page in OneNote.

## PARAMETERS

### -Id
Page Id?

```yaml
Type: String
Parameter Sets: Id
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Name
Page Name?

```yaml
Type: String
Parameter Sets: Name
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -NewWindow
Displays the Page in a new OneNote window, instead of re-using an existing window if one exists.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String[]

## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
