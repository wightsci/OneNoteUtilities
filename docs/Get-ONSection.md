---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Get-ONSection

## SYNOPSIS
Gets OneNote Section

## SYNTAX

### Name (Default)
```
Get-ONSection -Section <String[]> [<CommonParameters>]
```

### Id
```
Get-ONSection -Id <String[]> [<CommonParameters>]
```

## DESCRIPTION
Returns OneNote XML Schema based elements representing a Section

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Id
Section Id?

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

### -Section
The Section name to query.
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
