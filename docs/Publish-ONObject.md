---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Publish-ONObject

## SYNOPSIS
Publishes a OneNote page in an external file format.

## SYNTAX

```
Publish-ONObject [-Id] <String[]> [-Format] <String[]> [-Path] <String[]> [<CommonParameters>]
```

## DESCRIPTION
Publishes a OneNote page. Available formats are:

- MHTML files (.mht) - (OneNote 2013 or newer)
- Adobe Acrobat PDF files (.pdf)
- XML Paper Specification (XPS) files (.xps)
- OneNote Package files (.onepkg)
- Microsoft Word documents (.doc or .docx)
- Microsoft Windows Enhanced Metafiles (.emf)
- HTML files (.html)

## EXAMPLES

### Example 1
```powershell
PS C:\> Publish-ONObject -Id '{C19F5E9B-C37B-0C25-0FC7-55FCE4E36F7B}{26}{E18372038253285132566191417462735909894706105}' -Format PDF -Path C:\Users\User\Desktop\Chapter1.pdf
```

This command creates a PDF version of the specified page at the specified location.

## PARAMETERS

### -Format
One of the valid publishing file formats.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: Type
Accepted values: PDF, XPS, DOC, EMF, ONEPKG, MHT, HTML

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Id
The Id of the OneNote page to be published.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: Identity

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Path
The full path of the file to be created.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: FilePath

Required: True
Position: 2
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
