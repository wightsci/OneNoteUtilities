---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Add-ONElement

## SYNOPSIS
Adds a OneNote XML Schema element as a child of another

## SYNTAX

```
Add-ONElement [-Element] <Object> [-Parent] <Object> [<CommonParameters>]
```

## DESCRIPTION
Adds an already created OneNote XML Schema element as a child of another - using the XML DOM AppendChild method.
Note that no explicit checking of validity of the resulting XML is undertaken.

## EXAMPLES

### EXAMPLE 1
```
$myPage = Get-ONPage -Page 'Amazon.co.uk - Stuart'
```

$myOutline = New-ONElement -Element "Outline" -Document $myPage
$myOEChildren  = New-ONElement -Element "OEChildren" -Document $myPage
$myOE = New-ONElement -Element "OE" -Document $myPage
$myT = New-ONElement -Element "T" -Document $myPage
$myT.InnerText = "Hello There yyyxxxxxyyy !"
Add-ONElement -Element $myT -Parent $myOE
Add-ONElement -Element $myOE -Parent $myOEChildren
Add-ONElement -Element $myOEChildren -Parent $myOutline
Add-ONElement -Element $myOutLine -Parent $myPage

## PARAMETERS

### -Element
{{ Fill Element Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Parent
{{ Fill Parent Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
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
