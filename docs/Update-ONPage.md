---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Update-ONPage

## SYNOPSIS
Updates an existing OneNote page.

## SYNTAX

```
Update-ONPage [-PageContent] <Object> [<CommonParameters>]
```

## DESCRIPTION
Updates a OneNote page using the currently in-use schema.
The cmdlet automatically checks if the object passed to the cmdlet
is an XmlElement. If so, the OuterXML property is used.

## EXAMPLES

### EXAMPLE 1
```
Update-ONPage $myPage.OuterXML
```

In this example the OuterXML property of a OneNote XML page object is
passed to the Update-ONPage cmdlet.

### EXAMPLE 2
```
Update-ONPage $myPage
```

In this example a OneNote XML page object is passed to the Update-ONPage cmdlet.
The cmdlet automatically extracts the OuterXML property.

### EXAMPLE 3
```
$myPageXML = Get-ONPage -Page 'MyPage' | Select-Object  OuterXML
Update-ONPage $myPage
```

In this example a OneNote XML page's OuterXML property is passed to the Update-ONPage cmdlet.


## PARAMETERS

### -PageContent
An xml string containing the updated page content

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 0
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
