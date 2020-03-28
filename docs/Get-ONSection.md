---
external help file: OneNoteUtilities-help.xml
Module Name: OneNoteUtilities
online version:
schema: 2.0.0
---

# Get-ONSection

## SYNOPSIS
Gets one or more OneNote Sections

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
Returns OneNote XML Schema based elements representing one or more Sections.

## EXAMPLES

### Example 1
```powershell
PS C:\> Get-ONSection -Section 'Teacher Notes','Administration Notes'

name             : Teacher Notes
ID               : {1DE00420-3700-037A-3F54-97489F626533}{20}{B0}
path             : https://d.docs.live.net/816f7725bef99999/Documents/Real World Samples/Teacher Notes.one
lastModifiedTime : 2020-03-25T19:31:50.000Z
color            : #9595AA
isUnread         : true
Page             : {7th Grade Math, 1.10 - The Coordinate Plane, 1.2 - Variables and Expressions, Problem-Solving...}

name             : Administration Notes
ID               : {BACECC72-6805-0656-22E9-319A25A5247A}{22}{B0}
path             : https://d.docs.live.net/816f7725bef99999/Documents/Real World Samples/Administration Notes.one
lastModifiedTime : 2019-01-30T11:08:41.000Z
color            : #B7C997
Page             : {Staff and Faculty Notes examples, Calendars, Schedule and Academic Calendar, Fall Sports...}
```

This command returns the two Sections specified by name.

### Example 2
```powershell
PS C:\> Get-ONSection -Id '{70863F49-36D4-0CB0-1CBD-AF8C35E05883}{29}{B0}'

name             : Lesson Plans
ID               : {70863F49-36D4-0CB0-1CBD-AF8C35E05883}{29}{B0}
path             : https://d.docs.live.net/816f7725bef99999/Documents/Real World Samples/Lesson Plans.one
lastModifiedTime : 2020-03-25T19:31:49.000Z
color            : #D5A4BB
Page             : {3rd Grade Math, Egyptian numbers - Handout #2, Number System in China - Handout #8, Trace the Graph...}
```

This command returns the Section specified by the Id.

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
