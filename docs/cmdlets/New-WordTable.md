---
external help file: PSWordXml-help.xml
Module Name: PSWordXml
online version:
schema: 2.0.0
---

# New-WordTable

## SYNOPSIS
{{Fill in the Synopsis}}

## SYNTAX

### row (Default)
```
New-WordTable [-Row] <XmlElement> [-Headers] <String[]> [-ColumnWidths <Int32[]>] [-TableWidth <Int32>]
 [<CommonParameters>]
```

### array
```
New-WordTable [-DataArray] <PSObject> [-Headers] <String[]> [-ColumnWidths <Int32[]>] [-TableWidth <Int32>]
 [<CommonParameters>]
```

## DESCRIPTION
{{Fill in the Description}}

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -ColumnWidths
{{Fill ColumnWidths Description}}

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataArray
{{Fill DataArray Description}}

```yaml
Type: PSObject
Parameter Sets: array
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Headers
{{Fill Headers Description}}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
{{Fill Row Description}}

```yaml
Type: XmlElement
Parameter Sets: row
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -TableWidth
{{Fill TableWidth Description}}

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable.
For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.Management.Automation.PSObject
### System.Xml.XmlElement
## OUTPUTS

### System.Object
## NOTES

## RELATED LINKS
