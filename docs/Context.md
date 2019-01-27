---
external help file: Pester-help.xml
Module Name: Pester
online version:
schema: 2.0.0
---

# Context

## SYNOPSIS

Provides logical grouping of It blocks within a single Describe block.

## SYNTAX

```
Context [-Name] <String> [-Tag <String[]>] [[-Fixture] <ScriptBlock>] [<CommonParameters>]
```

## DESCRIPTION

Provides logical grouping of It blocks within a single Describe block.
Any Mocks defined inside a Context are removed at the end of the Context scope,
as are any files or folders added to the TestDrive during the Context block's
execution.
Any BeforeEach or AfterEach blocks defined inside a Context also only
apply to tests within that Context .

## EXAMPLES

### EXAMPLE 1

```
function Add-Numbers($a, $b) {
```

return $a + $b
}

Describe "Add-Numbers" {

    Context "when root does not exist" {
         It "..." { ...

}
}

    Context "when root does exist" {
        It "..." { ...

}
It "..." { ...
}
It "..." { ...
}
}
}

## PARAMETERS

### -Fixture

Script that is executed.
This may include setup specific to the context
and one or more It blocks that validate the expected outcomes.

```yaml
Type: ScriptBlock
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: $(Throw "No test script block is provided. (Have you put the open curly brace on the next line?)")
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name

The name of the Context.
This is a phrase describing a set of tests within a describe.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Tag

Optional parameter containing an array of strings.
When calling Invoke-Pester,
it is possible to specify a -Tag parameter which will only execute Context blocks
containing the same Tag.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: Tags

Required: False
Position: Named
Default value: @()
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

[Describe](Describe.md)

[It](It.md)

[AfterEach](AfterEach.md)

[BeforeEach](BeforeEach.md)

[about_Should](about_Should.md)

[about_Mocking](about_Mocking.md)

[about_TestDrive](about_TestDrive.md)
