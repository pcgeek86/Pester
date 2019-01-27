function GherkinStep {
    param(
        [Parameter(Mandatory = $True, Position = 0)]
        [String]$Name,

        [Parameter(Mandatory = $True, Position = 1)]
        [ScriptBlock]$Test
    )
    # We need to be able to look up where this step is defined
    $Definition = (& $SafeCommands["Get-PSCallStack"])[1]
    $RelativePath = & $SafeCommands["Resolve-Path"] $Definition.ScriptName -relative
    $Source = "{0}: line {1}" -f $RelativePath, $Definition.ScriptLineNumber

    $Script:GherkinSteps.${Name} = $Test | & $SafeCommands["Add-Member"] -MemberType NoteProperty -Name Source -Value $Source -PassThru
}

Set-Alias Given GherkinStep
Set-Alias When GherkinStep
Set-Alias Then GherkinStep
Set-Alias And GherkinStep
Set-Alias But GherkinStep
