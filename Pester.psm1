if ($PSVersionTable.PSVersion.Major -ge 3) {
    $script:IgnoreErrorPreference = 'Ignore'
    $outNullModule = 'Microsoft.PowerShell.Core'
    $outHostModule = 'Microsoft.PowerShell.Core'
}
else {
    $script:IgnoreErrorPreference = 'SilentlyContinue'
    $outNullModule = 'Microsoft.PowerShell.Utility'
    $outHostModule = $null
}

# Tried using $ExecutionState.InvokeCommand.GetCmdlet() here, but it does not trigger module auto-loading the way
# Get-Command does.  Since this is at import time, before any mocks have been defined, that's probably acceptable.
# If someone monkeys with Get-Command before they import Pester, they may break something.

# The -All parameter is required when calling Get-Command to ensure that PowerShell can find the command it is
# looking for. Otherwise, if you have modules loaded that define proxy cmdlets or that have cmdlets with the same
# name as the safe cmdlets, Get-Command will return null.
$safeCommandLookupParameters = @{
    CommandType = [System.Management.Automation.CommandTypes]::Cmdlet
    ErrorAction = [System.Management.Automation.ActionPreference]::Stop
}

if ($PSVersionTable.PSVersion.Major -gt 2) {
    $safeCommandLookupParameters['All'] = $true
}

$script:SafeCommands = @{
    'Add-Member'           = Get-Command -Name Add-Member           -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Add-Type'             = Get-Command -Name Add-Type             -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Compare-Object'       = Get-Command -Name Compare-Object       -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Export-ModuleMember'  = Get-Command -Name Export-ModuleMember  -Module Microsoft.PowerShell.Core       @safeCommandLookupParameters
    'ForEach-Object'       = Get-Command -Name ForEach-Object       -Module Microsoft.PowerShell.Core       @safeCommandLookupParameters
    'Format-Table'         = Get-Command -Name Format-Table         -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Get-Alias'            = Get-Command -Name Get-Alias            -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Get-ChildItem'        = Get-Command -Name Get-ChildItem        -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Get-Command'          = Get-Command -Name Get-Command          -Module Microsoft.PowerShell.Core       @safeCommandLookupParameters
    'Get-Content'          = Get-Command -Name Get-Content          -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Get-Date'             = Get-Command -Name Get-Date             -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Get-Item'             = Get-Command -Name Get-Item             -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Get-ItemProperty'     = Get-Command -Name Get-ItemProperty     -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Get-Location'         = Get-Command -Name Get-Location         -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Get-Member'           = Get-Command -Name Get-Member           -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Get-Module'           = Get-Command -Name Get-Module           -Module Microsoft.PowerShell.Core       @safeCommandLookupParameters
    'Get-PSDrive'          = Get-Command -Name Get-PSDrive          -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Get-PSCallStack'      = Get-Command -Name Get-PSCallStack      -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Get-Unique'           = Get-Command -Name Get-Unique           -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Get-Variable'         = Get-Command -Name Get-Variable         -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Group-Object'         = Get-Command -Name Group-Object         -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Import-LocalizedData' = Get-Command -Name Import-LocalizedData -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Import-Module'        = Get-Command -Name Import-Module        -Module Microsoft.PowerShell.Core       @safeCommandLookupParameters
    'Join-Path'            = Get-Command -Name Join-Path            -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Measure-Object'       = Get-Command -Name Measure-Object       -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'New-Item'             = Get-Command -Name New-Item             -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'New-ItemProperty'     = Get-Command -Name New-ItemProperty     -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'New-Module'           = Get-Command -Name New-Module           -Module Microsoft.PowerShell.Core       @safeCommandLookupParameters
    'New-Object'           = Get-Command -Name New-Object           -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'New-PSDrive'          = Get-Command -Name New-PSDrive          -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'New-Variable'         = Get-Command -Name New-Variable         -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Out-Host'             = Get-Command -Name Out-Host             -Module $outHostModule                  @safeCommandLookupParameters
    'Out-File'             = Get-Command -Name Out-File             -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Out-Null'             = Get-Command -Name Out-Null             -Module $outNullModule                  @safeCommandLookupParameters
    'Out-String'           = Get-Command -Name Out-String           -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Pop-Location'         = Get-Command -Name Pop-Location         -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Push-Location'        = Get-Command -Name Push-Location        -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Remove-Item'          = Get-Command -Name Remove-Item          -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Remove-PSBreakpoint'  = Get-Command -Name Remove-PSBreakpoint  -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Remove-PSDrive'       = Get-Command -Name Remove-PSDrive       -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Remove-Variable'      = Get-Command -Name Remove-Variable      -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Resolve-Path'         = Get-Command -Name Resolve-Path         -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Select-Object'        = Get-Command -Name Select-Object        -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Set-Content'          = Get-Command -Name Set-Content          -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Set-Location'         = Get-Command -Name Set-Location         -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Set-PSBreakpoint'     = Get-Command -Name Set-PSBreakpoint     -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Set-StrictMode'       = Get-Command -Name Set-StrictMode       -Module Microsoft.PowerShell.Core       @safeCommandLookupParameters
    'Set-Variable'         = Get-Command -Name Set-Variable         -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Sort-Object'          = Get-Command -Name Sort-Object          -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Split-Path'           = Get-Command -Name Split-Path           -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Start-Sleep'          = Get-Command -Name Start-Sleep          -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Test-Path'            = Get-Command -Name Test-Path            -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
    'Where-Object'         = Get-Command -Name Where-Object         -Module Microsoft.PowerShell.Core       @safeCommandLookupParameters
    'Write-Error'          = Get-Command -Name Write-Error          -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Write-Host'           = Get-Command -Name Write-Host           -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Write-Progress'       = Get-Command -Name Write-Progress       -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Write-Verbose'        = Get-Command -Name Write-Verbose        -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
    'Write-Warning'        = Get-Command -Name Write-Warning        -Module Microsoft.PowerShell.Utility    @safeCommandLookupParameters
}

# Not all platforms have Get-WmiObject (Nano or PSCore 6.0.0-beta.x on Linux)
# Get-CimInstance is preferred, but we can use Get-WmiObject if it exists
# Moreover, it shouldn't really be fatal if neither of those cmdlets
# exist
if ( Get-Command -ea SilentlyContinue Get-CimInstance ) {
    $script:SafeCommands['Get-CimInstance'] = Get-Command -Name Get-CimInstance -Module CimCmdlets @safeCommandLookupParameters
}
elseif ( Get-command -ea SilentlyContinue Get-WmiObject ) {
    $script:SafeCommands['Get-WmiObject'] = Get-Command -Name Get-WmiObject   -Module Microsoft.PowerShell.Management @safeCommandLookupParameters
}
elseif ( Get-Command -ea SilentlyContinue uname -Type Application ) {
    $script:SafeCommands['uname'] = Get-Command -Name uname -Type Application | Select-Object -First 1
    if ( Get-Command -ea SilentlyContinue id -Type Application ) {
        $script:SafeCommands['id'] = Get-Command -Name id -Type Application | Select-Object -First 1
    }
}
else {
    Write-Warning "OS Information retrieval is not possible, reports will contain only partial system data"
}

# little sanity check to make sure we don't blow up a system with a typo up there
# (not that I've EVER done that by, for example, mapping New-Item to Remove-Item...)

foreach ($keyValuePair in $script:SafeCommands.GetEnumerator()) {
    if ($keyValuePair.Key -ne $keyValuePair.Value.Name) {
        throw "SafeCommands entry for $($keyValuePair.Key) does not hold a reference to the proper command."
    }
}

$script:AssertionOperators = & $SafeCommands['New-Object'] 'Collections.Generic.Dictionary[string,object]'([StringComparer]::InvariantCultureIgnoreCase)
$script:AssertionAliases = & $SafeCommands['New-Object'] 'Collections.Generic.Dictionary[string,object]'([StringComparer]::InvariantCultureIgnoreCase)
$script:AssertionDynamicParams = & $SafeCommands['New-Object'] System.Management.Automation.RuntimeDefinedParameterDictionary
$script:DisableScopeHints = $true

function Count-Scopes {
    param(
        [Parameter(Mandatory = $true)]
        $ScriptBlock
    )

    if ($script:DisableScopeHints) {
        return 0
    }

    # automatic variable that can help us count scopes must be constant a must not be all scopes
    # from the standard ones only Error seems to be that, let's ensure it is like that everywhere run
    # other candidate variables can be found by this code
    # Get-Variable  | where { -not ($_.Options -band [Management.Automation.ScopedItemOptions]"AllScope") -and $_.Options -band $_.Options -band [Management.Automation.ScopedItemOptions]"Constant" }

    # get-variable steps on it's toes and recurses when we mock it in a test
    # and we are also invoking this in user scope so we need to pass the reference
    # to the safely captured function in the user scope
    $safeGetVariable = $script:SafeCommands['Get-Variable']
    $sb = {
        param($safeGetVariable)
        $err = (& $safeGetVariable -Name Error).Options
        if ($err -band "AllScope" -or (-not ($err -band "Constant"))) {
            throw "Error variable is set to AllScope, or is not marked as constant cannot use it to count scopes on this platform."
        }

        $scope = 0
        while ($null -eq (& $safeGetVariable -Name Error -Scope $scope -ErrorAction SilentlyContinue)) {
            $scope++
        }

        $scope - 1 # because we are in a function
    }

    $flags = [System.Reflection.BindingFlags]'Instance,NonPublic'
    $property = [scriptblock].GetProperty('SessionStateInternal', $flags)
    $ssi = $property.GetValue($ScriptBlock, $null)
    $property.SetValue($sb, $ssi, $null)

    &$sb $safeGetVariable
}

function Write-ScriptBlockInvocationHint {
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $ScriptBlock,

        [Parameter(Mandatory = $true)]
        [String]
        $Hint
    )

    if ($script:DisableScopeHints) {
        return
    }

    $scope = Get-ScriptBlockHint $ScriptBlock
    $count = Count-Scopes -ScriptBlock $ScriptBlock

    Write-Hint "Invoking scriptblock from location '$Hint' in state '$scope', $count scopes deep:
    {
        $ScriptBlock
    }`n`n"
}

function Write-Hint {
    param (
        $Hint
    )

    if ($script:DisableScopeHints) {
        return
    }

    Write-Host -ForegroundColor Cyan $Hint
}

function Test-Hint {
    param (
        [Parameter(Mandatory = $true)]
        $InputObject
    )

    if ($script:DisableScopeHints) {
        return $true
    }

    $property = $InputObject | Get-Member -Name Hint -MemberType NoteProperty
    if ($null -eq $property) {
        return $false
    }

    Test-NullOrWhiteSpace $property.Value
}

function Set-Hint {
    param(
        [Parameter(Mandatory = $true)]
        [String] $Hint,

        [Parameter(Mandatory = $true)]
        $InputObject,

        [Switch] $Force
    )

    if ($script:DisableScopeHints) {
        return
    }

    if ($InputObject | Get-Member -Name Hint -MemberType NoteProperty) {
        $hintIsNotSet = Test-NullOrWhiteSpace $InputObject.Hint
        if ($Force -or $hintIsNotSet) {
            $InputObject.Hint = $Hint
        }
    }
    else {
        # do not change this to be called without the pipeline, it will throw: Cannot evaluate parameter 'InputObject' because its argument is specified as a script block and there is no input. A script block cannot be evaluated without input.
        $InputObject | Add-Member -Name Hint -Value $Hint -MemberType NoteProperty
    }
}

function Set-SessionStateHint {
    param(
        [Parameter(Mandatory = $true)]
        [String] $Hint,

        [Parameter(Mandatory = $true)]
        [Management.Automation.SessionState] $SessionState,

        [Switch] $PassThru
    )

    if ($script:DisableScopeHints) {
        if ($PassThru) {
            return $SessionState
        }
        return
    }

    # in all places where we capture SessionState we mark its internal state with a hint
    # the internal state does not change and we use it to invoke scriptblock in diferent
    # states, setting the hint on SessionState is only secondary to make is easier to debug
    $flags = [System.Reflection.BindingFlags]'Instance,NonPublic'
    $internalSessionState = $SessionState.GetType().GetProperty('Internal', $flags).GetValue($SessionState, $null)
    if ($null -eq $internalSessionState) {
        throw "SessionState does not have any internal SessionState, this should never happen."
    }

    $hashcode = $internalSessionState.GetHashCode()
    # optionally sets the hint if there was none, so the hint from the
    # function that first captured this session state is preserved
    Set-Hint -Hint "$Hint ($hashcode))" -InputObject $internalSessionState
    # the public session state should always depend on the internal state
    Set-Hint -Hint $internalSessionState.Hint -InputObject $SessionState -Force

    if ($PassThru) {
        $SessionState
    }
}

function Get-SessionStateHint {
    param(
        [Parameter(Mandatory = $true)]
        [Management.Automation.SessionState] $SessionState
    )

    if ($script:DisableScopeHints) {
        return
    }

    # the hint is also attached to the session state object, but sessionstate objects are recreated while
    # the internal state stays static so to see the hint on object that we receive via $PSCmdlet.SessionState we need
    # to look at the InternalSessionState. the internal state should be never null so just looking there is enough
    $flags = [System.Reflection.BindingFlags]'Instance,NonPublic'
    $internalSessionState = $SessionState.GetType().GetProperty('Internal', $flags).GetValue($SessionState, $null)
    if (Test-Hint $internalSessionState) {
        $internalSessionState.Hint
    }
}

function Set-ScriptBlockHint {
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $ScriptBlock,

        [string] $Hint
    )

    if ($script:DisableScopeHints) {
        return
    }

    $flags = [System.Reflection.BindingFlags]'Instance,NonPublic'
    $internalSessionState = $ScriptBlock.GetType().GetProperty('SessionStateInternal', $flags).GetValue($ScriptBlock, $null)
    if ($null -eq $internalSessionState) {
        if (Test-Hint -InputObject $ScriptBlock) {
            # the scriptblock already has a hint and there is not internal state
            # so the hint on the scriptblock is enough
            # if there was an internal state we would try to copy the hint from it
            # onto the scriptblock to keep them in sync
            return
        }

        if ($null -eq $Hint) {
            throw "Cannot set ScriptBlock hint because it is unbound ScriptBlock (with null internal state) and no -Hint was provided."
        }

        # adds hint on the ScriptBlock
        # the internal session state is null so we must attach the hint directly
        # on the scriptblock
        Set-Hint -Hint "$Hint (Unbound)" -InputObject $ScriptBlock -Force
    }
    else {
        if (Test-Hint -InputObject $internalSessionState) {
            # there already is hint on the internal state, we take it and sync
            # it with the hint on the object
            Set-Hint -Hint $internalSessionState.Hint -InputObject $ScriptBlock -Force
            return
        }

        if ($null -eq $Hint) {
            throw "Cannot set ScriptBlock hint because it's internal state does not have any Hint and no external -Hint was provided."
        }

        $hashcode = $internalSessionState.GetHashCode()
        $Hint = "$Hint - ($hashCode)"
        Set-Hint -Hint $Hint -InputObject $internalSessionState -Force
        Set-Hint -Hint $Hint -InputObject $ScriptBlock -Force
    }
}

function Get-ScriptBlockHint {
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $ScriptBlock
    )

    if ($script:DisableScopeHints) {
        return
    }

    # the hint is also attached to the scriptblock object, but not all scriptblocks are tagged by us,
    # the internal state stays static so to see the hint on object that we receive we need to look at the InternalSessionState
    $flags = [System.Reflection.BindingFlags]'Instance,NonPublic'
    $internalSessionState = $ScriptBlock.GetType().GetProperty('SessionStateInternal', $flags).GetValue($ScriptBlock, $null)


    if ($null -ne $internalSessionState -and (Test-Hint $internalSessionState)) {
        return $internalSessionState.Hint
    }

    if (Test-Hint $ScriptBlock) {
        return $ScriptBlock.Hint
    }

    "Unknown unbound ScriptBlock"
}


function Test-NullOrWhiteSpace {
    param (
        [string]$String
    )

    $String -match "^\s*$"
}

function Assert-ValidAssertionName {
    param(
        [string]$Name
    )

    if ($Name -notmatch '^\S+$') {
        throw "Assertion name '$name' is invalid, assertion name must be a single word."
    }
}

function Assert-ValidAssertionAlias {
    param([string[]]$Alias)
    if ($Alias -notmatch '^\S+$') {
        throw "Assertion alias '$string' is invalid, assertion alias must be a single word."
    }
}

function Add-AssertionOperator {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Name,

        [Parameter(Mandatory = $true)]
        [scriptblock] $Test,

        [ValidateNotNullOrEmpty()]
        [AllowEmptyCollection()]
        [string[]] $Alias = @(),

        [Parameter()]
        [string] $InternalName,

        [switch] $SupportsArrayInput
    )

    $entry = New-Object psobject -Property @{
        Test               = $Test
        SupportsArrayInput = [bool]$SupportsArrayInput
        Name               = $Name
        Alias              = $Alias
        InternalName       = If ($InternalName) {
            $InternalName
        }
        Else {
            $Name
        }
    }
    if (Test-AssertionOperatorIsDuplicate -Operator $entry) {
        # This is an exact duplicate of an existing assertion operator.
        return
    }

    $namesToCheck = @(
        $Name
        $Alias
    )

    Assert-AssertionOperatorNameIsUnique -Name $namesToCheck

    $script:AssertionOperators[$Name] = $entry

    foreach ($string in $Alias | Where { -not (Test-NullOrWhiteSpace $_)}) {
        Assert-ValidAssertionAlias -Alias $string
        $script:AssertionAliases[$string] = $Name
    }

    Add-AssertionDynamicParameterSet -AssertionEntry $entry
}

function Test-AssertionOperatorIsDuplicate {
    param (
        [psobject] $Operator
    )

    $existing = $script:AssertionOperators[$Operator.Name]
    if (-not $existing) {
        return $false
    }

    return $Operator.SupportsArrayInput -eq $existing.SupportsArrayInput -and
    $Operator.Test.ToString() -eq $existing.Test.ToString() -and
    -not (Compare-Object $Operator.Alias $existing.Alias)
}
function Assert-AssertionOperatorNameIsUnique {
    param (
        [string[]] $Name
    )

    foreach ($string in $name | Where { -not (Test-NullOrWhiteSpace $_)}) {
        Assert-ValidAssertionName -Name $string

        if ($script:AssertionOperators.ContainsKey($string)) {
            throw "Assertion operator name '$string' has been added multiple times."
        }

        if ($script:AssertionAliases.ContainsKey($string)) {
            throw "Assertion operator name '$string' already exists as an alias for operator '$($script:AssertionAliases[$key])'"
        }
    }
}

function Add-AssertionDynamicParameterSet {
    param (
        [object] $AssertionEntry
    )

    ${function:__AssertionTest__} = $AssertionEntry.Test
    $commandInfo = Get-Command __AssertionTest__ -CommandType Function
    $metadata = [System.Management.Automation.CommandMetadata]$commandInfo

    $attribute = New-Object Management.Automation.ParameterAttribute
    $attribute.ParameterSetName = $AssertionEntry.Name
    $attribute.Mandatory = $true

    $attributeCollection = New-Object Collections.ObjectModel.Collection[Attribute]
    $null = $attributeCollection.Add($attribute)
    if (-not (Test-NullOrWhiteSpace $AssertionEntry.Alias)) {
        Assert-ValidAssertionAlias -Alias $AssertionEntry.Alias
        $attribute = New-Object System.Management.Automation.AliasAttribute($AssertionEntry.Alias)
        $attributeCollection.Add($attribute)
    }

    $dynamic = New-Object System.Management.Automation.RuntimeDefinedParameter($AssertionEntry.Name, [switch], $attributeCollection)
    $null = $script:AssertionDynamicParams.Add($AssertionEntry.Name, $dynamic)

    if ($script:AssertionDynamicParams.ContainsKey('Not')) {
        $dynamic = $script:AssertionDynamicParams['Not']
    }
    else {
        $dynamic = New-Object System.Management.Automation.RuntimeDefinedParameter('Not', [switch], (New-Object System.Collections.ObjectModel.Collection[Attribute]))
        $null = $script:AssertionDynamicParams.Add('Not', $dynamic)
    }

    $attribute = New-Object System.Management.Automation.ParameterAttribute
    $attribute.ParameterSetName = $AssertionEntry.Name
    $attribute.Mandatory = $false
    $null = $dynamic.Attributes.Add($attribute)

    $i = 1
    foreach ($parameter in $metadata.Parameters.Values) {
        # common parameters that are already defined
        if ($parameter.Name -eq 'ActualValue' -or $parameter.Name -eq 'Not' -or $parameter.Name -eq 'Negate') {
            continue
        }


        if ($script:AssertionOperators.ContainsKey($parameter.Name) -or $script:AssertionAliases.ContainsKey($parameter.Name)) {
            throw "Test block for assertion operator $($AssertionEntry.Name) contains a parameter named $($parameter.Name), which conflicts with another assertion operator's name or alias."
        }

        foreach ($alias in $parameter.Aliases) {
            if ($script:AssertionOperators.ContainsKey($alias) -or $script:AssertionAliases.ContainsKey($alias)) {
                throw "Test block for assertion operator $($AssertionEntry.Name) contains a parameter named $($parameter.Name) with alias $alias, which conflicts with another assertion operator's name or alias."
            }
        }

        if ($script:AssertionDynamicParams.ContainsKey($parameter.Name)) {
            $dynamic = $script:AssertionDynamicParams[$parameter.Name]
        }
        else {
            # We deliberately use a type of [object] here to avoid conflicts between different assertion operators that may use the same parameter name.
            # We also don't bother to try to copy transformation / validation attributes here for the same reason.
            # Because we'll be passing these parameters on to the actual test function later, any errors will come out at that time.

            # few years later: using [object] causes problems with switch params (in my case -PassThru), because then we cannot use them without defining a value
            # so for switches we must prefer the conflicts over type
            if ([switch] -eq $parameter.ParameterType) {
                $type = [switch]
            }
            else {
                $type = [object]
            }

            $dynamic = New-Object System.Management.Automation.RuntimeDefinedParameter($parameter.Name, $type, (New-Object System.Collections.ObjectModel.Collection[Attribute]))
            $null = $script:AssertionDynamicParams.Add($parameter.Name, $dynamic)
        }

        $attribute = New-Object Management.Automation.ParameterAttribute
        $attribute.ParameterSetName = $AssertionEntry.Name
        $attribute.Mandatory = $false
        $attribute.Position = ($i++)

        $null = $dynamic.Attributes.Add($attribute)
    }
}

function Get-AssertionOperatorEntry {
    param (
        [string] $Name
    )

    return $script:AssertionOperators[$Name]
}

function Get-AssertionDynamicParams {
    return $script:AssertionDynamicParams
}

$Script:PesterRoot = & $SafeCommands['Split-Path'] -Path $MyInvocation.MyCommand.Path
"$PesterRoot\Functions\*.ps1", "$PesterRoot\Functions\Assertions\*.ps1" |
    & $script:SafeCommands['Resolve-Path'] |
    & $script:SafeCommands['Where-Object'] { -not ($_.ProviderPath.ToLower().Contains(".tests.")) } |
    & $script:SafeCommands['ForEach-Object'] { . $_.ProviderPath }

if (& $script:SafeCommands['Test-Path'] "$PesterRoot\Dependencies") {
    # sub-modules
    & $script:SafeCommands['Get-ChildItem'] "$PesterRoot\Dependencies\*\*.psm1" |
        & $script:SafeCommands['ForEach-Object'] { & $script:SafeCommands['Import-Module'] $_.FullName -Force -DisableNameChecking }
}

Add-Type -TypeDefinition @"
using System;

namespace Pester
{
    [Flags]
    public enum OutputTypes
    {
        None = 0,
        Default = 1,
        Passed = 2,
        Failed = 4,
        Pending = 8,
        Skipped = 16,
        Inconclusive = 32,
        Describe = 64,
        Context = 128,
        Summary = 256,
        Header = 512,
        All = Default | Passed | Failed | Pending | Skipped | Inconclusive | Describe | Context | Summary | Header,
        Fails = Default | Failed | Pending | Skipped | Inconclusive | Describe | Context | Summary | Header
    }
}
"@

function Has-Flag {
    param
    (
        [Parameter(Mandatory = $true)]
        [Pester.OutputTypes]
        $Setting,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Pester.OutputTypes]
        $Value
    )

    0 -ne ($Setting -band $Value)
}

function Invoke-Pester {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param(
        [Parameter(Position = 0, Mandatory = 0)]
        [Alias('Path', 'relative_path')]
        [object[]]$Script = '.',

        [Parameter(Position = 1, Mandatory = 0)]
        [Alias("Name")]
        [string[]]$TestName,

        [Parameter(Position = 2, Mandatory = 0)]
        [switch]$EnableExit,

        [Parameter(Position = 4, Mandatory = 0)]
        [Alias('Tags')]
        [string[]]$Tag,

        [string[]]$ExcludeTag,

        [switch]$PassThru,

        [object[]] $CodeCoverage = @(),

        [string] $CodeCoverageOutputFile,

        [ValidateSet('JaCoCo')]
        [String]$CodeCoverageOutputFileFormat = "JaCoCo",

        [Switch]$Strict,

        [Parameter(Mandatory = $true, ParameterSetName = 'NewOutputSet')]
        [string] $OutputFile,

        [Parameter(ParameterSetName = 'NewOutputSet')]
        [ValidateSet('NUnitXml')]
        [string] $OutputFormat = 'NUnitXml',

        [Switch]$Quiet,

        [object]$PesterOption,

        [Pester.OutputTypes]$Show = 'All'
    )
    begin {
        # Ensure when running Pester that we're using RSpec strings
        & $script:SafeCommands['Import-LocalizedData'] -BindingVariable Script:ReportStrings -BaseDirectory $PesterRoot -FileName RSpec.psd1 -ErrorAction SilentlyContinue

        # Fallback to en-US culture strings
        If ([String]::IsNullOrEmpty($ReportStrings)) {

            & $script:SafeCommands['Import-LocalizedData'] -BaseDirectory $PesterRoot -BindingVariable Script:ReportStrings -UICulture 'en-US' -FileName RSpec.psd1 -ErrorAction Stop

        }
    }

    end {
        if ($PSBoundParameters.ContainsKey('Quiet')) {
            & $script:SafeCommands['Write-Warning'] 'The -Quiet parameter has been deprecated; please use the new -Show parameter instead. To get no output use -Show None.'
            & $script:SafeCommands['Start-Sleep'] -Seconds 2

            if (!$PSBoundParameters.ContainsKey('Show')) {
                $Show = [Pester.OutputTypes]::None
            }
        }

        $script:mockTable = @{}
        Remove-MockFunctionsAndAliases
        $sessionState = Set-SessionStateHint -PassThru  -Hint "Caller - Captured in Invoke-Pester" -SessionState $PSCmdlet.SessionState
        $pester = New-PesterState -TestNameFilter $TestName -TagFilter $Tag -ExcludeTagFilter $ExcludeTag -SessionState $SessionState -Strict:$Strict -Show:$Show -PesterOption $PesterOption -RunningViaInvokePester

        try {
            Enter-CoverageAnalysis -CodeCoverage $CodeCoverage -PesterState $pester
            Write-PesterStart $pester $Script

            $invokeTestScript = {
                param (
                    [Parameter(Position = 0)]
                    [string] $Path,
                    [string] $Script,
                    [object[]] $Arguments = @(),
                    [System.Collections.IDictionary] $Parameters = @{}
                )

                if (-not [string]::IsNullOrEmpty($Path)) {
                    & $Path @Parameters @Arguments
                }
                elseif (-not [string]::IsNullOrEmpty($Script)) {
                    $scriptBlock = [scriptblock]::Create($Script)
                    Set-ScriptBlockHint -Hint "Unbound ScriptBlock from Invoke-Pester" -ScriptBlock $scriptBlock
                    Invoke-Command -ScriptBlock ($scriptBlock)
                }
            }

            Set-ScriptBlockScope -ScriptBlock $invokeTestScript -SessionState $sessionState
            $testScripts = @(ResolveTestScripts $Script)


            foreach ($testScript in $testScripts) {
                #Get test description for better output
                if (-not [string]::IsNullOrEmpty($testScript.Path)) {
                    $testDescription = $testScript.Path
                }
                elseif (-not [string]::IsNullOrEmpty($testScript.Script)) {
                    $testDescription = $testScript.Script
                }

                try {
                    $pester.EnterTestGroup($testDescription, 'Script')
                    Write-Describe $testDescription -CommandUsed Script
                    do {
                        $testOutput = & $invokeTestScript -Path $testScript.Path -Script $testScript.Script -Arguments $testScript.Arguments -Parameters $testScript.Parameters
                    } until ($true)
                }
                catch {
                    $firstStackTraceLine = $null
                    if (($_ | & $SafeCommands['Get-Member'] -Name ScriptStackTrace) -and $null -ne $_.ScriptStackTrace) {
                        $firstStackTraceLine = $_.ScriptStackTrace -split '\r?\n' | & $script:SafeCommands['Select-Object'] -First 1
                    }
                    $pester.AddTestResult("Error occurred in test script '$($testDescription)'", "Failed", $null, $_.Exception.Message, $firstStackTraceLine, $null, $null, $_)

                    # This is a hack to ensure that XML output is valid for now.  The test-suite names come from the Describe attribute of the TestResult
                    # objects, and a blank name is invalid NUnit XML.  This will go away when we promote test scripts to have their own test-suite nodes,
                    # planned for v4.0
                    $pester.TestResult[-1].Describe = "Error in $($testDescription)"

                    $pester.TestResult[-1] | Write-PesterResult
                }
                finally {
                    Exit-MockScope
                    $pester.LeaveTestGroup($testDescription, 'Script')
                }
            }

            $pester | Write-PesterReport
            $coverageReport = Get-CoverageReport -PesterState $pester
            Write-CoverageReport -CoverageReport $coverageReport
            if ((& $script:SafeCommands['Get-Variable'] -Name CodeCoverageOutputFile -ValueOnly -ErrorAction $script:IgnoreErrorPreference) `
                    -and (& $script:SafeCommands['Get-Variable'] -Name CodeCoverageOutputFileFormat -ValueOnly -ErrorAction $script:IgnoreErrorPreference) -eq 'JaCoCo') {
                $jaCoCoReport = Get-JaCoCoReportXml -PesterState $pester -CoverageReport $coverageReport
                $jaCoCoReport | & $SafeCommands['Out-File'] $CodeCoverageOutputFile -Encoding utf8
            }
            Exit-CoverageAnalysis -PesterState $pester
        }
        finally {
            Exit-MockScope
        }

        Set-PesterStatistics

        if (& $script:SafeCommands['Get-Variable'] -Name OutputFile -ValueOnly -ErrorAction $script:IgnoreErrorPreference) {
            Export-PesterResults -PesterState $pester -Path $OutputFile -Format $OutputFormat
        }

        if ($PassThru) {
            # Remove all runtime properties like current* and Scope
            $properties = @(
                "TagFilter", "ExcludeTagFilter", "TestNameFilter", "ScriptBlockFilter", "TotalCount", "PassedCount", "FailedCount", "SkippedCount", "PendingCount", 'InconclusiveCount', "Time", "TestResult"

                if ($CodeCoverage) {
                    @{ Name = 'CodeCoverage'; Expression = { $coverageReport } }
                }
            )

            $pester | & $script:SafeCommands['Select-Object'] -Property $properties
        }

        if ($EnableExit) {
            Exit-WithCode -FailedCount $pester.FailedCount
        }
    }
}

function New-PesterOption {
    [CmdletBinding()]
    param (
        [switch] $IncludeVSCodeMarker,

        [ValidateNotNullOrEmpty()]
        [string] $TestSuiteName = 'Pester',

        [switch] $Experimental,

        [switch] $ShowScopeHints,

        [hashtable[]] $ScriptBlockFilter
    )

    # in PowerShell 2 Add-Member can attach properties only to
    # PSObjects, I could work around this by capturing all instances
    # in checking them during runtime, but that would bring a lot of
    # object management problems - so let's just not allow this in PowerShell 2
    if ($Experimental -and $ShowScopeHints) {
        if ($PSVersionTable.PSVersion.Major -lt 3) {
            throw "Scope hints cannot be used on PowerShell 2 due to limitations of Add-Member."
        }

        $script:DisableScopeHints = $false
    }
    else {
        $script:DisableScopeHints = $true
    }

    return & $script:SafeCommands['New-Object'] psobject -Property @{
        IncludeVSCodeMarker = [bool] $IncludeVSCodeMarker
        TestSuiteName       = $TestSuiteName
        ShowScopeHints      = $ShowScopeHints
        Experimental        = $Experimental
        ScriptBlockFilter   = $ScriptBlockFilter
    }
}

function ResolveTestScripts {
    param (
        [object[]] $Path
    )

    $resolvedScriptInfo = @(
        foreach ($object in $Path) {
            if ($object -is [System.Collections.IDictionary]) {
                $unresolvedPath = Get-DictionaryValueFromFirstKeyFound -Dictionary $object -Key 'Path', 'p'
                $script = Get-DictionaryValueFromFirstKeyFound -Dictionary $object -Key 'Script'
                $arguments = @(Get-DictionaryValueFromFirstKeyFound -Dictionary $object -Key 'Arguments', 'args', 'a')
                $parameters = Get-DictionaryValueFromFirstKeyFound -Dictionary $object -Key 'Parameters', 'params'

                if ($null -eq $Parameters) {
                    $Parameters = @{}
                }

                if ($unresolvedPath -isnot [string] -or $unresolvedPath -notmatch '\S' -and ($script -isnot [string] -or $script -notmatch '\S')) {
                    throw 'When passing hashtables to the -Path parameter, the Path key is mandatory, and must contain a single string.'
                }

                if ($null -ne $parameters -and $parameters -isnot [System.Collections.IDictionary]) {
                    throw 'When passing hashtables to the -Path parameter, the Parameters key (if present) must be assigned an IDictionary object.'
                }
            }
            else {
                $unresolvedPath = [string] $object
                $script = [string] $object
                $arguments = @()
                $parameters = @{}
            }

            if (-not [string]::IsNullOrEmpty($unresolvedPath)) {
                if ($unresolvedPath -notmatch '[\*\?\[\]]' -and
                    (& $script:SafeCommands['Test-Path'] -LiteralPath $unresolvedPath -PathType Leaf) -and
                    (& $script:SafeCommands['Get-Item'] -LiteralPath $unresolvedPath) -is [System.IO.FileInfo]) {
                    $extension = [System.IO.Path]::GetExtension($unresolvedPath)
                    if ($extension -ne '.ps1') {
                        & $script:SafeCommands['Write-Error'] "Script path '$unresolvedPath' is not a ps1 file."
                    }
                    else {
                        & $script:SafeCommands['New-Object'] psobject -Property @{
                            Path       = $unresolvedPath
                            Script     = $null
                            Arguments  = $arguments
                            Parameters = $parameters
                        }
                    }
                }
                else {
                    # World's longest pipeline?

                    & $script:SafeCommands['Resolve-Path'] -Path $unresolvedPath |
                        & $script:SafeCommands['Where-Object'] { $_.Provider.Name -eq 'FileSystem' } |
                        & $script:SafeCommands['Select-Object'] -ExpandProperty ProviderPath |
                        & $script:SafeCommands['Get-ChildItem'] -Include *.Tests.ps1 -Recurse |
                        & $script:SafeCommands['Where-Object'] { -not $_.PSIsContainer } |
                        & $script:SafeCommands['Select-Object'] -ExpandProperty FullName -Unique |
                        & $script:SafeCommands['ForEach-Object'] {
                        & $script:SafeCommands['New-Object'] psobject -Property @{
                            Path       = $_
                            Script     = $null
                            Arguments  = $arguments
                            Parameters = $parameters
                        }
                    }
                }
            }
            elseif (-not [string]::IsNullOrEmpty($script)) {
                & $script:SafeCommands['New-Object'] psobject -Property @{
                    Path       = $null
                    Script     = $script
                    Arguments  = $arguments
                    Parameters = $parameters
                }
            }
        }
    )

    # Here, we have the option of trying to weed out duplicate file paths that also contain identical
    # Parameters / Arguments.  However, we already make sure that each object in $Path didn't produce
    # any duplicate file paths, and if the caller happens to pass in a set of parameters that produce
    # dupes, maybe that's not our problem.  For now, just return what we found.

    $resolvedScriptInfo
}

function Get-DictionaryValueFromFirstKeyFound {
    param (
        [System.Collections.IDictionary] $Dictionary,
        [object[]] $Key
    )

    foreach ($keyToTry in $Key) {
        if ($Dictionary.Contains($keyToTry)) {
            return $Dictionary[$keyToTry]
        }
    }
}

function Set-ScriptBlockScope {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock,

        [Parameter(Mandatory = $true, ParameterSetName = 'FromSessionState')]
        [System.Management.Automation.SessionState]
        $SessionState,

        [Parameter(Mandatory = $true, ParameterSetName = 'FromSessionStateInternal')]
        [AllowNull()]
        $SessionStateInternal
    )

    $flags = [System.Reflection.BindingFlags]'Instance,NonPublic'

    if ($PSCmdlet.ParameterSetName -eq 'FromSessionState') {
        $SessionStateInternal = $SessionState.GetType().GetProperty('Internal', $flags).GetValue($SessionState, $null)
    }

    $property = [scriptblock].GetProperty('SessionStateInternal', $flags)
    $scriptBlockSessionState = $property.GetValue($ScriptBlock, $null)

    if (-not $script:DisableScopeHints) {
        # hint can be attached on the internal state (preferable) when the state is there.
        # if we are given unbound scriptblock with null internal state then we hope that
        # the source cmdlet set the hint directly on the ScriptBlock,
        # otherwise the origin is unknown and the cmdlet that allowed this scriptblock in
        # should be found and add hint

        $hint = $scriptBlockSessionState.Hint
        if ($null -eq $hint) {
            if ($null -ne $ScriptBlock.Hint) {
                $hint = $ScriptBlock.Hint
            }
            else {
                $hint = 'Unknown unbound ScriptBlock'
            }
        }
        Write-Hint "Setting ScriptBlock state from source state '$hint' to '$($SessionStateInternal.Hint)'"
    }

    $property.SetValue($ScriptBlock, $SessionStateInternal, $null)

    if (-not $script:DisableScopeHints) {
        Set-ScriptBlockHint -ScriptBlock $ScriptBlock
    }
}

function Get-ScriptBlockScope {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock
    )


    $flags = [System.Reflection.BindingFlags]'Instance,NonPublic'
    $sessionStateInternal = [scriptblock].GetProperty('SessionStateInternal', $flags).GetValue($ScriptBlock, $null)
    if (-not $script:DisableScopeHints) {
        Write-Hint "Getting scope from ScriptBlock '$($sessionStateInternal.Hint)'"
    }
    $sessionStateInternal
}

function SafeGetCommand {
    return $script:SafeCommands['Get-Command']
}

function Set-PesterStatistics {
    param (
        $Node
    )
    if ($null -eq $Node) {
        $Node = $pester.TestActions
    }

    foreach ($action in $Node.Actions) {
        if ($action.Type -eq 'TestGroup') {
            Set-PesterStatistics -Node $action

            $Node.TotalCount += $action.TotalCount
            $Node.PassedCount += $action.PassedCount
            $Node.FailedCount += $action.FailedCount
            $Node.SkippedCount += $action.SkippedCount
            $Node.PendingCount += $action.PendingCount
            $Node.InconclusiveCount += $action.InconclusiveCount
        }
        elseif ($action.Type -eq 'TestCase') {
            $node.TotalCount++

            switch ($action.Result) {
                Passed {
                    $Node.PassedCount++; break;
                }
                Failed {
                    $Node.FailedCount++; break;
                }
                Skipped {
                    $Node.SkippedCount++; break;
                }
                Pending {
                    $Node.PendingCount++; break;
                }
                Inconclusive {
                    $Node.InconclusiveCount++; break;
                }
            }
        }
    }
}

function Contain-AnyStringLike {
    param (
        $Filter,
        $Collection
    )

    foreach ($item in $Collection) {
        foreach ($value in $Filter) {
            if ($item -like $value) {
                return $true
            }
        }
    }
    return $false
}

$snippetsDirectoryPath = "$PSScriptRoot\Snippets"
if ((& $script:SafeCommands['Test-Path'] -Path Variable:\psise) -and
    ($null -ne $psISE) -and
    ($PSVersionTable.PSVersion.Major -ge 3) -and
    (& $script:SafeCommands['Test-Path'] $snippetsDirectoryPath)) {
    Import-IseSnippet -Path $snippetsDirectoryPath
}

function Assert-VerifiableMocks {
    [CmdletBinding()]
    param()

    Throw "This command has been renamed to 'Assert-VerifiableMock' (without the 's' at the end), please update your code. For more information see: https://github.com/pester/Pester/wiki/Migrating-from-Pester-3-to-Pester-4"

}

Set-SessionStateHint -Hint Pester -SessionState $ExecutionContext.SessionState
# in the future rename the function to Add-ShouldOperator
Set-Alias -Name Add-ShouldOperator -Value Add-AssertionOperator

& $script:SafeCommands['Export-ModuleMember'] Describe, Context, It, In, Mock, Assert-VerifiableMock, Assert-VerifiableMocks, Assert-MockCalled, Set-TestInconclusive, Set-ItResult
& $script:SafeCommands['Export-ModuleMember'] New-Fixture, Get-TestDriveItem, Should, Invoke-Pester, Setup, InModuleScope, Invoke-Mock
& $script:SafeCommands['Export-ModuleMember'] BeforeEach, AfterEach, BeforeAll, AfterAll
& $script:SafeCommands['Export-ModuleMember'] Get-MockDynamicParameter, Set-DynamicParameterVariable
& $script:SafeCommands['Export-ModuleMember'] SafeGetCommand, New-PesterOption
& $script:SafeCommands['Export-ModuleMember'] Invoke-Gherkin, Find-GherkinStep, BeforeEachFeature, BeforeEachScenario, AfterEachFeature, AfterEachScenario, GherkinStep -Alias Given, When, Then, And, But
& $script:SafeCommands['Export-ModuleMember'] New-MockObject, Add-AssertionOperator, Get-ShouldOperator  -Alias Add-ShouldOperator
