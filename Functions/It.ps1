function It {
    [CmdletBinding(DefaultParameterSetName = 'Normal')]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Name,

        [Parameter(Position = 1)]
        [ScriptBlock] $Test = {},

        [System.Collections.IDictionary[]] $TestCases,

        [Parameter(ParameterSetName = 'Pending')]
        [Switch] $Pending,

        [Parameter(ParameterSetName = 'Skip')]
        [Alias('Ignore')]
        [Switch] $Skip
    )

    ItImpl -Pester $pester -OutputScriptBlock ${function:Write-PesterResult} @PSBoundParameters
}

function ItImpl {
    [CmdletBinding(DefaultParameterSetName = 'Normal')]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Name,
        [Parameter(Position = 1)]
        [ScriptBlock] $Test,
        [System.Collections.IDictionary[]] $TestCases,
        [Parameter(ParameterSetName = 'Pending')]
        [Switch] $Pending,

        [Parameter(ParameterSetName = 'Skip')]
        [Alias('Ignore')]
        [Switch] $Skip,

        $Pester,
        [scriptblock] $OutputScriptBlock
    )

    Assert-DescribeInProgress -CommandName It

    # Jumping through hoops to make strict mode happy.
    if ($PSCmdlet.ParameterSetName -ne 'Skip') {
        $Skip = $false
    }
    if ($PSCmdlet.ParameterSetName -ne 'Pending') {
        $Pending = $false
    }

    #unless Skip or Pending is specified you must specify a ScriptBlock to the Test parameter
    if (-not ($PSBoundParameters.ContainsKey('test') -or $Skip -or $Pending)) {
        throw 'No test script block is provided. (Have you put the open curly brace on the next line?)'
    }

    #the function is called with Pending or Skipped set the script block if needed
    if ($null -eq $Test) {
        $Test = {}
    }

    #mark empty Its as Pending
    if ($PSVersionTable.PSVersion.Major -le 2 -and
        $PSCmdlet.ParameterSetName -eq 'Normal' -and
        [String]::IsNullOrEmpty((Remove-Comments $Test.ToString()) -replace "\s")) {
        $Pending = $true
    }
    elseIf ($PSVersionTable.PSVersion.Major -gt 2) {
        #[String]::IsNullOrWhitespace is not available in .NET version used with PowerShell 2
        # AST is not available also
        $testIsEmpty =
        [String]::IsNullOrEmpty($Test.Ast.BeginBlock.Statements) -and
        [String]::IsNullOrEmpty($Test.Ast.ProcessBlock.Statements) -and
        [String]::IsNullOrEmpty($Test.Ast.EndBlock.Statements)

        if ($PSCmdlet.ParameterSetName -eq 'Normal' -and $testIsEmpty) {
            $Pending = $true
        }
    }

    $pendingSkip = @{}

    if ($PSCmdlet.ParameterSetName -eq 'Skip') {
        $pendingSkip['Skip'] = $Skip
    }
    else {
        $pendingSkip['Pending'] = $Pending
    }

    if ($null -ne $TestCases -and $TestCases.Count -gt 0) {
        foreach ($testCase in $TestCases) {
            $expandedName = [regex]::Replace($Name, '<([^>]+)>', {
                    $capture = $args[0].Groups[1].Value
                    if ($testCase.Contains($capture)) {
                        $value = $testCase[$capture]
                        # skip adding quotes to non-empty strings to avoid adding junk to the
                        # test name in case you want to expand captures like 'because' or test name
                        if ($value -isnot [string] -or [string]::IsNullOrEmpty($value)) {
                            Format-Nicely $value
                        }
                        else {
                            $value
                        }
                    }
                    else {
                        "<$capture>"
                    }
                })

            $splat = @{
                Name                   = $expandedName
                Scriptblock            = $Test
                Parameters             = $testCase
                ParameterizedSuiteName = $Name
                OutputScriptBlock      = $OutputScriptBlock
            }

            Invoke-Test @splat @pendingSkip
        }
    }
    else {
        Invoke-Test -Name $Name -ScriptBlock $Test @pendingSkip -OutputScriptBlock $OutputScriptBlock
    }
}

function Invoke-Test {
    [CmdletBinding(DefaultParameterSetName = 'Normal')]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Name,

        [Parameter(Mandatory = $true)]
        [ScriptBlock] $ScriptBlock,

        [scriptblock] $OutputScriptBlock,

        [System.Collections.IDictionary] $Parameters,
        [string] $ParameterizedSuiteName,

        [Parameter(ParameterSetName = 'Pending')]
        [Switch] $Pending,

        [Parameter(ParameterSetName = 'Skip')]
        [Alias('Ignore')]
        [Switch] $Skip
    )

    if ($null -eq $Parameters) {
        $Parameters = @{}
    }

    try {
        if ($Skip) {
            $Pester.AddTestResult($Name, "Skipped", $null)
        }
        elseif ($Pending) {
            $Pester.AddTestResult($Name, "Pending", $null)
        }
        else {
            #todo: disabling the progress for now, it adds a lot of overhead and breaks output on linux, we don't have a good way to disable it by default, or to show it after delay see: https://github.com/pester/Pester/issues/846
            # & $SafeCommands['Write-Progress'] -Activity "Running test '$Name'" -Status Processing

            $errorRecord = $null
            try {
                $pester.EnterTest()
                Invoke-TestCaseSetupBlocks

                do {
                    Write-ScriptBlockInvocationHint -Hint "It" -ScriptBlock $ScriptBlock
                    $null = & $ScriptBlock @Parameters
                } until ($true)
            }
            catch {
                $errorRecord = $_
            }
            finally {
                #guarantee that the teardown action will run and prevent it from failing the whole suite
                try {
                    if (-not ($Skip -or $Pending)) {
                        Invoke-TestCaseTeardownBlocks
                    }
                }
                catch {
                    $errorRecord = $_
                }

                $pester.LeaveTest()
            }

            $result = ConvertTo-PesterResult -Name $Name -ErrorRecord $errorRecord
            $orderedParameters = Get-OrderedParameterDictionary -ScriptBlock $ScriptBlock -Dictionary $Parameters
            $Pester.AddTestResult( $result.Name, $result.Result, $null, $result.FailureMessage, $result.StackTrace, $ParameterizedSuiteName, $orderedParameters, $result.ErrorRecord )
            #todo: disabling progress reporting see above & $SafeCommands['Write-Progress'] -Activity "Running test '$Name'" -Completed -Status Processing
        }
    }
    finally {
        Exit-MockScope -ExitTestCaseOnly
    }

    if ($null -ne $OutputScriptBlock) {
        $Pester.testresult[-1] | & $OutputScriptBlock
    }
}

function Get-OrderedParameterDictionary {
    [OutputType([System.Collections.IDictionary])]
    param (
        [scriptblock] $ScriptBlock,
        [System.Collections.IDictionary] $Dictionary
    )

    $parameters = Get-ParameterDictionary -ScriptBlock $ScriptBlock

    $orderedDictionary = & $SafeCommands['New-Object'] System.Collections.Specialized.OrderedDictionary

    foreach ($parameterName in $parameters.Keys) {
        $value = $null
        if ($Dictionary.ContainsKey($parameterName)) {
            $value = $Dictionary[$parameterName]
        }

        $orderedDictionary[$parameterName] = $value
    }

    return $orderedDictionary
}

function Get-ParameterDictionary {
    param (
        [scriptblock] $ScriptBlock
    )

    $guid = [Guid]::NewGuid().Guid

    try {
        & $SafeCommands['Set-Content'] function:\$guid $ScriptBlock
        $metadata = [System.Management.Automation.CommandMetadata](& $SafeCommands['Get-Command'] -Name $guid -CommandType Function)

        return $metadata.Parameters
    }
    finally {
        if (& $SafeCommands['Test-Path'] function:\$guid) {
            & $SafeCommands['Remove-Item'] function:\$guid
        }
    }
}
