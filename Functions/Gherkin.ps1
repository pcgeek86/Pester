if (($PSVersionTable.ContainsKey('PSEdition')) -and ($PSVersionTable.PSEdition -eq 'Core')) {
    & $SafeCommands["Add-Type"] -Path "${Script:PesterRoot}/lib/Gherkin/core/Gherkin.dll"
}
else {
    & $SafeCommands["Import-Module"] -Name "${Script:PesterRoot}/lib/Gherkin/legacy/Gherkin.dll"
}

$GherkinSteps = @{}
$GherkinHooks = @{
    BeforeEachFeature  = @()
    BeforeEachScenario = @()
    AfterEachFeature   = @()
    AfterEachScenario  = @()
}

function Invoke-GherkinHook {
    <#
        .SYNOPSIS
        Internal function to run the various gherkin hooks

        .PARAMETER Hook
        The name of the hook to run

        .PARAMETER Name
        The name of the feature or scenario the hook is being invoked for

        .PARAMETER Tags
        Tags for filtering hooks
    #>
    [CmdletBinding()]
    param([string]$Hook, [string]$Name, [string[]]$Tags)

    if ($GherkinHooks.${Hook}) {
        foreach ($GherkinHook in $GherkinHooks.${Hook}) {
            if ($GherkinHook.Tags -and $Tags) {
                :tags foreach ($hookTag in $GherkinHook.Tags) {
                    foreach ($testTag in $Tags) {
                        if ($testTag -match "^($hookTag)$") {
                            & $hook.Script $Name
                            break :tags
                        }
                    }
                }
            }
            elseif ($GherkinHook.Tags) {
                # If the hook has tags, it can't run if the step doesn't
            }
            else {
                & $GherkinHook.Script $Name
            }
        } # @{ Tags = $Tags; Script = $Test }
    }
}

function Invoke-Gherkin {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param(
        [Parameter(Mandatory = $True, ParameterSetName = "RetestFailed")]
        [switch]$FailedLast,

        [Parameter(Position = 0, Mandatory = $False)]
        [Alias('Script', 'relative_path')]
        [string]$Path = $Pwd,

        [Parameter(Position = 1, Mandatory = $False)]
        [Alias("Name", "TestName")]
        [string[]]$ScenarioName,

        [Parameter(Position = 2, Mandatory = $False)]
        [switch]$EnableExit,

        [Parameter(Position = 4, Mandatory = $False)]
        [Alias('Tags')]
        [string[]]$Tag,

        [string[]]$ExcludeTag,

        [object[]] $CodeCoverage = @(),

        [Switch]$Strict,

        [string] $OutputFile,

        [ValidateSet('NUnitXml')]
        [string] $OutputFormat = 'NUnitXml',

        [Switch]$Quiet,

        [object]$PesterOption,

        [Pester.OutputTypes]$Show = 'All',

        [switch]$PassThru
    )
    begin {
        & $SafeCommands["Import-LocalizedData"] -BindingVariable Script:ReportStrings -BaseDirectory $PesterRoot -FileName Gherkin.psd1 -ErrorAction SilentlyContinue

        #Fallback to en-US culture strings
        If ([String]::IsNullOrEmpty($ReportStrings)) {

            & $SafeCommands["Import-LocalizedData"] -BaseDirectory $PesterRoot -BindingVariable Script:ReportStrings -UICulture 'en-US' -FileName Gherkin.psd1 -ErrorAction Stop

        }

        # Make sure broken tests don't leave you in space:
        $CWD = [Environment]::CurrentDirectory
        $Location = & $SafeCommands["Get-Location"]
        [Environment]::CurrentDirectory = & $SafeCommands["Get-Location"] -PSProvider FileSystem

        $script:GherkinSteps = @{}
        $script:GherkinHooks = @{
            BeforeEachFeature  = @()
            BeforeEachScenario = @()
            AfterEachFeature   = @()
            AfterEachScenario  = @()
        }
    }
    end {
        if ($PSBoundParameters.ContainsKey('Quiet')) {
            & $SafeCommands["Write-Warning"] 'The -Quiet parameter has been deprecated; please use the new -Show parameter instead. To get no output use -Show None.'
            & $SafeCommands["Start-Sleep"] -Seconds 2

            if (!$PSBoundParameters.ContainsKey('Show')) {
                $Show = [Pester.OutputTypes]::None
            }
        }

        if ($PSCmdlet.ParameterSetName -eq "RetestFailed" -and $FailedLast) {
            $ScenarioName = $script:GherkinFailedLast
            if (!$ScenarioName) {
                throw "There are no existing failed tests to re-run."
            }
        }
        $sessionState = Set-SessionStateHint -PassThru  -Hint "Caller - Captured in Invoke-Gherkin" -SessionState $PSCmdlet.SessionState
        $pester = New-PesterState -TagFilter $Tag -ExcludeTagFilter $ExcludeTag -TestNameFilter $ScenarioName -SessionState $sessionState -Strict:$Strict  -Show $Show -PesterOption $PesterOption |
            & $SafeCommands["Add-Member"] -MemberType NoteProperty -Name Features -Value (& $SafeCommands["New-Object"] System.Collections.Generic.List[PSObject] ) -PassThru |
            & $SafeCommands["Add-Member"] -MemberType ScriptProperty -Name FailedScenarios -PassThru -Value {
            $Names = $this.TestResult | & $SafeCommands["Group-Object"] Describe |
                & $SafeCommands["Where-Object"] { $_.Group |
                    & $SafeCommands["Where-Object"] { -not $_.Passed } } |
                & $SafeCommands["Select-Object"] -ExpandProperty Name
            $this.Features | Select-Object -ExpandProperty Scenarios | & $SafeCommands["Where-Object"] { $Names -contains $_.Name }
        } |
            & $SafeCommands["Add-Member"] -MemberType ScriptProperty -Name PassedScenarios -PassThru -Value {
            $Names = $this.TestResult | & $SafeCommands["Group-Object"] Describe |
                & $SafeCommands["Where-Object"] { -not ($_.Group |
                        & $SafeCommands["Where-Object"] { -not $_.Passed }) } |
                & $SafeCommands["Select-Object"] -ExpandProperty Name
            $this.Features | Select-Object -ExpandProperty Scenarios | & $SafeCommands["Where-Object"] { $Names -contains $_.Name }
        }

        Write-PesterStart $pester $Path

        Enter-CoverageAnalysis -CodeCoverage $CodeCoverage -PesterState $pester

        foreach ($FeatureFile in & $SafeCommands["Get-ChildItem"] $Path -Filter "*.feature" -Recurse ) {
            Invoke-GherkinFeature $FeatureFile -Pester $pester
        }

        # Remove all the steps
        $Script:GherkinSteps.Clear()

        $Location | & $SafeCommands["Set-Location"]
        [Environment]::CurrentDirectory = $CWD

        $pester | Write-PesterReport
        $coverageReport = Get-CoverageReport -PesterState $pester
        Write-CoverageReport -CoverageReport $coverageReport
        Exit-CoverageAnalysis -PesterState $pester

        if (& $SafeCommands["Get-Variable"]-Name OutputFile -ValueOnly -ErrorAction $script:IgnoreErrorPreference) {
            Export-PesterResults -PesterState $pester -Path $OutputFile -Format $OutputFormat
        }

        if ($PassThru) {
            # Remove all runtime properties like current* and Scope
            $properties = @(
                "Path", "Features", "TagFilter", "TestNameFilter", "TotalCount", "PassedCount", "FailedCount", "Time", "TestResult", "PassedScenarios", "FailedScenarios"

                if ($CodeCoverage) {
                    @{ Name = 'CodeCoverage'; Expression = { $coverageReport } }
                }
            )
            $result = $pester | & $SafeCommands["Select-Object"] -Property $properties
            $result.PSTypeNames.Insert(0, "Pester.Gherkin.Results")
            $result
        }
        $script:GherkinFailedLast = @($pester.FailedScenarios.Name)
        if ($EnableExit) {
            Exit-WithCode -FailedCount $pester.FailedCount
        }
    }
}

function Import-GherkinSteps {
    <#
        .SYNOPSIS
        Internal function for importing the script steps from a directory tree
        .PARAMETER StepPath
        The folder which contains step files
        .PARAMETER Pester
        Pester
    #>

    [CmdletBinding()]
    param(

        [Alias("PSPath")]
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipelineByPropertyName = $True)]
        $StepPath,

        [PSObject]$Pester
    )
    begin {
        # Remove all existing steps
        $Script:GherkinSteps.Clear()
        # Remove all existing hooks
        $Script:GherkinHooks.Clear()
    }
    process {
        $StepFiles = & $SafeCommands["Get-ChildItem"] $StepPath -Filter "*.?teps.ps1" -Include "*.[sS]teps.ps1" -Recurse

        foreach ($StepFile in $StepFiles) {
            $invokeTestScript = {
                [CmdletBinding()]
                param (
                    [Parameter(Position = 0)]
                    [string] $Path
                )

                & $Path
            }

            Set-ScriptBlockScope -ScriptBlock $invokeTestScript -SessionState $Pester.SessionState

            & $invokeTestScript $StepFile.FullName
        }

        & $SafeCommands["Write-Verbose"] "Loaded $($Script:GherkinSteps.Count) step definitions from $(@($StepFiles).Count) steps file(s)"
    }
}

function Import-GherkinFeature {
    <#
        .SYNOPSIS
        Internal function to import a Gherkin feature file. Wraps Gherkin.Parse

        .PARAMETER Path
        The path to the feature file to import

        .PARAMETER Pester
        Internal Pester object. For internal use only
    #>
    [CmdletBinding()]
    param($Path, [PSObject]$Pester)
    $Background = $null

    $parser = & $SafeCommands["New-Object"] Gherkin.Parser
    $Feature = $parser.Parse($Path).Feature | Convert-Tags
    $Scenarios = $(
        :scenarios foreach ($Child in $Feature.Children) {
            $null = & $SafeCommands["Add-Member"] -MemberType "NoteProperty" -InputObject $Child.Location -Name "Path" -Value $Path
            foreach ($Step in $Child.Steps) {
                $null = & $SafeCommands["Add-Member"] -MemberType "NoteProperty" -InputObject $Step.Location -Name "Path" -Value $Path
            }

            switch ($Child.Keyword.Trim()) {
                { (Test-Keyword $_ 'scenario' $Feature.Language) -or (Test-Keyword $_ 'scenarioOutline' $Feature.Language) } {
                    $Scenario = Convert-Tags -InputObject $Child -BaseTags $Feature.Tags
                }
                { Test-Keyword $_ 'background' $Feature.Language } {
                    $Background = Convert-Tags -InputObject $Child -BaseTags $Feature.Tags
                    continue scenarios
                }
                default {
                    & $SafeCommands["Write-Warning"] "Unexpected Feature Child: $_"
                }
            }

            if ( $Scenario -is [Gherkin.Ast.ScenarioOutline] ) {
                # If there is no example set name, the following index will be included in the scenario name
                $ScenarioIndex = 0
                foreach ($ExampleSet in $Scenario.Examples) {
                    ${Column Names} = @($ExampleSet.TableHeader.Cells | & $SafeCommands["Select-Object"] -ExpandProperty Value)
                    $NamesPattern = "<(?:" + (${Column Names} -join "|") + ")>"
                    # If there is an example set name, the following index will be included in the scenario name
                    $ExampleSetIndex = 0
                    foreach ($Example in $ExampleSet.TableBody) {
                        $ScenarioIndex++
                        $ExampleSetIndex++
                        $Steps = foreach ($Step in $Scenario.Steps) {
                            [string]$StepText = $Step.Text
                            if ($StepText -match $NamesPattern) {
                                for ($n = 0; $n -lt ${Column Names}.Length; $n++) {
                                    $Name = ${Column Names}[$n]
                                    if ($Example.Cells[$n].Value -and $StepText -match "<${Name}>") {
                                        $StepText = $StepText -replace "<${Name}>", $Example.Cells[$n].Value
                                    }
                                }
                            }
                            if ($StepText -ne $Step.Text) {
                                & $SafeCommands["New-Object"] Gherkin.Ast.Step $Step.Location, $Step.Keyword.Trim(), $StepText, $Step.Argument
                            }
                            else {
                                $Step
                            }
                        }
                        $ScenarioName = $Scenario.Name
                        if ($ExampleSet.Name) {
                            # Include example set name and index of example
                            $ScenarioName = $ScenarioName + " [$($ExampleSet.Name.Trim()) $ExampleSetIndex]"
                        }
                        else {
                            # Only include index of scenario
                            $ScenarioName = $ScenarioName + " [$ScenarioIndex]"
                        }
                        & $SafeCommands["New-Object"] Gherkin.Ast.Scenario $ExampleSet.Tags, $Scenario.Location, $Scenario.Keyword.Trim(), $ScenarioName, $Scenario.Description, $Steps | Convert-Tags $Scenario.Tags
                    }
                }
            }
            else {
                $Scenario
            }
        }
    )

    & $SafeCommands["Add-Member"] -MemberType NoteProperty -InputObject $Feature -Name Scenarios -Value $Scenarios -Force
    return $Feature, $Background, $Scenarios
}

function Invoke-GherkinFeature {
    <#
        .SYNOPSIS
        Internal function to (parse and) run a whole feature file
    #>
    [CmdletBinding()]
    param(
        [Alias("PSPath")]
        [Parameter(Mandatory = $True, Position = 0, ValueFromPipelineByPropertyName = $True)]
        [IO.FileInfo]$FeatureFile,

        [PSObject]$Pester
    )
    # Make sure broken tests don't leave you in space:
    $CWD = [Environment]::CurrentDirectory
    $Location = & $SafeCommands["Get-Location"]
    [Environment]::CurrentDirectory = & $SafeCommands["Get-Location"] -PSProvider FileSystem

    try {
        $Parent = & $SafeCommands["Split-Path"] $FeatureFile.FullName
        Import-GherkinSteps -StepPath $Parent -Pester $pester
        $Feature, $Background, $Scenarios = Import-GherkinFeature -Path $FeatureFile.FullName -Pester $Pester
    }
    catch [Gherkin.ParserException] {
        & $SafeCommands["Write-Error"] -Exception $_.Exception -Message "Skipped '$($FeatureFile.FullName)' because of parser error.`n$(($_.Exception.Errors | & $SafeCommands["Select-Object"] -Expand Message) -join "`n`n")"
        continue
    }

    # To create a more user-friendly test report, we use the feature name for the test group
    $Pester.EnterTestGroup($Feature.Name, 'Script')

    $null = $Pester.Features.Add($Feature)
    Invoke-GherkinHook BeforeEachFeature $Feature.Name $Feature.Tags

    # Test the name filter first, since it will probably return one single item
    if ($Pester.TestNameFilter) {
        $Scenarios = foreach ($nameFilter in $Pester.TestNameFilter) {
            $Scenarios | & $SafeCommands["Where-Object"] { $_.Name -like $NameFilter }
        }
        $Scenarios = $Scenarios | & $SafeCommands["Get-Unique"]
    }

    # if($Pester.TagFilter -and @(Compare-Object $Tags $Pester.TagFilter -IncludeEqual -ExcludeDifferent).count -eq 0) {return}
    if ($Pester.TagFilter) {
        $Scenarios = $Scenarios | & $SafeCommands["Where-Object"] { & $SafeCommands["Compare-Object"] $_.Tags $Pester.TagFilter -IncludeEqual -ExcludeDifferent }
    }

    # if($Pester.ExcludeTagFilter -and @(Compare-Object $Tags $Pester.ExcludeTagFilter -IncludeEqual -ExcludeDifferent).count -gt 0) {return}
    if ($Pester.ExcludeTagFilter) {
        $Scenarios = $Scenarios | & $SafeCommands["Where-Object"] { !(& $SafeCommands["Compare-Object"] $_.Tags $Pester.ExcludeTagFilter -IncludeEqual -ExcludeDifferent) }
    }

    if ($Scenarios) {
        Write-Describe (New-Object PSObject -Property @{Name = "$($Feature.Keyword): $($Feature.Name)"; Description = $Feature.Description })
    }

    try {
        foreach ($Scenario in $Scenarios) {
            Invoke-GherkinScenario $Pester $Scenario $Background $Feature.Language
        }
    }
    catch {
        $firstStackTraceLine = $_.ScriptStackTrace -split '\r?\n' | & $SafeCommands["Select-Object"] -First 1
        $Pester.AddTestResult("Error occurred in test script '$($Feature.Path)'", "Failed", $null, $_.Exception.Message, $firstStackTraceLine, $null, $null, $_)

        # This is a hack to ensure that XML output is valid for now.  The test-suite names come from the Describe attribute of the TestResult
        # objects, and a blank name is invalid NUnit XML.  This will go away when we promote test scripts to have their own test-suite nodes,
        # planned for v4.0
        $Pester.TestResult[-1].Describe = "Error in $($Feature.Path)"

        $Pester.TestResult[-1] | Write-PesterResult
    }
    finally {
        $Location | & $SafeCommands["Set-Location"]
        [Environment]::CurrentDirectory = $CWD
    }

    Invoke-GherkinHook AfterEachFeature $Feature.Name $Feature.Tags

    $Pester.LeaveTestGroup($Feature.Name, 'Script')

}

function Invoke-GherkinScenario {
    <#
        .SYNOPSIS
        Internal function to (parse and) run a single scenario
    #>
    [CmdletBinding()]
    param(
        $Pester, $Scenario, $Background, $Language
    )
    $Pester.EnterTestGroup($Scenario.Name, 'Scenario')
    try {
        # We just display 'Scenario', also for 'Scenario Outline' or 'Scenario Template'
        # Thus we use the translation of 'scenario' instead of $Scenario.Keyword
        Write-Context (New-Object PSObject -Property @{Name = "$(Get-Translation 'scenario' $Language): $($Scenario.Name)"; Description = $Scenario.Description })

        $script:mockTable = @{}

        # Create a clean variable scope in each scenario
        $script:GherkinScenarioScope = New-Module Scenario {       $a = 4
        }
        $script:GherkinSessionState = Set-SessionStateHint -PassThru -Hint Scenario -SessionState $Script:GherkinScenarioScope.SessionState

        #Wait-Debugger

        New-TestDrive
        Invoke-GherkinHook BeforeEachScenario $Scenario.Name $Scenario.Tags

        $testResultIndexStart = $Pester.TestResult.Count

        # If there's a background, run that before the test, but after hooks
        if ($Background) {
            foreach ($Step in $Background.Steps) {
                # Run Background steps -Background so they don't output in each scenario
                Invoke-GherkinStep -Step $Step -Pester $Pester -Scenario $GherkinSessionState -Visible -TestResultIndexStart $testResultIndexStart
            }
        }

        foreach ($Step in $Scenario.Steps) {
            Invoke-GherkinStep -Step $Step -Pester $Pester -Scenario $GherkinSessionState -Visible -TestResultIndexStart $testResultIndexStart
        }

        Invoke-GherkinHook AfterEachScenario $Scenario.Name $Scenario.Tags
    }
    catch {
        $firstStackTraceLine = $_.ScriptStackTrace -split '\r?\n' | & $SafeCommands["Select-Object"] -First 1
        $Pester.AddTestResult("Error occurred in scenario '$($Scenario.Name)'", "Failed", $null, $_.Exception.Message, $firstStackTraceLine, $null, $null, $_)

        # This is a hack to ensure that XML output is valid for now.  The test-suite names come from the Describe attribute of the TestResult
        # objects, and a blank name is invalid NUnit XML.  This will go away when we promote test scripts to have their own test-suite nodes,
        # planned for v4.0
        $Pester.TestResult[-1].Describe = "Error in $($Scenario.Name)"

        $Pester.TestResult[-1] | Write-PesterResult
    }

    Remove-TestDrive
    $Pester.LeaveTestGroup($Scenario.Name, 'Scenario')
    Exit-MockScope
}

function Find-GherkinStep {
    [CmdletBinding()]
    param(

        [string]$Step,

        [string]$BasePath = $Pwd
    )

    $OriginalGherkinSteps = $Script:GherkinSteps
    try {
        Import-GherkinSteps $BasePath -Pester $PSCmdlet

        $KeyWord, $StepText = $Step -split "(?<=^(?:Given|When|Then|And|But))\s+"
        if (!$StepText) {
            $StepText = $KeyWord
        }

        & $SafeCommands["Write-Verbose"] "Searching for '$StepText' in $($Script:GherkinSteps.Count) steps"
        $(
            foreach ($StepCommand in $Script:GherkinSteps.Keys) {
                & $SafeCommands["Write-Verbose"] "... $StepCommand"
                if ($StepText -match "^${StepCommand}$") {
                    & $SafeCommands["Write-Verbose"] "Found match: $StepCommand"
                    $StepCommand | & $SafeCommands["Add-Member"] -MemberType NoteProperty -Name MatchCount -Value $Matches.Count -PassThru
                }
            }
        ) | & $SafeCommands["Sort-Object"] MatchCount | & $SafeCommands["Select-Object"] @{
            Name       = 'Step'
            Expression = { $Step }
        }, @{
            Name       = 'Source'
            Expression = { $Script:GherkinSteps["$_"].Source }
        }, @{
            Name       = 'Implementation'
            Expression = { $Script:GherkinSteps["$_"] }
        } -First 1

        # $StepText = "{0} {1} {2}" -f $Step.Keyword.Trim(), $Step.Text, $Script:GherkinSteps[$StepCommand].Source

    }
    finally {
        $Script:GherkinSteps = $OriginalGherkinSteps
    }
}

function Invoke-GherkinStep {
    <#
        .SYNOPSIS
        Internal function to  run a single gherkin step, given the text from the feature file

        .PARAMETER Step
        The text of the step for matching against regex patterns in step implementations

        .PARAMETER Visible
        If Visible is true, the results of this step will be shown in the test report

        .PARAMETER Pester
        Pester state object. For internal use only

        .PARAMETER ScenarioState
        Gherkin state object. For internal use only

        .PARAMETER TestResultIndexStart
        Used to hold the test result index of the first step of the current scenario. For internal use only
    #>
    [CmdletBinding()]
    param (
        $Step,

        [Switch]$Visible,

        $Pester,

        $ScenarioState,

        [int] $TestResultIndexStart
    )
    if ($Step -is [string]) {
        $KeyWord, $StepText = $Step -split "(?<=^(?:Given|When|Then|And|But))\s+"
        if (!$StepText) {
            $StepText = $KeyWord
            $Keyword = "Step"
        }
        $Step = @{ Text = $StepText; Keyword = $Keyword }
    }
    $DisplayText = "{0} {1}" -f $Step.Keyword.Trim(), $Step.Text

    $PesterErrorRecord = $null
    $Elapsed = $null
    $NamedArguments = @{}

    try {
        #  Pick the match with the least grouping wildcards in it...
        $StepCommand = $(
            foreach ($StepCommand in $Script:GherkinSteps.Keys) {
                if ($Step.Text -match "^${StepCommand}$") {
                    $StepCommand | & $SafeCommands["Add-Member"] -MemberType NoteProperty -Name MatchCount -Value $Matches.Count -PassThru
                }
            }
        ) | & $SafeCommands["Sort-Object"] MatchCount | & $SafeCommands["Select-Object"] -First 1

        $previousStepsNotSuccessful = $false
        # Iterate over the test results of the previous steps
        for ($i = $TestResultIndexStart; $i -lt ($Pester.TestResult.Count); $i++) {
            $previousTestResult = $Pester.TestResult[$i].Result
            if ($previousTestResult -eq "Failed" -or $previousTestResult -eq "Inconclusive") {
                $previousStepsNotSuccessful = $true
                break
            }
        }
        if (!$StepCommand -or $previousStepsNotSuccessful) {
            $skipMessage = if (!$StepCommand) {
                "Could not find implementation for step!"
            }
            else {
                "Step skipped (previous step did not pass)"
            }
            $PesterErrorRecord = New-PesterErrorRecord -Result Inconclusive -Message $skipMessage -File $Step.Location.Path -Line $Step.Location.Line -LineText $DisplayText
        }
        else {
            $NamedArguments, $Parameters = Get-StepParameters $Step $StepCommand
            $watch = & $SafeCommands["New-Object"] System.Diagnostics.Stopwatch
            $watch.Start()
            try {
                # Invoke-GherkinHook BeforeStep $Step.Text $Step.Tags

                if ($NamedArguments.Count) {
                    if ($NamedArguments.ContainsKey("Table")) {
                        $DisplayText += "..."
                    }
                    $ScriptBlock = { . $Script:GherkinSteps.$StepCommand @NamedArguments @Parameters }
                }
                else {
                    $ScriptBlock = { . $Script:GherkinSteps.$StepCommand @Parameters }
                }
                Set-ScriptBlockScope -ScriptBlock $Script:GherkinSteps.$StepCommand -SessionState $ScenarioState

                Write-ScriptBlockInvocationHint -Hint "Invoke-Gherkin step" -ScriptBlock $Script:GherkinSteps.$StepCommand
                $null = & $ScriptBlock
            }
            catch {
                $PesterErrorRecord = $_
            }
            $watch.Stop()
            $Elapsed = $watch.Elapsed
        }
    }
    catch {
        $PesterErrorRecord = $_
    }

    if ($Pester -and $Visible) {
        for ($p = 0; $p -lt $Parameters.Count; $p++) {
            $NamedArguments."Unnamed-$p" = $Parameters[$p]
        }

        # Normally, PesterErrorRecord is an ErrorRecord. Sometimes, it's an exception which HAS A ErrorRecord
        if ($PesterErrorRecord.ErrorRecord) {
            $PesterErrorRecord = $PesterErrorRecord.ErrorRecord
        }

        ${Pester Result} = ConvertTo-PesterResult -ErrorRecord $PesterErrorRecord

        # For Gherkin, we want to show the step, but not pretend to be a StackTrace
        if (${Pester Result}.Result -eq 'Inconclusive') {
            ${Pester Result}.StackTrace = "At " + $Step.Keyword.Trim() + ', ' + $Step.Location.Path + ': line ' + $Step.Location.Line
        }
        else {
            # Unless we really are a StackTrace...
            ${Pester Result}.StackTrace += "`nFrom " + $Step.Location.Path + ': line ' + $Step.Location.Line
        }
        $Pester.AddTestResult($DisplayText, ${Pester Result}.Result, $Elapsed, ${Pester Result}.FailureMessage, ${Pester Result}.StackTrace, $null, $NamedArguments, $PesterErrorRecord)
        $Pester.TestResult[-1] | Write-PesterResult
    }
}

function Get-StepParameters {
    <#
        .SYNOPSIS
        Internal function for determining parameters for a step implementation
        .PARAMETER Step
        The parsed step from the feature file

        .PARAMETER CommandName
        The text of the best matching step
    #>
    param(
        $Step,
        $CommandName
    )
    $Null = $Step.Text -match $CommandName

    $NamedArguments = @{}
    $Parameters = @{}
    foreach ($kv in $Matches.GetEnumerator()) {
        switch ($kv.Name -as [int]) {
            0 {
            } # toss zero (where it matches the whole string)
            $null {
                $NamedArguments.($kv.Name) = $ExecutionContext.InvokeCommand.ExpandString($kv.Value)
            }
            default {
                $Parameters.([int]$kv.Name) = $ExecutionContext.InvokeCommand.ExpandString($kv.Value)
            }
        }
    }
    $Parameters = @($Parameters.GetEnumerator() | & $SafeCommands["Sort-Object"] Name | & $SafeCommands["Select-Object"] -ExpandProperty Value)

    # TODO: Convert parsed tables to tables....
    if ($Step.Argument -is [Gherkin.Ast.DataTable]) {
        $NamedArguments.Table = $Step.Argument.Rows | ConvertTo-HashTableArray
    }
    if ($Step.Argument -is [Gherkin.Ast.DocString]) {
        # trim empty matches if we're attaching DocStringArgument
        $Parameters = @( $Parameters | & $SafeCommands["Where-Object"] { $_.Length } ) + $Step.Argument.Content
    }

    return @($NamedArguments, $Parameters)
}

function Convert-Tags {
    <#
        .SYNOPSIS
        Internal function for tagging Gherkin feature files (including inheritance from the feature)
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject,

        [Parameter(Position = 0)]
        [string[]]$BaseTags = @()
    )
    process {
        # Adapt the Gherkin .Tags property to the way we prefer it...
        [string[]]$Tags = foreach ($tag in $InputObject.Tags | & $SafeCommands['Where-Object'] { $_ }) {
            $tag.Name.TrimStart("@")
        }
        & $SafeCommands["Add-Member"] -MemberType NoteProperty -InputObject $InputObject -Name Tags -Value ([string[]]($Tags + $BaseTags)) -Force
        $InputObject
    }
}

function ConvertTo-HashTableArray {
    <#
        .SYNOPSIS
            Internal function for converting Gherkin AST tables to arrays of hashtables for splatting
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [Gherkin.Ast.TableRow[]]$InputObject
    )
    begin {
        ${Column Names} = @()
        ${Result Table} = @()
    }
    process {
        # Convert the first table row into headers:
        ${InputObject Rows} = @($InputObject)
        if (!${Column Names}) {
            & $SafeCommands["Write-Verbose"] "Reading Names from Header"
            ${InputObject Header}, ${InputObject Rows} = ${InputObject Rows}
            ${Column Names} = @(${InputObject Header}.Cells | & $SafeCommands["Select-Object"] -ExpandProperty Value)
        }

        if ( $null -ne ${InputObject Rows} ) {
            & $SafeCommands["Write-Verbose"] "Processing $(${InputObject Rows}.Length) Rows"
            foreach (${InputObject row} in ${InputObject Rows}) {
                ${Pester Result} = @{}
                for ($n = 0; $n -lt ${Column Names}.Length; $n++) {
                    ${Pester Result}.Add(${Column Names}[$n], ${InputObject row}.Cells[$n].Value)
                }
                ${Result Table} += @(${Pester Result})
            }
        }
    }
    end {
        ${Result Table}
    }
}

function Get-Translations($TranslationKey, $Language) {
    <#
        .SYNOPSIS
        Internal function to get all translations for a translation key and language

        .PARAMETER TranslationKey
        The key name inside the language in gherkin-languages.json, e.g. 'scenarioOutline'

        .PARAMETER Language
        The used language, e.g. 'en'

        .OUTPUTS
        System.String[] an array of all the translations
    #>
    if (-not (Test-Path variable:Script:GherkinLanguagesJson)) {
        $Script:GherkinLanguagesJson = ConvertFrom-Json2 (Get-Content "${Script:PesterRoot}/lib/Gherkin/gherkin-languages.json" | Out-String)
        # We override the fixed values for 'Describe' and 'Context' of Gherkin.psd1 or Output.ps1 since the language aware keywords
        # (e.g. 'Feature'/'Funktionalität' or 'Scenario'/'Szenario') are provided by Gherkin.dll and we do not want to duplicate them.
        $Script:ReportStrings.Describe = "{0}" # instead of 'Feature: {0}'  or 'Describing {0}'
        $Script:ReportStrings.Context = "{0}" # instead of 'Scenario: {0}' or 'Context {0}'
    }
    $foundTranslations = $Script:GherkinLanguagesJson."$Language"."$TranslationKey"
    if (-not $foundTranslations) {
        Write-Warning "Translation key '$TranslationKey' is invalid"
    }
    return , $foundTranslations
}

function ConvertFrom-Json2([string] $jsonString) {
    <#
        .SYNOPSIS
        Internal function to convert from JSON even for PowerShell 2

        .PARAMETER jsonString
        The JSON content as string

        .OUTPUTS
        the JSON content as array
    #>
    if ($PSVersionTable.PSVersion.Major -le 2) {
        # On PowerShell <= 2 we use JavaScriptSerializer
        Add-Type -Assembly System.Web.Extensions
        return , (New-Object System.Web.Script.Serialization.JavaScriptSerializer).DeserializeObject($jsonString)
    }
    else {
        # On PowerShell > 2 we use the built-in ConvertFrom-Json cmdlet
        return ConvertFrom-Json $jsonString
    }
}

function Get-Translation($TranslationKey, $Language, $Index = -1) {
    <#
        .SYNOPSIS
        Internal function to get the first translation for a translation key and language

        .PARAMETER TranslationKey
        The key name inside the language in gherkin-languages.json, e.g. 'scenarioOutline'

        .PARAMETER Language
        The used language, e.g. 'en'

        .PARAMETER Index
        The index in the array of JSON values
        If -1 is used for Index (the default value), this function will choose the most common translation of the JSON values

        .OUTPUTS
        System.String the chosen translation
    #>
    $translations = (Get-Translations $TranslationKey $Language)
    if (-not $translations) {
        return
    }
    if ($Index -lt 0 -or $Index -ge $translations.Length) {
        # Fallback: if the index is not in range, we choose the most common translation
        # Normally, the most common translation will be found at index one, but under some keys the index is zero.
        $Index = if ($TranslationKey -eq "scenarioOutline" -or $TranslationKey -eq "feature" -or $TranslationKey -eq "examples") {
            0
        }
        else {
            1
        }
    }
    return $translations[$Index]
}

function Test-Keyword($Keyword, $TranslationKey, $Language) {
    <#
        .SYNOPSIS
        Internal function to check if the given keyword matches one of the translations for a translation key and language

        .PARAMETER Keyword
        The keyword, e.g. 'Scenario Outline'

        .PARAMETER TranslationKey
        The key name inside the language in gherkin-languages.json, e.g. 'scenarioOutline'

        .PARAMETER Language
        The used language, e.g. 'en'

        .OUTPUTS
        System.Boolean true, if the keyword matches one of the translations, false otherwise
    #>
    return (Get-Translations $TranslationKey $Language) -contains $Keyword
}
