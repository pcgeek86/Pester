function Get-ShouldOperator {
    [CmdletBinding()]
    param ()

    # Use a dynamic parameter to create a dynamic ValidateSet
    # Define parameter -Name and tab-complete all current values of $AssertionOperators
    # Discovers included assertions (-Be, -Not) and any registered by the user via Add-AssertionOperator
    # https://martin77s.wordpress.com/2014/06/09/dynamic-validateset-in-a-dynamic-parameter/
    DynamicParam {
        $ParameterName = 'Name'

        $RuntimeParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        $ParameterAttribute = New-Object System.Management.Automation.ParameterAttribute

        $AttributeCollection.Add($ParameterAttribute)

        $arrSet = $AssertionOperators.Values |
            Select-Object -Property Name, Alias |
            ForEach-Object { $_.Name; $_.Alias }

        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($arrSet)

        $AttributeCollection.Add($ValidateSetAttribute)

        $RuntimeParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
        $RuntimeParameterDictionary.Add($ParameterName, $RuntimeParameter)
        return $RuntimeParameterDictionary
    }

    BEGIN {
        # Bind the parameter to a friendly variable
        $Name = $PsBoundParameters[$ParameterName]
    }

    END {
        if ($Name) {
            $operator = $AssertionOperators.Values | Where-Object { $Name -eq $_.Name -or $_.Alias -contains $Name }
            $help = Get-Help $operator.InternalName -Examples -ErrorAction SilentlyContinue

            if (($help | Measure-Object).Count -ne 1) {
                Write-Warning ("No help found for Should operator '{0}'" -f ((Get-AssertionOperatorEntry $Name).InternalName))
            }
            else {
                $help
            }
        }
        else {
            $AssertionOperators.Keys | ForEach-Object {
                $aliases = (Get-AssertionOperatorEntry $_).Alias

                # Return name and alias(es) for all registered Should operators
                New-Object -TypeName PSObject -Property @{
                    Name  = $_
                    Alias = $aliases -join ', '
                }
            }
        }
    }
}
