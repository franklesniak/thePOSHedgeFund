# .SYNOPSIS
# Demo PowerShell script to automate a trading strategy via QuantConnect.
#
# .DESCRIPTION
#   - Authenticates with QuantConnect API using user ID and API token.
#   - Creates a new project and adds an algorithm (C#).
#   - Compiles, backtests, and deploys live paper trading on QuantConnect.
#   - Logs results, including profit/loss, BTC comparison, alpha, beta.
#
# .PARAMETER AKVEntraIdTenantId
# Specifies the tenant ID to use when authenticating to Entra ID to access the Azure Key
# Vault. The default tenant ID is the one used in Frank and Blake's demo.
#
# .PARAMETER AKVSubscriptionId
# Specifies the subscription ID that holds the Azure Key Vault. The default subscription
# ID is the one used in Frank and Blake's demo.
#
# .PARAMETER AKVName
# Specifies the name of the Azure Key Vault that holds the API key (secret). The default
# Key Vault name is the one used in Frank and Danny's demo.
#
# .PARAMETER AKVUserIDSecretName
# Specifies the name of the QuantConnect user ID secret in the Azure Key Vault. The
# secret must contain the user ID authorized to access QuantConnect.
#
# .PARAMETER AKVTokenSecretName
# Specifies the name of the QuantConnect token secret in the Azure Key Vault. The
# secret must contain the token for the user authorized to access QuantConnect.
#
# .PARAMETER DoNotCheckForModuleUpdates
# If supplied, the script will skip the check for PowerShell module updates. This can
# speed up the script's execution time, but it is not recommended unless the user knows
# that the computer's modules are already up-to-date.
#
# .NOTES
# This script is for educational/demo purposes.
#
# Version: 0.1.20250507.0

#region License ###################################################################
# Copyright (c) 2025 Frank Lesniak and Blake Cherry
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of
# this software and associated documentation files (the "Software"), to deal in the
# Software without restriction, including without limitation the rights to use,
# copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
# Software, and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
# FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
# COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
# AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
# WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#endregion License ####################################################################

param(
    [string]$AKVEntraIdTenantId = '',
    [string]$AKVSubscriptionId = '',
    [string]$AKVName = '',
    [string]$AKVUserIDSecretName = '',
    [string]$AKVTokenSecretName = '',
    [string]$CSharpFilePath = '',
    [switch]$DoNotCheckForModuleUpdates = $false
)

#region Function Definitions #######################################################
function Get-PSVersion {
    # .SYNOPSIS
    # Returns the version of PowerShell that is running.
    #
    # .DESCRIPTION
    # The function outputs a [version] object representing the version of
    # PowerShell that is running.
    #
    # On versions of PowerShell greater than or equal to version 2.0, this
    # function returns the equivalent of $PSVersionTable.PSVersion
    #
    # PowerShell 1.0 does not have a $PSVersionTable variable, so this
    # function returns [version]('1.0') on PowerShell 1.0.
    #
    # .EXAMPLE
    # $versionPS = Get-PSVersion
    # # $versionPS now contains the version of PowerShell that is running.
    # # On versions of PowerShell greater than or equal to version 2.0,
    # # this function returns the equivalent of $PSVersionTable.PSVersion.
    #
    # .INPUTS
    # None. You can't pipe objects to Get-PSVersion.
    #
    # .OUTPUTS
    # System.Version. Get-PSVersion returns a [version] value indiciating
    # the version of PowerShell that is running.
    #
    # .NOTES
    # Version: 1.0.20250106.0

    #region License ####################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining
    # a copy of this software and associated documentation files (the
    # "Software"), to deal in the Software without restriction, including
    # without limitation the rights to use, copy, modify, merge, publish,
    # distribute, sublicense, and/or sell copies of the Software, and to
    # permit persons to whom the Software is furnished to do so, subject to
    # the following conditions:
    #
    # The above copyright notice and this permission notice shall be
    # included in all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
    # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
    # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
    # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
    # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ####################################################

    if (Test-Path variable:\PSVersionTable) {
        return ($PSVersionTable.PSVersion)
    } else {
        return ([version]('1.0'))
    }
}

function Get-PowerShellModuleUsingHashtable {
    # .SYNOPSIS
    # Gets a list of installed PowerShell modules for each entry in a hashtable.
    #
    # .DESCRIPTION
    # The Get-PowerShellModuleUsingHashtable function steps through each entry in
    # the supplied hashtable and gets a list of installed PowerShell modules for
    # each entry. The list of installed PowerShell modules for each entry is stored
    # in the value of the hashtable entry for that module (as an array).
    #
    # .PARAMETER ReferenceToHashtable
    # This parameter is required; it is a reference (memory pointer) to a
    # hashtable. The referenced hashtable must have keys that are the names of
    # PowerShell modules and values that are initialized to be enpty arrays (@()).
    # After running this function, the list of installed PowerShell modules for
    # each entry will is stored in the value of the hashtable entry as a populated
    # array.
    #
    # .PARAMETER DoNotCheckPowerShellVersion
    # This parameter is optional. If this switch is present, the function will not
    # check the version of PowerShell that is running. This is useful if you are
    # running this function in a script and the script has already validated that
    # the version of PowerShell supports Get-Module -ListAvailable.
    #
    # .EXAMPLE
    # $hashtableModuleNameToInstalledModules = @{}
    # $hashtableModuleNameToInstalledModules.Add('PnP.PowerShell', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Authentication', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Groups', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Users', @())
    # $refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
    # Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules
    #
    # This example gets the list of installed PowerShell modules for each of the
    # four modules listed in the hashtable. The list of each respective module is
    # stored in the value of the hashtable entry for that module.
    #
    # .INPUTS
    # None. You can't pipe objects to Get-PowerShellModuleUsingHashtable.
    #
    # .OUTPUTS
    # None. This function does not generate any output. The list of installed
    # PowerShell modules for each key in the referenced hashtable is stored in the
    # respective entry's value.
    #
    # .NOTES
    # This function also supports the use of a positional parameter instead of a
    # named parameter. If a positional parameter is used intead of named
    # parameters, then the first and only positional parameters must be a reference
    # (memory pointer) to a hashtable. The referenced hashtable must have keys that
    # are the names of PowerShell modules and values that are initialized to be
    # enpty arrays (@()). After running this function, the list of installed
    # PowerShell modules for each entry will is stored in the value of the
    # hashtable entry as a populated array.
    #
    # Version: 1.1.20250216.0

    #region License ############################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy
    # of this software and associated documentation files (the "Software"), to deal
    # in the Software without restriction, including without limitation the rights
    # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    # copies of the Software, and to permit persons to whom the Software is
    # furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in
    # all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ############################################################

    param (
        [ref]$ReferenceToHashtable = ([ref]$null),
        [switch]$DoNotCheckPowerShellVersion
    )

    function Get-PSVersion {
        # .SYNOPSIS
        # Returns the version of PowerShell that is running.
        #
        # .DESCRIPTION
        # The function outputs a [version] object representing the version of
        # PowerShell that is running.
        #
        # On versions of PowerShell greater than or equal to version 2.0, this
        # function returns the equivalent of $PSVersionTable.PSVersion
        #
        # PowerShell 1.0 does not have a $PSVersionTable variable, so this
        # function returns [version]('1.0') on PowerShell 1.0.
        #
        # .EXAMPLE
        # $versionPS = Get-PSVersion
        # # $versionPS now contains the version of PowerShell that is running.
        # # On versions of PowerShell greater than or equal to version 2.0,
        # # this function returns the equivalent of $PSVersionTable.PSVersion.
        #
        # .INPUTS
        # None. You can't pipe objects to Get-PSVersion.
        #
        # .OUTPUTS
        # System.Version. Get-PSVersion returns a [version] value indiciating
        # the version of PowerShell that is running.
        #
        # .NOTES
        # Version: 1.0.20250106.0

        #region License ####################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining
        # a copy of this software and associated documentation files (the
        # "Software"), to deal in the Software without restriction, including
        # without limitation the rights to use, copy, modify, merge, publish,
        # distribute, sublicense, and/or sell copies of the Software, and to
        # permit persons to whom the Software is furnished to do so, subject to
        # the following conditions:
        #
        # The above copyright notice and this permission notice shall be
        # included in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
        # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
        # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
        # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
        # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
        # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
        # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ####################################################

        if (Test-Path variable:\PSVersionTable) {
            return ($PSVersionTable.PSVersion)
        } else {
            return ([version]('1.0'))
        }
    }

    #region Process input ######################################################
    # Validate that the required parameter was supplied:
    if ($null -eq $ReferenceToHashtable) {
        $strMessage = 'The Get-PowerShellModuleUsingHashtable function requires a parameter (-ReferenceToHashtable), which must reference a hashtable.'
        Write-Error -Message $strMessage
        return
    }
    if ($null -eq $ReferenceToHashtable.Value) {
        $strMessage = 'The Get-PowerShellModuleUsingHashtable function requires a parameter (-ReferenceToHashtable), which must reference a hashtable.'
        Write-Error -Message $strMessage
        return
    }
    if ($ReferenceToHashtable.Value.GetType().FullName -ne 'System.Collections.Hashtable') {
        $strMessage = 'The Get-PowerShellModuleUsingHashtable function requires a parameter (-ReferenceToHashtable), which must reference a hashtable.'
        Write-Error -Message $strMessage
        return $false
    }

    $boolCheckForPowerShellVersion = $true
    if ($null -ne $DoNotCheckPowerShellVersion) {
        if ($DoNotCheckPowerShellVersion.IsPresent) {
            $boolCheckForPowerShellVersion = $false
        }
    }
    #endregion Process input ######################################################

    #region Verify environment #################################################
    if ($boolCheckForPowerShellVersion) {
        $versionPS = Get-PSVersion
        if ($versionPS.Major -lt 2) {
            Write-Error ('The Get-PowerShellModuleUsingHashtable function requires PowerShell version 2.0 or greater.')
            return
        }
    }
    #endregion Verify environment #################################################

    $VerbosePreferenceAtStartOfFunction = $VerbosePreference

    $arrModulesToGet = @(($ReferenceToHashtable.Value).Keys)
    $intCountOfModules = $arrModulesToGet.Count

    for ($intCounter = 0; $intCounter -lt $intCountOfModules; $intCounter++) {
        Write-Verbose ('Checking for ' + $arrModulesToGet[$intCounter] + ' module...')
        $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
        ($ReferenceToHashtable.Value).Item($arrModulesToGet[$intCounter]) = @(Get-Module -Name ($arrModulesToGet[$intCounter]) -ListAvailable)
        $VerbosePreference = $VerbosePreferenceAtStartOfFunction
    }

    return
}

function Test-PowerShellModuleInstalledUsingHashtable {
    # .SYNOPSIS
    # Tests to see if a PowerShell module is installed based on entries in a
    # hashtable. If the PowerShell module is not installed, an error or warning
    # message may optionally be displayed.
    #
    # .DESCRIPTION
    # The Test-PowerShellModuleInstalledUsingHashtable function steps through each
    # entry in the supplied hashtable and, if there are any modules not installed,
    # it optionally throws an error or warning for each module that is not
    # installed. If all modules are installed, the function returns $true;
    # otherwise, if any module is not installed, the function returns $false.
    #
    # .PARAMETER ReferenceToHashtableOfInstalledModules
    # This parameter is required; it is a reference to a hashtable. The hashtable
    # must have keys that are the names of PowerShell modules with each hashtable
    # entry's value (in the key-value pair) populated with arrays of
    # ModuleInfoGrouping objects (i.e., the object returned from Get-Module).
    #
    # .PARAMETER ReferenceToHashtableOfCustomNotInstalledMessages
    # This parameter is optional; if supplied, it is a reference to a hashtable.
    # The hashtable must have keys that are custom error or warning messages
    # (string) to be displayed if one or more modules are not installed. The value
    # for each key must be an array of PowerShell module names (strings) relevant
    # to that error or warning message.
    #
    # If this parameter is not supplied, or if a custom error or warning message is
    # not supplied in the hashtable for a given module, the script will default to
    # using the following message:
    #
    # <MODULENAME> module not found. Please install it and then try again.
    # You can install the <MODULENAME> PowerShell module from the PowerShell
    # Gallery by running the following command:
    # Install-Module <MODULENAME>;
    #
    # If the installation command fails, you may need to upgrade the version of
    # PowerShellGet. To do so, run the following commands, then restart PowerShell:
    # Set-ExecutionPolicy Bypass -Scope Process -Force;
    # [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
    # Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;
    # Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;
    #
    # .PARAMETER ThrowErrorIfModuleNotInstalled
    # This parameter is optional; it is a switch parameter. If this parameter is
    # specified, an error is thrown for each module that is not installed. If this
    # parameter is not specified, no error is thrown.
    #
    # .PARAMETER ThrowWarningIfModuleNotInstalled
    # This parameter is optional; it is a switch parameter. If this parameter is
    # specified, a warning is thrown for each module that is not installed. If this
    # parameter is not specified, or if the ThrowErrorIfModuleNotInstalled
    # parameter was specified, no warning is thrown.
    #
    # .PARAMETER ReferenceToArrayOfMissingModules
    # This parameter is optional; if supplied, it is a reference to an array. The
    # array must be initialized to be empty. If any modules are not installed, the
    # names of those modules are added to the array.
    #
    # .EXAMPLE
    # $hashtableModuleNameToInstalledModules = @{}
    # $hashtableModuleNameToInstalledModules.Add('PnP.PowerShell', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Authentication', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Groups', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Users', @())
    # $refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
    # Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules
    #
    # $hashtableCustomNotInstalledMessageToModuleNames = @{}
    # $strGraphNotInstalledMessage = 'Microsoft.Graph.Authentication, Microsoft.Graph.Groups, and/or Microsoft.Graph.Users modules were not found. Please install the full Microsoft.Graph module and then try again.' + [System.Environment]::NewLine + 'You can install the Microsoft.Graph PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module Microsoft.Graph;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    # $hashtableCustomNotInstalledMessageToModuleNames.Add($strGraphNotInstalledMessage, @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users'))
    # $refhashtableCustomNotInstalledMessageToModuleNames = [ref]$hashtableCustomNotInstalledMessageToModuleNames
    #
    # $boolResult = Test-PowerShellModuleInstalledUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames -ThrowErrorIfModuleNotInstalled
    #
    # This example checks to see if the PnP.PowerShell,
    # Microsoft.Graph.Authentication, Microsoft.Graph.Groups, and
    # Microsoft.Graph.Users modules are installed. If any of these modules are not
    # installed, an error is thrown for the PnP.PowerShell module or the group of
    # Microsoft.Graph modules, respectively, and $boolResult is set to $false. If
    # all modules are installed, $boolResult is set to $true.
    #
    # .INPUTS
    # None. You can't pipe objects to Test-PowerShellModuleInstalledUsingHashtable.
    #
    # .OUTPUTS
    # System.Boolean. Test-PowerShellModuleInstalledUsingHashtable returns a
    # boolean value indiciating whether all modules were installed. $true means
    # that every module specified in the referenced hashtable (i.e., the one
    # referenced in the ReferenceToHashtableOfInstalledModules parameter) was
    # installed; $false means that at least one module was not installed.
    #
    # .NOTES
    # Version: 2.0.20250216.0

    #region License ############################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy
    # of this software and associated documentation files (the "Software"), to deal
    # in the Software without restriction, including without limitation the rights
    # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    # copies of the Software, and to permit persons to whom the Software is
    # furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in
    # all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ############################################################

    param (
        [ref]$ReferenceToHashtableOfInstalledModules = ([ref]$null),
        [ref]$ReferenceToHashtableOfCustomNotInstalledMessages = ([ref]$null),
        [switch]$ThrowErrorIfModuleNotInstalled,
        [switch]$ThrowWarningIfModuleNotInstalled,
        [ref]$ReferenceToArrayOfMissingModules = ([ref]$null)
    )

    #region Process input ######################################################
    # Validate that the required parameter was supplied:
    if ($null -eq $ReferenceToHashtableOfInstalledModules) {
        $strMessage = 'The Test-PowerShellModuleUpdatesAvailableUsingHashtable function requires a parameter (-ReferenceToHashtable), which must reference a hashtable.'
        Write-Error -Message $strMessage
        return
    }
    if ($null -eq $ReferenceToHashtableOfInstalledModules.Value) {
        $strMessage = 'The Test-PowerShellModuleUpdatesAvailableUsingHashtable function requires a parameter (-ReferenceToHashtable), which must reference a hashtable.'
        Write-Error -Message $strMessage
        return
    }
    if ($ReferenceToHashtableOfInstalledModules.Value.GetType().FullName -ne 'System.Collections.Hashtable') {
        $strMessage = 'The Test-PowerShellModuleUpdatesAvailableUsingHashtable function requires a parameter (-ReferenceToHashtable), which must reference a hashtable.'
        Write-Error -Message $strMessage
        return
    }

    $boolThrowErrorForMissingModule = $false
    if ($null -ne $ThrowErrorIfModuleNotInstalled) {
        if ($ThrowErrorIfModuleNotInstalled.IsPresent) {
            $boolThrowErrorForMissingModule = $true
        }
    }
    $boolThrowWarningForMissingModule = $false
    if (-not $boolThrowErrorForMissingModule) {
        if ($null -ne $ThrowWarningIfModuleNotInstalled) {
            if ($ThrowWarningIfModuleNotInstalled.IsPresent) {
                $boolThrowWarningForMissingModule = $true
            }
        }
    }
    #endregion Process input ######################################################

    $boolResult = $true

    $hashtableMessagesToThrowForMissingModule = @{}
    $hashtableModuleNameToCustomMessageToThrowForMissingModule = @{}
    if ($null -ne $ReferenceToHashtableOfCustomNotInstalledMessages) {
        if ($null -ne $ReferenceToHashtableOfCustomNotInstalledMessages.Value) {
            if ($ReferenceToHashtableOfCustomNotInstalledMessages.Value.GetType().FullName -ne 'System.Collections.Hashtable') {
                $arrMessages = @(($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Keys)
                foreach ($strMessage in $arrMessages) {
                    $hashtableMessagesToThrowForMissingModule.Add($strMessage, $false)

                    ($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Item($strMessage) | ForEach-Object {
                        $hashtableModuleNameToCustomMessageToThrowForMissingModule.Add($_, $strMessage)
                    }
                }
            }
        }
    }

    $arrModuleNames = @(($ReferenceToHashtableOfInstalledModules.Value).Keys)
    foreach ($strModuleName in $arrModuleNames) {
        $arrInstalledModules = @(($ReferenceToHashtableOfInstalledModules.Value).Item($strModuleName))
        if ($arrInstalledModules.Count -eq 0) {
            $boolResult = $false

            if ($hashtableModuleNameToCustomMessageToThrowForMissingModule.ContainsKey($strModuleName) -eq $true) {
                $strMessage = $hashtableModuleNameToCustomMessageToThrowForMissingModule.Item($strModuleName)
                $hashtableMessagesToThrowForMissingModule.Item($strMessage) = $true
            } else {
                $strMessage = $strModuleName + ' module not found. Please install it and then try again.' + [System.Environment]::NewLine + 'You can install the ' + $strModuleName + ' PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module ' + $strModuleName + ';' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                $hashtableMessagesToThrowForMissingModule.Add($strMessage, $true)
            }

            if ($null -ne $ReferenceToArrayOfMissingModules) {
                if ($null -ne $ReferenceToArrayOfMissingModules.Value) {
                    ($ReferenceToArrayOfMissingModules.Value) += $strModuleName
                }
            }
        }
    }

    if ($boolThrowErrorForMissingModule) {
        $arrMessages = @($hashtableMessagesToThrowForMissingModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingModule.Item($strMessage)) {
                Write-Error $strMessage
            }
        }
    } elseif ($boolThrowWarningForMissingModule) {
        $arrMessages = @($hashtableMessagesToThrowForMissingModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingModule.Item($strMessage)) {
                Write-Warning $strMessage
            }
        }
    }

    return $boolResult
}

function Test-PowerShellModuleUpdatesAvailableUsingHashtable {
    # .SYNOPSIS
    # Tests to see if updates are available for a PowerShell module based on
    # entries in a hashtable. If updates are available for a PowerShell module, an
    # error or warning message may optionally be displayed.
    #
    # .DESCRIPTION
    # The Test-PowerShellModuleUpdatesAvailableUsingHashtable function steps
    # through each entry in the supplied hashtable and, if there are updates
    # available, it optionally throws an error or warning for each module that has
    # updates available. If all modules are installed and up to date, the function
    # returns $true; otherwise, if any module is not installed or not up to date,
    # the function returns $false.
    #
    # .PARAMETER ReferenceToHashtableOfInstalledModules
    # This parameter is required; it is a reference to a hashtable. The hashtable
    # must have keys that are the names of PowerShell modules with each key's value
    # populated with arrays of ModuleInfoGrouping objects (the result of
    # Get-Module).
    #
    # .PARAMETER ThrowErrorIfModuleNotInstalled
    # This parameter is optional; if supplied, an error is thrown for each module
    # that is not installed. If this parameter is not specified, no error is
    # thrown.
    #
    # .PARAMETER ThrowWarningIfModuleNotInstalled
    # This parameter is optional; if supplied, a warning is thrown for each module
    # that is not installed. If this parameter is not specified, or if the
    # ThrowErrorIfModuleNotInstalled parameter was specified, no warning is thrown.
    #
    # .PARAMETER ThrowErrorIfModuleNotUpToDate
    # This parameter is optional; if supplied, an error is thrown for each module
    # that is not up to date. If this parameter is not specified, no error is
    # thrown.
    #
    # .PARAMETER ThrowWarningIfModuleNotUpToDate
    # This parameter is optional; if supplied, a warning is thrown for each module
    # that is not up to date. If this parameter is not specified, or if the
    # ThrowErrorIfModuleNotUpToDate parameter was specified, no warning is thrown.
    #
    # .PARAMETER ReferenceToHashtableOfCustomNotInstalledMessages
    # This parameter is optional; if supplied, it is a reference to a hashtable.
    # The hashtable must have keys that are custom error or warning messages
    # (each key is a string object) to be displayed if one or more modules are not
    # installed. The value for each key must be an array of PowerShell module names
    # (strings) relevant to that error or warning message.
    #
    # If this parameter is not supplied, or if a custom error or warning message is
    # not supplied in the hashtable for a given module, the script will default to
    # using the following message:
    #
    # <MODULENAME> module not found. Please install it and then try again.
    # You can install the <MODULENAME> PowerShell module from the PowerShell
    # Gallery by running the following command:
    # Install-Module <MODULENAME>;
    #
    # If the installation command fails, you may need to upgrade the version of
    # PowerShellGet. To do so, run the following commands, then restart PowerShell:
    # Set-ExecutionPolicy Bypass -Scope Process -Force;
    # [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
    # Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;
    # Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;
    #
    # .PARAMETER ReferenceToHashtableOfCustomNotUpToDateMessages
    # This parameter is optional; if supplied, it is a reference to a hashtable.
    # The hashtable must have keys that are custom error or warning messages
    # (string) to be displayed if one or more modules are not up to date. The value
    # for each key must be an array of PowerShell module names (strings) relevant
    # to that error or warning message.
    #
    # If this parameter is not supplied, or if a custom error or warning message is
    # not supplied in the hashtable for a given module, the script will default to
    # using the following message:
    #
    # A newer version of the <MODULENAME> PowerShell module is available. Please
    # consider updating it by running the following command:
    # Install-Module <MODULENAME> -Force;
    #
    # If the installation command fails, you may need to upgrade the version of
    # PowerShellGet. To do so, run the following commands, then restart PowerShell:
    # Set-ExecutionPolicy Bypass -Scope Process -Force;
    # [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
    # Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;
    # Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;
    #
    # .PARAMETER ReferenceToArrayOfMissingModules
    # This parameter is optional; if supplied, it is a reference to an array. The
    # array must be initialized to be empty. If any modules are not installed, the
    # names of those modules are added to the array.
    #
    # .PARAMETER ReferenceToArrayOfOutOfDateModules
    # This parameter is optional; if supplied, it is a reference to an array. The
    # array must be initialized to be empty. If any modules are not up to date, the
    # names of those modules are added to the array.
    #
    # .PARAMETER DoNotCheckPowerShellVersion
    # This parameter is optional. If this switch is present, the function will not
    # check the version of PowerShell that is running. This is useful if you are
    # running this function in a script and the script has already validated that
    # the version of PowerShell supports Find-Module.
    #
    # .EXAMPLE
    # $hashtableModuleNameToInstalledModules = @{}
    # $hashtableModuleNameToInstalledModules.Add('PnP.PowerShell', @())
    # $refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
    # Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules
    #
    # $hashtableCustomNotInstalledMessageToModuleNames = @{}
    # $refhashtableCustomNotInstalledMessageToModuleNames = [ref]$hashtableCustomNotInstalledMessageToModuleNames
    #
    # $hashtableCustomNotUpToDateMessageToModuleNames = @{}
    # $refhashtableCustomNotUpToDateMessageToModuleNames = [ref]$hashtableCustomNotUpToDateMessageToModuleNames
    #
    # $boolResult = Test-PowerShellModuleUpdatesAvailableUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ThrowWarningIfModuleNotUpToDate -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames -ReferenceToHashtableOfCustomNotUpToDateMessages $refhashtableCustomNotUpToDateMessageToModuleNames
    #
    # This example checks to see if the PnP.PowerShell module is installed. If it
    # is not installed, an error is thrown and $boolResult is set to $false. If it
    # is installed but not up to date, a warning message is thrown and $boolResult
    # is set to false. If PnP.PowerShell is installed and up to date, $boolResult
    # is set to $true.
    #
    # .EXAMPLE
    # $hashtableModuleNameToInstalledModules = @{}
    # $hashtableModuleNameToInstalledModules.Add('PnP.PowerShell', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Authentication', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Groups', @())
    # $hashtableModuleNameToInstalledModules.Add('Microsoft.Graph.Users', @())
    # $refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
    # Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules
    #
    # $hashtableCustomNotInstalledMessageToModuleNames = @{}
    # $strGraphNotInstalledMessage = 'Microsoft.Graph.Authentication, Microsoft.Graph.Groups, and/or Microsoft.Graph.Users modules were not found. Please install the full Microsoft.Graph module and then try again.' + [System.Environment]::NewLine + 'You can install the Microsoft.Graph PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module Microsoft.Graph;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    # $hashtableCustomNotInstalledMessageToModuleNames.Add($strGraphNotInstalledMessage, @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users'))
    # $refhashtableCustomNotInstalledMessageToModuleNames = [ref]$hashtableCustomNotInstalledMessageToModuleNames
    #
    # $hashtableCustomNotUpToDateMessageToModuleNames = @{}
    # $strGraphNotUpToDateMessage = 'A newer version of the Microsoft.Graph.Authentication, Microsoft.Graph.Groups, and/or Microsoft.Graph.Users modules was found. Please consider updating it by running the following command:' + [System.Environment]::NewLine + 'Install-Module Microsoft.Graph -Force;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    # $hashtableCustomNotUpToDateMessageToModuleNames.Add($strGraphNotUpToDateMessage, @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users'))
    # $refhashtableCustomNotUpToDateMessageToModuleNames = [ref]$hashtableCustomNotUpToDateMessageToModuleNames
    #
    # $boolResult = Test-PowerShellModuleUpdatesAvailableUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ThrowWarningIfModuleNotUpToDate -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames -ReferenceToHashtableOfCustomNotUpToDateMessages $refhashtableCustomNotUpToDateMessageToModuleNames
    #
    # This example checks to see if the PnP.PowerShell,
    # Microsoft.Graph.Authentication, Microsoft.Graph.Groups, and
    # Microsoft.Graph.Users modules are installed. If any of these modules are not
    # installed, an error is thrown for the PnP.PowerShell module or the group of
    # Microsoft.Graph modules, respectively, and $boolResult is set to $false. If
    # any of these modules are installed but not up to date, a warning message is
    # thrown for the PnP.PowerShell module or the group of Microsoft.Graph modules,
    # respectively, and $boolResult is set to false. If all modules are installed
    # and up to date, $boolResult is set to $true.
    #
    # .INPUTS
    # None. You can't pipe objects to
    # Test-PowerShellModuleUpdatesAvailableUsingHashtable.
    #
    # .OUTPUTS
    # System.Boolean. Test-PowerShellModuleUpdatesAvailableUsingHashtable returns a
    # boolean value indiciating whether all modules are installed and up to date.
    # If all modules are installed and up to date, the function returns $true;
    # otherwise, if any module is not installed or not up to date, the function
    # returns $false.
    #
    # .NOTES
    # Requires PowerShell v5.0 or later.
    #
    # Version: 2.1.20250218.0

    #region License ############################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy
    # of this software and associated documentation files (the "Software"), to deal
    # in the Software without restriction, including without limitation the rights
    # to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    # copies of the Software, and to permit persons to whom the Software is
    # furnished to do so, subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in
    # all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ############################################################

    param (
        [ref]$ReferenceToHashtableOfInstalledModules = ([ref]$null),
        [switch]$ThrowErrorIfModuleNotInstalled,
        [switch]$ThrowWarningIfModuleNotInstalled,
        [switch]$ThrowErrorIfModuleNotUpToDate,
        [switch]$ThrowWarningIfModuleNotUpToDate,
        [ref]$ReferenceToHashtableOfCustomNotInstalledMessages = ([ref]$null),
        [ref]$ReferenceToHashtableOfCustomNotUpToDateMessages = ([ref]$null),
        [ref]$ReferenceToArrayOfMissingModules = ([ref]$null),
        [ref]$ReferenceToArrayOfOutOfDateModules = ([ref]$null),
        [switch]$DoNotCheckPowerShellVersion
    )

    function Get-PSVersion {
        # .SYNOPSIS
        # Returns the version of PowerShell that is running.
        #
        # .DESCRIPTION
        # The function outputs a [version] object representing the version of
        # PowerShell that is running.
        #
        # On versions of PowerShell greater than or equal to version 2.0, this
        # function returns the equivalent of $PSVersionTable.PSVersion
        #
        # PowerShell 1.0 does not have a $PSVersionTable variable, so this
        # function returns [version]('1.0') on PowerShell 1.0.
        #
        # .EXAMPLE
        # $versionPS = Get-PSVersion
        # # $versionPS now contains the version of PowerShell that is running.
        # # On versions of PowerShell greater than or equal to version 2.0,
        # # this function returns the equivalent of $PSVersionTable.PSVersion.
        #
        # .INPUTS
        # None. You can't pipe objects to Get-PSVersion.
        #
        # .OUTPUTS
        # System.Version. Get-PSVersion returns a [version] value indiciating
        # the version of PowerShell that is running.
        #
        # .NOTES
        # Version: 1.0.20250106.0

        #region License ####################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining
        # a copy of this software and associated documentation files (the
        # "Software"), to deal in the Software without restriction, including
        # without limitation the rights to use, copy, modify, merge, publish,
        # distribute, sublicense, and/or sell copies of the Software, and to
        # permit persons to whom the Software is furnished to do so, subject to
        # the following conditions:
        #
        # The above copyright notice and this permission notice shall be
        # included in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
        # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
        # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
        # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
        # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
        # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
        # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ####################################################

        if (Test-Path variable:\PSVersionTable) {
            return ($PSVersionTable.PSVersion)
        } else {
            return ([version]('1.0'))
        }
    }

    function Convert-StringToFlexibleVersion {
        # .SYNOPSIS
        # Converts a string to a version object. However, when the string contains
        # characters not allowed in a version object, this function will attempt to
        # convert the string to a version object by removing the characters that are
        # not allowed, identifying the portions of the version object that are
        # not allowed, which can be evaluated further if needed.
        #
        # .DESCRIPTION
        # First attempts to convert a string to a version object. If the string
        # contains characters not allowed in a version object, this function will
        # iteratively attempt to convert the string to a version object by removing
        # period-separated substrings, working right to left, until the version is
        # successfully converted. Then, for the portions that could not be
        # converted, the function will select the numerical-only portions of the
        # problematic substrings and use those to generate a "best effort" version
        # object. The leftover portions of the substrings that could not be
        # converted will be returned by reference.
        #
        # .PARAMETER ReferenceToVersionObject
        # This parameter is required; it is a reference to a System.Version object
        # that will be used to store the version object that is generated from the
        # string. If the string is successfully converted to a version object, the
        # version object will be stored in this reference. If one or more portions
        # of the string could not be converted to a version object, the version
        # object will be generated from the portions that could be converted, and
        # the portions that could not be converted will be stored in the
        # other reference parameters.
        #
        # .PARAMETER ReferenceArrayOfLeftoverStrings
        # This parameter is required; it is a reference to an array of five
        # elements. Each element is a string; One or more of the elements may be
        # modified if the string could not be converted to a version object. If the
        # string could not be converted to a version object, any portions of the
        # string that exceed the major, minor, build, and revision version portions
        # will be stored in the elements of the array.
        #
        # The first element of the array will be modified if the major version
        # portion of the string could not be converted to a version object. If the
        # major version portion of the string could not be converted to a version
        # object, the left-most numerical-only portion of the major version will be
        # used to generate the version object. The remaining portion of the major
        # version will be stored in the first element of the array.
        #
        # The second element of the array will be modified if the minor version
        # portion of the string could not be converted to a version object. If the
        # minor version portion of the string could not be converted to a version
        # object, the left-most numerical-only portion of the minor version will be
        # used to generate the version object. The remaining portion of the minor
        # version will be stored in second element of the array.
        #
        # If the major version portion of the string could not be converted to a
        # version object, the entire minor version portion of the string will be
        # stored in the second element, and no portion of the supplied minor
        # version reference will be used to generate the version object.
        #
        # The third element of the array will be modified if the build version
        # portion of the string could not be converted to a version object. If the
        # build version portion of the string could not be converted to a version
        # object, the left-most numerical-only portion of the build version will be
        # used to generate the version object. The remaining portion of the build
        # version will be stored in the third element of the array.
        #
        # If the major or minor version portions of the string could not be
        # converted to a version object, the entire build version portion of the
        # string will be stored in the third element, and no portion of the
        # supplied build version reference will be used to generate the version
        # object.
        #
        # The fourth element of the array will be modified if the revision version
        # portion of the string could not be converted to a version object. If the
        # revision version portion of the string could not be converted to a
        # version object, the left-most numerical-only portion of the revision
        # version will be used to generate the version object. The remaining
        # portion of the revision version will be stored in the fourth element of
        # the array.
        #
        # If the major, minor, or build version portions of the string could not be
        # converted to a version object, the entire revision version portion of the
        # string will be stored in the fourth element, and no portion of the
        # supplied revision version reference will be used to generate the version
        # object.
        #
        # The fifth element of the array will be modified if the string could not
        # be converted to a version object. If the string could not be converted to
        # a version object, any portions of the string that exceed the major,
        # minor, build, and revision version portions will be stored in the string
        # reference.
        #
        # For example, if the string is '1.2.3.4.5', the fifth element in the array
        # will be '5'. If the string is '1.2.3.4.5.6', the fifth element of the
        # array will be '5.6'.
        #
        # .PARAMETER StringToConvert
        # This parameter is required; it is string that will be converted to a
        # version object. If the string contains characters not allowed in a
        # version object, this function will attempt to convert the string to a
        # version object by removing the characters that are not allowed,
        # identifying the portions of the version object that are not allowed,
        # which can be evaluated further if needed.
        #
        # .PARAMETER PSVersion
        # This parameter is optional; it is a version object that represents the
        # version of PowerShell that is running the script. If this parameter is
        # supplied, it will improve the performance of the function by allowing it
        # to skip the determination of the PowerShell engine version.
        #
        # .EXAMPLE
        # $version = $null
        # $arrLeftoverStrings = @('', '', '', '', '')
        # $strVersion = '1.2.3.4'
        # $intReturnCode = Convert-StringToFlexibleVersion -ReferenceToVersionObject ([ref]$version) -ReferenceArrayOfLeftoverStrings ([ref]$arrLeftoverStrings) -StringToConvert $strVersion
        # # $intReturnCode will be 0 because the string is in a valid format for a
        # # version object.
        # # $version will be a System.Version object with Major=1, Minor=2,
        # # Build=3, Revision=4.
        # # All strings in $arrLeftoverStrings will be empty.
        #
        # .EXAMPLE
        # $version = $null
        # $arrLeftoverStrings = @('', '', '', '', '')
        # $strVersion = '1.2.3.4-beta3'
        # $intReturnCode = Convert-StringToFlexibleVersion -ReferenceToVersionObject ([ref]$version) -ReferenceArrayOfLeftoverStrings ([ref]$arrLeftoverStrings) -StringToConvert $strVersion
        # # $intReturnCode will be 4 because the string is not in a valid format
        # # for a version object. The 4 indicates that the revision version portion
        # # of the string could not be converted to a version object.
        # # $version will be a System.Version object with Major=1, Minor=2,
        # # Build=3, Revision=4.
        # # $arrLeftoverStrings[3] will be '-beta3'. All other elements of
        # # $arrLeftoverStrings will be empty.
        #
        # .EXAMPLE
        # $version = $null
        # $arrLeftoverStrings = @('', '', '', '', '')
        # $strVersion = '1.2.2147483700.4'
        # $intReturnCode = Convert-StringToFlexibleVersion -ReferenceToVersionObject ([ref]$version) -ReferenceArrayOfLeftoverStrings ([ref]$arrLeftoverStrings) -StringToConvert $strVersion
        # # $intReturnCode will be 3 because the string is not in a valid format
        # # for a version object. The 3 indicates that the build version portion of
        # # the string could not be converted to a version object (the value
        # # exceeds the maximum value for a version element - 2147483647).
        # # $version will be a System.Version object with Major=1, Minor=2,
        # # Build=2147483647, Revision=-1.
        # # $arrLeftoverStrings[2] will be '53' (2147483700 - 2147483647) and
        # # $arrLeftoverStrings[3] will be '4'. All other elements of
        # # $arrLeftoverStrings will be empty.
        #
        # .EXAMPLE
        # $version = $null
        # $arrLeftoverStrings = @('', '', '', '', '')
        # $strVersion = '1.2.2147483700-beta5.4'
        # $intReturnCode = Convert-StringToFlexibleVersion -ReferenceToVersionObject ([ref]$version) -ReferenceArrayOfLeftoverStrings ([ref]$arrLeftoverStrings) -StringToConvert $strVersion
        # # $intReturnCode will be 3 because the string is not in a valid format
        # # for a version object. The 3 indicates that the build version portion of
        # # the string could not be converted to a version object (the value
        # # exceeds the maximum value for a version element - 2147483647).
        # # $version will be a System.Version object with Major=1, Minor=2,
        # # Build=2147483647, Revision=-1.
        # # $arrLeftoverStrings[2] will be '53-beta5' (2147483700 - 2147483647)
        # # plus the non-numeric portion of the string ('-beta5') and
        # # $arrLeftoverStrings[3] will be '4'. All other elements of
        # # $arrLeftoverStrings will be empty.
        #
        # .EXAMPLE
        # $version = $null
        # $arrLeftoverStrings = @('', '', '', '', '')
        # $strVersion = '1.2.3.4.5'
        # $intReturnCode = Convert-StringToFlexibleVersion -ReferenceToVersionObject ([ref]$version) -ReferenceArrayOfLeftoverStrings ([ref]$arrLeftoverStrings) -StringToConvert $strVersion
        # # $intReturnCode will be 5 because the string is in a valid format for a
        # # version object. The 5 indicates that there were excess portions of the
        # # string that could not be converted to a version object.
        # # $version will be a System.Version object with Major=1, Minor=2,
        # # Build=3, Revision=4.
        # # $arrLeftoverStrings[4] will be '5'. All other elements of
        # # $arrLeftoverStrings will be empty.
        #
        # .INPUTS
        # None. You can't pipe objects to Convert-StringToFlexibleVersion.
        #
        # .OUTPUTS
        # System.Int32. Convert-StringToFlexibleVersion returns an integer value
        # indicating whether the string was successfully converted to a version
        # object. The return value is as follows:
        # 0: The string was successfully converted to a version object.
        # 1: The string could not be converted to a version object because the
        #    major version portion of the string contained characters that made it
        #    impossible to convert to a version object. With these characters
        #    removed, the major version portion of the string was converted to a
        #    version object.
        # 2: The string could not be converted to a version object because the
        #    minor version portion of the string contained characters that made it
        #    impossible to convert to a version object. With these characters
        #    removed, the minor version portion of the string was converted to a
        #    version object.
        # 3: The string could not be converted to a version object because the
        #    build version portion of the string contained characters that made it
        #    impossible to convert to a version object. With these characters
        #    removed, the build version portion of the string was converted to a
        #    version object.
        # 4: The string could not be converted to a version object because the
        #    revision version portion of the string contained characters that made
        #    it impossible to convert to a version object. With these characters
        #    removed, the revision version portion of the string was converted to a
        #    version object.
        # 5: The string was successfully converted to a version object, but there
        #    were excess portions of the string that could not be converted to a
        #    version object.
        # -1: The string could not be converted to a version object because the
        #     string did not begin with numerical characters.
        #
        # .NOTES
        # This function also supports the use of positional parameters instead of
        # named parameters. If positional parameters are used instead of named
        # parameters, then three or four positional parameters are required:
        #
        # The first positional parameter is a reference to a System.Version object
        # that will be used to store the version object that is generated from the
        # string. If the string is successfully converted to a version object, the
        # version object will be stored in this reference. If one or more portions
        # of the string could not be converted to a version object, the version
        # object will be generated from the portions that could be converted, and
        # the portions that could not be converted will be stored in the
        # other reference parameters.
        #
        # The second positional parameter is a reference to an array of five
        # elements. Each element is a string; One or more of the elements may be
        # modified if the string could not be converted to a version object. If the
        # string could not be converted to a version object, any portions of the
        # string that exceed the major, minor, build, and revision version portions
        # will be stored in the elements of the array.
        #
        # The first element of the array will be modified if the major version
        # portion of the string could not be converted to a version object. If the
        # major version portion of the string could not be converted to a version
        # object, the left-most numerical-only portion of the major version will be
        # used to generate the version object. The remaining portion of the major
        # version will be stored in the first element of the array.
        #
        # The second element of the array will be modified if the minor version
        # portion of the string could not be converted to a version object. If the
        # minor version portion of the string could not be converted to a version
        # object, the left-most numerical-only portion of the minor version will be
        # used to generate the version object. The remaining portion of the minor
        # version will be stored in second element of the array.
        #
        # If the major version portion of the string could not be converted to a
        # version object, the entire minor version portion of the string will be
        # stored in the second element, and no portion of the supplied minor
        # version reference will be used to generate the version object.
        #
        # The third element of the array will be modified if the build version
        # portion of the string could not be converted to a version object. If the
        # build version portion of the string could not be converted to a version
        # object, the left-most numerical-only portion of the build version will be
        # used to generate the version object. The remaining portion of the build
        # version will be stored in the third element of the array.
        #
        # If the major or minor version portions of the string could not be
        # converted to a version object, the entire build version portion of the
        # string will be stored in the third element, and no portion of the
        # supplied build version reference will be used to generate the version
        # object.
        #
        # The fourth element of the array will be modified if the revision version
        # portion of the string could not be converted to a version object. If the
        # revision version portion of the string could not be converted to a
        # version object, the left-most numerical-only portion of the revision
        # version will be used to generate the version object. The remaining
        # portion of the revision version will be stored in the fourth element of
        # the array.
        #
        # If the major, minor, or build version portions of the string could not be
        # converted to a version object, the entire revision version portion of the
        # string will be stored in the fourth element, and no portion of the
        # supplied revision version reference will be used to generate the version
        # object.
        #
        # The fifth element of the array will be modified if the string could not
        # be converted to a version object. If the string could not be converted to
        # a version object, any portions of the string that exceed the major,
        # minor, build, and revision version portions will be stored in the string
        # reference.
        #
        # For example, if the string is '1.2.3.4.5', the fifth element in the array
        # will be '5'. If the string is '1.2.3.4.5.6', the fifth element of the
        # array will be '5.6'.
        #
        # The third positional parameter is string that will be converted to a
        # version object. If the string contains characters not allowed in a
        # version object, this function will attempt to convert the string to a
        # version object by removing the characters that are not allowed,
        # identifying the portions of the version object that are not allowed,
        # which can be evaluated further if needed.
        #
        # If supplied, the fourth positional parameter is a version object that
        # represents the version of PowerShell that is running the script. If this
        # parameter is supplied, it will improve the performance of the function by
        # allowing it to skip the determination of the PowerShell engine version.
        #
        # Version: 1.0.20250218.0

        #region License ########################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining a
        # copy of this software and associated documentation files (the
        # "Software"), to deal in the Software without restriction, including
        # without limitation the rights to use, copy, modify, merge, publish,
        # distribute, sublicense, and/or sell copies of the Software, and to permit
        # persons to whom the Software is furnished to do so, subject to the
        # following conditions:
        #
        # The above copyright notice and this permission notice shall be included
        # in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
        # OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
        # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN
        # NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
        # DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
        # OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
        # USE OR OTHER DEALINGS IN THE SOFTWARE.
        #endregion License ########################################################

        param (
            [ref]$ReferenceToVersionObject = ([ref]$null),
            [ref]$ReferenceArrayOfLeftoverStrings = ([ref]$null),
            [string]$StringToConvert = '',
            [version]$PSVersion = ([version]'0.0')
        )

        function Convert-StringToVersionSafely {
            # .SYNOPSIS
            # Attempts to convert a string to a System.Version object.
            #
            # .DESCRIPTION
            # Attempts to convert a string to a System.Version object. If the
            # string cannot be converted to a System.Version object, the function
            # suppresses the error and returns $false. If the string can be
            # converted to a version object, the function returns $true and passes
            # the version object by reference to the caller.
            #
            # .PARAMETER ReferenceToVersionObject
            # This parameter is required; it is a reference to a System.Version
            # object that will be used to store the converted version object if the
            # conversion is successful.
            #
            # .PARAMETER StringToConvert
            # This parameter is required; it is a string that is to be converted to
            # a System.Version object.
            #
            # .EXAMPLE
            # $version = $null
            # $strVersion = '1.2.3.4'
            # $boolSuccess = Convert-StringToVersionSafely -ReferenceToVersionObject ([ref]$version) -StringToConvert $strVersion
            # # $boolSuccess will be $true, indicating that the conversion was
            # # successful.
            # # $version will contain a System.Version object with major version 1,
            # # minor version 2, build version 3, and revision version 4.
            #
            # .EXAMPLE
            # $version = $null
            # $strVersion = '1'
            # $boolSuccess = Convert-StringToVersionSafely -ReferenceToVersionObject ([ref]$version) -StringToConvert $strVersion
            # # $boolSuccess will be $false, indicating that the conversion was
            # # unsuccessful.
            # # $version is undefined in this instance.
            #
            # .INPUTS
            # None. You can't pipe objects to Convert-StringToVersionSafely.
            #
            # .OUTPUTS
            # System.Boolean. Convert-StringToVersionSafely returns a boolean value
            # indiciating whether the process completed successfully. $true means
            # the conversion completed successfully; $false means there was an
            # error.
            #
            # .NOTES
            # This function also supports the use of positional parameters instead
            # of named parameters. If positional parameters are used instead of
            # named parameters, then two positional parameters are required:
            #
            # The first positional parameter is a reference to a System.Version
            # object that will be used to store the converted version object if the
            # conversion is successful.
            #
            # The second positional parameter is a string that is to be converted
            # to a System.Version object.
            #
            # Version: 1.0.20250215.0

            #region License ####################################################
            # Copyright (c) 2025 Frank Lesniak
            #
            # Permission is hereby granted, free of charge, to any person obtaining
            # a copy of this software and associated documentation files (the
            # "Software"), to deal in the Software without restriction, including
            # without limitation the rights to use, copy, modify, merge, publish,
            # distribute, sublicense, and/or sell copies of the Software, and to
            # permit persons to whom the Software is furnished to do so, subject to
            # the following conditions:
            #
            # The above copyright notice and this permission notice shall be
            # included in all copies or substantial portions of the Software.
            #
            # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
            # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
            # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
            # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
            # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
            # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
            # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
            # SOFTWARE.
            #endregion License ####################################################

            param (
                [ref]$ReferenceToVersionObject = ([ref]$null),
                [string]$StringToConvert = ''
            )

            #region FunctionsToSupportErrorHandling ############################
            function Get-ReferenceToLastError {
                # .SYNOPSIS
                # Gets a reference (memory pointer) to the last error that
                # occurred.
                #
                # .DESCRIPTION
                # Returns a reference (memory pointer) to $null ([ref]$null) if no
                # errors on on the $error stack; otherwise, returns a reference to
                # the last error that occurred.
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work:
                # $refLastKnownError = Get-ReferenceToLastError
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will suppress
                # # error output. Terminating errors will not output anything, kick
                # # to the empty trap statement and then continue on. Likewise, non-
                # # terminating errors will also not output anything, but they do not
                # # kick to the trap statement; they simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # $refNewestCurrentError = Get-ReferenceToLastError
                #
                # $boolErrorOccurred = $false
                # if (($null -ne $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #     # Both not $null
                #     if (($refLastKnownError.Value) -ne ($refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # } else {
                #     # One is $null, or both are $null
                #     # NOTE: $refLastKnownError could be non-null, while
                #     # $refNewestCurrentError could be null if $error was cleared;
                #     # this does not indicate an error.
                #     #
                #     # So:
                #     # If both are null, no error.
                #     # If $refLastKnownError is null and $refNewestCurrentError is
                #     # non-null, error.
                #     # If $refLastKnownError is non-null and $refNewestCurrentError
                #     # is null, no error.
                #     #
                #     if (($null -eq $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Get-ReferenceToLastError.
                #
                # .OUTPUTS
                # System.Management.Automation.PSReference ([ref]).
                # Get-ReferenceToLastError returns a reference (memory pointer) to
                # the last error that occurred. It returns a reference to $null
                # ([ref]$null) if there are no errors on on the $error stack.
                #
                # .NOTES
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################

                if ($Error.Count -gt 0) {
                    return ([ref]($Error[0]))
                } else {
                    return ([ref]$null)
                }
            }

            function Test-ErrorOccurred {
                # .SYNOPSIS
                # Checks to see if an error occurred during a time period, i.e.,
                # during the execution of a command.
                #
                # .DESCRIPTION
                # Using two references (memory pointers) to errors, this function
                # checks to see if an error occurred based on differences between
                # the two errors.
                #
                # To use this function, you must first retrieve a reference to the
                # last error that occurred prior to the command you are about to
                # run. Then, run the command. After the command completes, retrieve
                # a reference to the last error that occurred. Pass these two
                # references to this function to determine if an error occurred.
                #
                # .PARAMETER ReferenceToEarlierError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time,
                # ReferenceToEarlierError must be a reference to $null
                # ([ref]$null).
                #
                # .PARAMETER ReferenceToLaterError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time, ReferenceToLaterError
                # must be a reference to $null ([ref]$null).
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work
                # if ($Error.Count -gt 0) {
                #     $refLastKnownError = ([ref]($Error[0]))
                # } else {
                #     $refLastKnownError = ([ref]$null)
                # }
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will
                # # suppress error output. Terminating errors will not output
                # # anything, kick to the empty trap statement and then continue
                # # on. Likewise, non- terminating errors will also not output
                # # anything, but they do not kick to the trap statement; they
                # # simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # if ($Error.Count -gt 0) {
                #     $refNewestCurrentError = ([ref]($Error[0]))
                # } else {
                #     $refNewestCurrentError = ([ref]$null)
                # }
                #
                # if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                #     # Error occurred
                # } else {
                #     # No error occurred
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Test-ErrorOccurred.
                #
                # .OUTPUTS
                # System.Boolean. Test-ErrorOccurred returns a boolean value
                # indicating whether an error occurred during the time period in
                # question. $true indicates an error occurred; $false indicates no
                # error occurred.
                #
                # .NOTES
                # This function also supports the use of positional parameters
                # instead of named parameters. If positional parameters are used
                # instead of named parameters, then two positional parameters are
                # required:
                #
                # The first positional parameter is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time, the first
                # positional parameter must be a reference to $null ([ref]$null).
                #
                # The second positional parameter is a reference (memory pointer)
                # to a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time,
                # ReferenceToLaterError must be a reference to $null ([ref]$null).
                #
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################
                param (
                    [ref]$ReferenceToEarlierError = ([ref]$null),
                    [ref]$ReferenceToLaterError = ([ref]$null)
                )

                # TODO: Validate input

                $boolErrorOccurred = $false
                if (($null -ne $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                    # Both not $null
                    if (($ReferenceToEarlierError.Value) -ne ($ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                } else {
                    # One is $null, or both are $null
                    # NOTE: $ReferenceToEarlierError could be non-null, while
                    # $ReferenceToLaterError could be null if $error was cleared;
                    # this does not indicate an error.
                    # So:
                    # - If both are null, no error.
                    # - If $ReferenceToEarlierError is null and
                    #   $ReferenceToLaterError is non-null, error.
                    # - If $ReferenceToEarlierError is non-null and
                    #   $ReferenceToLaterError is null, no error.
                    if (($null -eq $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                }

                return $boolErrorOccurred
            }
            #endregion FunctionsToSupportErrorHandling ############################

            trap {
                # Intentionally left empty to prevent terminating errors from
                # halting processing
            }

            # Retrieve the newest error on the stack prior to doing work
            $refLastKnownError = Get-ReferenceToLastError

            # Store current error preference; we will restore it after we do the
            # work of this function
            $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

            # Set ErrorActionPreference to SilentlyContinue; this will suppress
            # error output. Terminating errors will not output anything, kick to
            # the empty trap statement and then continue on. Likewise, non-
            # terminating errors will also not output anything, but they do not
            # kick to the trap statement; they simply continue on.
            $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

            $ReferenceToVersionObject.Value = [version]$StringToConvert

            # Restore the former error preference
            $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

            # Retrieve the newest error on the error stack
            $refNewestCurrentError = Get-ReferenceToLastError

            if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                # Error occurred; return failure indicator:
                return $false
            } else {
                # No error occurred; return success indicator:
                return $true
            }
        }

        function Split-StringOnLiteralString {
            # .SYNOPSIS
            # Splits a string into an array using a literal string as the splitter.
            #
            # .DESCRIPTION
            # Splits a string using a literal string (as opposed to regex). The
            # function is designed to be backward-compatible with all versions of
            # PowerShell and has been tested successfully on PowerShell v1. This
            # function behaves more like VBScript's Split() function than other
            # string splitting-approaches in PowerShell while avoiding the use of
            # RegEx.
            #
            # .PARAMETER StringToSplit
            # This parameter is required; it is the string to be split into an
            # array.
            #
            # .PARAMETER Splitter
            # This parameter is required; it is the string that will be used to
            # split the string specified in the StringToSplit parameter.
            #
            # .EXAMPLE
            # $result = Split-StringOnLiteralString -StringToSplit 'What do you think of this function?' -Splitter ' '
            # # $result.Count is 7
            # # $result[2] is 'you'
            #
            # .EXAMPLE
            # $result = Split-StringOnLiteralString 'What do you think of this function?' ' '
            # # $result.Count is 7
            #
            # .EXAMPLE
            # $result = Split-StringOnLiteralString -StringToSplit 'foo' -Splitter ' '
            # # $result.GetType().FullName is System.Object[]
            # # $result.Count is 1
            #
            # .EXAMPLE
            # $result = Split-StringOnLiteralString -StringToSplit 'foo' -Splitter ''
            # # $result.GetType().FullName is System.Object[]
            # # $result.Count is 5 because of how .NET handles a split using an
            # # empty string:
            # # $result[0] is ''
            # # $result[1] is 'f'
            # # $result[2] is 'o'
            # # $result[3] is 'o'
            # # $result[4] is ''
            #
            # .INPUTS
            # None. You can't pipe objects to Split-StringOnLiteralString.
            #
            # .OUTPUTS
            # System.String[]. Split-StringOnLiteralString returns an array of
            # strings, with each string being an element of the resulting array
            # from the split operation. This function always returns an array, even
            # when there is zero elements or one element in it.
            #
            # .NOTES
            # This function also supports the use of positional parameters instead
            # of named parameters. If positional parameters are used instead of
            # named parameters, then two positional parameters are required:
            #
            # The first positional parameter is the string to be split into an
            # array.
            #
            # The second positional parameter is the string that will be used to
            # split the string specified in the first positional parameter.
            #
            # Also, please note that if -StringToSplit (or the first positional
            # parameter) is $null, then the function will return an array with one
            # element, which is an empty string. This is because the function
            # converts $null to an empty string before splitting the string.
            #
            # Version: 3.0.20250211.1

            #region License ####################################################
            # Copyright (c) 2025 Frank Lesniak
            #
            # Permission is hereby granted, free of charge, to any person obtaining
            # a copy of this software and associated documentation files (the
            # "Software"), to deal in the Software without restriction, including
            # without limitation the rights to use, copy, modify, merge, publish,
            # distribute, sublicense, and/or sell copies of the Software, and to
            # permit persons to whom the Software is furnished to do so, subject to
            # the following conditions:
            #
            # The above copyright notice and this permission notice shall be
            # included in all copies or substantial portions of the Software.
            #
            # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
            # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
            # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
            # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
            # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
            # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
            # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
            # SOFTWARE.
            #endregion License ####################################################

            param (
                [string]$StringToSplit = '',
                [string]$Splitter = ''
            )

            $strSplitterInRegEx = [regex]::Escape($Splitter)
            $result = @([regex]::Split($StringToSplit, $strSplitterInRegEx))

            # The following code forces the function to return an array, always,
            # even when there are zero or one elements in the array
            $intElementCount = 1
            if ($null -ne $result) {
                if ($result.GetType().FullName.Contains('[]')) {
                    if (($result.Count -ge 2) -or ($result.Count -eq 0)) {
                        $intElementCount = $result.Count
                    }
                }
            }
            $strLowercaseFunctionName = $MyInvocation.InvocationName.ToLower()
            $boolArrayEncapsulation = $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ')') -or $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ' ')
            if ($boolArrayEncapsulation) {
                return ($result)
            } elseif ($intElementCount -eq 0) {
                return (, @())
            } elseif ($intElementCount -eq 1) {
                return (, (, $StringToSplit))
            } else {
                return ($result)
            }
        }

        function Convert-StringToInt32Safely {
            # .SYNOPSIS
            # Attempts to convert a string to a System.Int32.
            #
            # .DESCRIPTION
            # Attempts to convert a string to a System.Int32. If the string
            # cannot be converted to a System.Int32, the function suppresses the
            # error and returns $false. If the string can be converted to an
            # int32, the function returns $true and passes the int32 by
            # reference to the caller.
            #
            # .PARAMETER ReferenceToInt32
            # This parameter is required; it is a reference to a System.Int32
            # object that will be used to store the converted int32 object if the
            # conversion is successful.
            #
            # .PARAMETER StringToConvert
            # This parameter is required; it is a string that is to be converted to
            # a System.Int32 object.
            #
            # .EXAMPLE
            # $int = $null
            # $strInt = '1234'
            # $boolSuccess = Convert-StringToInt32Safely -ReferenceToInt32 ([ref]$int) -StringToConvert $strInt
            # # $boolSuccess will be $true, indicating that the conversion was
            # # successful.
            # # $int will contain a System.Int32 object equal to 1234.
            #
            # .EXAMPLE
            # $int = $null
            # $strInt = 'abc'
            # $boolSuccess = Convert-StringToInt32Safely -ReferenceToInt32 ([ref]$int) -StringToConvert $strInt
            # # $boolSuccess will be $false, indicating that the conversion was
            # # unsuccessful.
            # # $int will be undefined in this case.
            #
            # .INPUTS
            # None. You can't pipe objects to Convert-StringToInt32Safely.
            #
            # .OUTPUTS
            # System.Boolean. Convert-StringToInt32Safely returns a boolean value
            # indiciating whether the process completed successfully. $true means
            # the conversion completed successfully; $false means there was an
            # error.
            #
            # .NOTES
            # This function also supports the use of positional parameters instead
            # of named parameters. If positional parameters are used instead of
            # named parameters, then two positional parameters are required:
            #
            # The first positional parameter is a reference to a System.Int32
            # object that will be used to store the converted int32 object if the
            # conversion is successful.
            #
            # The second positional parameter is a string that is to be converted
            # to a System.Int32 object.
            #
            # Version: 1.0.20250215.0

            #region License ####################################################
            # Copyright (c) 2025 Frank Lesniak
            #
            # Permission is hereby granted, free of charge, to any person obtaining
            # a copy of this software and associated documentation files (the
            # "Software"), to deal in the Software without restriction, including
            # without limitation the rights to use, copy, modify, merge, publish,
            # distribute, sublicense, and/or sell copies of the Software, and to
            # permit persons to whom the Software is furnished to do so, subject to
            # the following conditions:
            #
            # The above copyright notice and this permission notice shall be
            # included in all copies or substantial portions of the Software.
            #
            # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
            # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
            # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
            # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
            # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
            # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
            # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
            # SOFTWARE.
            #endregion License ####################################################

            param (
                [ref]$ReferenceToInt32 = ([ref]$null),
                [string]$StringToConvert = ''
            )

            #region FunctionsToSupportErrorHandling ############################
            function Get-ReferenceToLastError {
                # .SYNOPSIS
                # Gets a reference (memory pointer) to the last error that
                # occurred.
                #
                # .DESCRIPTION
                # Returns a reference (memory pointer) to $null ([ref]$null) if no
                # errors on on the $error stack; otherwise, returns a reference to
                # the last error that occurred.
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work:
                # $refLastKnownError = Get-ReferenceToLastError
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will suppress
                # # error output. Terminating errors will not output anything, kick
                # # to the empty trap statement and then continue on. Likewise, non-
                # # terminating errors will also not output anything, but they do not
                # # kick to the trap statement; they simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # $refNewestCurrentError = Get-ReferenceToLastError
                #
                # $boolErrorOccurred = $false
                # if (($null -ne $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #     # Both not $null
                #     if (($refLastKnownError.Value) -ne ($refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # } else {
                #     # One is $null, or both are $null
                #     # NOTE: $refLastKnownError could be non-null, while
                #     # $refNewestCurrentError could be null if $error was cleared;
                #     # this does not indicate an error.
                #     #
                #     # So:
                #     # If both are null, no error.
                #     # If $refLastKnownError is null and $refNewestCurrentError is
                #     # non-null, error.
                #     # If $refLastKnownError is non-null and $refNewestCurrentError
                #     # is null, no error.
                #     #
                #     if (($null -eq $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Get-ReferenceToLastError.
                #
                # .OUTPUTS
                # System.Management.Automation.PSReference ([ref]).
                # Get-ReferenceToLastError returns a reference (memory pointer) to
                # the last error that occurred. It returns a reference to $null
                # ([ref]$null) if there are no errors on on the $error stack.
                #
                # .NOTES
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################

                if ($Error.Count -gt 0) {
                    return ([ref]($Error[0]))
                } else {
                    return ([ref]$null)
                }
            }

            function Test-ErrorOccurred {
                # .SYNOPSIS
                # Checks to see if an error occurred during a time period, i.e.,
                # during the execution of a command.
                #
                # .DESCRIPTION
                # Using two references (memory pointers) to errors, this function
                # checks to see if an error occurred based on differences between
                # the two errors.
                #
                # To use this function, you must first retrieve a reference to the
                # last error that occurred prior to the command you are about to
                # run. Then, run the command. After the command completes, retrieve
                # a reference to the last error that occurred. Pass these two
                # references to this function to determine if an error occurred.
                #
                # .PARAMETER ReferenceToEarlierError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time,
                # ReferenceToEarlierError must be a reference to $null
                # ([ref]$null).
                #
                # .PARAMETER ReferenceToLaterError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time, ReferenceToLaterError
                # must be a reference to $null ([ref]$null).
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work
                # if ($Error.Count -gt 0) {
                #     $refLastKnownError = ([ref]($Error[0]))
                # } else {
                #     $refLastKnownError = ([ref]$null)
                # }
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will
                # # suppress error output. Terminating errors will not output
                # # anything, kick to the empty trap statement and then continue
                # # on. Likewise, non- terminating errors will also not output
                # # anything, but they do not kick to the trap statement; they
                # # simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # if ($Error.Count -gt 0) {
                #     $refNewestCurrentError = ([ref]($Error[0]))
                # } else {
                #     $refNewestCurrentError = ([ref]$null)
                # }
                #
                # if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                #     # Error occurred
                # } else {
                #     # No error occurred
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Test-ErrorOccurred.
                #
                # .OUTPUTS
                # System.Boolean. Test-ErrorOccurred returns a boolean value
                # indicating whether an error occurred during the time period in
                # question. $true indicates an error occurred; $false indicates no
                # error occurred.
                #
                # .NOTES
                # This function also supports the use of positional parameters
                # instead of named parameters. If positional parameters are used
                # instead of named parameters, then two positional parameters are
                # required:
                #
                # The first positional parameter is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time, the first
                # positional parameter must be a reference to $null ([ref]$null).
                #
                # The second positional parameter is a reference (memory pointer)
                # to a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time,
                # ReferenceToLaterError must be a reference to $null ([ref]$null).
                #
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################
                param (
                    [ref]$ReferenceToEarlierError = ([ref]$null),
                    [ref]$ReferenceToLaterError = ([ref]$null)
                )

                # TODO: Validate input

                $boolErrorOccurred = $false
                if (($null -ne $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                    # Both not $null
                    if (($ReferenceToEarlierError.Value) -ne ($ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                } else {
                    # One is $null, or both are $null
                    # NOTE: $ReferenceToEarlierError could be non-null, while
                    # $ReferenceToLaterError could be null if $error was cleared;
                    # this does not indicate an error.
                    # So:
                    # - If both are null, no error.
                    # - If $ReferenceToEarlierError is null and
                    #   $ReferenceToLaterError is non-null, error.
                    # - If $ReferenceToEarlierError is non-null and
                    #   $ReferenceToLaterError is null, no error.
                    if (($null -eq $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                }

                return $boolErrorOccurred
            }
            #endregion FunctionsToSupportErrorHandling ############################

            trap {
                # Intentionally left empty to prevent terminating errors from
                # halting processing
            }

            # Retrieve the newest error on the stack prior to doing work
            $refLastKnownError = Get-ReferenceToLastError

            # Store current error preference; we will restore it after we do the
            # work of this function
            $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

            # Set ErrorActionPreference to SilentlyContinue; this will suppress
            # error output. Terminating errors will not output anything, kick to
            # the empty trap statement and then continue on. Likewise, non-
            # terminating errors will also not output anything, but they do not
            # kick to the trap statement; they simply continue on.
            $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

            $ReferenceToInt32.Value = [int32]$StringToConvert

            # Restore the former error preference
            $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

            # Retrieve the newest error on the error stack
            $refNewestCurrentError = Get-ReferenceToLastError

            if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                # Error occurred; return failure indicator:
                return $false
            } else {
                # No error occurred; return success indicator:
                return $true
            }
        }

        function Convert-StringToInt64Safely {
            # .SYNOPSIS
            # Attempts to convert a string to a System.Int64.
            #
            # .DESCRIPTION
            # Attempts to convert a string to a System.Int64. If the string
            # cannot be converted to a System.Int64, the function suppresses the
            # error and returns $false. If the string can be converted to an
            # int64, the function returns $true and passes the int64 by
            # reference to the caller.
            #
            # .PARAMETER ReferenceToInt64
            # This parameter is required; it is a reference to a System.Int64
            # object that will be used to store the converted int64 object if the
            # conversion is successful.
            #
            # .PARAMETER StringToConvert
            # This parameter is required; it is a string that is to be converted to
            # a System.Int64 object.
            #
            # .EXAMPLE
            # $int = $null
            # $strInt = '1234'
            # $boolSuccess = Convert-StringToInt64Safely -ReferenceToInt64 ([ref]$int) -StringToConvert $strInt
            # # $boolSuccess will be $true, indicating that the conversion was
            # # successful.
            # # $int will contain a System.Int64 object equal to 1234.
            #
            # .EXAMPLE
            # $int = $null
            # $strInt = 'abc'
            # $boolSuccess = Convert-StringToInt64Safely -ReferenceToInt64 ([ref]$int) -StringToConvert $strInt
            # # $boolSuccess will be $false, indicating that the conversion was
            # # unsuccessful.
            # # $int will be undefined in this case.
            #
            # .INPUTS
            # None. You can't pipe objects to Convert-StringToInt64Safely.
            #
            # .OUTPUTS
            # System.Boolean. Convert-StringToInt64Safely returns a boolean value
            # indiciating whether the process completed successfully. $true means
            # the conversion completed successfully; $false means there was an
            # error.
            #
            # .NOTES
            # This function also supports the use of positional parameters instead
            # of named parameters. If positional parameters are used instead of
            # named parameters, then two positional parameters are required:
            #
            # The first positional parameter is a reference to a System.Int64
            # object that will be used to store the converted int64 object if the
            # conversion is successful.
            #
            # The second positional parameter is a string that is to be converted
            # to a System.Int64 object.
            #
            # Version: 1.0.20250215.0

            #region License ####################################################
            # Copyright (c) 2025 Frank Lesniak
            #
            # Permission is hereby granted, free of charge, to any person obtaining
            # a copy of this software and associated documentation files (the
            # "Software"), to deal in the Software without restriction, including
            # without limitation the rights to use, copy, modify, merge, publish,
            # distribute, sublicense, and/or sell copies of the Software, and to
            # permit persons to whom the Software is furnished to do so, subject to
            # the following conditions:
            #
            # The above copyright notice and this permission notice shall be
            # included in all copies or substantial portions of the Software.
            #
            # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
            # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
            # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
            # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
            # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
            # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
            # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
            # SOFTWARE.
            #endregion License ####################################################

            param (
                [ref]$ReferenceToInt64 = ([ref]$null),
                [string]$StringToConvert = ''
            )

            #region FunctionsToSupportErrorHandling ############################
            function Get-ReferenceToLastError {
                # .SYNOPSIS
                # Gets a reference (memory pointer) to the last error that
                # occurred.
                #
                # .DESCRIPTION
                # Returns a reference (memory pointer) to $null ([ref]$null) if no
                # errors on on the $error stack; otherwise, returns a reference to
                # the last error that occurred.
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work:
                # $refLastKnownError = Get-ReferenceToLastError
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will suppress
                # # error output. Terminating errors will not output anything, kick
                # # to the empty trap statement and then continue on. Likewise, non-
                # # terminating errors will also not output anything, but they do not
                # # kick to the trap statement; they simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # $refNewestCurrentError = Get-ReferenceToLastError
                #
                # $boolErrorOccurred = $false
                # if (($null -ne $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #     # Both not $null
                #     if (($refLastKnownError.Value) -ne ($refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # } else {
                #     # One is $null, or both are $null
                #     # NOTE: $refLastKnownError could be non-null, while
                #     # $refNewestCurrentError could be null if $error was cleared;
                #     # this does not indicate an error.
                #     #
                #     # So:
                #     # If both are null, no error.
                #     # If $refLastKnownError is null and $refNewestCurrentError is
                #     # non-null, error.
                #     # If $refLastKnownError is non-null and $refNewestCurrentError
                #     # is null, no error.
                #     #
                #     if (($null -eq $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Get-ReferenceToLastError.
                #
                # .OUTPUTS
                # System.Management.Automation.PSReference ([ref]).
                # Get-ReferenceToLastError returns a reference (memory pointer) to
                # the last error that occurred. It returns a reference to $null
                # ([ref]$null) if there are no errors on on the $error stack.
                #
                # .NOTES
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################

                if ($Error.Count -gt 0) {
                    return ([ref]($Error[0]))
                } else {
                    return ([ref]$null)
                }
            }

            function Test-ErrorOccurred {
                # .SYNOPSIS
                # Checks to see if an error occurred during a time period, i.e.,
                # during the execution of a command.
                #
                # .DESCRIPTION
                # Using two references (memory pointers) to errors, this function
                # checks to see if an error occurred based on differences between
                # the two errors.
                #
                # To use this function, you must first retrieve a reference to the
                # last error that occurred prior to the command you are about to
                # run. Then, run the command. After the command completes, retrieve
                # a reference to the last error that occurred. Pass these two
                # references to this function to determine if an error occurred.
                #
                # .PARAMETER ReferenceToEarlierError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time,
                # ReferenceToEarlierError must be a reference to $null
                # ([ref]$null).
                #
                # .PARAMETER ReferenceToLaterError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time, ReferenceToLaterError
                # must be a reference to $null ([ref]$null).
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work
                # if ($Error.Count -gt 0) {
                #     $refLastKnownError = ([ref]($Error[0]))
                # } else {
                #     $refLastKnownError = ([ref]$null)
                # }
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will
                # # suppress error output. Terminating errors will not output
                # # anything, kick to the empty trap statement and then continue
                # # on. Likewise, non- terminating errors will also not output
                # # anything, but they do not kick to the trap statement; they
                # # simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # if ($Error.Count -gt 0) {
                #     $refNewestCurrentError = ([ref]($Error[0]))
                # } else {
                #     $refNewestCurrentError = ([ref]$null)
                # }
                #
                # if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                #     # Error occurred
                # } else {
                #     # No error occurred
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Test-ErrorOccurred.
                #
                # .OUTPUTS
                # System.Boolean. Test-ErrorOccurred returns a boolean value
                # indicating whether an error occurred during the time period in
                # question. $true indicates an error occurred; $false indicates no
                # error occurred.
                #
                # .NOTES
                # This function also supports the use of positional parameters
                # instead of named parameters. If positional parameters are used
                # instead of named parameters, then two positional parameters are
                # required:
                #
                # The first positional parameter is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time, the first
                # positional parameter must be a reference to $null ([ref]$null).
                #
                # The second positional parameter is a reference (memory pointer)
                # to a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time,
                # ReferenceToLaterError must be a reference to $null ([ref]$null).
                #
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################
                param (
                    [ref]$ReferenceToEarlierError = ([ref]$null),
                    [ref]$ReferenceToLaterError = ([ref]$null)
                )

                # TODO: Validate input

                $boolErrorOccurred = $false
                if (($null -ne $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                    # Both not $null
                    if (($ReferenceToEarlierError.Value) -ne ($ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                } else {
                    # One is $null, or both are $null
                    # NOTE: $ReferenceToEarlierError could be non-null, while
                    # $ReferenceToLaterError could be null if $error was cleared;
                    # this does not indicate an error.
                    # So:
                    # - If both are null, no error.
                    # - If $ReferenceToEarlierError is null and
                    #   $ReferenceToLaterError is non-null, error.
                    # - If $ReferenceToEarlierError is non-null and
                    #   $ReferenceToLaterError is null, no error.
                    if (($null -eq $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                }

                return $boolErrorOccurred
            }
            #endregion FunctionsToSupportErrorHandling ############################

            trap {
                # Intentionally left empty to prevent terminating errors from
                # halting processing
            }

            # Retrieve the newest error on the stack prior to doing work
            $refLastKnownError = Get-ReferenceToLastError

            # Store current error preference; we will restore it after we do the
            # work of this function
            $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

            # Set ErrorActionPreference to SilentlyContinue; this will suppress
            # error output. Terminating errors will not output anything, kick to
            # the empty trap statement and then continue on. Likewise, non-
            # terminating errors will also not output anything, but they do not
            # kick to the trap statement; they simply continue on.
            $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

            $ReferenceToInt64.Value = [int64]$StringToConvert

            # Restore the former error preference
            $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

            # Retrieve the newest error on the error stack
            $refNewestCurrentError = Get-ReferenceToLastError

            if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                # Error occurred; return failure indicator:
                return $false
            } else {
                # No error occurred; return success indicator:
                return $true
            }
        }

        function Get-PSVersion {
            # .SYNOPSIS
            # Returns the version of PowerShell that is running.
            #
            # .DESCRIPTION
            # The function outputs a [version] object representing the version of
            # PowerShell that is running.
            #
            # On versions of PowerShell greater than or equal to version 2.0, this
            # function returns the equivalent of $PSVersionTable.PSVersion
            #
            # PowerShell 1.0 does not have a $PSVersionTable variable, so this
            # function returns [version]('1.0') on PowerShell 1.0.
            #
            # .EXAMPLE
            # $versionPS = Get-PSVersion
            # # $versionPS now contains the version of PowerShell that is running.
            # # On versions of PowerShell greater than or equal to version 2.0,
            # # this function returns the equivalent of $PSVersionTable.PSVersion.
            #
            # .INPUTS
            # None. You can't pipe objects to Get-PSVersion.
            #
            # .OUTPUTS
            # System.Version. Get-PSVersion returns a [version] value indiciating
            # the version of PowerShell that is running.
            #
            # .NOTES
            # Version: 1.0.20250106.0

            #region License ####################################################
            # Copyright (c) 2025 Frank Lesniak
            #
            # Permission is hereby granted, free of charge, to any person obtaining
            # a copy of this software and associated documentation files (the
            # "Software"), to deal in the Software without restriction, including
            # without limitation the rights to use, copy, modify, merge, publish,
            # distribute, sublicense, and/or sell copies of the Software, and to
            # permit persons to whom the Software is furnished to do so, subject to
            # the following conditions:
            #
            # The above copyright notice and this permission notice shall be
            # included in all copies or substantial portions of the Software.
            #
            # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
            # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
            # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
            # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
            # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
            # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
            # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
            # SOFTWARE.
            #endregion License ####################################################

            if (Test-Path variable:\PSVersionTable) {
                return ($PSVersionTable.PSVersion)
            } else {
                return ([version]('1.0'))
            }
        }

        function Convert-StringToBigIntegerSafely {
            # .SYNOPSIS
            # Attempts to convert a string to a System.Numerics.BigInteger object.
            #
            # .DESCRIPTION
            # Attempts to convert a string to a System.Numerics.BigInteger object.
            # If the string cannot be converted to a System.Numerics.BigInteger
            # object, the function suppresses the error and returns $false. If the
            # string can be converted to a bigint object, the function returns
            # $true and passes the bigint object by reference to the caller.
            #
            # .PARAMETER ReferenceToBigIntegerObject
            # This parameter is required; it is a reference to a
            # System.Numerics.BigInteger object that will be used to store the
            # converted bigint object if the conversion is successful.
            #
            # .PARAMETER StringToConvert
            # This parameter is required; it is a string that is to be converted to
            # a System.Numerics.BigInteger object.
            #
            # .EXAMPLE
            # $bigint = $null
            # $strBigInt = '100000000000000000000000000000'
            # $boolSuccess = Convert-StringToBigIntegerSafely -ReferenceToBigIntegerObject ([ref]$bigint) -StringToConvert $strBigInt
            # # $boolSuccess will be $true, indicating that the conversion was
            # # successful.
            # # $bigint will contain a System.Numerics.BigInteger object equal to
            # # 100000000000000000000000000000.
            #
            # .EXAMPLE
            # $bigint = $null
            # $strBigInt = 'abc'
            # $boolSuccess = Convert-StringToBigIntegerSafely -ReferenceToBigIntegerObject ([ref]$bigint) -StringToConvert $strBigInt
            # # $boolSuccess will be $false, indicating that the conversion was
            # # unsuccessful.
            # # $bigint will be undefined in this case.
            #
            # .INPUTS
            # None. You can't pipe objects to Convert-StringToBigIntegerSafely.
            #
            # .OUTPUTS
            # System.Boolean. Convert-StringToBigIntegerSafely returns a boolean
            # value indiciating whether the process completed successfully. $true
            # means the conversion completed successfully; $false means there was
            # an error.
            #
            # .NOTES
            # This function also supports the use of positional parameters instead
            # of named parameters. If positional parameters are used instead of
            # named parameters, then two positional parameters are required:
            #
            # The first positional parameter is a reference to a
            # System.Numerics.BigInteger object that will be used to store the
            # converted bigint object if the conversion is successful.
            #
            # The second positional parameter is a string that is to be converted
            # to a System.Numerics.BigInteger object.
            #
            # Version: 1.0.20250216.0

            #region License ####################################################
            # Copyright (c) 2025 Frank Lesniak
            #
            # Permission is hereby granted, free of charge, to any person obtaining
            # a copy of this software and associated documentation files (the
            # "Software"), to deal in the Software without restriction, including
            # without limitation the rights to use, copy, modify, merge, publish,
            # distribute, sublicense, and/or sell copies of the Software, and to
            # permit persons to whom the Software is furnished to do so, subject to
            # the following conditions:
            #
            # The above copyright notice and this permission notice shall be
            # included in all copies or substantial portions of the Software.
            #
            # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
            # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
            # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
            # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
            # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
            # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
            # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
            # SOFTWARE.
            #endregion License ####################################################

            param (
                [ref]$ReferenceToBigIntegerObject = ([ref]$null),
                [string]$StringToConvert = ''
            )

            #region FunctionsToSupportErrorHandling ############################
            function Get-ReferenceToLastError {
                # .SYNOPSIS
                # Gets a reference (memory pointer) to the last error that
                # occurred.
                #
                # .DESCRIPTION
                # Returns a reference (memory pointer) to $null ([ref]$null) if no
                # errors on on the $error stack; otherwise, returns a reference to
                # the last error that occurred.
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work:
                # $refLastKnownError = Get-ReferenceToLastError
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will suppress
                # # error output. Terminating errors will not output anything, kick
                # # to the empty trap statement and then continue on. Likewise, non-
                # # terminating errors will also not output anything, but they do not
                # # kick to the trap statement; they simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # $refNewestCurrentError = Get-ReferenceToLastError
                #
                # $boolErrorOccurred = $false
                # if (($null -ne $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #     # Both not $null
                #     if (($refLastKnownError.Value) -ne ($refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # } else {
                #     # One is $null, or both are $null
                #     # NOTE: $refLastKnownError could be non-null, while
                #     # $refNewestCurrentError could be null if $error was cleared;
                #     # this does not indicate an error.
                #     #
                #     # So:
                #     # If both are null, no error.
                #     # If $refLastKnownError is null and $refNewestCurrentError is
                #     # non-null, error.
                #     # If $refLastKnownError is non-null and $refNewestCurrentError
                #     # is null, no error.
                #     #
                #     if (($null -eq $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Get-ReferenceToLastError.
                #
                # .OUTPUTS
                # System.Management.Automation.PSReference ([ref]).
                # Get-ReferenceToLastError returns a reference (memory pointer) to
                # the last error that occurred. It returns a reference to $null
                # ([ref]$null) if there are no errors on on the $error stack.
                #
                # .NOTES
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################

                if ($Error.Count -gt 0) {
                    return ([ref]($Error[0]))
                } else {
                    return ([ref]$null)
                }
            }

            function Test-ErrorOccurred {
                # .SYNOPSIS
                # Checks to see if an error occurred during a time period, i.e.,
                # during the execution of a command.
                #
                # .DESCRIPTION
                # Using two references (memory pointers) to errors, this function
                # checks to see if an error occurred based on differences between
                # the two errors.
                #
                # To use this function, you must first retrieve a reference to the
                # last error that occurred prior to the command you are about to
                # run. Then, run the command. After the command completes, retrieve
                # a reference to the last error that occurred. Pass these two
                # references to this function to determine if an error occurred.
                #
                # .PARAMETER ReferenceToEarlierError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time,
                # ReferenceToEarlierError must be a reference to $null
                # ([ref]$null).
                #
                # .PARAMETER ReferenceToLaterError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time, ReferenceToLaterError
                # must be a reference to $null ([ref]$null).
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work
                # if ($Error.Count -gt 0) {
                #     $refLastKnownError = ([ref]($Error[0]))
                # } else {
                #     $refLastKnownError = ([ref]$null)
                # }
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will
                # # suppress error output. Terminating errors will not output
                # # anything, kick to the empty trap statement and then continue
                # # on. Likewise, non- terminating errors will also not output
                # # anything, but they do not kick to the trap statement; they
                # # simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # if ($Error.Count -gt 0) {
                #     $refNewestCurrentError = ([ref]($Error[0]))
                # } else {
                #     $refNewestCurrentError = ([ref]$null)
                # }
                #
                # if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                #     # Error occurred
                # } else {
                #     # No error occurred
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Test-ErrorOccurred.
                #
                # .OUTPUTS
                # System.Boolean. Test-ErrorOccurred returns a boolean value
                # indicating whether an error occurred during the time period in
                # question. $true indicates an error occurred; $false indicates no
                # error occurred.
                #
                # .NOTES
                # This function also supports the use of positional parameters
                # instead of named parameters. If positional parameters are used
                # instead of named parameters, then two positional parameters are
                # required:
                #
                # The first positional parameter is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time, the first
                # positional parameter must be a reference to $null ([ref]$null).
                #
                # The second positional parameter is a reference (memory pointer)
                # to a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time,
                # ReferenceToLaterError must be a reference to $null ([ref]$null).
                #
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################
                param (
                    [ref]$ReferenceToEarlierError = ([ref]$null),
                    [ref]$ReferenceToLaterError = ([ref]$null)
                )

                # TODO: Validate input

                $boolErrorOccurred = $false
                if (($null -ne $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                    # Both not $null
                    if (($ReferenceToEarlierError.Value) -ne ($ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                } else {
                    # One is $null, or both are $null
                    # NOTE: $ReferenceToEarlierError could be non-null, while
                    # $ReferenceToLaterError could be null if $error was cleared;
                    # this does not indicate an error.
                    # So:
                    # - If both are null, no error.
                    # - If $ReferenceToEarlierError is null and
                    #   $ReferenceToLaterError is non-null, error.
                    # - If $ReferenceToEarlierError is non-null and
                    #   $ReferenceToLaterError is null, no error.
                    if (($null -eq $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                }

                return $boolErrorOccurred
            }
            #endregion FunctionsToSupportErrorHandling ############################

            trap {
                # Intentionally left empty to prevent terminating errors from
                # halting processing
            }

            # Retrieve the newest error on the stack prior to doing work
            $refLastKnownError = Get-ReferenceToLastError

            # Store current error preference; we will restore it after we do the
            # work of this function
            $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

            # Set ErrorActionPreference to SilentlyContinue; this will suppress
            # error output. Terminating errors will not output anything, kick to
            # the empty trap statement and then continue on. Likewise, non-
            # terminating errors will also not output anything, but they do not
            # kick to the trap statement; they simply continue on.
            $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

            $ReferenceToBigIntegerObject.Value = [System.Numerics.BigInteger]$StringToConvert

            # Restore the former error preference
            $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

            # Retrieve the newest error on the error stack
            $refNewestCurrentError = Get-ReferenceToLastError

            if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                # Error occurred; return failure indicator:
                return $false
            } else {
                # No error occurred; return success indicator:
                return $true
            }
        }

        function Convert-StringToDoubleSafely {
            # .SYNOPSIS
            # Attempts to convert a string to a System.Double.
            #
            # .DESCRIPTION
            # Attempts to convert a string to a System.Double. If the string
            # cannot be converted to a System.Double, the function suppresses the
            # error and returns $false. If the string can be converted to an
            # double, the function returns $true and passes the double by
            # reference to the caller.
            #
            # .PARAMETER ReferenceToDouble
            # This parameter is required; it is a reference to a System.Double
            # object that will be used to store the converted double object if the
            # conversion is successful.
            #
            # .PARAMETER StringToConvert
            # This parameter is required; it is a string that is to be converted to
            # a System.Double object.
            #
            # .EXAMPLE
            # $double = $null
            # $str = '100000000000000000000000'
            # $boolSuccess = Convert-StringToDoubleSafely -ReferenceToDouble ([ref]$double) -StringToConvert $str
            # # $boolSuccess will be $true, indicating that the conversion was
            # # successful.
            # # $double will contain a System.Double object equal to 1E+23
            #
            # .EXAMPLE
            # $double = $null
            # $str = 'abc'
            # $boolSuccess = Convert-StringToDoubleSafely -ReferenceToDouble ([ref]$double) -StringToConvert $str
            # # $boolSuccess will be $false, indicating that the conversion was
            # # unsuccessful.
            # # $double will undefined in this case.
            #
            # .INPUTS
            # None. You can't pipe objects to Convert-StringToDoubleSafely.
            #
            # .OUTPUTS
            # System.Boolean. Convert-StringToDoubleSafely returns a boolean value
            # indiciating whether the process completed successfully. $true means
            # the conversion completed successfully; $false means there was an
            # error.
            #
            # .NOTES
            # This function also supports the use of positional parameters instead
            # of named parameters. If positional parameters are used instead of
            # named parameters, then two positional parameters are required:
            #
            # The first positional parameter is a reference to a System.Double
            # object that will be used to store the converted double object if the
            # conversion is successful.
            #
            # The second positional parameter is a string that is to be converted
            # to a System.Double object.
            #
            # Version: 1.0.20250216.0

            #region License ####################################################
            # Copyright (c) 2025 Frank Lesniak
            #
            # Permission is hereby granted, free of charge, to any person obtaining
            # a copy of this software and associated documentation files (the
            # "Software"), to deal in the Software without restriction, including
            # without limitation the rights to use, copy, modify, merge, publish,
            # distribute, sublicense, and/or sell copies of the Software, and to
            # permit persons to whom the Software is furnished to do so, subject to
            # the following conditions:
            #
            # The above copyright notice and this permission notice shall be
            # included in all copies or substantial portions of the Software.
            #
            # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
            # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
            # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
            # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
            # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
            # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
            # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
            # SOFTWARE.
            #endregion License ####################################################

            param (
                [ref]$ReferenceToDouble = ([ref]$null),
                [string]$StringToConvert = ''
            )

            #region FunctionsToSupportErrorHandling ############################
            function Get-ReferenceToLastError {
                # .SYNOPSIS
                # Gets a reference (memory pointer) to the last error that
                # occurred.
                #
                # .DESCRIPTION
                # Returns a reference (memory pointer) to $null ([ref]$null) if no
                # errors on on the $error stack; otherwise, returns a reference to
                # the last error that occurred.
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work:
                # $refLastKnownError = Get-ReferenceToLastError
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will suppress
                # # error output. Terminating errors will not output anything, kick
                # # to the empty trap statement and then continue on. Likewise, non-
                # # terminating errors will also not output anything, but they do not
                # # kick to the trap statement; they simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # $refNewestCurrentError = Get-ReferenceToLastError
                #
                # $boolErrorOccurred = $false
                # if (($null -ne $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #     # Both not $null
                #     if (($refLastKnownError.Value) -ne ($refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # } else {
                #     # One is $null, or both are $null
                #     # NOTE: $refLastKnownError could be non-null, while
                #     # $refNewestCurrentError could be null if $error was cleared;
                #     # this does not indicate an error.
                #     #
                #     # So:
                #     # If both are null, no error.
                #     # If $refLastKnownError is null and $refNewestCurrentError is
                #     # non-null, error.
                #     # If $refLastKnownError is non-null and $refNewestCurrentError
                #     # is null, no error.
                #     #
                #     if (($null -eq $refLastKnownError.Value) -and ($null -ne $refNewestCurrentError.Value)) {
                #         $boolErrorOccurred = $true
                #     }
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Get-ReferenceToLastError.
                #
                # .OUTPUTS
                # System.Management.Automation.PSReference ([ref]).
                # Get-ReferenceToLastError returns a reference (memory pointer) to
                # the last error that occurred. It returns a reference to $null
                # ([ref]$null) if there are no errors on on the $error stack.
                #
                # .NOTES
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################

                if ($Error.Count -gt 0) {
                    return ([ref]($Error[0]))
                } else {
                    return ([ref]$null)
                }
            }

            function Test-ErrorOccurred {
                # .SYNOPSIS
                # Checks to see if an error occurred during a time period, i.e.,
                # during the execution of a command.
                #
                # .DESCRIPTION
                # Using two references (memory pointers) to errors, this function
                # checks to see if an error occurred based on differences between
                # the two errors.
                #
                # To use this function, you must first retrieve a reference to the
                # last error that occurred prior to the command you are about to
                # run. Then, run the command. After the command completes, retrieve
                # a reference to the last error that occurred. Pass these two
                # references to this function to determine if an error occurred.
                #
                # .PARAMETER ReferenceToEarlierError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time,
                # ReferenceToEarlierError must be a reference to $null
                # ([ref]$null).
                #
                # .PARAMETER ReferenceToLaterError
                # This parameter is required; it is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred.
                #
                # If no error was on the stack at this time, ReferenceToLaterError
                # must be a reference to $null ([ref]$null).
                #
                # .EXAMPLE
                # # Intentionally empty trap statement to prevent terminating
                # # errors from halting processing
                # trap { }
                #
                # # Retrieve the newest error on the stack prior to doing work
                # if ($Error.Count -gt 0) {
                #     $refLastKnownError = ([ref]($Error[0]))
                # } else {
                #     $refLastKnownError = ([ref]$null)
                # }
                #
                # # Store current error preference; we will restore it after we do
                # # some work:
                # $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference
                #
                # # Set ErrorActionPreference to SilentlyContinue; this will
                # # suppress error output. Terminating errors will not output
                # # anything, kick to the empty trap statement and then continue
                # # on. Likewise, non- terminating errors will also not output
                # # anything, but they do not kick to the trap statement; they
                # # simply continue on.
                # $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
                #
                # # Do something that might trigger an error
                # Get-Item -Path 'C:\MayNotExist.txt'
                #
                # # Restore the former error preference
                # $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference
                #
                # # Retrieve the newest error on the error stack
                # if ($Error.Count -gt 0) {
                #     $refNewestCurrentError = ([ref]($Error[0]))
                # } else {
                #     $refNewestCurrentError = ([ref]$null)
                # }
                #
                # if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                #     # Error occurred
                # } else {
                #     # No error occurred
                # }
                #
                # .INPUTS
                # None. You can't pipe objects to Test-ErrorOccurred.
                #
                # .OUTPUTS
                # System.Boolean. Test-ErrorOccurred returns a boolean value
                # indicating whether an error occurred during the time period in
                # question. $true indicates an error occurred; $false indicates no
                # error occurred.
                #
                # .NOTES
                # This function also supports the use of positional parameters
                # instead of named parameters. If positional parameters are used
                # instead of named parameters, then two positional parameters are
                # required:
                #
                # The first positional parameter is a reference (memory pointer) to
                # a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack earlier in time, i.e., prior to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time, the first
                # positional parameter must be a reference to $null ([ref]$null).
                #
                # The second positional parameter is a reference (memory pointer)
                # to a System.Management.Automation.ErrorRecord that represents the
                # newest error on the stack later in time, i.e., after to running
                # the command for which you wish to determine whether an error
                # occurred. If no error was on the stack at this time,
                # ReferenceToLaterError must be a reference to $null ([ref]$null).
                #
                # Version: 2.0.20250215.0

                #region License ################################################
                # Copyright (c) 2025 Frank Lesniak
                #
                # Permission is hereby granted, free of charge, to any person
                # obtaining a copy of this software and associated documentation
                # files (the "Software"), to deal in the Software without
                # restriction, including without limitation the rights to use,
                # copy, modify, merge, publish, distribute, sublicense, and/or sell
                # copies of the Software, and to permit persons to whom the
                # Software is furnished to do so, subject to the following
                # conditions:
                #
                # The above copyright notice and this permission notice shall be
                # included in all copies or substantial portions of the Software.
                #
                # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
                # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
                # OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
                # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
                # HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
                # WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
                # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
                # OTHER DEALINGS IN THE SOFTWARE.
                #endregion License ################################################
                param (
                    [ref]$ReferenceToEarlierError = ([ref]$null),
                    [ref]$ReferenceToLaterError = ([ref]$null)
                )

                # TODO: Validate input

                $boolErrorOccurred = $false
                if (($null -ne $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                    # Both not $null
                    if (($ReferenceToEarlierError.Value) -ne ($ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                } else {
                    # One is $null, or both are $null
                    # NOTE: $ReferenceToEarlierError could be non-null, while
                    # $ReferenceToLaterError could be null if $error was cleared;
                    # this does not indicate an error.
                    # So:
                    # - If both are null, no error.
                    # - If $ReferenceToEarlierError is null and
                    #   $ReferenceToLaterError is non-null, error.
                    # - If $ReferenceToEarlierError is non-null and
                    #   $ReferenceToLaterError is null, no error.
                    if (($null -eq $ReferenceToEarlierError.Value) -and ($null -ne $ReferenceToLaterError.Value)) {
                        $boolErrorOccurred = $true
                    }
                }

                return $boolErrorOccurred
            }
            #endregion FunctionsToSupportErrorHandling ############################

            trap {
                # Intentionally left empty to prevent terminating errors from
                # halting processing
            }

            # Retrieve the newest error on the stack prior to doing work
            $refLastKnownError = Get-ReferenceToLastError

            # Store current error preference; we will restore it after we do the
            # work of this function
            $actionPreferenceFormerErrorPreference = $global:ErrorActionPreference

            # Set ErrorActionPreference to SilentlyContinue; this will suppress
            # error output. Terminating errors will not output anything, kick to
            # the empty trap statement and then continue on. Likewise, non-
            # terminating errors will also not output anything, but they do not
            # kick to the trap statement; they simply continue on.
            $global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

            $ReferenceToDouble.Value = [double]$StringToConvert

            # Restore the former error preference
            $global:ErrorActionPreference = $actionPreferenceFormerErrorPreference

            # Retrieve the newest error on the error stack
            $refNewestCurrentError = Get-ReferenceToLastError

            if (Test-ErrorOccurred -ReferenceToEarlierError $refLastKnownError -ReferenceToLaterError $refNewestCurrentError) {
                # Error occurred; return failure indicator:
                return $false
            } else {
                # No error occurred; return success indicator:
                return $true
            }
        }

        $ReferenceArrayOfLeftoverStrings.Value = @('', '', '', '', '')

        $boolResult = Convert-StringToVersionSafely -ReferenceToVersionObject $ReferenceToVersionObject -StringToConvert $StringToConvert

        if ($boolResult) {
            return 0
        }

        # If we are still here, the conversion was not successful.

        $arrVersionElements = Split-StringOnLiteralString -StringToSplit $StringToConvert -Splitter '.'
        $intCountOfVersionElements = $arrVersionElements.Count

        if ($intCountOfVersionElements -lt 2) {
            # You can't have a version with less than two elements
            return -1
        }

        if ($intCountOfVersionElements -ge 5) {
            $strExcessVersionElements = [string]::join('.', $arrVersionElements[4..($intCountOfVersionElements - 1)])
        } else {
            $strExcessVersionElements = ''
        }

        if ($intCountOfVersionElements -ge 3) {
            $intElementInQuestion = 3
        } else {
            $intElementInQuestion = $intCountOfVersionElements
        }

        $boolConversionSuccessful = $false

        # See if excess elements are our only problem
        if (-not [string]::IsNullOrEmpty($strExcessVersionElements)) {
            $strAttemptedVersion = [string]::join('.', $arrVersionElements[0..$intElementInQuestion])
            $boolResult = Convert-StringToVersionSafely -ReferenceToVersionObject $ReferenceToVersionObject -StringToConvert $strAttemptedVersion
            if ($boolResult) {
                # Conversion successful; the only problem was the excess elements
                $boolConversionSuccessful = $true
                $intReturnValue = 5
                ($ReferenceArrayOfLeftoverStrings.Value)[4] = $strExcessVersionElements
            }
        }

        while ($intElementInQuestion -gt 0 -and -not $boolConversionSuccessful) {
            $strAttemptedVersion = [string]::join('.', $arrVersionElements[0..($intElementInQuestion - 1)])
            $boolResult = $false
            if ($intElementInQuestion -gt 1) {
                $boolResult = Convert-StringToVersionSafely -ReferenceToVersionObject $ReferenceToVersionObject -StringToConvert $strAttemptedVersion
            }
            if ($boolResult -or $intElementInQuestion -eq 1) {
                # Conversion successful or we're on the second element
                # See if we can trim out non-numerical characters
                $strRegexFirstNumericalCharacters = '^\d+'
                $strFirstNumericalCharacters = [regex]::Match($arrVersionElements[$intElementInQuestion], $strRegexFirstNumericalCharacters).Value
                if ([string]::IsNullOrEmpty($strFirstNumericalCharacters)) {
                    # No numerical characters found
                    ($ReferenceArrayOfLeftoverStrings.Value)[$intElementInQuestion] = $arrVersionElements[$intElementInQuestion]
                    for ($intCounterA = $intElementInQuestion + 1; $intCounterA -le 3; $intCounterA++) {
                        ($ReferenceArrayOfLeftoverStrings.Value)[$intCounterA] = $arrVersionElements[$intCounterA]
                    }
                    $boolConversionSuccessful = $true
                    $intReturnValue = $intElementInQuestion + 1
                    ($ReferenceArrayOfLeftoverStrings.Value)[4] = $strExcessVersionElements
                } else {
                    # Numerical characters found
                    $boolResult = Convert-StringToInt32Safely -ReferenceToInt32 ([ref]$null) -StringToConvert $strFirstNumericalCharacters
                    if ($boolResult) {
                        # Append the first numerical characters to the version
                        $strAttemptedVersionNew = $strAttemptedVersion + '.' + $strFirstNumericalCharacters
                        $boolResult = Convert-StringToVersionSafely -ReferenceToVersionObject $ReferenceToVersionObject -StringToConvert $strAttemptedVersionNew
                        if ($boolResult) {
                            # Conversion successful
                            $strExcessCharactersInThisElement = ($arrVersionElements[$intElementInQuestion]).Substring($strFirstNumericalCharacters.Length)
                            ($ReferenceArrayOfLeftoverStrings.Value)[$intElementInQuestion] = $strExcessCharactersInThisElement
                            for ($intCounterA = $intElementInQuestion + 1; $intCounterA -le 3; $intCounterA++) {
                                ($ReferenceArrayOfLeftoverStrings.Value)[$intCounterA] = $arrVersionElements[$intCounterA]
                            }
                            $boolConversionSuccessful = $true
                            $intReturnValue = $intElementInQuestion + 1
                            ($ReferenceArrayOfLeftoverStrings.Value)[4] = $strExcessVersionElements
                        } else {
                            # Conversion was not successful even though we just
                            # tried converting using numbers we know are
                            # convertable to an int32. This makes no sense.
                            # Throw warning:
                            $strMessage = 'Conversion of string "' + $strAttemptedVersionNew + '" to a version object failed even though "' + $strAttemptedVersion + '" converted to a version object just fine, and we proved that "' + $strFirstNumericalCharacters + '" was converted to an int32 object successfully. This should not be possible!'
                            Write-Warning -Message $strMessage
                        }
                    } else {
                        # The string of numbers could not be converted to an int32;
                        # this is probably because the represented number is too
                        # large.
                        # Try converting to int64:
                        $int64 = $null
                        $boolResult = Convert-StringToInt64Safely -ReferenceToInt64 ([ref]$int64) -StringToConvert $strFirstNumericalCharacters
                        if ($boolResult) {
                            # Converted to int64 but not int32
                            $intRemainder = $int64 - [int32]::MaxValue
                            $strAttemptedVersionNew = $strAttemptedVersion + '.' + [int32]::MaxValue
                            $boolResult = Convert-StringToVersionSafely -ReferenceToVersionObject $ReferenceToVersionObject -StringToConvert $strAttemptedVersionNew
                            if ($boolResult) {
                                # Conversion successful
                                $strExcessCharactersInThisElement = ($arrVersionElements[$intElementInQuestion]).Substring($strFirstNumericalCharacters.Length)
                                ($ReferenceArrayOfLeftoverStrings.Value)[$intElementInQuestion] = ([string]$intRemainder) + $strExcessCharactersInThisElement
                                for ($intCounterA = $intElementInQuestion + 1; $intCounterA -le 3; $intCounterA++) {
                                    ($ReferenceArrayOfLeftoverStrings.Value)[$intCounterA] = $arrVersionElements[$intCounterA]
                                }
                                $boolConversionSuccessful = $true
                                $intReturnValue = $intElementInQuestion + 1
                                ($ReferenceArrayOfLeftoverStrings.Value)[4] = $strExcessVersionElements
                            } else {
                                # Conversion was not successful even though we just
                                # tried converting using numbers we know are
                                # convertable to an int32. This makes no sense.
                                # Throw warning:
                                $strMessage = 'Conversion of string "' + $strAttemptedVersionNew + '" to a version object failed even though "' + $strAttemptedVersion + '" converted to a version object just fine, and we know that "' + ([string]([int32]::MaxValue)) + '" is a valid int32 number. This should not be possible!'
                                Write-Warning -Message $strMessage
                            }
                        } else {
                            # Conversion to int64 failed; this is probably because
                            # the represented number is too large.
                            if ($PSVersion -eq ([version]'0.0')) {
                                $versionPS = Get-PSVersion
                            } else {
                                $versionPS = $PSVersion
                            }

                            if ($versionPS.Major -ge 3) {
                                # Use bigint
                                $bigint = $null
                                $boolResult = Convert-StringToBigIntegerSafely -ReferenceToBigIntegerObject ([ref]$bigint) -StringToConvert $strFirstNumericalCharacters
                                if ($boolResult) {
                                    # Converted to bigint but not int32 or
                                    # int64
                                    $bigintRemainder = $bigint - [int32]::MaxValue
                                    $strAttemptedVersionNew = $strAttemptedVersion + '.' + [int32]::MaxValue
                                    $boolResult = Convert-StringToVersionSafely -ReferenceToVersionObject $ReferenceToVersionObject -StringToConvert $strAttemptedVersionNew
                                    if ($boolResult) {
                                        # Conversion successful
                                        $strExcessCharactersInThisElement = ($arrVersionElements[$intElementInQuestion]).Substring($strFirstNumericalCharacters.Length)
                                        ($ReferenceArrayOfLeftoverStrings.Value)[$intElementInQuestion] = ([string]$bigintRemainder) + $strExcessCharactersInThisElement
                                        for ($intCounterA = $intElementInQuestion + 1; $intCounterA -le 3; $intCounterA++) {
                                            ($ReferenceArrayOfLeftoverStrings.Value)[$intCounterA] = $arrVersionElements[$intCounterA]
                                        }
                                        $boolConversionSuccessful = $true
                                        $intReturnValue = $intElementInQuestion + 1
                                        ($ReferenceArrayOfLeftoverStrings.Value)[4] = $strExcessVersionElements
                                    } else {
                                        # Conversion was not successful even though
                                        # we just tried converting using numbers we
                                        # know are convertable to an int32. This
                                        # makes no sense. Throw warning:
                                        $strMessage = 'Conversion of string "' + $strAttemptedVersionNew + '" to a version object failed even though "' + $strAttemptedVersion + '" converted to a version object just fine, and we know that "' + ([string]([int32]::MaxValue)) + '" is a valid int32 number. This should not be possible!'
                                        Write-Warning -Message $strMessage
                                    }
                                } else {
                                    # Conversion to bigint failed; given that we
                                    # know that the string is all numbers, this
                                    # should not be possible. Throw warning
                                    $strMessage = 'The string "' + $strFirstNumericalCharacters + '" could not be converted to an int32, int64, or bigint number. This should not be possible!'
                                    Write-Warning -Message $strMessage
                                }
                            } else {
                                # Use double
                                $double = $null
                                $boolResult = Convert-StringToDoubleSafely -ReferenceToDouble ([ref]$double) -StringToConvert $strFirstNumericalCharacters
                                if ($boolResult) {
                                    # Converted to double but not int32 or
                                    # int64
                                    $doubleRemainder = $double - [int32]::MaxValue
                                    $strAttemptedVersionNew = $strAttemptedVersion + '.' + [int32]::MaxValue
                                    $boolResult = Convert-StringToVersionSafely -ReferenceToVersionObject $ReferenceToVersionObject -StringToConvert $strAttemptedVersionNew
                                    if ($boolResult) {
                                        # Conversion successful
                                        $strExcessCharactersInThisElement = ($arrVersionElements[$intElementInQuestion]).Substring($strFirstNumericalCharacters.Length)
                                        ($ReferenceArrayOfLeftoverStrings.Value)[$intElementInQuestion] = ([string]$doubleRemainder) + $strExcessCharactersInThisElement
                                        for ($intCounterA = $intElementInQuestion + 1; $intCounterA -le 3; $intCounterA++) {
                                            ($ReferenceArrayOfLeftoverStrings.Value)[$intCounterA] = $arrVersionElements[$intCounterA]
                                        }
                                        $boolConversionSuccessful = $true
                                        $intReturnValue = $intElementInQuestion + 1
                                        ($ReferenceArrayOfLeftoverStrings.Value)[4] = $strExcessVersionElements
                                    } else {
                                        # Conversion was not successful even though
                                        # we just tried converting using numbers we
                                        # know are convertable to an int32. This
                                        # makes no sense. Throw warning:
                                        $strMessage = 'Conversion of string "' + $strAttemptedVersionNew + '" to a version object failed even though "' + $strAttemptedVersion + '" converted to a version object just fine, and we know that "' + ([string]([int32]::MaxValue)) + '" is a valid int32 number. This should not be possible!'
                                        Write-Warning -Message $strMessage
                                    }
                                } else {
                                    # Conversion to double failed; given that we
                                    # know that the string is all numbers, this
                                    # should not be possible unless the string of
                                    # numbers exceeded the maximum size allowed
                                    # for a double. This is possible, so don't
                                    # throw a warning.
                                    # Treat like no numerical characters found
                                    ($ReferenceArrayOfLeftoverStrings.Value)[$intElementInQuestion] = $arrVersionElements[$intElementInQuestion]
                                    for ($intCounterA = $intElementInQuestion + 1; $intCounterA -le 3; $intCounterA++) {
                                        ($ReferenceArrayOfLeftoverStrings.Value)[$intCounterA] = $arrVersionElements[$intCounterA]
                                    }
                                    $boolConversionSuccessful = $true
                                    $intReturnValue = $intElementInQuestion + 1
                                    ($ReferenceArrayOfLeftoverStrings.Value)[4] = $strExcessVersionElements
                                }
                            }
                        }
                    }
                }
            }
            $intElementInQuestion--
        }

        if (-not $boolConversionSuccessful) {
            # Conversion was not successful
            return -1
        } else {
            return $intReturnValue
        }
    }

    #region Process input ######################################################
    # Validate that the required parameter was supplied:
    if ($null -eq $ReferenceToHashtableOfInstalledModules) {
        $strMessage = 'The parameter $ReferenceToHashtableOfInstalledModules must be a reference to a hashtable. The hashtable must have keys that are the names of PowerShell modules with each key''s value populated with arrays of ModuleInfoGrouping objects (the result of Get-Module).'
        Write-Error -Message $strMessage
        return $false
    }
    if ($null -eq $ReferenceToHashtableOfInstalledModules.Value) {
        $strMessage = 'The parameter $ReferenceToHashtableOfInstalledModules must be a reference to a hashtable. The hashtable must have keys that are the names of PowerShell modules with each key''s value populated with arrays of ModuleInfoGrouping objects (the result of Get-Module).'
        Write-Error -Message $strMessage
        return $false
    }
    if ($ReferenceToHashtableOfInstalledModules.Value.GetType().FullName -ne 'System.Collections.Hashtable') {
        $strMessage = 'The parameter $ReferenceToHashtableOfInstalledModules must be a reference to a hashtable. The hashtable must have keys that are the names of PowerShell modules with each key''s value populated with arrays of ModuleInfoGrouping objects (the result of Get-Module).'
        Write-Error -Message $strMessage
        return $false
    }

    $boolThrowErrorForMissingModule = $false
    if ($null -ne $ThrowErrorIfModuleNotInstalled) {
        if ($ThrowErrorIfModuleNotInstalled.IsPresent -eq $true) {
            $boolThrowErrorForMissingModule = $true
        }
    }

    $boolThrowWarningForMissingModule = $false
    if (-not $boolThrowErrorForMissingModule) {
        if ($null -ne $ThrowWarningIfModuleNotInstalled) {
            if ($ThrowWarningIfModuleNotInstalled.IsPresent -eq $true) {
                $boolThrowWarningForMissingModule = $true
            }
        }
    }

    $boolThrowErrorForOutdatedModule = $false
    if ($null -ne $ThrowErrorIfModuleNotUpToDate) {
        if ($ThrowErrorIfModuleNotUpToDate.IsPresent -eq $true) {
            $boolThrowErrorForOutdatedModule = $true
        }
    }

    $boolThrowWarningForOutdatedModule = $false
    if (-not $boolThrowErrorForOutdatedModule) {
        if ($null -ne $ThrowWarningIfModuleNotUpToDate) {
            if ($ThrowWarningIfModuleNotUpToDate.IsPresent -eq $true) {
                $boolThrowWarningForOutdatedModule = $true
            }
        }
    }

    $boolCheckPowerShellVersion = $true
    if ($null -ne $DoNotCheckPowerShellVersion) {
        if ($DoNotCheckPowerShellVersion.IsPresent -eq $true) {
            $boolCheckPowerShellVersion = $false
        }
    }
    #endregion Process input ######################################################

    #region Verify environment #################################################
    if ($boolCheckPowerShellVersion) {
        $versionPS = Get-PSVersion
        if ($versionPS.Major -lt 5) {
            $strMessage = 'Test-PowerShellModuleUpdatesAvailableUsingHashtable requires PowerShell version 5.0 or newer.'
            Write-Warning -Message $strMessage
            return $false
        }
    } else {
        $versionPS = [version]'5.0'
    }
    #endregion Verify environment #################################################

    $VerbosePreferenceAtStartOfFunction = $VerbosePreference

    $boolResult = $true

    $hashtableMessagesToThrowForMissingModule = @{}
    $hashtableModuleNameToCustomMessageToThrowForMissingModule = @{}
    if ($null -ne $ReferenceToHashtableOfCustomNotInstalledMessages) {
        if ($null -ne $ReferenceToHashtableOfCustomNotInstalledMessages.Value) {
            if ($ReferenceToHashtableOfCustomNotInstalledMessages.Value.GetType().FullName -eq 'System.Collections.Hashtable') {
                foreach ($strMessage in @(($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Keys)) {
                    $hashtableMessagesToThrowForMissingModule.Add($strMessage, $false)

                    ($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Item($strMessage) | ForEach-Object {
                        $hashtableModuleNameToCustomMessageToThrowForMissingModule.Add($_, $strMessage)
                    }
                }
            }
        }
    }

    $hashtableMessagesToThrowForOutdatedModule = @{}
    $hashtableModuleNameToCustomMessageToThrowForOutdatedModule = @{}
    if ($null -ne $ReferenceToHashtableOfCustomNotUpToDateMessages) {
        if ($null -ne $ReferenceToHashtableOfCustomNotUpToDateMessages.Value) {
            if ($ReferenceToHashtableOfCustomNotUpToDateMessages.Value.GetType().FullName -eq 'System.Collections.Hashtable') {
                foreach ($strMessage in @(($ReferenceToHashtableOfCustomNotUpToDateMessages.Value).Keys)) {
                    $hashtableMessagesToThrowForOutdatedModule.Add($strMessage, $false)

                    ($ReferenceToHashtableOfCustomNotUpToDateMessages.Value).Item($strMessage) | ForEach-Object {
                        $hashtableModuleNameToCustomMessageToThrowForOutdatedModule.Add($_, $strMessage)
                    }
                }
            }
        }
    }

    foreach ($strModuleName in @(($ReferenceToHashtableOfInstalledModules.Value).Keys)) {
        if (@(($ReferenceToHashtableOfInstalledModules.Value).Item($strModuleName)).Count -eq 0) {
            # Module is not installed
            $boolResult = $false

            if ($hashtableModuleNameToCustomMessageToThrowForMissingModule.ContainsKey($strModuleName) -eq $true) {
                $strMessage = $hashtableModuleNameToCustomMessageToThrowForMissingModule.Item($strModuleName)
                $hashtableMessagesToThrowForMissingModule.Item($strMessage) = $true
            } else {
                $strMessage = $strModuleName + ' module not found. Please install it and then try again.' + [System.Environment]::NewLine + 'You can install the ' + $strModuleName + ' PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module ' + $strModuleName + ';' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                $hashtableMessagesToThrowForMissingModule.Add($strMessage, $true)
            }

            if ($null -ne $ReferenceToArrayOfMissingModules) {
                if ($null -ne $ReferenceToArrayOfMissingModules.Value) {
                    ($ReferenceToArrayOfMissingModules.Value) += $strModuleName
                }
            }
        } else {
            # Module is installed
            $versionNewestInstalledModule = (@(($ReferenceToHashtableOfInstalledModules.Value).Item($strModuleName)) | ForEach-Object { [version]($_.Version) } | Sort-Object)[-1]

            $arrModuleNewestInstalledModule = @(@(($ReferenceToHashtableOfInstalledModules.Value).Item($strModuleName)) | Where-Object { ([version]($_.Version)) -eq $versionNewestInstalledModule })

            # In the event there are multiple installations of the same version, reduce to a
            # single instance of the module
            if ($arrModuleNewestInstalledModule.Count -gt 1) {
                $moduleNewestInstalled = @($arrModuleNewestInstalledModule | Select-Object -Unique)[0]
            } else {
                $moduleNewestInstalled = $arrModuleNewestInstalledModule[0]
            }

            $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
            $moduleNewestAvailable = Find-Module -Name $strModuleName -ErrorAction SilentlyContinue
            $VerbosePreference = $VerbosePreferenceAtStartOfFunction

            if ($null -ne $moduleNewestAvailable) {
                $versionNewestModuleInPSGallery = $null
                $arrLeftoverStrings = @('', '', '', '', '')
                $intReturnCode = Convert-StringToFlexibleVersion -ReferenceToVersionObject ([ref]$versionNewestModuleInPSGallery) -ReferenceArrayOfLeftoverStrings ([ref]$arrLeftoverStrings) -StringToConvert $moduleNewestAvailable.Version -PSVersion $versionPS
                if ($intReturnCode -ge 0) {
                    # Conversion of the string version object from Find-Module was
                    # successful
                    if ($versionNewestModuleInPSGallery -gt $moduleNewestInstalled.Version) {
                        # A newer version is available
                        $boolResult = $false

                        if ($hashtableModuleNameToCustomMessageToThrowForOutdatedModule.ContainsKey($strModuleName) -eq $true) {
                            $strMessage = $hashtableModuleNameToCustomMessageToThrowForOutdatedModule.Item($strModuleName)
                            $hashtableMessagesToThrowForOutdatedModule.Item($strMessage) = $true
                        } else {
                            $strMessage = 'A newer version of the ' + $strModuleName + ' PowerShell module is available. Please consider updating it by running the following command:' + [System.Environment]::NewLine + 'Install-Module ' + $strModuleName + ' -Force;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                            $hashtableMessagesToThrowForOutdatedModule.Add($strMessage, $true)
                        }

                        if ($null -ne $ReferenceToArrayOfOutOfDateModules) {
                            if ($null -ne $ReferenceToArrayOfOutOfDateModules.Value) {
                                ($ReferenceToArrayOfOutOfDateModules.Value) += $strModuleName
                            }
                        }
                    }
                } else {
                    # Conversion of the string version object from Find-Module
                    # failed; this should not happen - throw a warning
                    $strMessage = 'When searching the PowerShell Gallery for the newest version of the module "' + $strModuleName + '", the conversion of its version string "' + $moduleNewestAvailable.Version + '" to a version object failed. This should not be possible!'
                    Write-Warning -Message $strMessage
                }
            } else {
                # Couldn't find the module in the PowerShell Gallery
            }
        }
    }

    if ($boolThrowErrorForMissingModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingModule.Item($strMessage) -eq $true) {
                Write-Error $strMessage
            }
        }
    } elseif ($boolThrowWarningForMissingModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingModule.Item($strMessage) -eq $true) {
                Write-Warning $strMessage
            }
        }
    }

    if ($boolThrowErrorForOutdatedModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForOutdatedModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForOutdatedModule.Item($strMessage) -eq $true) {
                Write-Error $strMessage
            }
        }
    } elseif ($boolThrowWarningForOutdatedModule -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForOutdatedModule.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForOutdatedModule.Item($strMessage) -eq $true) {
                Write-Warning $strMessage
            }
        }
    }
    return $boolResult
}

function Get-QCAuthHeaders {
    param(
        [System.Security.SecureString]$UserID,
        [System.Security.SecureString]$Token
    )
    # Generate a timestamp (seconds since Unix epoch)
    $timestamp = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    # Create the string "<ApiToken>:<Timestamp>"
    $tokenTimeStamp = ([System.Net.NetworkCredential]::new('', $Token).Password) + ':' + $timestamp
    # Compute SHA256 hash of the token:timestamp string
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($tokenTimestamp)
    $hashBytes = $sha256.ComputeHash($bytes)
    # Convert hash bytes to hex string
    $hashString = ($hashBytes | ForEach-Object { $_.ToString('x2') }) -join ''
    # Create the authentication string "<UserId>:<Hash>" and Base64 encode it
    $authString = ([System.Net.NetworkCredential]::new('', $UserID).Password) + ':' + $hashString
    $authBytes = [System.Text.Encoding]::UTF8.GetBytes($authString)
    $authBase64 = [Convert]::ToBase64String($authBytes)
    # Return headers required by QC API
    return @{
        'Authorization' = 'Basic ' + $authBase64
        'Timestamp' = $timestamp.ToString()
    }
}
#endregion Function Definitions #######################################################

$versionPS = Get-PSVersion

#region Quit if PowerShell version is unsupported by Az Module #####################
if ($versionPS -lt [version]'5.1') {
    Write-Warning 'This script requires PowerShell v5.1 or higher. Please upgrade to PowerShell v5.1 or higher and try again.'
    return # Quit script
}
#endregion Quit if PowerShell version is unsupported by Az Module #####################

#region Validate Input #############################################################
if ([string]::IsNullOrEmpty($AKVEntraIdTenantId)) {
    Write-Error 'The parameter AKVEntraIdTenantId must be specified. It is the ID of the Entra ID tenant that includes the Azure Key Vault resource. It must be specified as a string in GUID format, e.g.: ''a11a5a95-33be-428c-a534-0f4712e5110b'''
    return # Quit script
}

if ([string]::IsNullOrEmpty($AKVSubscriptionId)) {
    Write-Error 'The parameter AKVSubscriptionId must be specified. It is the Azure Subscription ID of the subscription that includes the Azure Key Vault. It must be specified as a string in GUID format, e.g.: ''574615a1-23fc-4220-91e7-e86622a070fa'''
    return # Quit script
}

if ([string]::IsNullOrEmpty($AKVName)) {
    Write-Error -Message 'The parameter AKVName must be specified. It is the name of the Azure Key Vault. It must be specified as a string, e.g.: ''MyKeyVault'''
    return # Quit script
}

if ([string]::IsNullOrEmpty($AKVUserIDSecretName)) {
    Write-Error -Message 'The parameter AKVUserIDSecretName must be specified. It is the name of the secret in the Azure Key Vault that contains the QuantConnect User ID. It must be specified as a string, e.g.: ''QuantConnectUserID'''
    return # Quit script
}

if ([string]::IsNullOrEmpty($AKVTokenSecretName)) {
    Write-Error -Message 'The parameter AKVTokenSecretName must be specified. It is the name of the secret in the Azure Key Vault that contains the QuantConnect API token. It must be specified as a string, e.g.: ''QuantConnectAPIToken'''
    return # Quit script
}

if ([string]::IsNullOrEmpty($CSharpFilePath)) {
    Write-Error -Message 'The parameter CSharpFilePath must be specified. It is the path to the C# file to be uploaded to QuantConnect and compiled. It must be specified as a string, e.g.: ''.\MyAlgorithm.cs'''
    return # Quit script
}
if (-not (Test-Path $CSharpFilePath -PathType Leaf)) {
    Write-Error -Message 'No file existed at the path specified for the parameter CSharpFilePath. Please check the path and try again.'
    return # Quit script
}
#endregion Validate Input #############################################################

#region Check for required PowerShell Modules ######################################
$hashtableModuleNameToInstalledModules = @{}
$hashtableModuleNameToInstalledModules.Add('Az.Accounts', @())
$hashtableModuleNameToInstalledModules.Add('Az.KeyVault', @())
$hashtableModuleNameToInstalledModules.Add('Microsoft.PowerShell.SecretManagement', @())
$hashtableModuleNameToInstalledModules.Add('Microsoft.PowerShell.SecretStore', @())
$refHashtableModuleNameToInstalledModules = [ref]$hashtableModuleNameToInstalledModules
Get-PowerShellModuleUsingHashtable -ReferenceToHashtable $refHashtableModuleNameToInstalledModules

$hashtableCustomNotInstalledMessageToModuleNames = @{}

$strAzNotInstalledMessage = 'Az.Accounts and/or Az.KeyVault modules were not found. Please install the full Az module and then try again.' + [System.Environment]::NewLine + 'You can install the Az PowerShell module from the PowerShell Gallery by running the following command:' + [System.Environment]::NewLine + 'Install-Module Az;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
$hashtableCustomNotInstalledMessageToModuleNames.Add($strAzNotInstalledMessage, @('Az.Accounts', 'Az.KeyVault'))

$refhashtableCustomNotInstalledMessageToModuleNames = [ref]$hashtableCustomNotInstalledMessageToModuleNames
$boolResult = Test-PowerShellModuleInstalledUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames

if ($boolResult -eq $false) {
    return # Quit script
}
#endregion Check for required PowerShell Modules ######################################

#region Check for PowerShell module updates ########################################
if ($DoNotCheckForModuleUpdates.IsPresent -eq $false) {
    Write-Verbose 'Checking for module updates...'
    $hashtableCustomNotUpToDateMessageToModuleNames = @{}

    $strAzNotUpToDateMessage = 'A newer version of the Az.Accounts and/or Az.KeyVault modules was found. Please consider updating it by running the following command:' + [System.Environment]::NewLine + 'Install-Module Az -Force;' + [System.Environment]::NewLine + [System.Environment]::NewLine + 'If the installation command fails, you may need to upgrade the version of PowerShellGet. To do so, run the following commands, then restart PowerShell:' + [System.Environment]::NewLine + 'Set-ExecutionPolicy Bypass -Scope Process -Force;' + [System.Environment]::NewLine + '[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;' + [System.Environment]::NewLine + 'Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force;' + [System.Environment]::NewLine + 'Install-Module PowerShellGet -MinimumVersion 2.2.4 -SkipPublisherCheck -Force -AllowClobber;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotUpToDateMessageToModuleNames.Add($strAzNotUpToDateMessage, @('Az.Accounts', 'Az.KeyVault'))

    $refhashtableCustomNotUpToDateMessageToModuleNames = [ref]$hashtableCustomNotUpToDateMessageToModuleNames
    $boolResult = Test-PowerShellModuleUpdatesAvailableUsingHashtable -ReferenceToHashtableOfInstalledModules $refHashtableModuleNameToInstalledModules -ThrowErrorIfModuleNotInstalled -ThrowWarningIfModuleNotUpToDate -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToModuleNames -ReferenceToHashtableOfCustomNotUpToDateMessages $refhashtableCustomNotUpToDateMessageToModuleNames
}
#endregion Check for PowerShell module updates ########################################

#region Connect to Azure ###########################################################
$PSAzureContext = Get-AzContext
if ($null -eq $PSAzureContext) {
    # Connect to Azure without caching credentials to disk:
    [void](Connect-AzAccount -Tenant $AKVEntraIdTenantId -Subscription $AKVSubscriptionId -Scope Process)
    $PSAzureContext = Get-AzContext
    if ($null -eq $PSAzureContext) {
        Write-Warning 'No Azure context found. Please connect to Azure and try again.'
        return # Quit script
    }
}
#endregion Connect to Azure ###########################################################

#region Get QuantConnect User ID and Token from Azure Key Vault ####################
$arrSecretVaults = @(@(Get-SecretVault) | Where-Object { $_.Name -eq ($AKVName + '-AKV') })
if ($arrSecretVaults.Count -eq 0) {
    # Secret vault is not registered

    #TODO: store the connection results in a variable, and make sure the connection was successful?

    # Set the Azure Key Vault as the default secret store
    $parameters = @{
        Name = ($AKVName + '-AKV')
        ModuleName = 'Az.KeyVault'
        VaultParameters = @{
            AZKVaultName = $AKVName
            SubscriptionId = $AKVSubscriptionId
        }
        DefaultVault = $true
    }

    # Register the Azure Key Vault in the SecretManagement module
    Register-SecretVault @parameters
}

# Get the secret stored in AKV
$secureStringUserID = $null
$secureStringToken = $null
$secureStringUserID = Get-Secret -Name $AKVUserIDSecretName -Vault ($AKVName + '-AKV')
$secureStringToken = Get-Secret -Name $AKVTokenSecretName -Vault ($AKVName + '-AKV')

if ($null -eq $secureStringUserID -or $null -eq $secureStringToken) {
    Write-Error 'Unable to retrieve secrets from Azure Key Vault. Please check the secret names and ensure you have sufficient access to the secrets in the key vault and try again.'
    return # Quit script
}

# NOTE: to convert the secure string to a plain text string, use syntax like:
# $strUserID = ([System.Net.NetworkCredential]::new("", $secureStringUserID).Password)
# $strAPIToken = ([System.Net.NetworkCredential]::new("", $secureStringToken).Password)

#endregion Get QuantConnect User ID and Token from Azure Key Vault ####################

#region Authenticate to QuantConnect ###############################################
# Base API URL
$strQCAPIBaseURL = 'https://www.quantconnect.com/api/v2'

# Test authentication with a simple request (optional, to verify credentials).
# We can call an endpoint like reading projects (which requires auth).
try {
    $hashtableTestHeaders = Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken
    $psCustomObjectProjectList = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/projects/read') -Headers $hashtableTestHeaders
    # If unauthorized, the API usually returns 401 or success=false.
    if (-not $psCustomObjectProjectList.success) {
        Write-Error 'Authentication failed! Check UserId/ApiToken.'
        return
    }
    Write-Host ('Authenticated successfully. You have ' + $psCustomObjectProjectList.projects.Count + ' projects in QuantConnect.')
} catch {
    Write-Error -Message ('Error during authentication: ' + $_.Exception.Message)
    return
}
#endregion Authenticate to QuantConnect ###############################################

#region Create a New Project in QuantConnect for our Algorithm #####################
# Use a unique name with timestamp:
$strProjectName = 'PS_AutoTrade_Demo_' + (Get-Date -UFormat %s)
$hashtableCreateProjectBody = @{ 'name' = $strProjectName; 'language' = 'C#' }

Write-Host ('Creating project "' + $strProjectName + '"...')
$psCustomObjectProjectCreationResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/projects/create') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableCreateProjectBody | ConvertTo-Json)
if ($psCustomObjectProjectCreationResponse.success -and $psCustomObjectProjectCreationResponse.projects) {
    $int64ProjectID = $psCustomObjectProjectCreationResponse.projects[0].projectId
    Write-Host ('-> Project created. ProjectId = ' + $int64ProjectID)
} else {
    Write-Error ('Project creation failed: ' + $psCustomObjectProjectCreationResponse.errors)
    return
}
#endregion Create a New Project in QuantConnect for our Algorithm #####################

#region Upload C# Algorithm Code to the QuantConnect Project #######################
# Read file from $CSharpFilePath and store it in $strCSharpAlgorithm:
$strCSharpAlgorithm = Get-Content -Path $CSharpFilePath -Raw -ErrorAction SilentlyContinue
if ($null -eq $strCSharpAlgorithm) {
    Write-Error ('Failed to read C# algorithm file: ' + $CSharpFilePath)
    return
}

if ($strCSharpAlgorithm.Length -eq 0) {
    Write-Error ('C# algorithm file is empty: ' + $CSharpFilePath)
    return
}

Write-Host 'Uploading algorithm code to project...'
$strFileName = 'Main.cs'

$hashtableReadFileBody = @{
    'projectId' = $int64ProjectID
    'name' = $strFileName
}
$psCustomObjectReadFileResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/files/read') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableReadFileBody | ConvertTo-Json -Depth 5)
if ($psCustomObjectReadFileResponse.success) {
    Write-Host '-> Algorithm code file already exists in project. Will update existing file.'
    $hashtableUpdateFileBody = @{
        'projectId' = $int64ProjectID
        'name' = $strFileName
        'content' = $strCSharpAlgorithm
    }
    $psCustomObjectAddFileResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/files/update') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableUpdateFileBody | ConvertTo-Json -Depth 5)
    if ($psCustomObjectAddFileResponse.success) {
        Write-Host '-> Algorithm code file updated in project.'
    } else {
        Write-Error ('Failed to update code file: ' + $psCustomObjectAddFileResponse.errors)
        return
    }
} else {
    Write-Host '-> Algorithm code file does not exist in project. Uploading new file.'
    $hashtableAddFileBody = @{
        'projectId' = $int64ProjectID
        'name' = $strFileName
        'content' = $strCSharpAlgorithm
    }
    $psCustomObjectAddFileResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/files/create') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableAddFileBody | ConvertTo-Json -Depth 5)
    if ($psCustomObjectAddFileResponse.success) {
        Write-Host '-> Algorithm code file added to project.'
    } else {
        Write-Error ('Failed to add code file: ' + $psCustomObjectAddFileResponse.errors)
        return
    }
}
#endregion Upload C# Algorithm Code to the QuantConnect Project #######################

#region Compiling the Project's Code on the QuantConnect Platform ##################
Write-Host 'Compiling the project...'
$hashtableCompileRequestBody = @{ 'projectId' = $int64ProjectID }
$psCustomObjectCompileRequestResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/compile/create') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableCompileRequestBody | ConvertTo-Json)
if (-not $psCustomObjectCompileRequestResponse.success) {
    Write-Error ('Compilation request failed: ' + $psCustomObjectCompileRequestResponse.errors)
    return
}
$strCompileID = $psCustomObjectCompileRequestResponse.compileId

# The compile is asynchronous; we need to poll until it's completed
Write-Host 'Waiting for compilation to complete...'
$boolCompileCompleted = $false
while (-not $boolCompileCompleted) {
    $hashtableCompileStatusBody = @{ 'projectId' = $int64ProjectID; 'compileId' = $strCompileID }
    $psCustomObjectCompileStatusResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/compile/read') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableCompileStatusBody | ConvertTo-Json)
    if (-not $psCustomObjectCompileStatusResponse.success) {
        Write-Error ('Failed to get compile status: ' + $psCustomObjectCompileStatusResponse.errors)
        return
    }

    if ($psCustomObjectCompileStatusResponse.state -ne 'InQueue') {
        $boolCompileCompleted = $true
    } else {
        Start-Sleep -Seconds (Get-Random -Minimum 2 -Maximum 9)
    }
}

if ($psCustomObjectCompileStatusResponse.state -ne 'BuildSuccess') {
    Write-Host 'Compilation failed:'
    foreach ($line in $psCustomObjectCompileStatusResponse.logs) {
        Write-Host $line
    }
    return
} else {
    Write-Host '-> Compilation successful.'
    if (-not [string]::IsNullOrEmpty($psCustomObjectCompileStatusResponse.logs)) {
        Write-Host 'Compilation logs:'
        foreach ($line in $psCustomObjectCompileStatusResponse.logs) {
            Write-Host $line
        }
    }
}
#endregion Compiling the Project's Code on the QuantConnect Platform ##################

#region Run a Backtest on the Compiled Project #####################################
$strBacktestName = 'PS_Backtest_Demo_' + (Get-Date -UFormat %s)
Write-Host ('Starting backtest "' + $strBacktestName + '"...')
$hashtableBacktestRequestBody = @{
    'projectId' = $int64ProjectID
    'compileId' = $strCompileID
    'backtestName' = $strBacktestName
    # TODO: look into sending parameters to the backtest (e.g., start/end date, cash,
    # etc.). See:
    # https://www.quantconnect.com/docs/v2/cloud-platform/api-reference/backtest-management/create-backtest
}
$psCustomObjectBacktestResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/backtests/create') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableBacktestRequestBody | ConvertTo-Json)
if (-not $psCustomObjectBacktestResponse.success) {
    Write-Error ('Failed to start backtest "' + $strBacktestName + '": ' + $psCustomObjectBacktestResponse.errors)
    return
}
$strBacktestID = $psCustomObjectBacktestResponse.backtest.backtestId
Write-Host ('-> Backtest started. BacktestId = "' + $strBacktestID + '"')
# Backtests are asynchronous; need to poll for completion:
$boolBacktestCompleted = $false
Write-Host 'Waiting for backtest to finish...'
while (-not $boolBacktestCompleted) {
    $hashtableBacktestReadBody = @{
        'projectId' = $int64ProjectID
        'backtestId' = $strBacktestID
    }
    $psCustomObjectBacktestReadResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/backtests/read') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableBacktestReadBody | ConvertTo-Json)
    if (-not $psCustomObjectBacktestReadResponse.success) {
        Write-Error ('Failed to read backtest status: ' + $psCustomObjectBacktestReadResponse.errors)
        return
    }
    if ($psCustomObjectBacktestReadResponse.backtest.completed) {
        $boolBacktestCompleted = $true
    } else {
        # Backtest is still running; wait a bit before checking again
        Start-Sleep -Seconds (Get-Random -Minimum 2 -Maximum 9)
    }
}
Write-Host "-> Backtest completed."
if ($psCustomObjectBacktestReadResponse.backtest.status -ne 'Completed.') {
    Write-Warning ('Backtest failed with status "' + $psCustomObjectBacktestReadResponse.backtest.status + '". Error: ' + $psCustomObjectBacktestReadResponse.backtest.error)
    Write-Host ('Stack trace: ' + $psCustomObjectBacktestReadResponse.backtest.stacktrace)
    return
} else {
    Write-Host ('-> Backtest completed successfully.')
    # Check if the backtest has stats
    if ($psCustomObjectBacktestReadResponse.backtest.statistics) {
        $psCustomObjectBacktestStatistics = $psCustomObjectBacktestReadResponse.backtest.statistics
        Write-Host ('-> Backtest statistics available.')
    } else {
        Write-Warning ('No backtest statistics found.')
        return
    }
}

# Display key backtest statistics
Write-Host "Backtest Results Summary:"
$listStatistics = New-Object System.Collections.Generic.List[string]
@($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | ForEach-Object { $_.Name } | Sort-Object | ForEach-Object { [void]($listStatistics.Add($_)) }

# Surface a key statistics first
$strStatisticName = 'Start Equity'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
    try {
        $doubleStartEquity = [double]$strStatisticValue
    } catch {
        Write-Error ('Failed to convert Start Equity to double: ' + $_.Exception.Message)
        $doubleStartEquity = 0.0
    }
} else {
    Write-Warning ('Start Equity statistic not found.')
    $doubleStartEquity = 0.0
}
$strStatisticName = 'End Equity'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
    try {
        $doubleEndEquity = [double]$strStatisticValue
    } catch {
        Write-Error ('Failed to convert End Equity to double: ' + $_.Exception.Message)
        $doubleEndEquity = 0.0
    }
} else {
    Write-Warning ('End Equity statistic not found.')
    $doubleEndEquity = 0.0
}
$strStatisticName = 'Net Profit'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
}
$strStatisticName = 'Total Fees'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
}
$strStatisticName = 'Compounding Annual Return'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
}
$strStatisticName = 'Win Rate'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
}
$strStatisticName = 'Average Win'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
}
$strStatisticName = 'Loss Rate'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
}
$strStatisticName = 'Average Loss'
if ($listStatistics -contains $strStatisticName) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
    [void]($listStatistics.Remove($strStatisticName))
}

# Display the remaining statistics
foreach ($strStatisticName in $listStatistics) {
    $strStatisticValue = @($psCustomObjectBacktestReadResponse.backtest.statistics.PSObject.Properties) | Where-Object { $_.Name -eq $strStatisticName } | Select-Object -ExpandProperty Value
    Write-Host ("   {0}: {1}" -f $strStatisticName, $strStatisticValue)
}
#endregion Run a Backtest on the Compiled Project #####################################

#region Start Live Paper Trading on the Compiled Project ###########################
if ($doubleEndEquity -lt $doubleStartEquity) {
    Write-Warning ('Backtest ended with a loss. Do you wish to start paper trading anyway (Y/N)?')
    $strUserInput = Read-Host
    if ($strUserInput -ne 'Y' -and $strUserInput -ne 'yes') {
        Write-Host 'Paper trading aborted.'
        return
    }
}
Write-Host "Deploying live algorithm in paper trading mode..."
# Build the live deployment request
# Note: For paper trading, brokerage = "QuantConnectBrokerage" (the internal paper broker)
$hashtableProjectNodeReadBody = @{
    'projectId' = $int64ProjectID
}
$psCustomObjectProjectNodeReadResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/projects/nodes/read') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableProjectNodeReadBody | ConvertTo-Json)
if (-not $psCustomObjectProjectNodeReadResponse.success) {
    Write-Error ('Failed to read project nodes: ' + $psCustomObjectProjectNodeReadResponse.errors)
    return
}
$strNodeID = (@($psCustomObjectProjectNodeReadResponse.nodes.live) | Where-Object { -not $_.busy } | Sort-Object -Property 'ram' -Descending | Select-Object -First 1).id
if (-not [string]::IsNullOrEmpty($strNodeID)) {
    Write-Host ('-> Using nodeId: ' + $strNodeID)
} else {
    Write-Warning ('No available nodes found for live paper trading. Quitting!')
    return
}

$hashtableQuantConnectBrokerageCashSettings = @{
    'amount' = $doubleStartEquity
    'currency' = 'USD'
}
$hashtableBrokerageSettings = @{
    'brokerage' = 'QuantConnectBrokerage'
    'cash' = $hashtableQuantConnectBrokerageCashSettings
}
$hashtableLiveTradingCreateBody = @{
    'versionId' = -1 # -1 means latest version (master)
    'projectId' = $int64ProjectID
    'compileId' = $strCompileID
    'brokerage' = $hashtableBrokerageSettings
    'nodeId' = $strNodeID
}
$psCustomObjectLiveTradingCreateResponse = Invoke-RestMethod -Method POST -Uri ($strQCAPIBaseURL + '/live/create') -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) -Body ($hashtableLiveTradingCreateBody | ConvertTo-Json)
if ($psCustomObjectLiveTradingCreateResponse.success -ne $true) {
    Write-Error "Live deployment failed: $($psCustomObjectLiveTradingCreateResponse.errors)"
    return
}

########################### end of developed code; more to come later!
return # Quit script


$deployId = $psCustomObjectLiveTradingCreateResponse.live.deployId
Write-Host "-> Live algorithm deployed! DeployId = $deployId"
Write-Host "(Live trading is now running on QuantConnect's servers in paper mode.)"

# Allow some time to accumulate live data (for demo, wait ~1 minute to gather logs or any trade).
Start-Sleep -Seconds 60
Write-Host "Fetching live trading logs..."
$logReq = @{ "projectId" = $int64ProjectID; "deployId" = $deployId }
$logResp = Invoke-RestMethod -Method POST -Uri "$strQCAPIBaseURL/live/read" `
           -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) `
           -Body ($logReq | ConvertTo-Json)
if ($logResp.success -and $logResp.live) {
    $liveLogs = $logResp.live.logs
    Write-Host "Live Algorithm Logs:"
    if ($liveLogs) {
        $liveLogs | ForEach-Object { Write-Host "   $_" }
        $finalResult = $liveLogs | Where-Object { $_ -match "Final Results" } -First 4
        if ($finalResult) { Write-Host "FINAL RESULT:" -ForegroundColor Green; $finalResult | Write-Host }
    } else {
        Write-Host "   (No logs yet or no trading activity.)"
    }
    # We could also fetch live statistics similar to backtest if needed (liveResults, etc.)
    # Fetch live statistics for alpha/beta
    $liveStats = $logResp.live.statistics
    if ($liveStats -and $liveStats.ContainsKey("Alpha")) {
        Write-Host "Live Trading Metrics:" -ForegroundColor Green
        Write-Host "   Alpha (Excess Return): $($liveStats.Alpha)"
        Write-Host "   Beta (Market Sensitivity): $($liveStats.Beta)"
    }
} else {
    Write-Warning "Could not retrieve live logs. Response: $($logResp | ConvertTo-Json)"
}
#endregion Start Live Paper Trading on the Compiled Project ###########################

### 8. Stop the live algorithm (to clean up after demo) ###
Write-Host "Stopping live algorithm (paper trading)..."
$stopReq = @{ "projectId" = $int64ProjectID; "deployId" = $deployId }
$stopResp = Invoke-RestMethod -Method POST -Uri "$strQCAPIBaseURL/live/stop" `
            -Headers (Get-QCAuthHeaders -UserID $secureStringUserID -Token $secureStringToken) `
            -Body ($stopReq | ConvertTo-Json)
if ($stopResp.success) {
    Write-Host "-> Live algorithm stopped."
} else {
    Write-Warning "Live algorithm stop failed or already stopped: $($stopResp.errors)"
}

Write-Host "=== Demo Script Completed ==="

# Fallback: Calculate alpha/beta via regression (uncomment if API lacks metrics)
<#
function Get-LinearRegression {
    param ($x, $y)
    $n = $x.Count
    if ($n -lt 2) { return @{ Slope = 0; Intercept = 0 } }
    $sumX = ($x | Measure-Object -Sum).Sum
    $sumY = ($y | Measure-Object -Sum).Sum
    $sumXY = 0; $sumXX = 0
    for ($i = 0; $i -lt $n; $i++) {
        $sumXY += $x[$i] * $y[$i]
        $sumXX += $x[$i] * $x[$i]
    }
    $slope = ($n * $sumXY - $sumX * $sumY) / ($n * $sumXX - $sumX * $sumX)
    $intercept = ($sumY - $slope * $sumX) / $n
    return @{ Slope = $slope; Intercept = $intercept }
}
if ($logResp.success -and $logResp.live) {
    $returnLogs = $liveLogs | Where-Object { $_ -match "Periodic Returns" }
    if ($returnLogs) {
        $portfolioReturns = @()
        $nzx50Returns = @()
        foreach ($log in $returnLogs) {
            if ($log -match "Portfolio=([\d.-]+), NZX50=([\d.-]+)") {
                $portfolioReturns += [decimal]$matches[1]
                $nzx50Returns += [decimal]$matches[2]
            }
        }
        if ($portfolioReturns.Count -gt 1) {
            $regression = Get-LinearRegression -x $nzx50Returns -y $portfolioReturns
            Write-Host "Calculated Metrics:" -ForegroundColor Green
            Write-Host "   Alpha (Excess Return): $($regression.Intercept:F4)"
            Write-Host "   Beta (Market Sensitivity): $($regression.Slope:F4)"
        }
    }
}
#>
