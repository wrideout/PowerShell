#requires -version 2.0

param([string[]]$PreCacheList)

if (!(Test-Path variable:\helpCache) -or $RefreshCache) {
    $SCRIPT:helpCache = @{}
}    

function Resolve-MemberOwnerType
{
    [CmdletBinding()]
    param
    (
        [system.management.automation.psmethod]$method
    )

    # TODO: support overloads, support interface definitions

    $PSCmdlet.WriteVerbose("Resolving $($method.name)'s owning Type.")
   
    # hackety-hack - this is prone to breaking in the future
    $targetType = [system.management.automation.psmethod].getfield("baseObject", "Instance,NonPublic").getvalue($method)
    
    # [system.runtimetype] is special-cased in powershell - you can't reference it?
    if (-not ($targetType.GetType().fullname -eq "System.RuntimeType"))
    {
        $targetType = $targetType.GetType()
    }

    if ($method.OverloadDefinitions -match "static")
    {
        $flags = "Static,Public"
    }
    else
    {
        $flags = "Instance,Public"
    }

    # FIXME: support overloads    
    $methodInfo = $targetType.GetMethods($flags) | ?{$_.Name -eq $method.Name}| select -first 1
    
    if (-not $methodInfo)
    {
        # this shouldn't happen.
        throw "Could not resolve owning type!"
    }
    
    $declaringType = $methodInfo.DeclaringType
    
    $PSCmdlet.WriteVerbose("Owning Type is $($targetType.fullname). Method declared on $($declaringType.fullname).")

    $declaringType
}

function Get-DocsLocation
{
    [CmdletBinding()]
    param
    (
        [type]$type,
        
        [switch]$Online,
        
        [switch]$Members,
        
        [switch]$Static
    )
    
    # get documentation filename, assembly location and assembly codebase
    $docFilename = [io.path]::changeextension([io.path]::getfilename($type.assembly.location), ".xml")
    $location = [io.path]::getdirectoryname($type.assembly.location)
    $codebase = (new-object uri $type.assembly.codebase).localpath
    
    $PSCmdlet.WriteVerbose("Documentation file is $docFilename")
    
    if (-not $Online.IsPresent)
    {
        # try localized location (typically newer than base framework dir)
        $frameworkDir = "${env:windir}\Microsoft.NET\framework\v2.0.50727"
        $lang = [system.globalization.cultureinfo]::CurrentUICulture.parent.name

        # I love looking at this. A Duff's Device for PowerShell.. well, maybe not.
        switch
            (
            "${frameworkdir}\${lang}\$docFilename",
            "${frameworkdir}\$docFilename",
            "$location\$docFilename",
            "$codebase\$docFilename"
            )
        {
            { test-path $_ } { $_; return; }
            
            default
            {
                # try next path
                continue;
            }        
        }       
    }

    # failed to find local docs, is it from MS?
    if ((Get-ObjectVendor $type) -like "*Microsoft*")
    {
        # drop locale - site will redirect to correct variation based on browser accept-lang
        $suffix = ""
        if ($Members.IsPresent)
        {
            $suffix = "_members"
        }
        
        new-object uri ("http://msdn.microsoft.com/library/{0}{1}.aspx" -f $type.fullname,$suffix)
        
        return
    }
    
    $PSCmdlet.WriteWarning("Sorry, I couldn't find any local documentation for ${type}.")
}

# Dig out something that might lead us to the vendor of this Object
function Get-ObjectVendor
{
    [CmdletBinding()]
    param
    (
        [type]$type,
        [switch]$CompanyOnly
    )

    $assembly = $type.assembly
    $attrib = $assembly.GetCustomAttributes([Reflection.AssemblyCompanyAttribute], $false) | select -first 1        
    
    if ($attrib.Company)
    {
        # try company
        $attrib.Company
        return
    }
    else
    {
        if ($CompanyOnly) { return }
        
        # try copyright
        $attrib = $assembly.GetCustomAttributes([Reflection.AssemblyCopyrightAttribute], $false) | select -first 1
        
        if ($attrib.Copyright)
        {
            $attrib.Copyright
            return
        }
    }
    $PSCmdlet.WriteVerbose("Assembly has no [AssemblyCompany] or [AssemblyCopyright] attributes.")
}

function Get-HelpSummary
{
        [CmdletBinding()]
        param
        (        
            [string]$file,
            [reflection.assembly]$assembly,
            [string]$selector
        )
        
        if ($helpCache.ContainsKey($assembly))
        {            
            $xml = $helpCache[$assembly]
            
            $PSCmdlet.WriteVerbose("Docs were found in the cache.")
        }
        else
        {
            # cache it
            Write-Progress -id 1 "Caching Help Documentation" $assembly.getname().name

            # cache this for future lookups. It's a giant pig. Oink.
            $xml = [xml](gc $file)
            
            $helpCache.Add($assembly, $xml)
            
            Write-Progress -id 1 "Caching Help Documentation" $assembly.getname().name -completed
        }

        $PSCmdlet.WriteVerbose("Selector is $selector")        

        # TODO: support overloads
        $summary = $xml.doc.members.SelectSingleNode("member[@name='$selector' or starts-with(@name,'$selector(')]").summary
        
        $summary
}

function Show-Help
{
@"    
    
   
SYNTAX

$((get-help get-objecthelp).split([char]13) | % { "$_" })
"@
}

function Get-ObjectHelp
{    
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
        [ValidateNotNull()]
        $Object,

        [Parameter()]
        [switch]$Online,
        
        [Parameter()]
        [switch]$Member,
        
        [Parameter()]
        [switch]$Static
    )
    
    process 
    {
        if ($Object -is [string])
        {
            $PSCmdlet.WriteVerbose("A string was passed - reparsing as expression.")
            
            # they probably meant to pass the string inside '(' and ').'
            try
            {
                # e.g. "[int]::gettype" was passed without being wrapped
                # in new evaluative parentheses.
                $Object = invoke-expression $Object
            }
            catch
            {
                if ($_.fullyqualifiederrorid -eq "TypeNotFound,Microsoft.PowerShell.Commands.InvokeExpressionCommand")
                {
                    $PSCmdlet.WriteWarning("I don't recognize the Type in ${InputObject}. Are you sure you've typed it correctly?")
                }
                else
                {            
                    $PSCmdlet.WriteWarning("A string was passed and was parsed as an expression, and failed. " +
                        "If you really meant to find help on strings, pass [string] instead.")
                }
                $PSCmdlet.WriteVerbose($_)
                
                return
            }
        }

        $type = $Object.GetType()    
        $PSCmdlet.WriteVerbose("InputObject Type is $($type.Fullname)")
        
        $selector = $null
        
        # won't work with $type; case statements don't match with type literals?
        switch ($type.FullName)
        {
            "System.RuntimeType"
            {
                $PSCmdlet.WriteVerbose("[runtimetype]")
            
                $type = $Object
                $selector = "T:$($type.FullName)"
                
                break;
            }

            "System.Management.Automation.PSMethod"
            {
                $PSCmdlet.WriteVerbose("[psmethod]")
                
                $type = Resolve-MemberOwnerType $Object
                
                # TODO: support overloaded methods
                $selector = "M:$($type.FullName).$($Object.Name)"            
                
                break;
            }

            default
            {
                $PSCmdlet.WriteVerbose("[object]")
                $selector = "T:$($type.FullName)"            
            }
        }
        
        # do we have an assembly help xml somewhere?
        $docs = Get-DocsLocation $type -Online:$Online.IsPresent -Members:$Member.IsPresent -Static:$Static.IsPresent

        if ($docs)
        {
            $PSCmdlet.WriteVerbose("Found $docs")
            
            if ($docs -is [uri])
            {
                # Could not find local xml, but object is from Microsoft. Offer to view MSDN.
                $title = "Microsoft Developer Network"
                $message = "No local help for $($type.fullname).`n`nDo you want to visit this object's documentation page on MSDN?"
                $options = [System.Management.Automation.Host.ChoiceDescription[]]("&Yes", "&No")

                $result = $host.ui.PromptForChoice($title, $message, $options, 0)
                
                if ($result -eq 0) {
                    [diagnostics.process]::Start("iexplore.exe", $docs) > $null
                }
                return
            }
                    
            # get summary, if possible
            $summary = Get-HelpSummary $docs $type.assembly $selector
                    
            if ($summary)
            {
                [string]::empty
                
                # TODO: parse out <see ...> tags and create a PromptForChoice list to lookup referenced type(s).
                if ($summary.selectnodes) {
                    $see = $summary.selectnodes("see")
                }
                
                if (($Object -eq 42) -and (!$PSCmdlet.Force)) {
                
                    "What do you get if you multiply six by nine?"
                    [string]::empty
                    "That's it. That's all there is."
                
                } else {
                
                    $text = & {
                        if ($summary.innerxml) {
                            $summary.innerxml.trim()
                        }
                        else
                        {
                            $summary.trim()
                        }
                    }
                    
                    # strip <see ... /> tags
                    $text -replace [regex]'<see.*?"?:(.*?)"\s/>', '$1'
                }

                if ((Test-Path Variable:\see) -and $see) {
                    #Show-References 
                    # TODO: list of <see cref="foo" /> types
                }
                
                [string]::empty                
            }
            else
            {
                Write-Host "While some local documentation was found, it was incomplete. Sorry!"            
            }
        }
        else 
        {
            Write-Host "Sorry, I couldn't find any local documentation for ${type}."
            
            $vendor = Get-ObjectVendor $type -CompanyOnly
            
            if ($vendor)
            {
                # needed for urlencode
                add-type -a system.web

                write-host "However, it looks like the vendor of this Object is '${vendor}.'"
                
                $title = "Bing Search"
                $message = "Do you want to search for this object's documentation?"
                $options = [System.Management.Automation.Host.ChoiceDescription[]]("&Yes", "&No")

                $result = $host.ui.PromptForChoice($title, $message, $options, 0)
                
                if ($result -eq 0) {
                    # encode our question
                    $q = [system.web.httputility]::urlencode(("`"{0}`" {1}" -f $vendor, $type))
                    
                    # fire up the browser
                    [diagnostics.process]::Start("http://www.bing.com/results.aspx?q=$q")
                }
            }
        }
    }    
}

# cache common assembly help
function Preload-Documentation
{       
    if ($SCRIPT:helpCache.Keys.Count -eq 0) {
        # mscorlib
        $file = Get-DocsLocation ([int])
        Get-HelpSummary $file ([int].assembly) "T:System.Int32" > $null
        
        # system
        $file = Get-DocsLocation ([regex])    
        Get-HelpSummary $file ([regex].assembly) "T:System.Regex" > $null
    }
}

<#
.ForwardHelpTargetName Get-Help
.ForwardHelpCategory Cmdlet
#>
function Get-Help {
    # our proxy command generated from [proxycommand]::create((gcm get-help))
    [CmdletBinding(DefaultParameterSetName='AllUsersView')]
    param(
        [Parameter(Position=0, ValueFromPipelineByPropertyName=$true)]
        [System.String]
        ${Name},

        [System.String]
        ${Path},

        [System.String[]]
        ${Category},

        [System.String[]]
        ${Component},

        [System.String[]]
        ${Functionality},

        [System.String[]]
        ${Role},

        [Parameter(ParameterSetName='DetailedView', Mandatory=$true)]
        [Switch]
        ${Detailed},

        [Parameter(ParameterSetName='AllUsersView')]
        [Switch]
        ${Full},

        [Parameter(ParameterSetName='Examples', Mandatory=$true)]
        [Switch]
        ${Examples},

        [Parameter(ParameterSetName='Parameters', Mandatory=$true)]
        [System.String]
        ${Parameter},
        
        [Parameter(ParameterSetName='ObjectHelp', ValueFromPipeline = $true, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        ${Object},

        [Parameter(ParameterSetName='ObjectHelp')]
        [Switch]
        ${Member},
        
        [Parameter(ParameterSetName='ObjectHelp')]
        [Switch]
        ${Static},        

        [switch]
        ${Online}
    )

    begin
    {
        try 
        {
            if ($PSCmdlet.ParameterSetName -eq "ObjectHelp") 
            {                                
                Preload-Documentation
                
                $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Get-ObjectHelp', [System.Management.Automation.CommandTypes]::Function)
                $scriptCmd = { & $wrappedCmd @PSBoundParameters }
                $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)                
            
            } 
            else 
            {
				# Working around a bug in PowerShell (try man -?) where it passes in the wrong category info for aliases.
                if ($Name)
                {
                    $isAlias = (Microsoft.PowerShell.Core\Get-Command $Name -ErrorAction 'SilentlyContinue').CommandType -eq 'Alias'
				    if ($isAlias)
				    {
				        $PSBoundParameters['Category'] = 'Alias'
				    }
                }

                $outBuffer = $null
                if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer) -and $outBuffer -gt 1024)
                {
                    $PSBoundParameters['OutBuffer'] = 1024
                }

                $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Microsoft.PowerShell.Core\Get-Help', [System.Management.Automation.CommandTypes]::Cmdlet)
                $scriptCmd = { & $wrappedCmd @PSBoundParameters }
                $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)   
            }
            $steppablePipeline.Begin($PSCmdlet)
        } 
        catch {
            throw
        }
    }

    process
    {
        try {        
            $steppablePipeline.Process($_)
        } catch {
            throw
        }
    }

    end
    {
        try {
            $steppablePipeline.End()
        } catch {
            throw
        }
    }
}

Export-ModuleMember Get-Help

<#
    NAME
    
        ObjectHelp Extensions Module 0.3 for PowerShell 2.0
     
    SYNOPSIS
    
         Get-Help -Object allows you to display usage and summary help for .NET Types and Members.
         
    DETAILED DESCRIPTION
    
        Get-Help -Object allows you to display usage and summary help for .NET Types and Members.
    
        If local documentation is not found and the object vendor is Microsoft, you will be directed
        to MSDN online to the correct page. If the vendor is not Microsoft and vendor information
        exists on the owning assembly, you will be prompted to search for information using Bing.
     
    TODO
     
         * localize strings into PSD1 file
         * Implement caching in hashtables. XMLDocuments are fat pigs.
         * Support getting property/field help
         * PowerTab integration?
         * Test with Strict Parser
             
    EXAMPLES

        # get help on a type
        PS> get-help -obj [int]

        # get help against live instances
        PS> $obj = new-object system.xml.xmldocument
        PS> get-help -obj `$obj

        or even:
        
        PS> get-help -obj 42
        
        # get help against methods
        PS> get-help -obj `$obj.Load

        # explictly try msdn
        PS> get-help -obj [regex] -online

        # go to msdn for regex's members
        PS> get-help -obj [regex] -online -member
        
        # pipe support
        PS> 1,[int],[string]::format | get-help -verbose
    
    CREDITS
    
        Author: Oisin Grehan (MVP)
        Blog  : http://www.nivot.org/
    
        Have fun!    
#>

# SIG # Begin signature block
# MIIfVQYJKoZIhvcNAQcCoIIfRjCCH0ICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUOgje6nQKRRkuhE67mtGDyjEY
# MU6gghqHMIIGbzCCBVegAwIBAgIQA4uW8HDZ4h5VpUJnkuHIOjANBgkqhkiG9w0B
# AQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVk
# IElEIENBLTEwHhcNMTIwNDA0MDAwMDAwWhcNMTMwNDE4MDAwMDAwWjBHMQswCQYD
# VQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRp
# bWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDGf7tj+/F8Q0mIJnRfituiDBM1pYivqtEwyjPdo9B2gRXW1tvhNC0FIG/BofQX
# Z7dN3iETYE4Jcq1XXniQO7XMLc15uGLZTzHc0cmMCAv8teTgJ+mn7ra9Depw8wXb
# 82jr+D8RM3kkwHsqfFKdphzOZB/GcvgUnE0R2KJDQXK6DqO+r9L9eNxHlRdwbJwg
# wav5YWPmj5mAc7b+njHfTb/hvE+LgfzFqEM7GyQoZ8no89SRywWpFs++42Pf6oKh
# qIXcBBDsREA0NxnNMHF82j0Ctqh3sH2D3WQIE3ome/SXN8uxb9wuMn3Y07/HiIEP
# kUkd8WPenFhtjzUmWSnGwHTPAgMBAAGjggM6MIIDNjAOBgNVHQ8BAf8EBAMCB4Aw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDCCAcQGA1UdIASC
# AbswggG3MIIBswYJYIZIAYb9bAcBMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3
# dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUF
# BwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUA
# cgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMA
# YwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQA
# IABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAA
# UABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkA
# bQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4A
# YwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYA
# ZQByAGUAbgBjAGUALjAfBgNVHSMEGDAWgBQVABIrE5iymQftHt+ivlcNK2cCzTAd
# BgNVHQ4EFgQUJqoP9EMNo5gXpV8S9PiSjqnkhDQwdwYIKwYBBQUHAQEEazBpMCQG
# CCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKG
# NWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENB
# LTEuY3J0MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDANBgkqhkiG9w0B
# AQUFAAOCAQEAvCT5g9lmKeYy6GdDbzfLaXlHl4tifmnDitXp13GcjqH52v4k498m
# bK/g0s0vxJ8yYdB2zERcy+WPvXhnhhPiummK15cnfj2EE1YzDr992ekBaoxuvz/P
# MZivhUgRXB+7ycJvKsrFxZUSDFM4GS+1lwp+hrOVPNxBZqWZyZVXrYq0xWzxFjOb
# vvA8rWBrH0YPdskbgkNe3R2oNWZtNV8hcTOgHArLRWmJmaX05mCs7ksBKGyRlK+/
# +fLFWOptzeUAtDnjsEWFuzG2wym3BFDg7gbFFOlvzmv8m7wkfR2H3aiObVCUNeZ8
# AB4TB5nkYujEj7p75UsZu62Y9rXC8YkgGDCCBpswggWDoAMCAQICEAoVPQh11uMo
# zhH2mVCPvBEwDQYJKoZIhvcNAQEFBQAwbzELMAkGA1UEBhMCVVMxFTATBgNVBAoT
# DERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEuMCwGA1UE
# AxMlRGlnaUNlcnQgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EtMTAeFw0xMjA5
# MTEwMDAwMDBaFw0xMzA5MTgxMjAwMDBaMGcxCzAJBgNVBAYTAlVTMQswCQYDVQQI
# EwJDTzEVMBMGA1UEBxMMRm9ydCBDb2xsaW5zMRkwFwYDVQQKExA2TDYgU29mdHdh
# cmUgTExDMRkwFwYDVQQDExA2TDYgU29mdHdhcmUgTExDMIIBIjANBgkqhkiG9w0B
# AQEFAAOCAQ8AMIIBCgKCAQEAvtSuQar5tMsJw1RaGhLz9ECpar95hZ4d0dHivIK2
# maFz8QQeSJbqQbouzWJWfgvncWIhfZs9wyJjCdHbW7xVSmK/GPI+mfTky66lP99W
# dfV6gY0WkBYkFvzTQ0s/P9+qS1PEfAb8CFZYx3Ti8GVSUVSS87/TZm1SS+lnCg4m
# Rlp+BM9FDaK8IA/UjUjl277qmVnfvB35ey4I81421hsl5uJsZ5ZB+C9PFvkIzhR4
# Eo7o7R13Erjiryran/aJb77YgjRueC+EZ8rCx+kDq5TsLAzYZQwfgaKXpFlvXdiF
# vdFD6Hf6j4QonmtwG1RDYS5Vp1O/d2y/aunhKW3Wr94kywIDAQABo4IDOTCCAzUw
# HwYDVR0jBBgwFoAUe2jOKarAF75JeuHlP9an90WPNTIwHQYDVR0OBBYEFPNABKbs
# Aid4soPC5f6eM7TdDs50MA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEF
# BQcDAzBzBgNVHR8EbDBqMDOgMaAvhi1odHRwOi8vY3JsMy5kaWdpY2VydC5jb20v
# YXNzdXJlZC1jcy0yMDExYS5jcmwwM6AxoC+GLWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0
# LmNvbS9hc3N1cmVkLWNzLTIwMTFhLmNybDCCAcQGA1UdIASCAbswggG3MIIBswYJ
# YIZIAYb9bAMBMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5j
# b20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIA
# QQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMA
# YQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4A
# YwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAA
# UwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAA
# QQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkA
# YQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIA
# YQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUA
# LjCBggYIKwYBBQUHAQEEdjB0MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wTAYIKwYBBQUHMAKGQGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRENvZGVTaWduaW5nQ0EtMS5jcnQwDAYDVR0TAQH/
# BAIwADANBgkqhkiG9w0BAQUFAAOCAQEAIK6UA1qKXKbv5fphePINxEahQyZCLaFY
# OO+Q2jcrnrXofOGqZOLz/M33cJErAOyZQvKANOKybsMlpzmkQpP8jJsNRXuDmEOl
# bilUkwssxSTHeLfKgfRbB5RMi7RhvWyzhoC+FELHI+99VDJAQzWYwokAsSHohUPj
# QsEn6sI2ITvxOKgZKurzzFTmFberEA55RoszUKRcP9E0aW6L94ysSpVmzcJxY8ZE
# ny91ACmlHCSzxjrON/nlikzFtDTRlLr//dAm/XPXNlpEA1gIqS4zqUapRyFP/VhW
# NgHsjMdwIHpRAgpLPkvobG+TMHobi2IqkzSt5SrrDVcaH7t+RpKlLTCCBqAwggWI
# oAMCAQICEAf0c2+v70CKH2ZA8mXRCsEwDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTExMDIxMDEyMDAwMFoXDTI2MDIxMDEyMDAwMFowbzELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEu
# MCwGA1UEAxMlRGlnaUNlcnQgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EtMTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJx8+aCPCsqJS1OaPOwZIn8M
# y/dIRNA/Im6aT/rO38bTJJH/qFKT53L48UaGlMWrF/R4f8t6vpAmHHxTL+WD57tq
# BSjMoBcRSxgg87e98tzLuIZARR9P+TmY0zvrb2mkXAEusWbpprjcBt6ujWL+RCeC
# qQPD/uYmC5NJceU4bU7+gFxnd7XVb2ZklGu7iElo2NH0fiHB5sUeyeCWuAmV+Uue
# rswxvWpaQqfEBUd9YCvZoV29+1aT7xv8cvnfPjL93SosMkbaXmO80LjLTBA1/FBf
# rENEfP6ERFC0jCo9dAz0eotyS+BWtRO2Y+k/Tkkj5wYW8CWrAfgoQebH1GQ7XasC
# AwEAAaOCA0AwggM8MA4GA1UdDwEB/wQEAwIBBjATBgNVHSUEDDAKBggrBgEFBQcD
# AzCCAcMGA1UdIASCAbowggG2MIIBsgYIYIZIAYb9bAMwggGkMDoGCCsGAQUFBwIB
# Fi5odHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRt
# MIIBZAYIKwYBBQUHAgIwggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABo
# AGkAcwAgAEMAZQByAHQAaQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0
# AGUAcwAgAGEAYwBjAGUAcAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBn
# AGkAQwBlAHIAdAAgAEMAUAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBs
# AHkAaQBuAGcAIABQAGEAcgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABp
# AGMAaAAgAGwAaQBtAGkAdAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABh
# AHIAZQAgAGkAbgBjAG8AcgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABi
# AHkAIAByAGUAZgBlAHIAZQBuAGMAZQAuMA8GA1UdEwEB/wQFMAMBAf8weQYIKwYB
# BQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20w
# QwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9j
# cmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4
# oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJv
# b3RDQS5jcmwwHQYDVR0OBBYEFHtozimqwBe+SXrh5T/Wp/dFjzUyMB8GA1UdIwQY
# MBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQCPJ3L2
# XadkkG25JWDGsBetHvmd7qFQgoTHKlU1rBhCuzbY3lPwkie/M1WvSlq4OA3r9OQ4
# DY38RegRPAw1V69KXHlBV94JpA8RkxA6fG40gz3xb/l0H4sRKsqbsO/TgJJEQ9El
# yTkyITHnKYLKxEGIyAeBp/1bIRd+HbqpY2jIctHil1lyQubYp7a6sy7JZK2TwuHl
# eKm36bs9MZKa3pGJJgoLXj5Oz2HrWmxkBT57q3+mWN0elcdeW8phnTR1pOUHSPhM
# 0k88EjW7XxFljf1yueiXIKUxh77LAx/LAzlu97I7OJV9cEW4VvWAcoEETIBzoK0t
# OfUCyOSF1TskL7NsMIIGzTCCBbWgAwIBAgIQBv35A5YDreoACus/J7u6GzANBgkq
# hkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5j
# MRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBB
# c3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEwMDAwMDAwWhcNMjExMTEwMDAwMDAw
# WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElE
# IENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDogi2Z+crCQpWl
# gHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc+BZZ3561JBXCmLm0d0ncicQK2q/L
# XmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJgv7HxzFIgHweog+SDlDJxofrNj/Y
# MMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWXfaPHQ90C6Ds97bFBo+0/vtuVSMTu
# HrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYKxI2+0s4Grq2Eb0iEm09AufFM8q+Y
# +/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8eCnRcQrhofrfVdwonVnwPYqQ/MhRg
# lf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8BAf8EBAMCAYYwOwYDVR0lBDQwMgYI
# KwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcDAwYIKwYBBQUHAwQGCCsGAQUFBwMI
# MIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgBhv1sAAEEMIIBpDA6BggrBgEFBQcC
# ARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20vc3NsLWNwcy1yZXBvc2l0b3J5Lmh0
# bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQA
# aABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUA
# dABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkA
# ZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUA
# bAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgA
# aQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAA
# YQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAA
# YgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwEgYDVR0TAQH/
# BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2VydHMuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCBgQYDVR0fBHow
# eDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJl
# ZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAdBgNVHQ4EFgQUFQASKxOYspkH7R7f
# or5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZI
# hvcNAQEFBQADggEBAEZQPsm3KCSnOB22WymvUs9S6TFHq1Zce9UNC0Gz7+x1H3Q4
# 8rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0ckeFlFbR67s2hHfMJKXzBBlVqefj
# 56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aGMWKoqDPkmzmnxPXOHXh2lCVz5Cqr
# z5x2S+1fwksW5EtwTACJHvzFebxMElf+X+EevAJdqP77BzhPDcZdkbkPZ0XN1oPt
# 55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3Fu4Pw718CqDuLAhVhSK46xgaTfwq
# Ia1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgpECYxggQ4MIIENAIBATCBgzBvMQsw
# CQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cu
# ZGlnaWNlcnQuY29tMS4wLAYDVQQDEyVEaWdpQ2VydCBBc3N1cmVkIElEIENvZGUg
# U2lnbmluZyBDQS0xAhAKFT0IddbjKM4R9plQj7wRMAkGBSsOAwIaBQCgeDAYBgor
# BgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEE
# MBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQ9
# xxjWnPDv17vo/mJ1Zz6uDqLu6TANBgkqhkiG9w0BAQEFAASCAQAOLPXTftZoJqLY
# 0BGxSyzg0ZiYF4hguluTBfVup01E6ob+OCcFOPrdNlJ+YJTU6Arr7KIbkwyDhR5k
# 4aVNVqlFBKD4vPQNuG9MZOZlOb9f3JcgYMTVeS+Oi2m0De7rx4mYx/rHEYiynKYn
# qD/qA5QU8M0Vre8N66R71z6M8oAegJSFEWPJm/Venb0q/0WQP0Db87v8CH07S9lG
# P0QlyLt3lCPrMjiHUbph9D9tgj8iXHwtUPzvYWWGHdu7zqsWC3jM8IiFA44wYwTQ
# ppY0CkTk6NY00e5FzXgWhFMKQ9j3Wg8xLLRWdLxgrr1rucP0urndBwfxzJTHFl3w
# zpsVBTcUoYICDzCCAgsGCSqGSIb3DQEJBjGCAfwwggH4AgEBMHYwYjELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEhMB8GA1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xAhADi5bw
# cNniHlWlQmeS4cg6MAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0xMjEyMTcwMTUxMTJaMCMGCSqGSIb3DQEJBDEW
# BBTaYD7dy0hx2BsIZwOOVHwxTxm/AjANBgkqhkiG9w0BAQEFAASCAQCD0qYn3Qd5
# PBUmm0YDhi2FQ5tzM+m4vCRW8ZAd5sh2Jl1drXkCkjYvSd3JTMYX+oOHkl3UQogD
# noOw8myCoaEewIJkEnlQeBwq4FGJAt8WEeYzvrpearxpbnk9eM1F8I3NBXSNx/AU
# qPkKsjZM64AZf2KOUOAaxqiNWxqKv+UXDzSMhsjTPP2s+E9txTHbLLch+rc8CHf9
# koxSESAjn75h5OG9Gb4TWJiwfyKxyybNZ0zpCd0qj0JLlTBP6oJfA0RLEQq2X8xa
# Tq0EWD+ryXKfO8KiDqykpVQyDazK+rw0W7zR2bVMiJGTaZ/PalefNAarb5lRrk/F
# 0eOk3j8J32zR
# SIG # End signature block
