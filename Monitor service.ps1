#Requires -Version 5.1
#Requires -Modules ImportExcel, Toolbox.HTML, Toolbox.EventLog

<# 
    .SYNOPSIS
        Stop a service, kill a process and start a service.

    .DESCRIPTION
        This script reads a .JSON input file and performs the requested actions.
        When 'SetServiceStartupType' is used, the service startup type will be 
        corrected when needed. Each entry in 'Action' will be performed in the
        order 'StopService', 'KillProcess', 'StartService'.

    .PARAMETER ImportFile
        A .JSON file containing the script arguments. See the Example.json file
        for more details on what input is accepted.

    .PARAMETER SetServiceStartType
        Configure the service with the correct startup type

        Valid values:
        - Automatic
        - Disabled
        - DelayedAutostart
        - Manual

    .PARAMETER Action
        Stop a service, kill a process or start a service on a specific computer

    .PARAMETER SendMail.To
        List of e-mail addresses where to send the e-mail too.
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Monitor\Monitor service\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Date         = 'ScriptStartTime'
                NoFormatting = $true
                Unique       = $True
            }
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        if (-not ($mailTo = $file.SendMail.To)) {
            throw "Input file '$ImportFile': No 'SendMail.To' addresses found."
        }

        if (-not ($Tasks = $file.Tasks)) {
            throw "Input file '$ImportFile': No 'Tasks' found."
        }

        foreach ($task in $Tasks) {
            #region Task properties
            $properties = $task.PSObject.Properties.Name
            
            @(
                'SetServiceStartupType',
                'ComputerName',
                'Execute'
            ) | Where-Object { $properties -notContains $_ } | 
            ForEach-Object {
                throw "Property '$_' not found in one of the 'Tasks'."
            }
            #endregion

            $actionInTask = $false

            #region Task.SetServiceStartupType properties
            $properties = $task.SetServiceStartupType.PSObject.Properties.Name

            $serviceNamesInStartupTypes = @()
        
            $serviceStartupTypes = @(
                'Automatic', 
                'DelayedAutostart', 
                'Disabled',
                'Manual'
            )

            foreach ($startupTypeName in $serviceStartupTypes) {
                if ($properties -notContains $startupTypeName) {
                    throw "Property 'SetServiceStartupType.$startupTypeName' not found in one of the 'Tasks'."
                }
                if ($task.SetServiceStartupType.$startupTypeName) {
                    $actionInTask = $true
                }

                foreach (
                    $serviceName in 
                    $task.SetServiceStartupType.$startupTypeName
                ) {
                    $serviceNamesInStartupTypes += $serviceName
                }
            }

            if (
                $duplicateServiceNamesInStartupType = $serviceNamesInStartupTypes | Group-Object | 
                Where-Object { ($_.Name) -and ($_.Count -gt 1) } 
            ) {
                throw "Service '$($duplicateServiceNamesInStartupType.Name)' can only have one StartupType."
            }
            #endregion

            #region Task.Execute properties
            $properties = $task.Execute.PSObject.Properties.Name
            
            @(
                'StopService', 
                'KillProcess',
                'StartService'
            ) | ForEach-Object {
                if ($properties -notContains $_) {
                    throw "Property 'Execute.$_' not found in one of the 'Tasks'."
                }
                if ($task.Execute.$_) {
                    $actionInTask = $true
                }
            }
            #endregion

            if (-not $task.ComputerName) {
                throw "Input file '$ImportFile': No 'ComputerName' found in one of the 'Tasks'."
            }

            if (-not $actionInTask) {
                throw "Input file '$ImportFile': Contains a task where properties 'SetServiceStartupType' and 'Execute' are both empty."
            }

            if (
                $duplicateComputers = $task.ComputerName | Group-Object | 
                Where-Object { $_.Count -gt 1 } 
            ) {
                throw "duplicate ComputerName '$($duplicateComputers.Name)' found in a single task"
            }

            $task.SetServiceStartupType.Disabled | 
            Where-Object { $task.Execute.StartService -contains $_ } |
            ForEach-Object {
                throw "Service '$_' cannot have StartupType 'Disabled' and 'StartService' at the same time"
            }
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        Foreach ($task in $Tasks) {
            foreach ($computerName in $task.ComputerName) {
                try {
                    $params = @{
                        ComputerName = $computerName
                        Name         = $s.Name 
                        ErrorAction  = 'Stop'
                    }
                    $service = Get-Service @params
                
                    Write-Verbose "'$computerName' Service '$($s.Name)' State '$($service.Status)' StartType '$($service.StartType)'"
        
                    if ($service.StartType -ne $s.StartType) {
                        if ($s.StartType -eq 'Disabled') {
                            Write-Verbose "'$computerName' set StartupType to 'Disabled'"
                            $service | Set-Service -StartupType 'Disabled'
                        }
                        if ($s.StartType -eq 'Automatic') {
                            Write-Verbose "'$computerName' set StartupType to 'Automatic'"
                            $service | Set-Service -StartupType 'Automatic'
                        }
                    }
                    
                    if ($service.Status -ne $s.Status) {
                        if ($s.Status -eq 'Stopped') {
                            Write-Verbose "'$computerName' stop service"
                            $service | Stop-Service
                        }
                        if ($s.Status -eq 'Running') {
                            Write-Verbose "'$computerName' start service"
                            $service | Start-Service
                        }
                    }
        
                    $service = Get-Service @params
                
                    Write-Verbose "'$computerName' Service '$($s.Name)' State '$($service.Status)' StartType '$($service.StartType)'"
                }
                catch {
                    Write-Warning "Failed stopping service '$($s.Name)' on '$computerName': $_"
                }   
            }
        }    
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {

    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}
