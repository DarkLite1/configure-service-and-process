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
        Function Test-DelayedAutoStartHC {
            Param (
                [parameter(Mandatory)]
                [String]$ComputerName,
                [parameter(Mandatory)]
                [alias('Name')]
                [String]$ServiceName
            )
        
            try {
                $params = @{
                    ComputerName = $ComputerName
                    ArgumentList = $ServiceName 
                    ErrorAction  = 'Stop'
                }
                Invoke-Command @params -ScriptBlock {
                    Param (
                        [parameter(Mandatory)]
                        [String]$ServiceName
                    )
            
                    $params = @{
                        Path        = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName" 
                        ErrorAction = 'Stop'
                    }
                    $property = Get-ItemProperty @params
            
                    if (
                    ($property.Start -eq 2) -and 
                    ($property.DelayedAutostart -eq 1)
                    ) {
                        $true
                    }
                    else {
                        $false
                    }
                }
            }
            catch {
                $M = $_
                $Error.RemoveAt(0)
                throw "Failed testing if 'DelayedAutostart' is set: $M"
            }
        }
        Function Set-DelayedAutoStartHC {
            Param (
                [parameter(Mandatory)]
                [String]$ComputerName,
                [parameter(Mandatory)]
                [alias('Name')]
                [String]$ServiceName
            )
    
            try {
                $params = @{
                    ComputerName = $ComputerName
                    ArgumentList = $ServiceName 
                    ErrorAction  = 'Stop'
                }
                Invoke-Command @params -ScriptBlock {
                    Param (
                        [parameter(Mandatory)]
                        [String]$ServiceName
                    )

                    $params = @{
                        Path        = "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName" 
                        ErrorAction = 'Stop'
                    }
                    Set-ItemProperty @params -Name 'Start' -Value 2
                    Set-ItemProperty @params -Name 'DelayedAutostart' -Value 1
                }
            }
            catch {
                $M = $_
                $Error.RemoveAt(0)
                throw "Failed to enable 'DelayedAutostart': $M"
            }
        }

        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start
        $Error.Clear()

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
                # Unique       = $true
            }
            $logFile = New-LogFileNameHC @logParams
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

        $accepted = @{
            serviceStartupTypes = @(
                'Automatic', 'DelayedAutoStart', 'Disabled', 'Manual'
            )
            executionTypes      = @(
                'StopService', 'KillProcess', 'StartService'
            )
        }
    
        #region Test and sanitize .json file properties
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

            #region SetServiceStartupType properties
            $properties = $task.SetServiceStartupType.PSObject.Properties.Name

            $serviceNamesInStartupTypes = @()
        
            foreach ($startupTypeName in $accepted.serviceStartupTypes) {
                if ($properties -notContains $startupTypeName) {
                    throw "Property 'SetServiceStartupType.$startupTypeName' not found in one of the 'Tasks'."
                }

                #region Remove empty values from arrays
                $task.SetServiceStartupType.$startupTypeName = 
                $task.SetServiceStartupType.$startupTypeName | 
                Where-Object { $_ }
                #endregion

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

            #region Execute properties
            $properties = $task.Execute.PSObject.Properties.Name
            
            foreach ($executionType in $accepted.executionTypes) {
                if ($properties -notContains $executionType) {
                    throw "Property 'Execute.$executionType' not found in one of the 'Tasks'."
                }

                #region Remove empty values from arrays
                $task.Execute.$executionType = $task.Execute.$executionType | 
                Where-Object { $_ }
                #endregion

                if ($task.Execute.$executionType) {
                    $actionInTask = $true
                }   
            }
            #endregion

            #region Test incorrect input
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

            foreach ($disabledService in $task.SetServiceStartupType.Disabled) {
                if ($task.Execute.StartService -contains $disabledService) {
                    throw "Service '$disabledService' cannot have StartupType 'Disabled' and 'StartService' at the same time"    
                }
            }
            #endregion
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
    $export = @{
        service = @()
        process = @()
    }

    Try {
        $i = 0
        Foreach ($task in $Tasks) {
            $i++
            foreach ($computerName in $task.ComputerName) {
                #region Test computer online
                if (-not (Test-Connection -ComputerName $computerName -Quiet)) {
                    $M = "Computer '$computerName' is offline"
                    Write-Verbose $M
                    Write-EventLog @EventErrorParams -Message $M
                    Continue
                }
                #endregion

                #region Set service startup type
                foreach ($startupTypeName in $accepted.serviceStartupTypes) {
                    foreach (
                        $serviceName in 
                        $task.SetServiceStartupType.$startupTypeName
                    ) {
                        try {
                            $result = [PSCustomObject]@{
                                Task         = $i
                                Part         = 'SetServiceStartupType'
                                Date         = Get-Date
                                ComputerName = $computerName
                                ServiceName  = $serviceName
                                DisplayName  = $null
                                StartupType  = $null
                                Status       = $null
                                Action       = $null
                                Error        = $null
                            }

                            $params = @{
                                ComputerName = $computerName
                                Name         = $serviceName 
                                ErrorAction  = 'Stop'
                            }
                            $service = Get-Service @params

                            #region Get service state before
                            $result.DisplayName = $service.DisplayName
                            $result.Status = $service.Status

                            $result.StartupType = if (
                                ($service.StartType -eq 'Automatic') -and
                                (Test-DelayedAutoStartHC @params)
                            ) {
                                'DelayedAutoStart'
                            }
                            else {
                                $service.StartType
                            }
                            #endregion

                            if ($startupTypeName -ne $result.StartupType) {
                                if ($startupTypeName -eq 'DelayedAutoStart') {
                                    Set-DelayedAutoStartHC @params
                                }
                                else {
                                    $setParams = @{
                                        StartupType = $startupTypeName 
                                        ErrorAction = 'Stop'
                                    }
                                    $service | Set-Service @setParams    
                                }

                                $result.StartupType = $startupTypeName

                                $result.Action = "updated StartupType from '$($result.StartupType)' to '$startupTypeName'"

                                $M = "'$computerName' service '$serviceName' action 'SetServiceStartupType': {0}" -f $result.Action
                                Write-Verbose $M
                                Write-EventLog @EventOutParams -Message $M
                            }
                        }
                        catch {
                            $result.Error = $_

                            $M = "'$computerName' service '$serviceName' action 'SetServiceStartupType': $_"
                            Write-Warning $M
                            Write-EventLog @EventErrorParams -Message $M

                            $Error.RemoveAt(0)
                        }
                        finally {
                            $export.service += $result
                        }
                    }
                }
                #endregion

                #region Stop service
                foreach ($serviceName in $task.Execute.StopService) {
                    try {
                        $result = [PSCustomObject]@{
                            Task         = $i
                            Part         = 'StopService'
                            Date         = Get-Date
                            ComputerName = $computerName
                            ServiceName  = $serviceName
                            DisplayName  = $null
                            StartupType  = $null
                            Status       = $null
                            Action       = $null
                            Error        = $null
                        }
    
                        $params = @{
                            ComputerName = $computerName
                            Name         = $serviceName 
                            ErrorAction  = 'Stop'
                        }
                        $service = Get-Service @params
    
                        #region Get service state before
                        $result.DisplayName = $service.DisplayName
                        $result.Status = $service.Status
    
                        $result.StartupType = if (
                            ($service.StartType -eq 'Automatic') -and
                            (Test-DelayedAutoStartHC @params)
                        ) {
                            'DelayedAutoStart'
                        }
                        else {
                            $service.StartType
                        }
                        #endregion
                        
                        if ($service.Status -ne 'Stopped') {
                            $service | Stop-Service -ErrorAction 'Stop'

                            $result.Status = 'Stopped'

                            $result.Action = "stopped service that was in state '$($service.Status)'"

                            $M = "'$computerName' service '$serviceName' action 'StopService': {0}" -f $result.Action
                            Write-Verbose $M
                            Write-EventLog @EventOutParams -Message $M
                        }
                    }
                    catch {
                        $result.Error = $_

                        $M = "'$computerName' service '$serviceName' action 'StopService': $_"
                        Write-Warning $M
                        Write-EventLog @EventErrorParams -Message $M

                        $Error.RemoveAt(0)
                    }
                    finally {
                        $export.service += $result
                    }
                }
                #endregion

                #region Kill process
                foreach ($processName in $task.Execute.KillProcess) {
                    $params = @{
                        ComputerName = $computerName
                        Name         = $processName 
                        ErrorAction  = 'Ignore'
                    }
                    $processes = Get-Process @params

                    if (-not $processes) {
                        $M = "'$computerName' process '$processName' action 'KillService': process not running"
                        Write-Verbose $M
                        Write-EventLog @EventVerboseParams -Message $M
                        Continue
                    }

                    foreach ($process in $processes) {
                        try {
                            $result = [PSCustomObject]@{
                                Task           = $i
                                Part           = 'KillProcess'
                                Date           = Get-Date
                                ComputerName   = $computerName
                                ProcessName    = $processName
                                Description    = $process.Description
                                Company        = $process.Company
                                Product        = $process.Product
                                ProductVersion = $process.ProductVersion
                                Id             = $process.Id
                                Action         = $null
                                Error          = $null
                            }

                            $process | Stop-Process -EA 'Stop'
                        }
                        catch {
                            $result.Error = $_
    
                            $M = "'$computerName' process '$processName' action 'KillProcess': $_"
                            Write-Warning $M
                            Write-EventLog @EventErrorParams -Message $M
    
                            $Error.RemoveAt(0)
                        }
                        finally {
                            $export.process += $result
                        }
                    }
                }
                #endregion 

                #region Start service
                foreach ($serviceName in $task.Execute.StartService) {
                    try {
                        $result = [PSCustomObject]@{
                            Task         = $i
                            Part         = 'StartService'
                            Date         = Get-Date
                            ComputerName = $computerName
                            ServiceName  = $serviceName
                            DisplayName  = $null
                            StartupType  = $null
                            Status       = $null
                            Action       = $null
                            Error        = $null
                        }
    
                        $params = @{
                            ComputerName = $computerName
                            Name         = $serviceName 
                            ErrorAction  = 'Stop'
                        }
                        $service = Get-Service @params
    
                        #region Get service state before
                        $result.DisplayName = $service.DisplayName
                        $result.Status = $service.Status
    
                        $result.StartupType = if (
                            ($service.StartType -eq 'Automatic') -and
                            (Test-DelayedAutoStartHC @params)
                        ) {
                            'DelayedAutoStart'
                        }
                        else {
                            $service.StartType
                        }
                        #endregion
                        
                        if ($service.Status -ne 'Running') {
                            $service | Start-Service -ErrorAction 'Stop'

                            $result.Status = 'Running'

                            $result.Action = "started service that was in state '$($service.Status)'"

                            $M = "'$computerName' service '$serviceName' action 'StartService': {0}" -f $result.Action
                            Write-Verbose $M
                            Write-EventLog @EventOutParams -Message $M
                        }
                    }
                    catch {
                        $result.Error = $_

                        $M = "'$computerName' service '$serviceName' action 'StartService': $_"
                        Write-Warning $M
                        Write-EventLog @EventErrorParams -Message $M

                        $Error.RemoveAt(0)
                    }
                    finally {
                        $export.service += $result
                    }
                }
                #endregion 
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
        $excelParams = @{
            Path         = "$logFile - Report.xlsx"
            AutoSize     = $true
            FreezeTopRow = $true
            Verbose      = $false
        }

        $mailParams = @{
            To        = $mailTo
            Bcc       = $ScriptAdmin
            Priority  = 'Normal'
            LogFolder = $logParams.LogFolder
            Header    = $sendMailHeader
            Save      = "$logFile - Mail.html"
        }

        #region Export results to Excel file
        if ($export.service) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Services'

            $export.service | Sort-Object -Property 'Date' | 
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }

        if ($export.process) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Processes'

            $export.process | Sort-Object -Property 'Date' | 
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        $count = @{
            service     = @{
                total  = $export.service.Count
                action = $export.service | Where-Object { $_.Action } | Measure-Object | Select-Object -ExpandProperty 'Count'
                error  = $export.service | Where-Object { $_.Error } | Measure-Object | Select-Object -ExpandProperty 'Count'
            }
            process     = @{
                total  = $export.process.Count
                action = $export.process | Where-Object { $_.Action } | Measure-Object | Select-Object -ExpandProperty 'Count'
                error  = $export.process | Where-Object { $_.Error } | Measure-Object | Select-Object -ExpandProperty 'Count'
            }
            systemError = ($Error.Exception.Message | Measure-Object).Count
        }

        #region Subject and Priority
        $mailParams.Subject = '{0} service{1}, {2} process{3}' -f
        $count.service.total,
        $(if ($count.service.total -ne 1) { 's' }),
        $count.process.total,
        $(if ($count.process.total -ne 1) { 'es' })

        if (
            $totalErrorCount = $count.systemError + $count.service.error + 
            $count.process.error
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -gt 1) { 's' }
            )
        }
        #endregion

        $systemErrorHtmlList = if ($count.systemError) {
            "<p>Detected <b>{0} error{1}:{2}</p>" -f $count.systemError, 
            $(
                if ($count.systemError -gt 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } | 
                ConvertTo-HtmlListHC
            )
        }
     
        #region Create HTML table
        $htmlTable = "
        <table>
            <tr>
                <th colspan=`"2`">Services</th>
            </tr>
            <tr>
                <td>Rows</td>
                <td>$($count.service.total)</td>
            </tr>
            <tr>
                <td>Actions</td>
                <td>$($count.service.action)</td>
            </tr>
            <tr>
                <td>Errors</td>
                <td>$($count.service.error)</td>
            </tr>
            <tr>
                <th colspan=`"2`">Processes</th>
            </tr>
            <tr>
                <td>Total</td>
                <td>$($count.process.total)</td>
            </tr>
            <tr>
                <td>Actions</td>
                <td>$($count.process.action)</td>
            </tr>
            <tr>
                <td>Errors</td>
                <td>$($count.process.error)</td>
            </tr>
        </table>" 
        #endregion
     
        #region Send mail
        $mailParams.Message = "
            <p>Manage services and processes: configure the service startup type, stop a service, stop a process, start a service.</p>
            $systemErrorHtmlList
            $htmlTable
            {0}" -f 
        $(
            if ($mailParams.Attachments) {
                '<p><i>* Check the attachment for details</i></p>'
            }
        )
     
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
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
