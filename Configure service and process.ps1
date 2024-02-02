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
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    $scriptBlock = {
        Param (
            [String[]]$SetServiceAutomatic,
            [String[]]$SetServiceDelayedAutoStart,
            [String[]]$SetServiceDisabled,
            [String[]]$SetServiceManual,
            [String[]]$ExecuteStopService,
            [String[]]$ExecuteKillProcess,
            [String[]]$ExecuteStartService
        )

        Function Test-DelayedAutoStartHC {
            Param (
                [parameter(Mandatory)]
                [alias('Name')]
                [String]$ServiceName
            )

            try {
                $params = @{
                    Path        = 'HKLM:\SYSTEM\CurrentControlSet\Services\{0}' -f $ServiceName
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
            catch {
                $M = $_; $Error.RemoveAt(0)
                throw "Failed testing if 'DelayedAutostart' is set: $M"
            }
        }
        Function Set-DelayedAutoStartHC {
            Param (
                [parameter(Mandatory)]
                [alias('Name')]
                [String]$ServiceName
            )

            try {
                Param (
                    [parameter(Mandatory)]
                    [String]$ServiceName
                )

                $params = @{
                    Path        = 'HKLM:\SYSTEM\CurrentControlSet\Services\{0}' -f $ServiceName
                    ErrorAction = 'Stop'
                }
                Set-ItemProperty @params -Name 'Start' -Value 2
                Set-ItemProperty @params -Name 'DelayedAutostart' -Value 1
            }
            catch {
                $M = $_; $Error.RemoveAt(0)
                throw "Failed to enable 'DelayedAutostart': $M"
            }
        }
        Function Set-ServiceStartupTypeHC {
            Param (
                [parameter(Mandatory)]
                [alias('Name')]
                [String]$ServiceName,
                [ValidateSet(
                    'Automatic', 'DelayedAutoStart',
                    'Disabled', 'Manual'
                )]
                [Parameter(Mandatory)]
                [String]$StartupTypeName
            )
            try {
                $result = [PSCustomObject]@{
                    Part        = 'SetServiceStartupType'
                    Date        = Get-Date
                    ServiceName = $ServiceName
                    DisplayName = $null
                    StartupType = $null
                    Status      = $null
                    Action      = $null
                    Error       = $null
                }

                #region Get service state
                $params = @{
                    Name        = $ServiceName
                    ErrorAction = 'Stop'
                }
                $service = Get-Service @params

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

                #region Set service startup type
                if ($StartupTypeName -ne $result.StartupType) {
                    if ($StartupTypeName -eq 'DelayedAutoStart') {
                        Set-DelayedAutoStartHC @params
                    }
                    else {
                        $setParams = @{
                            StartupType = $StartupTypeName
                            ErrorAction = 'Stop'
                        }
                        $service | Set-Service @setParams
                    }

                    $result.Action = "updated StartupType from '$($result.StartupType)' to '$StartupTypeName'"

                    $result.StartupType = $StartupTypeName
                }
                #endregion
            }
            catch {
                $result.Error = $_
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
        Function Stop-ProcessHC {
            Param (
                [parameter(Mandatory)]
                [alias('Name')]
                [String]$ProcessName
            )

            try {
                Get-Process -Name $ProcessName |
                Stop-Process -EA Stop -Force
            }
            catch {
                $M = $_; $Error.RemoveAt(0)
                throw "Failed to stop process '$ProcessName': $M"
            }
        }

        #region Set service startup type
        foreach ($serviceName in $SetServiceAutomatic) {
            $params = @{
                ServiceName     = $serviceName
                StartupTypeName = 'Automatic'
            }
            Set-ServiceStartupTypeHC @params
        }
        foreach ($serviceName in $SetServiceDelayedAutoStart) {
            $params = @{
                ServiceName     = $serviceName
                StartupTypeName = 'DelayedAutoStart'
            }
            Set-ServiceStartupTypeHC @params
        }
        foreach ($serviceName in $SetServiceDisabled) {
            $params = @{
                ServiceName     = $serviceName
                StartupTypeName = 'Disabled'
            }
            Set-ServiceStartupTypeHC @params
        }
        foreach ($serviceName in $SetServiceManual) {
            $params = @{
                ServiceName     = $serviceName
                StartupTypeName = 'Manual'
            }
            Set-ServiceStartupTypeHC @params
        }
        #endregion

        #region Stop service
        foreach ($serviceName in $ExecuteStopService) {
            try {
                $result = [PSCustomObject]@{
                    Part        = 'StopService'
                    Date        = Get-Date
                    ServiceName = $serviceName
                    DisplayName = $null
                    StartupType = $null
                    Status      = $null
                    Action      = $null
                    Error       = $null
                }

                $params = @{
                    Name        = $serviceName
                    ErrorAction = 'Stop'
                }
                $service = Get-Service @params

                #region Get service state
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

                #region Stop service
                if ($service.Status -ne 'Stopped') {
                    $service | Stop-Service -ErrorAction 'Stop' -Force

                    $result.Action = 'stopped service'
                    $result.Status = 'Stopped'
                }
                #endregion
            }
            catch {
                $result.Error = $_
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
        #endregion

        #region Kill process
        foreach ($processName in $ExecuteKillProcess) {
            $params = @{
                Name        = $processName
                ErrorAction = 'Ignore'
            }
            $processes = Get-Process @params

            foreach ($process in $processes) {
                try {
                    $result = [PSCustomObject]@{
                        Part        = 'KillProcess'
                        Date        = Get-Date
                        ProcessName = $processName
                        Description = $process.Description
                        Id          = $process.Id
                        Action      = $null
                        Error       = $null
                    }

                    Stop-ProcessHC -ProcessName $processName -EA Stop

                    $result.Action = 'stopped running process'
                }
                catch {
                    $result.Error = $_
                    $Error.RemoveAt(0)
                }
                finally {
                    $result
                }
            }
        }
        #endregion

        #region Start service
        foreach ($serviceName in $ExecuteStartService) {
            try {
                $result = [PSCustomObject]@{
                    Part        = 'StartService'
                    Date        = Get-Date
                    ServiceName = $serviceName
                    DisplayName = $null
                    StartupType = $null
                    Status      = $null
                    Action      = $null
                    Error       = $null
                }

                $params = @{
                    Name        = $serviceName
                    ErrorAction = 'Stop'
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

                #region Start service
                if ($service.Status -ne 'Running') {
                    $service | Start-Service -ErrorAction 'Stop'

                    $result.Action = 'started service'
                    $result.Status = 'Running'
                }
                #endregion
            }
            catch {
                $result.Error = $_
                $Error.RemoveAt(0)
            }
            finally {
                $result
            }
        }
        #endregion
    }

    Try {
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

        #region Test .json file
        try {
            if (-not ($mailTo = $file.SendMail.To)) {
                throw "Property 'SendMail.To' not found."
            }

            if (-not ($MaxConcurrentJobs = $file.MaxConcurrentJobs)) {
                throw "Property 'MaxConcurrentJobs' not found"
            }
            try {
                $null = $MaxConcurrentJobs.ToInt16($null)
            }
            catch {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$MaxConcurrentJobs' is not supported."
            }

            if (-not ($Tasks = $file.Tasks)) {
                throw "Property 'Tasks' not found."
            }

            foreach ($task in $Tasks) {
                #region Test Tasks properties
                $properties = $task.PSObject.Properties.Name

                @(
                    'SetServiceStartupType',
                    'ComputerName',
                    'Execute'
                ) | Where-Object { $properties -notContains $_ } |
                ForEach-Object {
                    throw "Property 'Tasks.$_' not found"
                }
                #endregion

                #region Test ComputerName
                if (-not $task.ComputerName) {
                    throw "Property 'Tasks.ComputerName' not found"
                }

                if (
                    $duplicateComputers = $task.ComputerName | Group-Object |
                    Where-Object { $_.Count -gt 1 }
                ) {
                    throw "duplicate ComputerName '$($duplicateComputers.Name)' found in a single task"
                }
                #endregion

                #region Test SetServiceStartupType
                $properties = $task.SetServiceStartupType.PSObject.Properties.Name

                foreach ($startupTypeName in $accepted.serviceStartupTypes) {
                    if ($properties -notContains $startupTypeName) {
                        throw "Property 'SetServiceStartupType.$startupTypeName' not found"
                    }
                }
                #endregion

                #region Test Execute
                $properties = $task.Execute.PSObject.Properties.Name

                foreach ($executionType in $accepted.executionTypes) {
                    if ($properties -notContains $executionType) {
                        throw "Property 'Execute.$executionType' not found"
                    }
                }
                #endregion

                #region Test StartupType Disabled and StartService
                foreach (
                    $disabledService in
                    $task.SetServiceStartupType.Disabled
                ) {
                    if ($task.Execute.StartService -contains $disabledService) {
                        throw "Service '$disabledService' cannot have StartupType 'Disabled' and 'StartService' at the same time"
                    }
                }
                #endregion

                #region Test Tasks have an action
                $actionInTask = $false

                foreach ($startupTypeName in $accepted.serviceStartupTypes) {
                    $task.SetServiceStartupType.$startupTypeName =
                    $task.SetServiceStartupType.$startupTypeName |
                    Where-Object { $_ }

                    if ($task.SetServiceStartupType.$startupTypeName) {
                        $actionInTask = $true
                    }
                }

                foreach ($executionType in $accepted.executionTypes) {
                    $task.Execute.$executionType = $task.Execute.$executionType |
                    Where-Object { $_ }

                    if ($task.Execute.$executionType) {
                        $actionInTask = $true
                    }
                }

                if (-not $actionInTask) {
                    throw "Input file '$ImportFile': Contains a task where properties 'SetServiceStartupType' and 'Execute' are both empty."
                }
                #endregion

                #region Test duplicate names in SetServiceStartupType
                $serviceNamesInStartupTypes = @()

                foreach ($startupTypeName in $accepted.serviceStartupTypes) {
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
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        #region Add properties
        $taskNumber = 0

        foreach ($task in $Tasks) {
            $taskNumber++
            $task | Add-Member -NotePropertyMembers @{
                Jobs       = @()
                TaskNumber = $taskNumber
            }

            foreach ($computerName in $task.ComputerName) {
                $task.Jobs += @{
                    ComputerName = $computerName
                    Session      = $null
                    Object       = $null
                    Results      = @()
                    Errors       = @()
                }
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
    $export = @{
        service = @()
        process = @()
    }

    Try {
        Foreach ($task in $Tasks) {
            $invokeParams = @{
                ScriptBlock  = $scriptBlock
                ArgumentList = $task.SetServiceStartupType.Automatic,
                $task.SetServiceStartupType.DelayedAutoStart,
                $task.SetServiceStartupType.Disabled,
                $task.SetServiceStartupType.Manual,
                $task.Execute.StopService,
                $task.Execute.KillProcess,
                $task.Execute.StartService
            }

            foreach ($job in $task.Jobs) {
                $M = "Start job on '{0}' with SetServiceStartupType.Automatic '{1}' SetServiceStartupType.DelayedAutoStart '{2}' SetServiceStartupType.Disabled '{3}' SetServiceStartupType.Manual '{4}' Execute.StopService '{5}' Execute.KillProcess '{6}' Execute.StartService '{7}'" -f
                $job.ComputerName,
                ($invokeParams.ArgumentList[0] -join ','),
                ($invokeParams.ArgumentList[1] -join ','),
                ($invokeParams.ArgumentList[2] -join ','),
                ($invokeParams.ArgumentList[3] -join ','),
                ($invokeParams.ArgumentList[4] -join ','),
                ($invokeParams.ArgumentList[5] -join ','),
                ($invokeParams.ArgumentList[6] -join ',')
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

                #region Start job
                $computerName = $job.ComputerName

                try {
                    $job.Session = New-PSSessionHC -ComputerName $computerName
                    $invokeParams += @{
                        Session = $job.Session
                        AsJob   = $true
                    }
                    $job.Object = Invoke-Command @invokeParams
                }
                catch {
                    $M = "Failed creating a session to '$computerName': $_"
                    $job.Errors += $M
                    $Error.RemoveAt(0)
                    Write-Warning $M
                    Write-EventLog @EventWarnParams -Message $M
                    Continue
                }
                #endregion

                #region Wait for max running jobs
                $waitJobParams = @{
                    Job        = $Tasks.Jobs.Object | Where-Object { $_ }
                    MaxThreads = $MaxConcurrentJobs
                }

                if ($waitJobParams.Job) {
                    Wait-MaxRunningJobsHC @waitJobParams
                }
                #endregion
            }
        }

        #region Wait for all jobs to finish
        $waitJobParams = @{
            Job = $Tasks.Jobs.Object | Where-Object { $_ }
        }
        if ($waitJobParams.Job) {
            Write-Verbose 'Wait for all jobs to finish'

            $null = Wait-Job @waitJobParams
        }
        #endregion

        #region Get job results and job errors
        foreach ($task in $Tasks) {
            foreach (
                $job in
                $task.Jobs | Where-Object { $_.Object }
            ) {
                $jobErrors = @()
                $receiveParams = @{
                    ErrorVariable = 'jobErrors'
                    ErrorAction   = 'SilentlyContinue'
                }
                $job.Results += $job.Object | Receive-Job @receiveParams

                foreach ($e in $jobErrors) {
                    $job.Errors += $e.ToString()
                    $Error.Remove($e)

                    $M = "Failed job on '{0}' with SetServiceStartupType.Automatic '{1}' SetServiceStartupType.DelayedAutoStart '{2}' SetServiceStartupType.Disabled '{3}' SetServiceStartupType.Manual '{4}' Execute.StopService '{5}' Execute.KillProcess '{6}' Execute.StartService '{7}': {8}" -f
                    $job.ComputerName,
                    ($invokeParams.ArgumentList[0] -join ','),
                    ($invokeParams.ArgumentList[1] -join ','),
                    ($invokeParams.ArgumentList[2] -join ','),
                    ($invokeParams.ArgumentList[3] -join ','),
                    ($invokeParams.ArgumentList[4] -join ','),
                    ($invokeParams.ArgumentList[5] -join ','),
                    ($invokeParams.ArgumentList[6] -join ','), $e.ToString()
                    Write-Verbose $M
                    Write-EventLog @EventErrorParams -Message $M
                }

                $job.Session | Remove-PSSession -ErrorAction Ignore
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
            Header    = $ScriptName
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
        $htmlTable = '<table>{0}{1}</table>' -f
        $(
            if ($count.service.total) {
                "<tr>
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
                </tr>"
            }
        ),
        $(
            if ($count.process.total) {
                "<tr>
                        <th colspan=`"2`">Processes</th>
                    </tr>
                    <tr>
                        <td>Rows</td>
                        <td>$($count.process.total)</td>
                    </tr>
                    <tr>
                        <td>Actions</td>
                        <td>$($count.process.action)</td>
                    </tr>
                    <tr>
                        <td>Errors</td>
                        <td>$($count.process.error)</td>
                    </tr>"
            }
        )
        #endregion

        #region Send mail
        $mailParams.Message = "
            <p>Manage services and processes: configure the service startup type, stop a service, stop a process, start a service.</p>
            $systemErrorHtmlList
            $htmlTable
            {0}" -f $(
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
