#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ScriptAdmin = 'admin@contoso.com'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Function Test-DelayedAutoStartHC {
        Param (
            [parameter(Mandatory)]
            [String]$ComputerName,
            [parameter(Mandatory)]
            [alias('Name')]
            [String]$ServiceName
        )
    }
    Function Set-DelayedAutoStartHC {
        Param (
            [parameter(Mandatory)]
            [String]$ComputerName,
            [parameter(Mandatory)]
            [alias('Name')]
            [String]$ServiceName
        )
    }
    Function Stop-ProcessHC {
        Param (
            [parameter(Mandatory)]
            [String]$ComputerName,
            [parameter(Mandatory)]
            [alias('Name')]
            [String]$ProcessName
        )
    }
    Mock Set-DelayedAutoStartHC
    Mock Test-DelayedAutoStartHC { $true }

    Mock Get-Service
    Mock Get-Process
    Mock Invoke-Command
    Mock Send-MailHC
    Mock Set-Service
    Mock Start-Process
    Mock Start-Service
    Mock Stop-ProcessHC
    Mock Stop-Service
    Mock Test-Connection { $true }
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach 'ScriptName', 'ImportFile' {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $mailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx::\notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like "*Failed creating the log folder 'xxx::\notExistingLocation'*")
        }
    }
    Context 'the ImportFile' {
        BeforeEach {
            $testJsonFile = @{
                MaxConcurrentJobs = 5
                Tasks             = @(
                    @{
                        ComputerName          = @('PC1')
                        SetServiceStartupType = @{
                            Automatic        = @()
                            DelayedAutostart = @()
                            Disabled         = @()
                            Manual           = @()
                        }
                        Execute               = @{
                            StopService  = @()
                            KillProcess  = @()
                            StartService = @()
                        }
                    }
                )
                SendMail          = @{
                    To = 'bob@contoso.com'
                }
            }
        }
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'is missing property' {
            It 'Tasks.<_>' -ForEach @(
                'ComputerName', 'SetServiceStartupType', 'Execute'
            ) {
                $testJsonFile.Tasks[0].Remove($_)
                $testJsonFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Property 'Tasks.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.SetServiceStartupType.<_>' -ForEach @(
                'Automatic', 'DelayedAutostart', 'Disabled', 'Manual'
            ) {
                $testJsonFile.Tasks[0].SetServiceStartupType.Remove($_)
                $testJsonFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Property 'SetServiceStartupType.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.Execute.<_>' -ForEach @(
                'StopService', 'KillProcess', 'StartService'
            ) {
                $testJsonFile.Tasks[0].Execute.Remove($_)
                $testJsonFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Property 'Execute.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.To' {
                $testJsonFile.SendMail.Remove('To')
                $testJsonFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Property 'SendMail.To' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
        Context 'is missing content for property' {
            It 'MaxConcurrentJobs' {
                $testJsonFile.MaxConcurrentJobs = 'a'
                $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*Property 'MaxConcurrentJobs' needs to be a number, the value 'a' is not supported*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'MaxConcurrentJobs' {
                $testJsonFile.MaxConcurrentJobs = $null
                $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*Property 'MaxConcurrentJobs' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks' {
                $testJsonFile.Tasks = @()
                $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*Property 'Tasks' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.ComputerName' {
                $testJsonFile.Tasks[0].ComputerName = @()
                $testJsonFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Property 'Tasks.ComputerName' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.SetServiceStartupType and Tasks.Execute' {
                $testJsonFile.Tasks[0].ComputerName = @('PC1')
                $testJsonFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*Contains a task where properties 'SetServiceStartupType' and 'Execute' are both empty.*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
        Context 'contains incorrect content' {
            BeforeEach {
                $testJsonFile.Tasks[0].Execute.KillProcess = @('chrome')
            }
            It 'duplicate ComputerName' {
                $testJsonFile.Tasks[0].ComputerName = @('PC1', 'PC1', 'PC2')
                $testJsonFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and
                ($Message -like "*duplicate ComputerName 'PC1' found in a single task*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'ServiceStartupType' {
                It 'is set to Disabled but StartService is used' {
                    $testJsonFile.Tasks[0].SetServiceStartupType.Disabled = @('x')
                    $testJsonFile.Tasks[0].Execute.StartService = @('x')
                    $testJsonFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*Service 'x' cannot have StartupType 'Disabled' and 'StartService' at the same time*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'contains the same service name for different startup types' {
                    $testJsonFile.Tasks[0].SetServiceStartupType.Manual = @('x')
                    $testJsonFile.Tasks[0].SetServiceStartupType.Disabled = @('x')
                    $testJsonFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*Service 'x' can only have one StartupType*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
} -Tag test
Describe 'a service startup type in SetServiceStartupType is' {
    Context 'corrected when it is incorrect' {
        BeforeEach {
            $testJsonFile = @{
                MaxConcurrentJobs = 5
                Tasks             = @(
                    @{
                        ComputerName          = @('PC1')
                        SetServiceStartupType = @{
                            Automatic        = @()
                            DelayedAutostart = @()
                            Disabled         = @()
                            Manual           = @()
                        }
                        Execute               = @{
                            StopService  = @()
                            KillProcess  = @()
                            StartService = @()
                        }
                    }
                )
                SendMail          = @{
                    To = 'bob@contoso.com'
                }
            }
        }
        It "actual '<actual>' expected '<expected>'" -ForEach @(
            @{ actual = 'Automatic'; expected = 'DelayedAutoStart' }
            # @{ actual = 'Automatic'; expected = 'Disabled' }
            # @{ actual = 'Automatic'; expected = 'Manual' }
            # @{ actual = 'DelayedAutoStart'; expected = 'Automatic' }
            # @{ actual = 'DelayedAutoStart'; expected = 'Disabled' }
            # @{ actual = 'DelayedAutoStart'; expected = 'Manual' }
            # @{ actual = 'Disabled'; expected = 'Automatic' }
            # @{ actual = 'Disabled'; expected = 'DelayedAutoStart' }
            # @{ actual = 'Disabled'; expected = 'Manual' }
            # @{ actual = 'Manual'; expected = 'Automatic' }
            # @{ actual = 'Manual'; expected = 'DelayedAutoStart' }
            # @{ actual = 'Manual'; expected = 'Disabled' }
        ) {
            Mock Get-Service {
                @{
                    Status      = 'Running'
                    StartType   = $actual
                    Name        = 'testService'
                    DisplayName = 'the display name'
                }
            }
            Mock Test-DelayedAutoStartHC { $true }

            if ($expected -eq 'DelayedAutoStart') {
                Mock Test-DelayedAutoStartHC { $false }
            }

            if ($actual -eq 'DelayedAutoStart') {
                Mock Get-Service {
                    @{
                        Status      = 'Running'
                        StartType   = 'Automatic'
                        Name        = 'testService'
                        DisplayName = 'the display name'
                    }
                }
            }

            $testJsonFile.Tasks[0].SetServiceStartupType.$expected = @(
                'testService'
            )
            $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

            .$testScript @testParams

            if (($actual -eq 'DelayedAutoStart') -or ($actual -eq 'Automatic')) {
                Should -Invoke Test-DelayedAutoStartHC -Times 1 -Exactly -ParameterFilter {
                    ($ServiceName -eq 'testService') -and
                    ($ComputerName -eq 'PC1')
                }
            }
            else {
                Should -Not -Invoke Test-DelayedAutoStartHC
            }

            if ($expected -eq 'DelayedAutoStart') {
                Should -Invoke Set-DelayedAutoStartHC -Times 1 -Exactly -ParameterFilter {
                    ($ServiceName -eq 'testService') -and
                    ($ComputerName -eq 'PC1')
                }
                Should -Not -Invoke Set-Service
            }
            else {
                Should -Not -Invoke Set-DelayedAutoStartHC
                Should -Invoke Set-Service -Times 1 -Exactly -ParameterFilter {
                    $StartupType -eq $expected
                }
            }
        }
    }
    Context 'ignored when it is correct' {
        BeforeEach {
            $testJsonFile = @{
                MaxConcurrentJobs = 5
                Tasks             = @(
                    @{
                        ComputerName          = @('PC1')
                        SetServiceStartupType = @{
                            Automatic        = @()
                            DelayedAutostart = @()
                            Disabled         = @()
                            Manual           = @()
                        }
                        Execute               = @{
                            StopService  = @()
                            KillProcess  = @()
                            StartService = @()
                        }
                    }
                )
                SendMail          = @{
                    To = 'bob@contoso.com'
                }
            }
        }
        It "actual '<actual>' expected '<expected>'"-ForEach @(
            @{ actual = 'Automatic'; expected = 'Automatic' }
            @{ actual = 'DelayedAutoStart'; expected = 'DelayedAutoStart' }
            @{ actual = 'Disabled'; expected = 'Disabled' }
            @{ actual = 'Manual'; expected = 'Manual' }
        ) {
            Mock Get-Service {
                @{
                    Status      = 'Running'
                    StartType   = $actual
                    Name        = 'testService'
                    DisplayName = 'the display name'
                }
            }
            Mock Test-DelayedAutoStartHC { $false }

            if ($actual -eq 'DelayedAutoStart') {
                Mock Test-DelayedAutoStartHC { $true }
                Mock Get-Service {
                    @{
                        Status      = 'Running'
                        StartType   = 'Automatic'
                        Name        = 'testService'
                        DisplayName = 'the display name'
                    }
                }
            }

            $testJsonFile.Tasks[0].SetServiceStartupType.$expected = @(
                'testService'
            )
            $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

            .$testScript @testParams

            if (($actual -eq 'DelayedAutoStart') -or ($actual -eq 'Automatic')) {
                Should -Invoke Test-DelayedAutoStartHC -Times 1 -Exactly -ParameterFilter {
                    ($ServiceName -eq 'testService') -and
                    ($ComputerName -eq 'PC1')
                }
            }
            else {
                Should -Not -Invoke Test-DelayedAutoStartHC
            }

            Should -Not -Invoke Set-Service
            Should -Not -Invoke Set-DelayedAutoStartHC

        }
    }
}
Describe 'a service in StopService is' {
    BeforeAll {
        $testJsonFile = @{
            MaxConcurrentJobs = 5
            Tasks             = @(
                @{
                    ComputerName          = @('PC1')
                    SetServiceStartupType = @{
                        Automatic        = @()
                        DelayedAutostart = @()
                        Disabled         = @()
                        Manual           = @()
                    }
                    Execute               = @{
                        StopService  = @('testService')
                        KillProcess  = @()
                        StartService = @()
                    }
                }
            )
            SendMail          = @{
                To = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams
    }
    It 'stopped when it is running' {
        $testService = New-MockObject -Type 'System.ServiceProcess.ServiceController' -Properties @{
            ServiceName = 'testService'
            MachineName = 'PC1'
            Status      = 'Running'
        }

        Mock Get-Service {
            $testService
        } -ParameterFilter {
            ($ComputerName -eq $computerName) -and
            ($Name -eq $serviceName)
        }

        .$testScript @testParams

        Should -Invoke Stop-Service -Times 1 -Exactly -ParameterFilter {
            $InputObject -eq $testService
        }
    }
    It 'ignored when it is not running' {
        $testService = New-MockObject -Type 'System.ServiceProcess.ServiceController' -Properties @{
            ServiceName = 'testService'
            MachineName = 'PC1'
            Status      = 'Stopped'
        }

        Mock Get-Service {
            $testService
        } -ParameterFilter {
            ($ComputerName -eq $computerName) -and
            ($Name -eq $serviceName)
        }

        .$testScript @testParams

        Should -Not -Invoke Stop-Service
    }
}
Describe 'a process in KillProcess is' {
    BeforeAll {
        $testJsonFile = @{
            MaxConcurrentJobs = 5
            Tasks             = @(
                @{
                    ComputerName          = @('PC1')
                    SetServiceStartupType = @{
                        Automatic        = @()
                        DelayedAutostart = @()
                        Disabled         = @()
                        Manual           = @()
                    }
                    Execute               = @{
                        StopService  = @()
                        KillProcess  = @('testProcess')
                        StartService = @()
                    }
                }
            )
            SendMail          = @{
                To = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams
    }
    It 'stopped when it is running' {
        $testProcess = New-MockObject -Type 'System.Diagnostics.Process' -Properties @{
            ProcessName = 'testProcess'
            Id          = 124
            MachineName = 'PC1'
        }

        Mock Get-Process {
            $testProcess
        } -ParameterFilter {
            ($ComputerName -eq $computerName) -and
            ($Name -eq $processName)
        }

        .$testScript @testParams

        Should -Invoke Stop-ProcessHC -Times 1 -Exactly -ParameterFilter {
            ($ComputerName -eq 'PC1') -and
            ($ProcessName -eq 'testProcess')
        }
    }
    It 'ignored when it is not running' {
        Mock Get-Process {}

        .$testScript @testParams

        Should -Not -Invoke Stop-ProcessHC
    }
}
Describe 'a service in StartService is' {
    BeforeAll {
        $testJsonFile = @{
            MaxConcurrentJobs = 5
            Tasks             = @(
                @{
                    ComputerName          = @('PC1')
                    SetServiceStartupType = @{
                        Automatic        = @()
                        DelayedAutostart = @()
                        Disabled         = @()
                        Manual           = @()
                    }
                    Execute               = @{
                        StopService  = @()
                        KillProcess  = @()
                        StartService = @('testService')
                    }
                }
            )
            SendMail          = @{
                To = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams
    }
    It 'started when it is not running' {
        $testService = New-MockObject -Type 'System.ServiceProcess.ServiceController' -Properties @{
            ServiceName = 'testService'
            MachineName = 'PC1'
            Status      = 'Stopped'
        }

        Mock Get-Service {
            $testService
        } -ParameterFilter {
            ($ComputerName -eq $computerName) -and
            ($Name -eq $serviceName)
        }

        .$testScript @testParams

        Should -Invoke Start-Service -Times 1 -Exactly -ParameterFilter {
            $InputObject -eq $testService
        }
    }
    It 'ignored when it is running' {
        $testService = New-MockObject -Type 'System.ServiceProcess.ServiceController' -Properties @{
            ServiceName = 'testService'
            MachineName = 'PC1'
            Status      = 'Running'
        }

        Mock Get-Service {
            $testService
        } -ParameterFilter {
            ($ComputerName -eq $computerName) -and
            ($Name -eq $serviceName)
        }

        .$testScript @testParams

        Should -Not -Invoke Start-Service
    }
}
Describe 'after the script runs' {
    BeforeAll {
        $testJsonFile = @{
            MaxConcurrentJobs = 5
            Tasks             = @(
                @{
                    ComputerName          = @('PC1')
                    SetServiceStartupType = @{
                        Automatic        = @()
                        DelayedAutostart = @()
                        Disabled         = @('testServiceManual')
                        Manual           = @()
                    }
                    Execute               = @{
                        StopService  = @('testServiceStopped')
                        KillProcess  = @('testProcessKilled')
                        StartService = @('testServiceStarted')
                    }
                }
            )
            SendMail          = @{
                To = 'bob@contoso.com'
            }
        }
        $testJsonFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

        $testData = @{
            Services  = @(
                @{
                    # SetServiceStartupType
                    ServiceName = $testJsonFile.Tasks[0].SetServiceStartupType.Disabled[0]
                    MachineName = $testJsonFile.Tasks[0].ComputerName[0]
                    Status      = 'Stopped'
                    StartType   = 'Manual'
                    DisplayName = 'display name 1'
                }
                @{
                    # StopService
                    ServiceName = $testJsonFile.Tasks[0].Execute.StopService[0]
                    MachineName = $testJsonFile.Tasks[0].ComputerName[0]
                    Status      = 'Running'
                    StartType   = 'Automatic'
                    DisplayName = 'display name 2'
                }
                @{
                    # StartService
                    ServiceName = $testJsonFile.Tasks[0].Execute.StartService[0]
                    MachineName = $testJsonFile.Tasks[0].ComputerName[0]
                    Status      = 'Stopped'
                    StartType   = 'Automatic'
                    DisplayName = 'display name 3'
                }
            )
            Processes = @(
                @{
                    # KillProcess
                    ProcessName = $testJsonFile.Tasks[0].Execute.KillProcess[0]
                    MachineName = $testJsonFile.Tasks[0].ComputerName[0]
                    Id          = 124
                }
            )
        }

        #region SetServiceStartupType
        Mock Get-Service {
            New-MockObject -Type 'System.ServiceProcess.ServiceController' -Properties $testData.Services[0]
        } -ParameterFilter {
            $Name -eq $testData.Services[0].ServiceName
        }
        #endregion

        #region StopService
        Mock Get-Service {
            New-MockObject -Type 'System.ServiceProcess.ServiceController' -Properties $testData.Services[1]
        } -ParameterFilter {
            $Name -eq $testData.Services[1].ServiceName
        }
        #endregion

        #region StartService
        Mock Get-Service {
            New-MockObject -Type 'System.ServiceProcess.ServiceController' -Properties $testData.Services[2]
        } -ParameterFilter {
            $Name -eq $testData.Services[2].ServiceName
        }
        #endregion

        #region KillProcess
        Mock Get-Process {
            New-MockObject -Type 'System.Diagnostics.Process' -Properties $testData.Processes[0]
        } -ParameterFilter {
            $Name -eq $testData.Processes[0].ProcessName
        }
        #endregion

        Mock Test-DelayedAutoStartHC { $true }

        .$testScript @testParams

        $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Report.xlsx'
    }
    Context 'an Excel file is created' {
        Context "with worksheet 'Services'" {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        # SetServiceStartupType
                        Task         = 1
                        Part         = 'SetServiceStartupType'
                        ComputerName = $testData.Services[0].MachineName
                        ServiceName  = $testData.Services[0].ServiceName
                        DisplayName  = $testData.Services[0].DisplayName
                        Status       = $testData.Services[0].Status
                        StartupType  = 'Disabled'
                        Action       = "updated StartupType from '$($testData.Services[0].StartType)' to 'Disabled'"
                        Error        = $null
                    }
                    @{
                        # StopService
                        Task         = 1
                        Part         = 'StopService'
                        ComputerName = $testData.Services[1].MachineName
                        ServiceName  = $testData.Services[1].ServiceName
                        DisplayName  = $testData.Services[1].DisplayName
                        Status       = 'Stopped'
                        StartupType  = 'DelayedAutoStart'
                        Action       = 'stopped service'
                        Error        = $null
                    }
                    @{
                        # StartService
                        Task         = 1
                        Part         = 'StartService'
                        ComputerName = $testData.Services[2].MachineName
                        ServiceName  = $testData.Services[2].ServiceName
                        DisplayName  = $testData.Services[2].DisplayName
                        Status       = 'Running'
                        StartupType  = 'DelayedAutoStart'
                        Action       = 'started service'
                        Error        = $null
                    }
                )

                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Services'
            }
            It 'in the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.ServiceName -eq $testRow.ServiceName
                    }
                    $actualRow.Task | Should -Be $testRow.Task
                    $actualRow.Part | Should -Be $testRow.Part
                    $actualRow.ComputerName | Should -Be $testRow.ComputerName
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.ServiceName | Should -Be $testRow.ServiceName
                    $actualRow.Status | Should -Be $testRow.Status
                    $actualRow.StartupType | Should -Be $testRow.StartupType
                    $actualRow.Action | Should -Be $testRow.Action
                    $actualRow.Error | Should -Be $testRow.Error
                    $actualRow.Date.ToString('yyyyMMdd HHmm') |
                    Should -Not -BeNullOrEmpty
                }
            }
        }
        Context "with worksheet 'Processes'" {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        Task         = $i
                        Part         = 'KillProcess'
                        ComputerName = $testData.Processes[0].MachineName
                        ProcessName  = $testData.Processes[0].ProcessName
                        Id           = $testData.Processes[0].Id
                        Action       = 'stopped running process'
                        Error        = $null
                    }
                )

                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Processes'
            }
            It 'in the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.ProcessName -eq $testRow.ProcessName
                    }
                    $actualRow.Task | Should -Be $testRow.Task
                    $actualRow.Part | Should -Be $testRow.Part
                    $actualRow.ComputerName | Should -Be $testRow.ComputerName
                    $actualRow.ProcessName | Should -Be $testRow.ProcessName
                    $actualRow.Id | Should -Be $testRow.Id
                    $actualRow.Action | Should -Be $testRow.Action
                    $actualRow.Error | Should -Be $testRow.Error
                    $actualRow.Date.ToString('yyyyMMdd HHmm') |
                    Should -Not -BeNullOrEmpty
                }
            }
        }
    }
    Context 'an e-mail is sent to the user' {
        BeforeAll {
            $testMail = @{
                Header      = $testParams.ScriptName
                To          = $testJsonFile.SendMail.To
                Bcc         = $ScriptAdmin
                Priority    = 'Normal'
                Subject     = '3 services, 1 process'
                Message     = "*<p>Manage services and processes: configure the service startup type, stop a service, stop a process, start a service.</p>*
                *<th*>Services</th>*
                *<td>Rows</td>*3*
                *<td>Actions</td>*3*
                *<td>Errors</td>*0*
                *<th*>Processes</th>*
                *<td>Rows</td>*1*
                *<td>Actions</td>*1*
                *<td>Errors</td>*0*Check the attachment for details*"
                Attachments = $testExcelLogFile.FullName
            }
        }
        It 'Send-MailHC is called with the correct arguments' {
            $mailParams.Header | Should -Be $testMail.Header
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Priority | Should -Be $testMail.Priority
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -Be $testMail.Attachments
        }
        It 'Send-MailHC is called once' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Header -eq $testMail.Header) -and
                ($To -eq $testMail.To) -and
                ($Bcc -eq $testMail.Bcc) -and
                ($Priority -eq $testMail.Priority) -and
                ($Subject -eq $testMail.Subject) -and
                ($Attachments -like $testMail.Attachments) -and
                ($Message -like $testMail.Message)
            }
        }
    }
}