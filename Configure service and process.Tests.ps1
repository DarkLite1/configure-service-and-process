#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testName = @{
        Process = 'notepad'
        Service = 'bits'
    }

    $testInputFile = @{
        SendMail          = @{
            To = 'bob@contoso.com'
        }
        MaxConcurrentJobs = 5
        Tasks             = @(
            @{
                ComputerName          = @($env:COMPUTERNAME)
                SetServiceStartupType = @{
                    Automatic        = @()
                    DelayedAutoStart = @()
                    Disabled         = @()
                    Manual           = @()
                }
                Execute               = @{
                    StopService  = @()
                    StopProcess  = @()
                    StartService = @()
                }
            }
        )
    }

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

    Mock Send-MailHC
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
                            StopProcess  = @()
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
                'StopService', 'StopProcess', 'StartService'
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
                $testJsonFile.Tasks[0].Execute.StopProcess = @('chrome')
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
}
Context 'SetServiceStartupType' {
    BeforeAll {
        Get-Service -Name $testName.Service |
        Set-Service -StartupType 'Disabled'
    }
    It '<_>' -ForEach @('Automatic', 'Manual', 'Disabled', 'DelayedAutoStart') {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].SetServiceStartupType.$_ = $testName.Service
        $testNewInputFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

        .$testScript @testParams

        if ($_ -eq 'DelayedAutoStart') {
            (Get-Service -Name $testName.Service).StartType |
            Should -Be 'Automatic'

            Test-DelayedAutoStartHC -ServiceName $testName.Service |
            Should -BeTrue
        }
        else {
            (Get-Service -Name $testName.Service).StartType | Should -Be $_
        }
    }
}
Context 'Execute' {
    BeforeAll {
        Get-Service -Name $testName.Service |
        Set-Service -StartupType 'Automatic'
    }
    It 'StopService' {
        Start-Service -Name $testName.Service
        (Get-Service -Name $testName.Service).Status | Should -Be 'Running'

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Execute.StopService = $testName.Service
        $testNewInputFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

        .$testScript @testParams

        (Get-Service -Name $testName.Service).Status | Should -Be 'Stopped'
    }
    It 'StartService' {
        Stop-Service -Name $testName.Service
        (Get-Service -Name $testName.Service).Status | Should -Be 'Stopped'

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Execute.StartService = $testName.Service
        $testNewInputFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

        .$testScript @testParams

        (Get-Service -Name $testName.Service).Status | Should -Be 'Running'
    }
    It 'StopProcess' {
        Start-Process -FilePath $testName.Process
        Get-Process -Name $testName.Process | Should -Not -BeNullOrEmpty

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Execute.StopProcess = $testName.Process
        $testNewInputFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

        .$testScript @testParams

        Get-Process -Name $testName.Process -EA Ignore | Should -BeNullOrEmpty
    }
}
Describe 'when the script runs successfully' {
    BeforeAll {
        Get-Service -Name $testName.Service |
        Set-Service -StartupType 'Automatic'
        Start-Service -Name $testName.Service
        Start-Process -FilePath $testName.Process

        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile.Tasks[0].Execute.StopProcess = $testName.Process
        $testNewInputFile.Tasks[0].Execute.StopService = $testName.Service
        $testNewInputFile.Tasks[0].SetServiceStartupType.DelayedAutoStart = $testName.Service

        $testNewInputFile | ConvertTo-Json -Depth 5 | Out-File @testOutParams

        .$testScript @testParams

        $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Report.xlsx'
    }
    Context 'an Excel file is created' {
        Context "with worksheet 'Services'" {
            BeforeAll {
                $testExportedExcelRows = @(
                    @{
                        Task         = 1
                        Request      = 'SetServiceStartupType'
                        ComputerName = $env:COMPUTERNAME
                        Name  = $testData.Services[0].Name
                        DisplayName  = $testData.Services[0].DisplayName
                        Status       = $testData.Services[0].Status
                        StartupType  = 'Disabled'
                        Action       = "updated StartupType from '$($testData.Services[0].StartType)' to 'Disabled'"
                        Error        = $null
                    }
                    @{
                        Task         = 1
                        Request      = 'StopService'
                        ComputerName = $env:COMPUTERNAME
                        Name  = $testData.Services[1].Name
                        DisplayName  = $testData.Services[1].DisplayName
                        Status       = 'Stopped'
                        StartupType  = 'DelayedAutoStart'
                        Action       = 'stopped service'
                        Error        = $null
                    }
                    @{
                        Task         = 1
                        Request      = 'StartService'
                        ComputerName = $env:COMPUTERNAME
                        Name  = $testData.Services[2].Name
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
                        $_.Name -eq $testRow.Name
                    }
                    $actualRow.Task | Should -Be $testRow.Task
                    $actualRow.Request | Should -Be $testRow.Request
                    $actualRow.ComputerName | Should -Be $testRow.ComputerName
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.Name | Should -Be $testRow.Name
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
                        Request      = 'StopProcess'
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
                    $actualRow.Request | Should -Be $testRow.Request
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