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
    Mock Set-DelayedAutoStartHC
    Mock Test-DelayedAutoStartHC { $true }

    Mock Get-Service
    Mock Invoke-Command
    Mock Set-Service
    Mock Start-Service
    Mock Stop-Service
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
                Tasks    = @(
                    @{
                        ComputerName          = @()
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
                SendMail = @{
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
            It 'Task.<_>' -ForEach @(
                'ComputerName', 'SetServiceStartupType', 'Execute'
            ) {
                $testJsonFile.Tasks[0].Remove($_)
                $testJsonFile | ConvertTo-Json -Depth 3 | 
                Out-File @testOutParams

                .$testScript @testParams
                        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Property '$_' not found in one of the 'Tasks'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Task.SetServiceStartupType.<_>' -ForEach @(
                'Automatic', 'DelayedAutostart', 'Disabled', 'Manual'
            ) {
                $testJsonFile.Tasks[0].SetServiceStartupType.Remove($_)
                $testJsonFile | ConvertTo-Json -Depth 3 | 
                Out-File @testOutParams

                .$testScript @testParams
                        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Property 'SetServiceStartupType.$_' not found in one of the 'Tasks'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            } 
            It 'Task.Execute.<_>' -ForEach @(
                'StopService', 'KillProcess', 'StartService'
            ) {
                $testJsonFile = @{
                    Tasks    = @(
                        @{
                            ComputerName          = @()
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
                    SendMail = @{
                        To = 'bob@contoso.com'
                    }
                }
                $testJsonFile.Tasks[0].Execute.Remove($_)
                $testJsonFile | ConvertTo-Json -Depth 3 | 
                Out-File @testOutParams

                .$testScript @testParams
                        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Property 'Execute.$_' not found in one of the 'Tasks'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            } 
            It 'SendMail.To' {
                $testJsonFile.SendMail.Remove('To')
                $testJsonFile | ConvertTo-Json -Depth 3 | 
                Out-File @testOutParams

                .$testScript @testParams
                        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*No 'SendMail.To' addresses found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
        Context 'is missing content for property' {
            It 'Tasks' {
                $testJsonFile.Tasks = @()
                $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams
    
                .$testScript @testParams
                            
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*No 'Tasks' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Task.ComputerName' {
                $testJsonFile.Tasks[0].ComputerName = @()
                $testJsonFile | ConvertTo-Json -Depth 3 | 
                Out-File @testOutParams

                .$testScript @testParams
                        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*No 'ComputerName' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Task.SetServiceStartupType and Task.Execute' {
                $testJsonFile.Tasks[0].ComputerName = @('PC1')
                $testJsonFile | ConvertTo-Json -Depth 3 | 
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
                $testJsonFile.Tasks[0].ComputerName = @('PC1')
                $testJsonFile.Tasks[0].Execute.KillProcess = @('chrome')
            }
            It 'duplicate ComputerName in one task' {
                $testJsonFile.Tasks[0].ComputerName = @('PC1', 'PC1', 'PC2')
                $testJsonFile | ConvertTo-Json -Depth 3 | 
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
                    $testJsonFile | ConvertTo-Json -Depth 3 | 
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
                    $testJsonFile | ConvertTo-Json -Depth 3 | 
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
Describe 'service startup types are corrected when they are incorrect' {
    BeforeEach {
        $testJsonFile = @{
            Tasks    = @(
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
            SendMail = @{
                To = 'bob@contoso.com'
            }
        }
    }
    It "actual '<actual>' expected '<expected>'"-ForEach @(
        @{ actual = 'Automatic'; expected = 'DelayedAutoStart' }
        @{ actual = 'Automatic'; expected = 'Disabled' }
        @{ actual = 'Automatic'; expected = 'Manual' }
        @{ actual = 'DelayedAutoStart'; expected = 'Automatic' }
        @{ actual = 'DelayedAutoStart'; expected = 'Disabled' }
        @{ actual = 'DelayedAutoStart'; expected = 'Manual' }
        @{ actual = 'Disabled'; expected = 'Automatic' }
        @{ actual = 'Disabled'; expected = 'DelayedAutoStart' }
        @{ actual = 'Disabled'; expected = 'Manual' }
        @{ actual = 'Manual'; expected = 'Automatic' }
        @{ actual = 'Manual'; expected = 'DelayedAutoStart' }
        @{ actual = 'Manual'; expected = 'Disabled' }
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
        $testJsonFile | ConvertTo-Json -Depth 3 | Out-File @testOutParams

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
    } -Tag test
}
