function Invoke-AutopilotOOBEcmd {
    [CmdletBinding()]
    param (
        [ValidateSet(
            'AutopilotDiagnostics',
            'AutopilotDiagnosticsOnline',
            'GetTpm',
            'ClearTpm',
            'InitializeTpm',
            'EventViewer'
        )]
        [string]$Action
    )
    #================================================
    #   Resources
    #================================================
    $MDMEventLog = @'
<ViewerConfig>
    <QueryConfig>
        <QueryParams>
            <UserQuery/>
        </QueryParams>
        <QueryNode>
            <Name LanguageNeutralValue="MDMDiagnosticsTool">MDMDiagnosticsTool</Name>
            <Description>MDMDiagnosticsTool</Description>
            <SuppressQueryExecutionErrors>1</SuppressQueryExecutionErrors>
            <QueryList>
                <Query>
                    <Select Path="Microsoft-Windows-AAD/Operational">*</Select>
                    <Select Path="Microsoft-Windows-AppXDeployment-Server/Operational">*</Select>
                    <Select Path="Microsoft-Windows-AssignedAccess/Admin">*</Select>
                    <Select Path="Microsoft-Windows-AssignedAccess/Operational">*</Select>
                    <Select Path="Microsoft-Windows-AssignedAccessBroker/Admin">*</Select>
                    <Select Path="Microsoft-Windows-AssignedAccessBroker/Operational">*</Select>
                    <Select Path="Microsoft-Windows-Crypto-NCrypt/Operational">*</Select>
                    <Select Path="Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin">*</Select>
                    <Select Path="Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Operational">*</Select>
                    <Select Path="Microsoft-Windows-ModernDeployment-Diagnostics-Provider/Autopilot">*</Select>
                    <Select Path="Microsoft-Windows-ModernDeployment-Diagnostics-Provider/ManagementService">*</Select>
                    <Select Path="Microsoft-Windows-Provisioning-Diagnostics-Provider/Admin">*</Select>
                    <Select Path="Microsoft-Windows-Shell-Core/Operational">*</Select>
                    <Select Path="Microsoft-Windows-User Device Registration/Admin">*</Select>
                </Query>
            </QueryList>
        </QueryNode>
    </QueryConfig>
</ViewerConfig>
'@
    #================================================
    #   Transcript
    #================================================
    $Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-$Action.log"
    Start-Transcript -Path (Join-Path "$env:SystemRoot\Temp" $Transcript) -ErrorAction Ignore
    #================================================
    #   Switch
    #================================================
    switch ($Action) {
        AutopilotDiagnostics {
            if (!(Get-Command Get-AutopilotDiagnostics -ErrorAction Ignore)) {
                Write-Host -ForegroundColor Cyan "Install-Script Get-AutopilotDiagnostics -Force -Verbose"
                Install-Script Get-AutopilotDiagnostics -Force -Verbose
            }
            if (Test-Path "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1") {
                & "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1"
            }
            else {
                Write-Warning "Unable to find $env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1"
            }
        }
        AutopilotDiagnosticsOnline {
            if (!(Get-Command Get-AutopilotDiagnostics -ErrorAction Ignore)) {
                Write-Host -ForegroundColor Cyan "Install-Script Get-AutopilotDiagnostics -Force -Verbose"
                Install-Script Get-AutopilotDiagnostics -Force -Verbose
            }
            if (Test-Path "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1") {
                & "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1" -Online
            }
            else {
                Write-Warning "Unable to find $env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1"
            }
        }
        GetTpm {
            Get-Tpm
        }
        ClearTpm {
            Clear-Tpm
            Write-Warning "Restart the computer to complete the process"
        }
        InitializeTpm {
            Initialize-Tpm -AllowClear -AllowPhysicalPresence
            Write-Warning "Restart the computer to complete the process"
        }
        EventViewer {
            $MDMEventLog | Set-Content -Path "$env:ProgramData\Microsoft\Event Viewer\Views\MDMDiagnosticsTool.xml" -Force
            Start-Sleep -Seconds 2
            try {
                Restart-Service -Name EventLog -Force -ErrorAction Ignore
            }
            catch {
                #Nothing
            }
            Show-EventLog
        }
    }
    Stop-Transcript
}