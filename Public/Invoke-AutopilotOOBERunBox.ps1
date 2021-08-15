function Invoke-AutopilotOOBERunBox {
    [CmdletBinding()]
    param (
        [ValidateSet(
            'Diagnostics',
            'DiagnosticsOnline'
        )]
        [string]$Action
    )
    
    $Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-$Action.log"
    Start-Transcript -Path (Join-Path "$env:SystemRoot\Temp" $Transcript) -ErrorAction Ignore

    switch ($Action) {

        Diagnostics {
                if (!(Get-Command Get-AutopilotDiagnostics -ErrorAction Ignore)) {
                    Write-Host -ForegroundColor Cyan "Install-Script Get-AutopilotDiagnostics -Force -Verbose"
                    Install-Script Get-AutopilotDiagnostics -Force -Verbose
                }
                & "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1"
            }
        DiagnosticsOnline {
                if (!(Get-Command Get-AutopilotDiagnostics -ErrorAction Ignore)) {
                    Write-Host -ForegroundColor Cyan "Install-Script Get-AutopilotDiagnostics -Force -Verbose"
                    Install-Script Get-AutopilotDiagnostics -Force -Verbose
                }
                & "$env:ProgramFiles\WindowsPowerShell\Scripts\Get-AutopilotDiagnostics.ps1" -Online
            }
    }
    Stop-Transcript
}