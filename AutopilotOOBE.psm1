function Start-AutopilotOOBE {
    [CmdletBinding()]
    param (
        [switch]$Assign,
        [string]$GroupTag,
        [string]$AddToGroup,
        [ValidateSet (
            'None',
            'Sysprep',
            'SysprepReboot',
            'SysprepShutdown'
        )]
        [string]$PostAction = 'None',
        [ValidateSet (
            'CommandPrompt',
            'PowerShell',
            'PowerShellISE',
            'WindowsExplorer',
            'WindowsSettings',
            'NetworkingWireless',
            'Restart',
            'Shutdown',
            'Sysprep',
            'SysprepReboot',
            'SysprepShutdown',
            'MDMDiag',
            'MDMDiagAutopilot',
            'MDMDiagAutopilotTPM'
        )]
        [string]$Run = 'PowerShell',
        [string]$Start = 'https://docs.microsoft.com/en-us/mem/autopilot/'
    )

    $Global:AutopilotOOBE = @{
        AddToGroup = $AddToGroup
        Assign = $Assign
        GroupTag = $GroupTag
        PostAction = $PostAction
        Run = $Run
        Start = $Start
    }

    & "$($MyInvocation.MyCommand.Module.ModuleBase)\AutopilotOOBE.ps1"
}

New-Alias -Name AutopilotOOBE -Value Start-AutopilotOOBE -Force -ErrorAction SilentlyContinue
Export-ModuleMember -Function Start-AutopilotOOBE -Alias AutopilotOOBE