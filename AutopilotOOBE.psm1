function Start-AutopilotOOBE {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$Profile,

        [ValidateSet (
            'GroupTag',
            'AddToGroup',
            'AssignedUser',
            'AssignedComputerName',
            'Assign'
        )]
        [string[]]$Disable,

        [string]$AddToGroup,
        [switch]$Assign,
        [string]$AssignedUser,
        [string]$AssignedUserExample = 'someone@example.com',
        [string]$AssignedComputerName,
        [string]$AssignedComputerNameExample = 'Azure AD Join Only',
        [string]$GroupTag,
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
        [string]$Docs = 'https://docs.microsoft.com/en-us/mem/autopilot/',
        [string]$Title = 'Join Autopilot OOBE'
    )
    #=======================================================================
    #   Profile OSDeploy
    #=======================================================================
    if ($Profile -in 'OSD','OSDeploy','OSDeploy.com') {
        $Assign = $true
        $AssignedUserExample = 'someone@osdeploy.com'
        $AssignedComputerName = 'OSD-' + ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
        $Disable = 'GroupTag'
        $AddToGroup = 'Administrators'
        $PostAction = 'SysprepReboot'
        $Run = 'NetworkingWireless'
        $Title = 'Welcome to OSDeploy.com Autopilot'
    }
    #=======================================================================
    #   Profile SeguraOSD
    #=======================================================================
    if ($Profile -match 'SeguraOSD') {
        $Assign = $true
        $AssignedUserExample = 'david@segura.org'
        $AssignedComputerName = ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
        $Disable = 'GroupTag'
        $AddToGroup = 'Twitter'
        $PostAction = 'SysprepReboot'
        $Run = 'PowerShell'
        $Title = 'Welcome to @SeguraOSD Autopilot'
    }
    #=======================================================================
    #   Profile Baker Hughes
    #=======================================================================
    if ($Profile -eq 'BH') {
        $Assign = $true
        $AssignedUserExample = 'first.last@bakerhughes.com'
        $AssignedComputerNameExample = 'AAD Only'
        $Disable = 'AddToGroup','AssignedComputerName'
        $GroupTag = 'Enterprise'
        $PostAction = 'SysprepReboot'
        $Run = 'NetworkingWireless'
        $Title = 'Welcome to Baker Hughes Autopilot'
    }
    #=======================================================================
    #   Set Global Variable
    #=======================================================================
    $Global:AutopilotOOBE = @{
        AddToGroup = $AddToGroup
        Assign = $Assign
        AssignedUser = $AssignedUser
        AssignedUserExample = $AssignedUserExample
        AssignedComputerName = $AssignedComputerName
        AssignedComputerNameExample = $AssignedComputerNameExample
        Disable = $Disable
        GroupTag = $GroupTag
        PostAction = $PostAction
        Run = $Run
        Docs = $Docs
        Title = $Title
    }
    #=======================================================================
    #   Launch
    #=======================================================================
    & "$($MyInvocation.MyCommand.Module.ModuleBase)\Forms\Join-AutopilotOOBE.ps1"
}
#=======================================================================
#   Create Alias
#=======================================================================
New-Alias -Name AutopilotOOBE -Value Start-AutopilotOOBE -Force -ErrorAction SilentlyContinue
Export-ModuleMember -Function Start-AutopilotOOBE -Alias AutopilotOOBE