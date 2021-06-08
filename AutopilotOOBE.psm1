function Start-AutopilotOOBE {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$CustomProfile,
        [switch]$Demo,

        [ValidateSet (
            'GroupTag',
            'AddToGroup',
            'AssignedUser',
            'AssignedComputerName',
            'Assign'
        )]
        [string[]]$Disabled,

        [ValidateSet (
            'GroupTag',
            'AddToGroup',
            'AssignedUser',
            'AssignedComputerName',
            'Assign'
        )]
        [string[]]$Hidden,

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
    if ($CustomProfile -in 'OSD','OSDeploy','OSDeploy.com') {
        $Title = 'OSDeploy.com Autopilot Enrollment'
        $Assign = $true
        $AssignedUserExample = 'someone@osdeploy.com'
        $AssignedComputerName = 'OSD-' + ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
        $Disabled = 'GroupTag'
        $AddToGroup = 'Administrators'
        $PostAction = 'SysprepReboot'
        $Run = 'NetworkingWireless'
    }
    #=======================================================================
    #   Profile SeguraOSD
    #=======================================================================
    if ($CustomProfile -match 'SeguraOSD') {
        $Title = '@SeguraOSD Autopilot Enrollment'
        $Assign = $true
        $AssignedUserExample = 'someone@segura.org'
        $AssignedComputerName = ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
        $Disabled = 'GroupTag'
        $AddToGroup = 'Twitter'
        $PostAction = 'SysprepReboot'
        $Run = 'PowerShell'
    }
    #=======================================================================
    #   Profile Baker Hughes
    #=======================================================================
    if ($CustomProfile -eq 'BH') {
        $Title = 'Baker Hughes Autopilot Enrollment'
        $Assign = $true
        $AssignedUserExample = 'first.last@bakerhughes.com'
        $AssignedComputerNameExample = 'Disabled for Hybrid Join'
        $Disabled = 'AddToGroup','AssignedComputerName'
        $Hidden = 'AddToGroup'
        $GroupTag = 'Enterprise'
        $PostAction = 'SysprepReboot'
        $Run = 'NetworkingWireless'
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
        Disabled = $Disabled
        Demo = $Demo
        GroupTag = $GroupTag
        Hidden = $Hidden
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