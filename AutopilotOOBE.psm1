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
            'PostAction',
            'Assign'
        )]
        [string[]]$Disabled,

        [ValidateSet (
            'GroupTag',
            'AddToGroup',
            'AssignedUser',
            'AssignedComputerName',
            'PostAction',
            'Assign',
            'Run',
            'Docs'
        )]
        [string[]]$Hidden,

        [string]$AddToGroup,
        [string[]]$AddToGroupOptions,
        [switch]$Assign,
        [string]$AssignedUser,
        [string]$AssignedUserExample = 'someone@example.com',
        [string]$AssignedComputerName,
        [string]$AssignedComputerNameExample = 'Azure AD Join Only',
        [string]$GroupTag,
        [string[]]$GroupTagOptions,
        [ValidateSet (
            'None',
            'Restart',
            'Shutdown',
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
            'SysprepAudit',
            'EventViewer',
            'MDMDiag',
            'MDMDiagAutopilot',
            'MDMDiagAutopilotTPM'
        )]
        [string]$Run = 'PowerShell',
        [string]$Docs,
        [string]$Title = 'Autopilot Manual Enrollment'
    )
    #=======================================================================
    #   Profile OSDeploy
    #=======================================================================
    if ($CustomProfile -in 'OSD','OSDeploy','OSDeploy.com') {
        $Title = 'OSDeploy Autopilot Enrollment'
        $AddToGroup = 'Administrators'
        $AssignedUserExample = 'someone@osdeploy.com'
        $AssignedComputerName = 'OSD-' + ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
        $PostAction = 'Shutdown'
        $Assign = $true
        $Run = 'PowerShell'
        $Docs = 'https://www.osdeploy.com/'
        $Hidden = 'GroupTag'
    }
    #=======================================================================
    #   Profile SeguraOSD
    #=======================================================================
    if ($CustomProfile -match 'SeguraOSD') {
        $Title = 'SeguraOSD Autopilot Enrollment'
        $GroupTag = 'Twitter'
        $AssignedComputerName = ((Get-CimInstance -ClassName Win32_BIOS).SerialNumber).Trim()
        $PostAction = 'Restart'
        $Assign = $true
        $Run = 'WindowsSettings'
        $Docs = 'https://twitter.com/SeguraOSD'
        $Hidden = 'AddToGroup','AssignedUser'
    }
    #=======================================================================
    #   Profile Baker Hughes
    #=======================================================================
    if ($CustomProfile -eq 'BH') {
        $Title = 'Baker Hughes Autopilot Enrollment'
        $Assign = $true
        $AssignedComputerNameExample = 'Disabled for Hybrid Join'
        $Hidden = 'AddToGroup','AssignedComputerName','AssignedUser'
        $GroupTag = 'Enterprise'
        $GroupTagOptions = 'Development','Enterprise'
        $PostAction = 'Restart'
        $Run = 'EventViewer'
    }
    #=======================================================================
    #   Profile HalfMan
    #=======================================================================
    if ($CustomProfile -eq 'HalfMan') {
        $Title = 'Autopilot Enrollment'
        $Hidden = 'GroupTag'
        $AddToGroup = 'Azr_crp_ent_modern_workplace_devices'
        $AddToGroupOptions = 'Azr_crp_ent_modern_workplace_devices'
    }
    #=======================================================================
    #   Set Global Variable
    #=======================================================================
    $Global:AutopilotOOBE = @{
        AddToGroup = $AddToGroup
        AddToGroupOptions = $AddToGroupOptions
        Assign = $Assign
        AssignedUser = $AssignedUser
        AssignedUserExample = $AssignedUserExample
        AssignedComputerName = $AssignedComputerName
        AssignedComputerNameExample = $AssignedComputerNameExample
        Disabled = $Disabled
        Demo = $Demo
        GroupTag = $GroupTag
        GroupTagOptions = $GroupTagOptions
        Hidden = $Hidden
        PostAction = $PostAction
        Run = $Run
        Docs = $Docs
        Title = $Title
    }
    #=======================================================================
    #   Launch
    #=======================================================================
    Install-Script Get-AutopilotDiagnostics -Force -Verbose
    & "$($MyInvocation.MyCommand.Module.ModuleBase)\Forms\Join-AutopilotOOBE.ps1"
}
#=======================================================================
#   Create Alias
#=======================================================================
New-Alias -Name AutopilotOOBE -Value Start-AutopilotOOBE -Force -ErrorAction SilentlyContinue
Export-ModuleMember -Function Start-AutopilotOOBE -Alias AutopilotOOBE