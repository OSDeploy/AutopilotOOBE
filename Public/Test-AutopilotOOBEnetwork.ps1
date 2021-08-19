function Test-AutopilotOOBEnetwork {
    [CmdletBinding()]
    param ()
    #================================================
    #   Initialize
    #================================================
    $Title = 'Test-AutopilotOOBEnetwork'
    $host.ui.RawUI.WindowTitle = $Title
    $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size(2000,2000)
    $host.ui.RawUI.BackgroundColor = ($bckgrnd = 'Black')
    Clear-Host
    #================================================
    #   Temp
    #================================================
    if (!(Test-Path "$env:SystemDrive\Temp")) {
        New-Item -Path "$env:SystemDrive\Temp" -ItemType Directory -Force
    }
    #================================================
    #   Transcript
    #================================================
    $Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-$Title.log"
    Start-Transcript -Path (Join-Path "$env:SystemDrive\Temp" $Transcript) -ErrorAction Ignore
    #=======================================================================
    #   Networking Requirements
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Test HTTPS Networking Requirements"
    Write-Host -ForegroundColor DarkGray "https://docs.microsoft.com/en-us/mem/autopilot/networking-requirements"

    $Global:ProgressPreference = 'SilentlyContinue'
    #=======================================================================
    #   Microsoft NCSI Connect Test
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan 'Microsoft NCSI Connect Test'

    $Urls = @(
        'http://www.msftconnecttest.com/connecttest.txt'
    )

    foreach ($Uri in $Urls){
        try {
            if ($null = Invoke-WebRequest -Uri $Uri -Method Head -UseBasicParsing -ErrorAction Stop) {
                Write-Host -ForegroundColor DarkCyan "PASS: $Uri"
            }
            else {
            }
        }
        catch {
            Write-Host -ForegroundColor Yellow "FAIL: $Uri"
        }
    }
    #=======================================================================
    #   PowerShell Gallery
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "PowerShell Gallery"

    $ComputerNames = @('powershellgallery.com')
    $Ports = @(443)

    foreach ($ComputerName in $ComputerNames){
        foreach ($Port in $Ports){
            try {
                if (Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -ErrorAction Stop -WarningAction 'Continue') {
                    Write-Host -ForegroundColor DarkCyan "PASS: $ComputerName [Port: $Port]"
                }
                else {
                    Write-Host -ForegroundColor Yellow "FAIL: $ComputerName [Port: $Port]"
                }
            }
            catch {}
            finally {}
        }
    }
    #=======================================================================
    #   Windows Autopilot Deployment Service
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "Windows Autopilot Deployment Service"

    $ComputerNames = @(
        'cs.dds.microsoft.com'
        'login.live.com'
        'ztd.dds.microsoft.com'
    )
    $Ports = @(443)

    foreach ($ComputerName in $ComputerNames){
        foreach ($Port in $Ports){
            try {
                if (Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -ErrorAction Stop -WarningAction 'Continue') {
                    Write-Host -ForegroundColor DarkCyan "PASS: $ComputerName [Port: $Port]"
                }
                else {
                    Write-Host -ForegroundColor Yellow "FAIL: $ComputerName [Port: $Port]"
                }
            }
            catch {}
            finally {}
        }
    }
    #=======================================================================
    #   Windows Activation
    #   https://support.microsoft.com/en-us/topic/windows-activation-or-validation-fails-with-error-code-0x8004fe33-a9afe65e-230b-c1ed-3414-39acd7fddf52
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan 'Windows Activation'

    $Urls = @(
        'http://crl.microsoft.com/pki/crl/products/MicProSecSerCA_2007-12-04.crl'
    )

    foreach ($Uri in $Urls){
        try {
            if ($null = Invoke-WebRequest -Uri $Uri -Method Head -UseBasicParsing -ErrorAction Stop) {
                Write-Host -ForegroundColor DarkCyan "PASS: $Uri"
            }
            else {
            }
        }
        catch {
            Write-Host -ForegroundColor Yellow "FAIL: $Uri"
        }
    }

    $ComputerNames = @(
        'activation.sls.microsoft.com'
        'activation-v2.sls.microsoft.com'
        'displaycatalog.mp.microsoft.com'
        'displaycatalog.md.mp.microsoft.com'
        'go.microsoft.com'
        'licensing.mp.microsoft.com'
        'licensing.md.mp.microsoft.com'
        'login.live.com'
        'purchase.md.mp.microsoft.com'
        'purchase.mp.microsoft.com'
        'validation.sls.microsoft.com'
        'validation-v2.sls.microsoft.com'
    )
    $Ports = @(443)

    foreach ($ComputerName in $ComputerNames){
        foreach ($Port in $Ports){
            try {
                if (Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -ErrorAction Stop -WarningAction 'Continue') {
                    Write-Host -ForegroundColor DarkCyan "PASS: $ComputerName [Port: $Port]"
                }
                else {
                    Write-Host -ForegroundColor Yellow "FAIL: $ComputerName [Port: $Port]"
                }
            }
            catch {}
            finally {}
        }
    }
    #=======================================================================
    #   Azure Active Directory | Office 365 IP Address and URL web service
    #   https://docs.microsoft.com/en-us/microsoft-365/enterprise/microsoft-365-ip-web-service?view=o365-worldwide
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "Office 365 IP Address and URL Web Service"

    $Urls = @()

    foreach ($Uri in $Urls){
        try {
            if ($null = Invoke-WebRequest -Uri $Uri -Method Head -UseBasicParsing -ErrorAction Stop) {
                Write-Host -ForegroundColor DarkCyan "PASS: $Uri"
            }
            else {
            }
        }
        catch {
            Write-Host -ForegroundColor Yellow "FAIL: $Uri"
        }
    }

    $ComputerNames = @(
        'broadcast.skype.com'
        'compliance.microsoft.com'
        'lync.com'
        'mail.protection.outlook.com'
        'msftidentity.com'
        'msidentity.com'
        'officeapps.live.com'
        'online.office.com'
        'outlook.office.com'
        'portal.cloudappsecurity.com'
        'protection.office.com'
        'protection.outlook.com'
        'security.microsoft.com'
        'sharepoint.com'
        'skypeforbusiness.com'
        'teams.microsoft.com'
        'account.activedirectory.windowsazure.com'
        'account.office.net'
        'accounts.accesscontrol.windows.net'
        'adminwebservice.microsoftonline.com'
        'api.passwordreset.microsoftonline.com'
        'autologon.microsoftazuread-sso.com'
        'becws.microsoftonline.com'
        'broadcast.skype.com'
        'clientconfig.microsoftonline-p.net'
        'companymanager.microsoftonline.com'
        'compliance.microsoft.com'
        'device.login.microsoftonline.com'
        'graph.microsoft.com'
        'graph.windows.net'
        'login.microsoft.com'
        'login.microsoftonline.com'
        'login.microsoftonline-p.com'
        'login.windows.net'
        'logincert.microsoftonline.com'
        'loginex.microsoftonline.com'
        'login-us.microsoftonline.com'
        'nexus.microsoftonline-p.com'
        'office.live.com'
        'outlook.office.com'
        'outlook.office365.com'
        'passwordreset.microsoftonline.com'
        'protection.office.com'
        'provisioningapi.microsoftonline.com'
        'security.microsoft.com'
        'smtp.office365.com'
        'teams.microsoft.com'
    )
    $Ports = @(443)

    foreach ($ComputerName in $ComputerNames){
        foreach ($Port in $Ports){
            try {
                if (Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -ErrorAction Stop -WarningAction 'Continue') {
                    Write-Host -ForegroundColor DarkCyan "PASS: $ComputerName [Port: $Port]"
                }
                else {
                    Write-Host -ForegroundColor Yellow "FAIL: $ComputerName [Port: $Port]"
                }
            }
            catch {}
            finally {}
        }
    }
    #=======================================================================
    #   Intune
    #   https://docs.microsoft.com/en-us/mem/intune/fundamentals/network-bandwidth-use
    #   https://docs.microsoft.com/en-us/mem/intune/fundamentals/intune-endpoints#access-for-managed-devices
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "Intune Access for managed devices"

    $ComputerNames = @(
        'login.microsoftonline.com'
        #'*.officeconfig.msocdn.com'
        'config.office.com'
        'graph.windows.net'
        'enterpriseregistration.windows.net'
        'portal.manage.microsoft.com'
        'm.manage.microsoft.com'
        'fef.msuc03.manage.microsoft.com'
        'wip.mam.manage.microsoft.com'
        'mam.manage.microsoft.com'
        #'*manage.microsoft.com'
    )
    $Ports = @(443)

    foreach ($ComputerName in $ComputerNames){
        foreach ($Port in $Ports){
            try {
                if (Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -ErrorAction Stop -WarningAction 'Continue') {
                    Write-Host -ForegroundColor DarkCyan "PASS: $ComputerName [Port: $Port]"
                }
                else {
                    Write-Host -ForegroundColor Yellow "FAIL: $ComputerName [Port: $Port]"
                }
            }
            catch {}
            finally {}
        }
    }
    #=======================================================================
    #   Intune
    #   https://docs.microsoft.com/en-us/mem/intune/fundamentals/network-bandwidth-use
    #   https://docs.microsoft.com/en-us/mem/intune/fundamentals/intune-endpoints#network-requirements-for-powershell-scripts-and-win32-apps
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "Intune Network requirements for PowerShell scripts and Win32 apps"

    $ComputerNames = @(
        'naprodimedatapri.azureedge.net'
        'naprodimedatasec.azureedge.net'
        'naprodimedatahotfix.azureedge.net'
        'euprodimedatapri.azureedge.net'
        'euprodimedatasec.azureedge.net'
        'euprodimedatahotfix.azureedge.net'
        'approdimedatapri.azureedge.net'
        'approdimedatasec.azureedge.net'
        'approdimedatahotfix.azureedge.net'

    )
    $Ports = @(443)

    foreach ($ComputerName in $ComputerNames){
        foreach ($Port in $Ports){
            try {
                if (Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -ErrorAction Stop -WarningAction 'Continue') {
                    Write-Host -ForegroundColor DarkCyan "PASS: $ComputerName [Port: $Port]"
                }
                else {
                    Write-Host -ForegroundColor Yellow "FAIL: $ComputerName [Port: $Port]"
                }
            }
            catch {}
            finally {}
        }
    }
    #=======================================================================
    #   Windows Update
    #   https://docs.microsoft.com/en-US/windows/deployment/update/windows-update-troubleshooting
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "Windows Update"

    $ComputerNames = @(
        'emdl.ws.microsoft.com'
        'dl.delivery.mp.microsoft.com'
        #'windowsupdate.com'
    )
    $Ports = @(80)

    foreach ($ComputerName in $ComputerNames){
        foreach ($Port in $Ports){
            try {
                if (Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -ErrorAction Stop -WarningAction 'Continue') {
                    Write-Host -ForegroundColor DarkCyan "PASS: $ComputerName [Port: $Port]"
                }
                else {
                    Write-Host -ForegroundColor Yellow "FAIL: $ComputerName [Port: $Port]"
                }
            }
            catch {}
            finally {}
        }
    }
    #=======================================================================
    #   Autopilot self-Deploying mode and Autopilot pre-provisioning
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "Autopilot self-Deploying mode and Autopilot pre-provisioning"

    $ComputerNames = @(
        'ekop.intel.com'
        'ekcert.spserv.microsoft.com'
        'ftpm.amd.com'
    )
    $Ports = @(443)

    foreach ($ComputerName in $ComputerNames){
        foreach ($Port in $Ports){
            try {
                if (Test-NetConnection -ComputerName $ComputerName -Port $Port -InformationLevel Quiet -ErrorAction Stop -WarningAction 'Continue') {
                    Write-Host -ForegroundColor DarkCyan "PASS: $ComputerName [Port: $Port]"
                }
                else {
                    Write-Host -ForegroundColor Yellow "FAIL: $ComputerName [Port: $Port]"
                }
            }
            catch {}
            finally {}
        }
    }
    #=======================================================================
    #   Complete
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    $Global:ProgressPreference = 'Continue'
    Stop-Transcript
}