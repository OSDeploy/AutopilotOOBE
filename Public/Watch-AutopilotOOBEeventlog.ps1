function Watch-AutopilotOOBEeventlog {
    [CmdletBinding()]
    param ()
    #================================================
    #   Initialize
    #================================================
    $Title = 'Watch-AutopilotOOBEeventlog'
    $host.ui.RawUI.WindowTitle = $Title
    $host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size(2000,2000)
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
    #================================================
    #   Main Variables
    #================================================
    $Monitor = $true
    $Results = @()
    $FormatEnumerationLimit = -1
    # This will go back 5 days in the logs.  Adjust as needed
    [DateTime]$StartTime = (Get-Date).AddDays(- 1)
    $ExcludeEventId = @(200,202,260,263,266,272)
    # Remove Line Wrap
    reg add HKCU\Console /v LineWrap /t REG_DWORD /d 0 /f
    #================================================
    #   LogName
    #   These are the WinEvent logs to monitor
    #================================================
    $LogName = @(
        'Microsoft-Windows-AAD/Operational'
        #'Microsoft-Windows-AppXDeploymentServer/Operational'
        'Microsoft-Windows-AssignedAccess/Admin'
        'Microsoft-Windows-AssignedAccess/Operational'
        'Microsoft-Windows-AssignedAccessBroker/Admin'
        'Microsoft-Windows-AssignedAccessBroker/Operational'
        'Microsoft-Windows-Crypto-NCrypt/Operational'
        'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'
        #'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Debug'
        'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Operational'
        'Microsoft-Windows-ModernDeployment-Diagnostics-Provider/Autopilot'
        'Microsoft-Windows-ModernDeployment-Diagnostics-Provider/ManagementService'
        'Microsoft-Windows-Provisioning-Diagnostics-Provider/Admin'
        'Microsoft-Windows-Shell-Core/Operational'
        'Microsoft-Windows-Time-Service/Operational'
        'Microsoft-Windows-User Device Registration/Admin'
    )
    #================================================
    #   FilterHashtable
    #================================================
    $FilterHashtable = @{
        StartTime = $StartTime
        LogName = $LogName
    }
    #================================================
    #   Get-WinEvent Results
    #================================================
    $Results = Get-WinEvent -FilterHashtable $FilterHashtable -ErrorAction Ignore | Sort-Object TimeCreated | Where-Object {$_.Id -notin $ExcludeEventId}
    $Results = $Results | Select-Object TimeCreated,LevelDisplayName,LogName,Id, @{Name='Message';Expression={ ($_.Message -Split '\n')[0]}}
    #================================================
    #   Display Results
    #================================================
    foreach ($Item in $Results) {
        if ($Item.LevelDisplayName -eq 'Error') {
            Write-Host "$($Item.TimeCreated) $($Item.LevelDisplayName) $($Item.Id) $($Item.Message)" -ForegroundColor Red
        }
        elseif ($Item.LevelDisplayName -eq 'Warning') {
            Write-Host "$($Item.TimeCreated) $($Item.LevelDisplayName) $($Item.Id) $($Item.Message)" -ForegroundColor Yellow
            
        }
        else {
            Write-Host "$($Item.TimeCreated) $($Item.LevelDisplayName) $($Item.Id) $($Item.Message)"
        }
    }
    #================================================
    #   Monitor New Events
    #================================================
    if ($Monitor) {
        Write-Host -ForegroundColor Cyan "Listening for new events"
        while ($true) {
            Start-Sleep -Seconds 1 | Out-Null
            #================================================
            #   Get-WinEvent NewResults
            #================================================
            $NewResults = Get-WinEvent -FilterHashtable $FilterHashtable -ErrorAction Ignore | Sort-Object TimeCreated | Where-Object {$_.Id -notin $ExcludeEventId} | Where-Object {$_.TimeCreated -notin $Results.TimeCreated}
            if ($NewResults) {
                $Results += $NewResults
            }
            $NewResults = $NewResults | Select-Object TimeCreated,LevelDisplayName,LogName,Id, @{Name='Message';Expression={ ($_.Message -Split '\n')[0]}}
            #================================================
            #   Display Results
            #================================================
            foreach ($Item in $NewResults) {
                if ($Item.LevelDisplayName -eq 'Error') {
                    Write-Host "$($Item.TimeCreated) $($Item.LevelDisplayName) $($Item.Id) $($Item.Message)" -ForegroundColor Red
                }
                elseif ($Item.LevelDisplayName -eq 'Warning') {
                    Write-Host "$($Item.TimeCreated) $($Item.LevelDisplayName) $($Item.Id) $($Item.Message)" -ForegroundColor Yellow
                    
                }
                else {
                    Write-Host "$($Item.TimeCreated) $($Item.LevelDisplayName) $($Item.Id) $($Item.Message)"
                }
            }
        }
    }
}