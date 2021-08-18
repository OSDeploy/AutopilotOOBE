function Watch-AutopilotOOBEevents {
    [CmdletBinding()]
    param (
        [switch]$Denoise
    )
    #================================================
    #   Initialize
    #================================================
    $Title = 'Watch-AutopilotOOBEevents'
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
    #================================================
    #   Main Variables
    #================================================
    $Monitor = $true
    $Results = @()
    $FormatEnumerationLimit = -1
    # This will go back 1 days in the logs.  Adjust as needed
    [DateTime]$StartTime = (Get-Date).AddDays(- 1)

    $CyanEvents = @(153,162,164,702,704)
    $DarkEvents = @(20,261,62171)

    if ($Denoise) {
        $ExcludeEventId = @(3,9,10,11,90,91)
        $ExcludeEventId += @(101,104,106,108,110,111,112,144)
        $ExcludeEventId += @(200,202,257,258,259,260,263,265,266,272)
        $ExcludeEventId += @(507,509,510,511,512,513,514,516,518,520,522,524,525)
        $ExcludeEventId += @(813)
        $ExcludeEventId += @(1000,1001,1100,1101,1102,1709)
        $ExcludeEventId += @(28017,28018,28019,28032,28115,28125)
        $ExcludeEventId += @(62144,62170,62460)
        $ExcludeEventId += @(705,1007)
    }
    else {
        $ExcludeEventId = @(200,202,260,263,266,272)
    }

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
            Write-Host "$($Item.TimeCreated)`tERROR:$($Item.Id)  `t$($Item.Message)" -ForegroundColor Red
        }
        elseif ($Item.LevelDisplayName -eq 'Warning') {
            Write-Host "$($Item.TimeCreated)`tWARN:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Yellow
            
        }
        elseif ($Item.Id -in $DarkEvents) {
            Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor DarkGray
        }
        elseif ($Item.Message -like "CloudExperienceHost*") {
            Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Magenta
        }
        elseif ($Item.Message -like "AutopilotManager*") {
            Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Green
        }
        elseif ($Item.Id -in $CyanEvents) {
            Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Cyan
        }
        else {
            Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Gray
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
                    Write-Host "$($Item.TimeCreated)`tERROR:$($Item.Id)  `t$($Item.Message)" -ForegroundColor Red
                }
                elseif ($Item.LevelDisplayName -eq 'Warning') {
                    Write-Host "$($Item.TimeCreated)`tWARN:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Yellow
                    
                }
                elseif ($Item.Id -in $DarkEvents) {
                    Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor DarkGray
                }
                elseif ($Item.Message -like "CloudExperienceHost*") {
                    Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Magenta
                }
                elseif ($Item.Message -like "AutopilotManager*") {
                    Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Green
                }
                elseif ($Item.Id -in $CyanEvents) {
                    Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Cyan
                }
                else {
                    Write-Host "$($Item.TimeCreated)`tINFO:$($Item.Id)   `t$($Item.Message)" -ForegroundColor Gray
                }
            }
        }
    }
}