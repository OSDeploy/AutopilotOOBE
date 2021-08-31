function Test-AutopilotOOBEconnection {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $True)]
        [string]$Uri = 'google.com'
    )
    
    begin {}
    
    process {
        $Params = @{
            Method = 'Head'
            Uri = $Uri
            UseBasicParsing = $True
        }

        try {
            Write-Verbose "Test-WebConnection OK: $Uri"
            Invoke-WebRequest @Params | Out-Null
            $true
        }
        catch {
            Write-Verbose "Test-WebConnection FAIL: $Uri"
            $false
        }
        finally {
            $Error.Clear()
        }
    }
    
    end {}
}