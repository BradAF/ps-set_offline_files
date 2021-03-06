function Get-OfflineStatus {

    # Return the top 4 items configured for offline files, and their 'Online/Offline' status
    Get-WmiObject `
        -Class win32_offlineFilesItem | 
            Select-Object `
                -First 4 `
                -Skip 1 `
                -Property `
                    @{
                        N="Path";
                        E={$_.ItemPath}
                    },
                    @{
                        N="ConnectState";
                        E={
                            switch ($_.ConnectionInfo.Connectstate) {
                                (1) {"Offline"}
                                (2) {"Online"}
                                (3) {"Transparently Cached"}
                                (4) {"Partly Transparently Cached"}
                                default {$_.ConnectionInfo.ConnectState}
                            }
                        }
                    },
                    @{
                        N="OfflineReason";
                        E={
                            switch ($_.Connectioninfo.OfflineReason) {    
                                (0) {"Unknown"}
                                (1) {"Not applicable"}
                                (2) {"Working offline"}
                                (3) {"Slow Connection"}
                                (4) {"Net disconnected"}
                                (5) {"Need to sync item"}
                                (6) {"Item suspended"}
                                default {$_.Connectioninfo.OfflineReason}
                            }
                        }
                    }
}

function Get-ItemOfflineStatus ($itemPath) {
    
    # Return the status of just a single file
    # It is much faster to just search through the top 4 each time for the specified item.
    Get-WmiObject -Class win32_offlineFilesItem | 
        Select-Object `
            -First 4 `
            -Skip  1 |
        Where-object ItemPath -eq $itemPath
}

## main

# Check the Offline Status of the top 4 items returned. If not Offline, set them Offline
Get-OfflineStatus | 
        Where-Object ConnectState -ne "Offline" |
        ForEach-Object {
            
            # Initialize attempt counter at 1
            $i = 0
            
            # Repeat attempt to set the item Offline 5 times, or until successful
            Do {
                Write-Warning "$($_.Path) is $($_.ConnectState). Attempting to set Offline..."
                ([WMIClass]"\\localhost\root\cimv2:Win32_OfflineFilesCache").TransitionOffline(
                    ($_.Path),
                    $true
                ) | Out-Null

                # Increment attempt counter
                $i++
            } Until (((Get-ItemOfflineStatus $_.Path).ConnectionInfo.Connectstate -eq 1) -or ($i -eq 5))
        } #ForEach-Object

# Just return a final list of items and their status
Get-OfflineStatus

if ($i -gt 0) {
    Write-Warning "Please restart any running applications or close and re-open any open documents to complete the process."
}

## end main