# Function to find the CLSID and ProgID by AppID
function Get-CLSIDAndProgIDFromAppID {
    param (
        [string]$AppID
    )
    
    # Define the registry path where AppID is linked to CLSID
    $CLSIDRegistryPath = "HKLM:\SOFTWARE\Classes\CLSID\"
    
    # Iterate through all CLSID keys and check for the AppID
    $CLSIDKeys = Get-ChildItem -Path $CLSIDRegistryPath
    
    $foundCLSID = foreach ($CLSIDKey in $CLSIDKeys) {
        $CLSIDKeyPath = "HKLM:\SOFTWARE\Classes\CLSID\$($CLSIDKey.PSChildName)"
        
        # Check if the AppID exists under the CLSID
        $currentAppID = Get-ItemProperty -Path $CLSIDKeyPath -Name "AppID" -ErrorAction SilentlyContinue
        
        if ($currentAppID -and $currentAppID.AppID -eq $AppID) {
            # Get the ProgID from the CLSID key itself
            $SubKeys = Get-ChildItem -Path $CLSIDKeyPath 

            $ProgID = "none"

            $foundProgID = foreach ($SubKey in $SubKeys) {
                if ($SubKey.PSChildName -eq "ProgID") {
                    $ProgIDPath = $CLSIDKeyPath + '\ProgID'
                    $ProgID = (Get-ItemProperty -Path $ProgIDPath).'(default)'
                }
            }

            # Return CLSID and ProgID
            [PSCustomObject]@{
                CLSID  = $CLSIDKey
                ProgID = $ProgID
            }
        }
    }
    
    if ($foundCLSID) {
        $foundCLSID
    } else {
        Write-Host "No CLSID found for AppID: $AppID"
    }
}

# Getting all DCOM AppIDs
$AppIDs = Get-CimInstance Win32_DCOMApplication | select -ExpandProperty AppId

# For each AppID, get the CLSID and ProgID
foreach ($AppID in $AppIDs) {
    $CLSIDAndProgID = Get-CLSIDAndProgIDFromAppID -AppID $AppID


    # Display the results
    if ($CLSIDAndProgID) {
        $CLSIDAndProgID | Format-Table -Property CLSID, ProgID
    }
}
