# DCOM Hunter

Not all CLSIDs are related to DCOM objects, and iterating through all CLSIDs to determine if they are DCOMs can be tedious due to the large number of CLSIDs available.

Instead this script does the reverse and gets all DCOM AppIDs, and from there does a reverse search for the corresponding CLSID and ProgID

Using the ProgID or CLSID, we can then initialize the COM object as such

Using the ProgID
```
$com = [activator]::CreateInstance([type]::GetTypeFromProgID(<ProgID>))
```

For example:
```
$com = [activator]::CreateInstance([type]::GetTypeFromProgID("MMC20.Application.1"))
```

Or using CLSID
```
$com = [Type]::GetTypeFromCLSID(<CLSID>)
```

Example:
```
$com = [Type]::GetTypeFromCLSID("49B2791A-B1AE-4C90-9B8E-E860BA07F889")
```
