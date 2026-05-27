 param
(
    # Where to stage the CAB/XML
    [Parameter()]
    [string]$WorkDir = "$env:TEMP\DellCatalog",

    # Force re-download even if a recent catalog already exists
    [Parameter()]
    [switch]$ForceRefresh,

    # Check for an update only. Does not install available update.
    [Parameter()]
    [switch]$Check
)

Start-Transcript -Path "C:\Windows\Logs\Software\Update-DellBios.log" -IncludeInvocationHeader -Append

$Restart = $false

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$ComputerSystem = Get-CimInstance -ClassName Win32_Computersystem
$ValidManufacturers = @("Dell Inc.")

If($ValidManufacturers -notcontains $ComputerSystem.Manufacturer)
{
    throw "This script is only for Dell workstations. It does not work for $($ComputerSystem.Manufacturer) systems."
}

$SystemID = $ComputerSystem.SystemSKUNumber
$ModelName = $ComputerSystem.Model

Write-Host "This systems is a $ModelName ($SystemID)"

If(-not(Test-Path $WorkDir))
{
    Write-Host "Creating $WorkDir"
    New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
}

$cabUrl   = 'https://downloads.dell.com/catalog/CatalogPC.cab'
$cabPath  = Join-Path $env:TEMP 'CatalogPC.cab'
$xmlPath  = Join-Path $WorkDir 'CatalogPC.xml'

# Refresh policy: re-use XML if <12 hours old, unless -ForceRefresh
$needsDownload = $true
if((Test-Path $xmlPath) -and (-not $ForceRefresh))
{
    $ageHours = (New-TimeSpan -Start (Get-Item $xmlPath).LastWriteTime -End (Get-Date)).TotalHours
    if($ageHours -lt 12)
    {
        Write-Host "Existing $xmlPath found with recent date. Skipping download."
        $needsDownload = $false
    }
}

if ($needsDownload)
{
    Write-Host "Downloading Dell cab file"
    # Download the CAB
    Invoke-WebRequest -Uri $cabUrl -OutFile $cabPath

    # Expand the CAB -> CatalogPC.xml
    # Use built-in expand.exe (works on all supported Windows)
    if(Test-Path $xmlPath)
    {
        Write-Host "Removing existing file $xmlPath"
        Remove-Item $xmlPath -Force
    }

    Write-Host "Expanding $cabPath to $WorkDir"
    $shell = New-Object -ComObject Shell.Application
    $src   = $shell.NameSpace($CabPath)
    $dst   = $shell.NameSpace($WorkDir)
    $dst.CopyHere($src.Items(), 16)
    Start-Sleep -Milliseconds 500

    if(-not(Test-Path $xmlPath))
    {
        throw "Failed to extract CatalogPC.xml from $cabPath"
    }
}

Write-Host "Loading $xmlPath"
# Load XML
[xml]$catalog = Get-Content -LiteralPath $xmlPath

# Helper to try parse [version] safely
function Convert-ToVersion([string]$v)
{
    try
    {
        return [version]$v
    }
    catch
    {
        return $null
    }
}

Write-Host "Checking for BIOS updates in the xml file"
# --- SystemID-first selection using correct BIOS predicate --
$biosNodes = $catalog.SelectNodes("//Model[@systemID='$systemId']/ancestor::SoftwareComponent[ComponentType/@value='BIOS']")

# Try via ancestor axis (robust to intermediate nodes)
If((-not $biosNodes) -or ($biosNodes.Count -eq 0))
{
    Write-Host "Checking for BIOS updates in the xml file via ancestor axis"
    $biosNodes = $catalog.SelectNodes("//SoftwareComponent[ComponentType/@value='BIOS' and .//SupportedSystems//Model[@systemID='$systemId']]")
}

# Fallback (descendant style) in case some catalogs nest different
If((-not $biosNodes) -or ($biosNodes.Count -eq 0))
{
    Write-Host "Checking for BIOS updates in the xml file using descendant style"
    $biosNodes = $catalog.SelectNodes("//SoftwareComponent[Category/@value='BI' and .//SupportedSystems//Model[@systemID='$systemId']]")
}

# Optional extra fallback using Category
If((-not $biosNodes) -or ($biosNodes.Count -eq 0))
{
    Write-Host "Checking for BIOS updates in the xml file using category"
    $biosNodes = $catalog.SelectNodes("//SoftwareComponent[Category/@value='BI' and .//SupportedSystems//Model[@systemID='$systemId']]")
}

Write-Host "$($biosNodes.Count) BIOS updates were found"

$biosEntries = foreach($sc in $biosNodes)
{
    # Convert to string once to allow regex match on any subnode text
    $ss = $sc.SupportedSystems
    $ssText = if($ss)
    {
        $ss.OuterXml
    }
    else
    {
        ''
    }

    if ($ssText -match [regex]::Escape($SystemID))
    {
        # Build a PowerShell object with details we care about
        [pscustomobject]@{
            Name          = $sc.name
            DellVersion   = $sc.dellVersion
            VendorVersion = $sc.vendorVersion
            ReleaseDate   = $sc.releaseDate
            Path          = $sc.path
            Url           = if ($sc.path) { "https://dl.dell.com/$($sc.path)" } else { $null }
            HashMD5       = $sc.hashMD5
            Raw           = $sc
        }
    }
}

if (-not $biosEntries)
{
    throw "No BIOS entries found in catalog for $ModelName with SystemID matching $SystemID."
}

# Pick the latest by version (prefer DellVersion, fall back to VendorVersion, finally by ReleaseDate)

$latest = $biosEntries |
    Sort-Object -Property @{Expression = { Convert-ToVersion $_.DellVersion }   ; Descending = $true},
                           @{Expression = { Convert-ToVersion $_.VendorVersion }; Descending = $true},
                           @{Expression = { [datetime]$_.ReleaseDate }           ; Descending = $true} |
    Select-Object -First 1


# Return a clean object
$BiosUpdateInfo = [pscustomobject]@{
    ModelFilter   = $ModelName
    Name          = $latest.Name
    DellVersion   = $latest.DellVersion
    VendorVersion = $latest.VendorVersion
    ReleaseDate   = $latest.ReleaseDate
    Url           = $latest.Url
    Md5           = $latest.HashMD5
    CatalogPath   = $xmlPath
}

[version]$BiosUpdateVersion = $BiosUpdateInfo.DellVersion
$BiosUpdateReleaseDate = $BiosUpdateInfo.ReleaseDate
[version]$CurrentBiosVersion = (Get-CIMInstance -ClassName Win32_BIOS).SMBIOSBIOSVersion
$CurrentBiosReleaseDate = (Get-CIMInstance -ClassName Win32_BIOS).ReleaseDate

Write-Host "Current BIOS version: $CurrentBiosVersion - Release date $CurrentBiosReleaseDate"
Write-Host "Update BIOS version: $BiosUpdateVersion - Release date $BiosUpdateReleaseDate"

If($CurrentBiosVersion -ge $BiosUpdateVersion)
{
    Write-Host "BIOS is up to date. No update needed."
}
else
{
    If($Check)
    {
        Write-Host "Check mode specified. Not proceeding with available BIOS update."
    }
    else
    {
        Write-Host "Proceeding with BIOS update."
        $BiosUpdateFile = Split-Path -Path $BiosUpdateInfo.Url -Leaf
        $BiosUpdatePath = "$env:TEMP\$BiosUpdateFile"
        Write-Host "Downloading $($BiosUpdateInfo.Url) to $BiosUpdatePath"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $BiosUpdateInfo.Url -OutFile $BiosUpdatePath -Headers @{ 'User-Agent' = 'Mozilla/5.0' }

        If($BiosUpdateInfo.Md5)
        {
            Write-Host "Verifying downloaded file"
            $BiosUpdateHash = (Get-FileHash -Algorithm MD5 -Path $BiosUpdatePath).Hash.ToLower()
            If($BiosUpdateInfo.Md5 -ne $BiosUpdateHash)
            {
                throw "MD5 mismatch for $BiosUpdatePath. Expected $($BiosUpdateInfo.Md5) but got $BiosUpdateHash"
            }
            else
            {
                Write-Host "Bios update file has been verified"
            }
        }

        Write-Host "Starting BIOS update $BiosUpdatePath with parameters /s /bls. The system will restart to complete the update."
        Start-Process -FilePath $BiosUpdatePath -ArgumentList "/s /bls" -NoNewWindow -PassThru -Wait
        $Restart = $true
    }
}

If($Restart)
{
    Write-Host "Restarting system to complete the update"
    Stop-Transcript
    Restart-Computer -Force
}
Else
{
    Stop-Transcript
}