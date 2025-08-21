# Justin Verstijnen Server Install Updates and Restart script
# Github page: https://github.com/JustinVerstijnen/JV-ServerInventoryReport
# Let's start!
Write-Host "Script made by..." -ForegroundColor DarkCyan
Write-Host "     _           _   _        __     __            _   _  _                  
    | |_   _ ___| |_(_)_ __   \ \   / /__ _ __ ___| |_(_)(_)_ __   ___ _ __  
 _  | | | | / __| __| | '_ \   \ \ / / _ \ '__/ __| __| || | '_ \ / _ \ '_ \ 
| |_| | |_| \__ \ |_| | | | |   \ V /  __/ |  \__ \ |_| || | | | |  __/ | | |
 \___/ \__,_|___/\__|_|_| |_|    \_/ \___|_|  |___/\__|_|/ |_| |_|\___|_| |_|
                                                       |__/                  " -ForegroundColor DarkCyan                                                       
[CmdletBinding()]
param([string]$OutputPath)

function Test-CommandExists {
  [CmdletBinding()] param([Parameter(Mandatory)][string]$Name)
  return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

function New-Alert {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Text,[ValidateSet("error","warn","info","ok")]$Type="error")
  $icon = switch ($Type) { "error" { "[ERROR]" } "warn" { "[WARN]" } "info" { "[INFO]" } "ok" { "[OK]" } }
  $enc = [System.Net.WebUtility]::HtmlEncode($Text)
  "<div class='alert $Type'><span class='ico'>$icon</span><span>$enc</span></div>"
}

function ConvertTo-HtmlTable {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][object]$InputObject,
    [string[]]$Properties,
    [string]$Title,
    [string]$TableId,
    [string]$Classes = "compact"
  )
  try {
    $pre  = $( if ($Title) { "<h3>$Title</h3>" } else { $null } )
    $frag = @($InputObject) | Select-Object -Property $Properties | ConvertTo-Html -Fragment -PreContent $pre
    $openTag = if ($TableId) { "<table id=""$TableId"" class=""$Classes"">" } else { "<table class=""$Classes"">" }
    ($frag -join "`n") -replace "<table>", $openTag
  } catch {
    New-Alert -Text "Could not render table: $Title. $($_.Exception.Message)" -Type error
  }
}

function ConvertTo-NameValueTable {
  [CmdletBinding()]
  param([Parameter(Mandatory)][object]$Object,[string[]]$Properties,[string]$Title="Overview")
  $rows = @()
  if ($Object -is [System.Collections.IDictionary]) {
    foreach($k in $Object.Keys){ $rows += [pscustomobject]@{ Name="$k"; Value=(($Object[$k] | Out-String).Trim()) } }
  } else {
    $props = if ($Properties) { $Properties } else { $Object.PSObject.Properties.Name }
    foreach($p in $props){ $rows += [pscustomobject]@{ Name="$p"; Value=(($Object.$p | Out-String).Trim()) } }
  }
  ConvertTo-HtmlTable -InputObject $rows -Title $Title -Properties "Name","Value"
}

function Format-Preformatted {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Text,[string]$Title)
  $enc = [System.Net.WebUtility]::HtmlEncode($Text)
  $pre = $( if ($Title) { "<h3>$Title</h3>" } else { "" } )
  "$pre<pre class=""raw"">$enc</pre>"
}

function Add-Section {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Id,[Parameter(Mandatory)][string]$Title,[Parameter(Mandatory)][string]$BodyHtml,[string]$Description)
  $descHtml = if ($Description) { "<div class=""desc"">$([System.Net.WebUtility]::HtmlEncode($Description))</div>" } else { "" }
@"
  <section id="$Id" class="tab-content" aria-labelledby="tab-$Id">
    <div class="section">
      <h2>$Title</h2>
      $descHtml
      $BodyHtml
    </div>
  </section>
"@
}

function Format-Percent { [CmdletBinding()] param([double]$Part,[double]$Whole) if($Whole -le 0){ return "n/a" } [math]::Round(($Part/$Whole)*100,2).ToString("0.##") + "%" }

function ConvertTo-DateTimeSafe {
  [CmdletBinding()] param([Parameter(Mandatory)][object]$Value)
  if ($null -eq $Value) { return $null }
  if ($Value -is [datetime]) { return $Value }
  $s = [string]$Value
  if ($s -match "^\d{14}\.\d{6}(\+|-)\d{3}$") { try { return [Management.ManagementDateTimeConverter]::ToDateTime($s) } catch {} }
  try { return [DateTime]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture) } catch { try { return [DateTime]::Parse($s) } catch { return $null } }
}

if (-not $OutputPath) {
  $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
  $desktop = [Environment]::GetFolderPath("Desktop")
  $OutputPath = Join-Path $desktop "Server-Inventory_$env:COMPUTERNAME_$stamp.html"
}
$reportSections = New-Object System.Collections.Generic.List[string]

$sectionDescriptions = @{
  system   = "This page shows a summary of the complete system and Windows information."
  network  = "This page shows the complete network configuration including the raw data."
  firewall = "This page shows the Windows Firewall configuration and status."
  storage  = "This page shows the storage information like volumes, total size, free space and raw data."
  apps     = "This page shows all installed applications."
  roles    = "This page shows the Windows Server roles, SQL and IIS information (if applicable)."
  services = "This page shows all Windows services with state, start mode, account and path."
  shares   = "This page shows all created SMB shares, share permissions and NTFS ACLs."
  printers = "This page shows all installed printers, ports, drivers and IP addresses."
}

# ===================== System Info =====================
try {
  $cs   = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
  $os   = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
  $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
  $proc = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1

  $installDt = ConvertTo-DateTimeSafe $os.InstallDate
  $bootDt    = ConvertTo-DateTimeSafe $os.LastBootUpTime
  $uptimeDays = if($bootDt){ [Math]::Round((New-TimeSpan -Start $bootDt -End (Get-Date)).TotalDays,1) } else { "(unknown)" }

  $winProdId = try { (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "ProductId" -ErrorAction Stop).ProductId } catch { $null }

  $sysSummary = [ordered]@{
    "Computer name"      = $env:COMPUTERNAME
    "Domain"             = $cs.Domain
    "Manufacturer"       = $cs.Manufacturer
    "Model"              = $cs.Model
    "Serial number"      = ($bios.SerialNumber | Out-String).Trim()
    "Windows product ID" = if($winProdId){ $winProdId } else { "(unknown)" }
    "OS"                 = $os.Caption
    "OS version"         = $os.Version
    "Installed on"       = if($installDt){ $installDt.ToString("yyyy-MM-dd HH:mm") } else { "(unknown)" }
    "Last boot"          = if($bootDt){ $bootDt.ToString("yyyy-MM-dd HH:mm") } else { "(unknown)" }
    "Uptime (days)"      = $uptimeDays
    "CPU model"          = $proc.Name
    "Physical cores"     = $proc.NumberOfCores
    "Logical processors" = $proc.NumberOfLogicalProcessors
    "Memory total (GB)"  = [Math]::Round($cs.TotalPhysicalMemory/1GB,2)
    "Memory free (GB)"   = [Math]::Round($os.FreePhysicalMemory*1KB/1GB,2)
  }

  $systeminfoRaw = try { (cmd /c systeminfo) -join "`r`n" } catch { "" }

  $topCards = @"
  <div class='grid'>
    <div class='card'><h4>Computer</h4><p>$env:COMPUTERNAME</p></div>
    <div class='card'><h4>OS</h4><p>$($os.Caption)</p></div>
    <div class='card'><h4>Version</h4><p>$($os.Version)</p></div>
    <div class='card'><h4>Uptime (days)</h4><p>$($uptimeDays)</p></div>
    <div class='card'><h4>CPU</h4><p>$($proc.Name)</p></div>
    <div class='card'><h4>RAM (GB)</h4><p>$([Math]::Round($cs.TotalPhysicalMemory/1GB,2)) total / $([Math]::Round($os.FreePhysicalMemory*1KB/1GB,2)) free</p></div>
  </div>
"@

  $sysHtml  = $topCards
  $sysHtml += ConvertTo-NameValueTable -Object $sysSummary -Title "Overview"
  $rawBlock = ""
  if ($systeminfoRaw) { $rawBlock += Format-Preformatted -Text $systeminfoRaw -Title "Raw data (systeminfo)" }
  $sysHtml += $rawBlock

  $reportSections.Add((Add-Section -Id "system" -Title "System Info" -BodyHtml $sysHtml -Description $sectionDescriptions.system))
} catch {
  $reportSections.Add((Add-Section -Id "system" -Title "System Info" -BodyHtml (New-Alert -Text "Failed to collect system info: $($_.Exception.Message)") -Description $sectionDescriptions.system))
}

# ===================== Network =====================
try {
  $ipconfigRaw = try { (ipconfig /all) -join "`r`n" } catch { "" }
  $adapterRows = @()
  if (Test-CommandExists Get-NetAdapter) {
    $ipcfg = Get-NetIPConfiguration -All -ErrorAction SilentlyContinue
    $binds = Get-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue | Select-Object Name, Enabled
    foreach ($c in $ipcfg) {
      $ipv4 = ($c.IPv4Address | ForEach-Object { $_.IPAddress }) -join ", "
      $ipv6 = ($c.IPv6Address | ForEach-Object { $_.IPAddress }) -join ", "
      $dns  = ($c.DnsServer.ServerAddresses) -join ", "
      $gw   = ($c.IPv4DefaultGateway.NextHop, $c.IPv6DefaultGateway.NextHop | Where-Object { $_ }) -join ", "
      $bind = $binds | Where-Object Name -eq $c.InterfaceAlias
      $dhcp = $null; try { $iface = Get-NetIPInterface -InterfaceIndex $c.InterfaceIndex -AddressFamily IPv4 -ErrorAction Stop; $dhcp=$iface.Dhcp } catch {}
      $adapterRows += [PSCustomObject]@{
        Interface   = $c.InterfaceAlias
        Index       = $c.InterfaceIndex
        Description = $c.NetAdapter.Description
        Status      = $c.NetAdapter.Status
        MAC         = $c.NetAdapter.MacAddress
        IPv4        = $ipv4
        IPv6        = $ipv6
        Gateway     = $gw
        DNS         = $dns
        DHCP        = $dhcp
        IPv6Enabled = if ($bind) { $bind.Enabled } else { $null }
      }
    }
  }
  $netHtml = ""
  if ($adapterRows) { $netHtml += ConvertTo-HtmlTable -InputObject $adapterRows -Title "Network Adapters" -Properties "Interface","Index","Description","Status","MAC","IPv4","IPv6","Gateway","DNS","DHCP","IPv6Enabled" }
  $rawBlock = ""
  if ($ipconfigRaw) { $rawBlock += Format-Preformatted -Text $ipconfigRaw -Title "Raw data (ipconfig /all)" }
  if ($rawBlock){ $netHtml += $rawBlock }
  $reportSections.Add((Add-Section -Id "network" -Title "Network Configuration" -BodyHtml $netHtml -Description $sectionDescriptions.network))
} catch {
  $reportSections.Add((Add-Section -Id "network" -Title "Network Configuration" -BodyHtml (New-Alert -Text "Failed to collect network info: $($_.Exception.Message)") -Description $sectionDescriptions.network))
}

# ===================== Firewall and Ports =====================
try {
  $fwHtml = ""
  $profiles = $null
  if (Test-CommandExists Get-NetFirewallProfile) {
    $profiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, NotifyOnListen, AllowInboundRules
    if ($profiles) { $fwHtml += ConvertTo-HtmlTable -InputObject $profiles -Title "Firewall Profiles" -Properties * -TableId "fw-profiles" }
  }
  if (Test-CommandExists Get-NetFirewallRule) {
    $customRules = Get-NetFirewallRule -PolicyStore ActiveStore -ErrorAction SilentlyContinue | Where-Object { -not $_.Group -and $_.PolicyStoreSourceType -eq "PersistentStore" }
    if ($customRules) {
      $portFilters = $customRules | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue | Select-Object Name, Protocol, LocalPort, RemotePort, DynamicTarget, Program
      if ($portFilters) { $fwHtml += ConvertTo-HtmlTable -InputObject $portFilters -Title "Custom Firewall Rules (PersistentStore, no Group)" -Properties * }
    }
  }
  $netstatRaw = try { (netstat -a -n -o) -join "`r`n" } catch { "" }
  $tcpListen = @()
  if (Test-CommandExists Get-NetTCPConnection) { $tcpListen = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess }
  if ($tcpListen) { $fwHtml += ConvertTo-HtmlTable -InputObject $tcpListen -Title "Listening TCP Ports (Get-NetTCPConnection)" -Properties * }

  $rawBlock = ""
  if ($profiles) { $rawBlock += Format-Preformatted -Text (($profiles | ConvertTo-Json -Depth 4) | Out-String) -Title "Raw data (Firewall profiles JSON)" }
  if ($netstatRaw) { $rawBlock += Format-Preformatted -Text $netstatRaw -Title "Raw data (netstat -a -n -o)" }
  if ($rawBlock){ $fwHtml += $rawBlock }

  $reportSections.Add((Add-Section -Id "firewall" -Title "Firewall and Ports" -BodyHtml $fwHtml -Description $sectionDescriptions.firewall))
} catch {
  $reportSections.Add((Add-Section -Id "firewall" -Title "Firewall and Ports" -BodyHtml (New-Alert -Text "Failed to collect firewall/port info: $($_.Exception.Message)") -Description $sectionDescriptions.firewall))
}

# ===================== Storage =====================
try {
  if (Test-CommandExists Get-Volume) {
    $vols = Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.DriveType -eq "Fixed" -and $_.FileSystem } |
      Select-Object DriveLetter, Path, FileSystem, HealthStatus,
        @{n="SizeGB";e={[math]::Round($_.Size/1GB,2)}},
        @{n="FreeGB";e={[math]::Round($_.SizeRemaining/1GB,2)}},
        @{n="Free%";e={Format-Percent $_.SizeRemaining $_.Size}}
  } else {
    $vols = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue |
      Select-Object @{n="DriveLetter";e={$_.DeviceID}}, @{n="Path";e={$_.ProviderName}}, FileSystem, @{n="HealthStatus";e={"n/a"}},
        @{n="SizeGB";e={[math]::Round($_.Size/1GB,2)}},
        @{n="FreeGB";e={[math]::Round($_.FreeSpace/1GB,2)}},
        @{n="Free%";e={Format-Percent $_.FreeSpace $_.Size}}
  }
  $stHtml = ConvertTo-HtmlTable -InputObject $vols -Title "Volumes" -Properties *

  # Root listings for C:\, D:\, E:\ (only if the drive exists)
  $rootLs = ""
  foreach ($drv in "C:","D:","E:") {
    try {
      $path = "$drv\"
      if (Test-Path $path) {
        $ls = Get-ChildItem -Force -LiteralPath $path -ErrorAction SilentlyContinue |
             Select-Object Mode, LastWriteTime, Length, Name |
             Format-Table -AutoSize | Out-String
        $rootLs += (Format-Preformatted -Text $ls -Title "Root listing $path")
      }
    } catch {}
  }
  if ($rootLs) { $stHtml += $rootLs }

  $rawBlock = Format-Preformatted -Text (($vols | ConvertTo-Json -Depth 4) | Out-String) -Title "Raw data (Volumes JSON)"
  $stHtml += $rawBlock

  $reportSections.Add((Add-Section -Id "storage" -Title "Storage" -BodyHtml $stHtml -Description $sectionDescriptions.storage))
} catch {
  $reportSections.Add((Add-Section -Id "storage" -Title "Storage" -BodyHtml (New-Alert -Text "Failed to collect storage info: $($_.Exception.Message)") -Description $sectionDescriptions.storage))
}

# ===================== Applications =====================
try {
  $uninstallKeys = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
  )
  $apps = @(
    foreach ($k in $uninstallKeys) {
      if (Test-Path $k) {
        Get-ChildItem $k -ErrorAction SilentlyContinue | ForEach-Object {
          try {
            $p = Get-ItemProperty $_.PsPath -ErrorAction Stop
            if ($p.DisplayName) {
              [PSCustomObject]@{
                Name        = $p.DisplayName
                Version     = $p.DisplayVersion
                Publisher   = $p.Publisher
                InstallDate = $p.InstallDate
                Uninstall   = $p.UninstallString
                Wow6432     = ($k -like "*WOW6432Node*")
              }
            }
          } catch {}
        }
      }
    }
  )
  $apps = @($apps) | Sort-Object Name, Version
  $appHtml = ConvertTo-HtmlTable -InputObject $apps -Title "Installed Software" -Properties "Name","Version","Publisher","InstallDate","Wow6432","Uninstall"
  $appHtml += Format-Preformatted -Text (($apps | ConvertTo-Json -Depth 4) | Out-String) -Title "Raw data (Installed software JSON)"
  $reportSections.Add((Add-Section -Id "apps" -Title "Applications" -BodyHtml $appHtml -Description $sectionDescriptions.apps))
} catch {
  $reportSections.Add((Add-Section -Id "apps" -Title "Applications" -BodyHtml (New-Alert -Text "Failed to collect application list: $($_.Exception.Message)") -Description $sectionDescriptions.apps))
}

# ===================== Roles / SQL / IIS =====================
$rolesHtml = ""
try {
  if (Test-CommandExists Get-WindowsFeature) {
    $roles = Get-WindowsFeature -ErrorAction SilentlyContinue | Where-Object Installed | Select-Object Name, DisplayName, Installed
    if ($roles) { $rolesHtml += ConvertTo-HtmlTable -InputObject $roles -Title "Installed Roles and Features" -Properties "Name","DisplayName","Installed" }
  } else {
    $rolesHtml += New-Alert -Text "Get-WindowsFeature is not available. Perhaps this is no Windows Server installation." -Type warn
  }
} catch { $rolesHtml += New-Alert -Text "Failed to collect roles: $($_.Exception.Message)" }

function Get-SqlInstanceNames {
  $instances = @()
  try {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"
    if (Test-Path $regPath) {
      $props = Get-ItemProperty $regPath
      foreach ($name in ($props.PSObject.Properties | Where-Object { $_.MemberType -eq "NoteProperty" }).Name) {
        $instances += if ($name -eq "MSSQLSERVER") { $env:COMPUTERNAME } else { "$env:COMPUTERNAME\$name" }
      }
    }
  } catch {}
  if (-not $instances) {
    Get-Service -Name "MSSQL*" -ErrorAction SilentlyContinue | ForEach-Object {
      if ($_.Name -eq "MSSQLSERVER") { $instances += $env:COMPUTERNAME } else { $instances += "$env:COMPUTERNAME$([System.IO.Path]::DirectorySeparatorChar)$($_.Name -replace ""^MSSQL\\$"","""")" }
    }
  }
  $instances | Select-Object -Unique
}

function Get-SqlDbMdfInfo {
  $out = @()
  $instances = Get-SqlInstanceNames
  if (-not $instances) { return $out }
  $smoLoaded = $false
  foreach ($asm in "Microsoft.SqlServer.Smo","Microsoft.SqlServer.ConnectionInfo","Microsoft.SqlServer.SmoExtended","Microsoft.SqlServer.Management.Sdk.Sfc") {
    try { Add-Type -AssemblyName $asm -ErrorAction Stop; $smoLoaded = $true } catch {}
  }
  foreach ($inst in $instances) {
    if ($smoLoaded) {
      try {
        $srv = New-Object Microsoft.SqlServer.Management.Smo.Server $inst
        foreach ($db in $srv.Databases) {
          try {
            $files = $db.EnumFiles() | Where-Object { $_.FileName -match "\.mdf$" }
            foreach ($f in $files) { $out += [PSCustomObject]@{ Instance=$inst; Database=$db.Name; MdfPath=$f.FileName } }
          } catch { $out += [PSCustomObject]@{ Instance=$inst; Database=$db.Name; MdfPath="(could not read MDF path)" } }
        }
        continue
      } catch {}
    }
    $sqlcmd = Get-Command sqlcmd.exe -ErrorAction SilentlyContinue
    if ($sqlcmd) {
      $query = "set nocount on; select DB_NAME(database_id) as DBName, physical_name from sys.master_files where type_desc='ROWS' and physical_name like '%.mdf' order by 1,2;"
      try {
        $raw = & $sqlcmd.Source -S $inst -E -Q $query -h -1 -W -s '|' 2>$null
        foreach ($line in $raw) {
          if ($line -match "\|") {
            $parts = $line -split "\|"
            $dbn=$parts[0].Trim(); $path=$parts[1].Trim()
            if ($dbn -and $path -and $path -match "\.mdf$") { $out += [PSCustomObject]@{ Instance=$inst; Database=$dbn; MdfPath=$path } }
          }
        }
      } catch {}
    }
  }
  return $out
}

try {
  $sqlData = Get-SqlDbMdfInfo
  if ($sqlData -and $sqlData.Count -gt 0) {
    $rolesHtml += ConvertTo-HtmlTable -InputObject $sqlData -Title "SQL Server Databases (.MDF)" -Properties "Instance","Database","MdfPath"
    $rolesHtml += Format-Preformatted -Text (($sqlData | ConvertTo-Json -Depth 4) | Out-String) -Title "Raw data (SQL .MDF JSON)"
  } else {
    $rolesHtml += New-Alert -Text "SQL Server is not installed on this server." -Type error
  }
} catch { $rolesHtml += New-Alert -Text "SQL detection error: $($_.Exception.Message)" }

try {
  $iisInstalled = $false
  if (Test-CommandExists Get-WindowsFeature) { $feat = Get-WindowsFeature -Name Web-Server -ErrorAction SilentlyContinue; $iisInstalled = [bool]($feat -and $feat.Installed) }
  if ($iisInstalled) {
    Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null
    $sites = Get-Website -ErrorAction SilentlyContinue

    $siteBindRows = @()
    foreach ($s in $sites) {
      $bindings = Get-WebBinding -Name $s.Name -ErrorAction SilentlyContinue
      foreach ($b in $bindings) {
        $proto = $b.protocol
        $info  = $b.bindingInformation
        $ip,$port,$hostHeader = $info -split ":"
        $siteBindRows += [PSCustomObject]@{
          Site         = $s.Name
          State        = $s.State
          AppPool      = $s.applicationPool
          Protocol     = $proto
          IP           = $ip
          Port         = $port
          HostHeader   = $hostHeader
          PhysicalRoot = $s.physicalPath
        }
      }
    }
    if ($siteBindRows) {
      $rolesHtml += ConvertTo-HtmlTable -InputObject $siteBindRows -Title "IIS Sites and Bindings" -Properties "Site","State","AppPool","Protocol","IP","Port","HostHeader","PhysicalRoot"
      $rolesHtml += Format-Preformatted -Text (($siteBindRows | ConvertTo-Json -Depth 4) | Out-String) -Title "Raw data (IIS Sites/Bindings JSON)"
    }

    $appRows = @()
    foreach ($s in $sites) {
      $apps = Get-WebApplication -Site $s.Name -ErrorAction SilentlyContinue
      foreach ($a in $apps) {
        $appRows += [PSCustomObject]@{
          Site         = $s.Name
          Application  = ($a.Path.TrimStart("/"))
          AppPool      = $a.ApplicationPool
          PhysicalPath = $a.PhysicalPath
        }
      }
    }
    if ($appRows) {
      $rolesHtml += ConvertTo-HtmlTable -InputObject $appRows -Title "IIS Applications" -Properties "Site","Application","AppPool","PhysicalPath"
      $rolesHtml += Format-Preformatted -Text (($appRows | ConvertTo-Json -Depth 4) | Out-String) -Title "Raw data (IIS Applications JSON)"
    }
  } else {
    $rolesHtml += New-Alert -Text "IIS (Web-Server) is not installed on this server." -Type error
  }
} catch { $rolesHtml += New-Alert -Text "IIS information error: $($_.Exception.Message)" }

$reportSections.Add((Add-Section -Id "roles" -Title "Server Roles / SQL / IIS" -BodyHtml $rolesHtml -Description $sectionDescriptions.roles))

# ===================== Services =====================
try {
  $svcs = Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | Select-Object Name, DisplayName, State, StartMode, StartName, PathName
  $svcHtml = ConvertTo-HtmlTable -InputObject $svcs -Title "All Services" -Properties "Name","DisplayName","State","StartMode","StartName","PathName"
  $svcHtml += Format-Preformatted -Text (($svcs | ConvertTo-Json -Depth 4) | Out-String) -Title "Raw data (Services JSON)"
  $reportSections.Add((Add-Section -Id "services" -Title "Services" -BodyHtml $svcHtml -Description $sectionDescriptions.services))
} catch {
  $reportSections.Add((Add-Section -Id "services" -Title "Services" -BodyHtml (New-Alert -Text "Failed to collect services: $($_.Exception.Message)") -Description $sectionDescriptions.services))
}

# ===================== Shares =====================
try {
  $sharesHtml = ""
  $shares = $sp = $ntfs = $null
  if (Test-CommandExists Get-SmbShare) {
    $shares = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { -not $_.Special }
    if ($shares) {
      $sharesHtml += ConvertTo-HtmlTable -InputObject $shares -Title "Shares (non-administrative)" -Properties "Name","Path","Description","CachingMode","EncryptData"
      $sp = foreach ($s in $shares) { Get-SmbShareAccess -Name $s.Name -ErrorAction SilentlyContinue | Select-Object @{n="Share";e={$s.Name}}, AccountName, AccessControlType, AccessRight }
      if ($sp) { $sharesHtml += ConvertTo-HtmlTable -InputObject $sp -Title "Share Permissions" -Properties "Share","AccountName","AccessControlType","AccessRight" }
      $ntfs = foreach ($s in $shares) {
        try { $acl = Get-Acl -Path $s.Path -ErrorAction Stop } catch { $acl = $null }
        if ($acl) {
          foreach ($ace in $acl.Access) { [PSCustomObject]@{ Path=$s.Path; Identity=$ace.IdentityReference; Rights=$ace.FileSystemRights; Inherited=$ace.IsInherited; Type=$ace.AccessControlType } }
        } else {
          [PSCustomObject]@{ Path=$s.Path; Identity="(no access)"; Rights="n/a"; Inherited="n/a"; Type="n/a" }
        }
      }
      if ($ntfs) { $sharesHtml += ConvertTo-HtmlTable -InputObject $ntfs -Title "NTFS Permissions" -Properties "Path","Identity","Rights","Inherited","Type" }
    } else {
      $sharesHtml += New-Alert -Text "No non-administrative shares found." -Type info
    }
  } else {
    $sharesHtml += New-Alert -Text "Get-SmbShare is not available on this system." -Type warn
  }
  $rawText = ""
  if ($shares) { $rawText += (($shares | ConvertTo-Json -Depth 4) | Out-String) + "`r`n" }
  if ($sp)     { $rawText += (($sp     | ConvertTo-Json -Depth 4) | Out-String) + "`r`n" }
  if ($ntfs)   { $rawText += (($ntfs   | ConvertTo-Json -Depth 4) | Out-String) }
  if ($rawText){ $sharesHtml += Format-Preformatted -Text $rawText -Title "Raw data (Shares JSON)" }

  $reportSections.Add((Add-Section -Id "shares" -Title "Shares" -BodyHtml $sharesHtml -Description $sectionDescriptions.shares))
} catch {
  $reportSections.Add((Add-Section -Id "shares" -Title "Shares" -BodyHtml (New-Alert -Text "Failed to collect share information: $($_.Exception.Message)") -Description $sectionDescriptions.shares))
}

# ===================== Printers =====================
try {
  $prtHtml = ""
  $rows = $null
  if (Test-CommandExists Get-Printer) {
    $printers = Get-Printer -ErrorAction SilentlyContinue
    $ports    = Get-PrinterPort -ErrorAction SilentlyContinue | Select-Object Name, PrinterHostAddress
    $drivers  = Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object Name, Manufacturer
    $rows = foreach ($p in $printers) {
      $port = $ports | Where-Object Name -eq $p.PortName | Select-Object -First 1
      $ip = $null
      if ($port -and $port.PrinterHostAddress) { $ip = $port.PrinterHostAddress }
      elseif ($p.PortName -match "^IP_(\d+\.\d+\.\d+\.\d+)$") { $ip = $Matches[1] }
      $drv = $drivers | Where-Object Name -eq $p.DriverName | Select-Object -First 1
      [PSCustomObject]@{ Name=$p.Name; Driver=$p.DriverName; DriverVendor=$drv.Manufacturer; Port=$p.PortName; IPAddress=$ip; Shared=$p.Shared; ShareName=$p.ShareName }
    }
    if ($rows) { $prtHtml += ConvertTo-HtmlTable -InputObject $rows -Title "Printers" -Properties "Name","Driver","DriverVendor","Port","IPAddress","Shared","ShareName" }
  } else {
    $prtHtml += New-Alert -Text "Printer cmdlets not available (PrintManagement module missing?)." -Type warn
  }
  if ($rows){ $prtHtml += Format-Preformatted -Text (($rows | ConvertTo-Json -Depth 4) | Out-String) -Title "Raw data (Printers JSON)" }
  $reportSections.Add((Add-Section -Id "printers" -Title "Printers" -BodyHtml $prtHtml -Description $sectionDescriptions.printers))
} catch {
  $reportSections.Add((Add-Section -Id "printers" -Title "Printers" -BodyHtml (New-Alert -Text "Failed to collect printer info: $($_.Exception.Message)") -Description $sectionDescriptions.printers))
}

# ===================== CSS =====================
$css = @"
:root{--hdrH:64px}

*{box-sizing:border-box}
html{font-family:Segoe UI,Arial;line-height:1.35}
body{margin:0;background:#F2F2F2;color:#111827}

/* Header: gradient from #8EAFDA to white, centered content */
header{
  position:sticky; top:0; z-index:30;
  background:linear-gradient(180deg,#8EAFDA 0%, #FFFFFF 100%);
  padding:16px 24px;
  color:#0b1220;
  display:flex; flex-direction:column; align-items:center; text-align:center; gap:6px;
}
header h1{margin:0;font-size:20px;color:#0b1220}
header .meta{opacity:.9;font-size:12px;color:#0b1220}

/* Logo */
header .brand{display:inline-block; line-height:0}
header .brand img{width:35px;height:35px;border-radius:6px}
header .brand:focus{outline:2px solid rgba(0,0,0,.2); outline-offset:3px}

/* Tabs under header; centered */
nav.tabs{
  position:sticky; top:var(--hdrH); z-index:25;
  display:flex; flex-wrap:wrap; gap:6px; align-content:flex-start; justify-content:center;
  padding:10px 12px; background:#ffffff; border-bottom:1px solid #e5e7eb;
  overflow:visible!important; max-height:none
}
nav.tabs a{padding:8px 12px;border-radius:10px;background:#f3f4f6;color:#111827;text-decoration:none;font-size:13px;transition:.15s}
nav.tabs a:hover{background:#e5e7eb}
nav.tabs a.active{background:#8EAFDA;color:#0b1220;box-shadow:0 0 0 1px rgba(0,0,0,.05) inset}

/* Content */
main{padding:18px}

/* Section cards */
.section{background:#ffffff;border:1px solid #e5e7eb;border-radius:14px;padding:16px;margin-bottom:16px;box-shadow:0 1px 0 rgba(0,0,0,.03) inset}
.section h2{margin-top:0;font-size:18px;color:#8EAFDA}
.desc{margin:8px 0 14px;font-size:13px;color:#334155;background:#f8fafc;border:1px solid #e5e7eb;border-radius:10px;padding:10px 12px}

/* Alerts */
.alert{display:flex;gap:8px;align-items:flex-start;border-radius:10px;padding:10px 12px;margin:8px 0}
.alert .ico{font-size:12px}
.alert.error{background:#fee2e2;color:#7f1d1d}
.alert.warn{background:#fef3c7;color:#78350f}
.alert.info{background:#e0f2fe;color:#1e3a8a}
.alert.ok{background:#dcfce7;color:#065f46}

/* Preformatted (raw blocks have a light blue tint) */
pre{background:#f8fafc;border:1px solid #e5e7eb;border-radius:10px;padding:12px;overflow:auto;max-height:15em;white-space:pre;color:#111827}
pre.raw{background:#EAF2FF;border-color:#CFE0FF}

/* Tables */
.tablewrap{overflow:auto}
table{border-collapse:collapse;width:100%;margin:8px 0;background:#ffffff}
th,td{border-bottom:1px solid #e5e7eb;padding:8px 10px;text-align:left;color:#111827}
th{position:sticky;top:0;background:#f1f5f9}
tr:hover{background:#f9fafb}
.compact th,.compact td{font-size:12px}

/* Cards */
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:8px}
.card{background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;padding:12px}
.card h4{margin:0 0 6px 0;font-size:12px;color:#475569;text-transform:uppercase;letter-spacing:.06em}
.card p{margin:0;font-size:14px;color:#111827}

/* Tab visibility */
.tab-content{display:none}
.tab-content.active{display:block}

/* Disabled firewall profile rows */
tr.danger{background:#fee2e2 !important;color:#991b1b !important}

footer{opacity:.7;font-size:12px;padding:12px 18px;color:#334155}
"@

# ===================== JS =====================
$js = @"
(function(){
  function setHeaderHeightVar(){
    var hdr = document.querySelector('header');
    if(!hdr) return;
    var h = Math.round(hdr.getBoundingClientRect().height);
    document.documentElement.style.setProperty('--hdrH', h + 'px');
  }

  var tabs = document.querySelectorAll('nav.tabs a');
  var secs = document.querySelectorAll('.tab-content');

  function highlightFirewallRows(){
    var tbl = document.getElementById('fw-profiles');
    if(!tbl) return;
    var rows = tbl.querySelectorAll('tr');
    if(rows.length < 2) return;

    var header = rows[0].querySelectorAll('th');
    var enabledIdx = -1;
    for (var i=0;i<header.length;i++){
      if(header[i].textContent.trim().toLowerCase() === 'enabled'){ enabledIdx = i; break; }
    }
    if(enabledIdx === -1) return;

    for (var r=1;r<rows.length;r++){
      var cells = rows[r].children;
      var val = (cells[enabledIdx]?.textContent || '').trim().toLowerCase();
      var disabled = (val === 'false' || val === '0' || val === 'no');
      rows[r].classList.toggle('danger', disabled);
    }
  }

  function activate(id){
    for (var i=0;i<secs.length;i++){ secs[i].classList.toggle('active', secs[i].id===id); }
    for (var j=0;j<tabs.length;j++){ tabs[j].classList.toggle('active', (tabs[j].getAttribute('href')==='#'+id)); }
    try{ history.replaceState(null,'','#'+id); }catch(e){}
    highlightFirewallRows();
  }

  for (var k=0;k<tabs.length;k++){
    tabs[k].addEventListener('click', function(e){ e.preventDefault(); activate(this.getAttribute('href').substring(1)); });
  }

  window.addEventListener('load', setHeaderHeightVar);
  window.addEventListener('resize', setHeaderHeightVar);
  setHeaderHeightVar();

  var wanted = (location.hash ? location.hash.substring(1) : 'system');
  if (!document.getElementById(wanted)) { wanted = 'system'; }
  activate(wanted);
  highlightFirewallRows();
})();
"@

# ===================== NAV + HTML =====================
$idsAndTitles = @(
  @{Id="system";Title="System Info"},
  @{Id="network";Title="Network"},
  @{Id="firewall";Title="Firewall and Ports"},
  @{Id="storage";Title="Storage"},
  @{Id="apps";Title="Applications"},
  @{Id="roles";Title="Server Roles / SQL / IIS"},
  @{Id="services";Title="Services"},
  @{Id="shares";Title="Shares"},
  @{Id="printers";Title="Printers"}
)
$nav = @()
foreach($it in $idsAndTitles){ $nav += ('<a id="tab-{0}" class="tab-link" href="#{0}">{1}</a>' -f $it.Id, $it.Title) }

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Server Inventory | $env:COMPUTERNAME</title>
  <style>$css</style>
</head>
<body>
  <header>
    <a class="brand" href="https://justinverstijnen.nl/" target="_blank" rel="noopener">
      <img src="https://justinverstijnen.nl/wp-content/uploads/2025/04/cropped-Logo-2.0-Transparant.png" alt="Logo" width="35" height="35" />
    </a>
    <h1>Server Inventory  $env:COMPUTERNAME</h1>
    <div class="meta">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")  User: $env:USERNAME  Domain: $env:USERDOMAIN</div>
  </header>
  <nav class="tabs">
    $($nav -join "`n")
  </nav>
  <main>
    $($reportSections -join "`n")
  </main>
  <footer><a href="https://github.com/JustinVerstijnen/JV-ServerInventoryReport/tree/main" target="_blank" rel="noopener">Report generated by JV-ServerInventoryReport.ps1</a></footer>
  <script>$js</script>
</body>
</html>
"@

try {
  $null = New-Item -Path (Split-Path $OutputPath) -ItemType Directory -Force -ErrorAction SilentlyContinue
  $html | Out-File -FilePath $OutputPath -Encoding UTF8
  Write-Host "Report written to: $OutputPath" -ForegroundColor Cyan
} catch {
  Write-Warning "Could not write report: $($_.Exception.Message)"
}
