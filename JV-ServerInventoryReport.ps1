<#
SYNOPSIS
  Genereert een HTML-inventarisatierapport met tabbladen voor (onbekende) Windows Servers.
DESCRIPTION
  Verzamelt systeem-, netwerk-, firewall-, storage-, applicatie-, rollen-, SQL-, IIS-, services-, shares- en printerinfo
  en schrijft een modern, stand-alone HTML-rapport met tabs. Alleen ingebouwde cmdlets + waar nodig WMI/klassieke tools.
PARAMETER OutputPath
  Volledig pad voor het HTML-rapport. Zonder parameter: Desktop\Server-Inventory_<COMPUTERNAME>_<timestamp>.html
NOTES
  PowerShell 5.1+. Sommige secties vereisen rollen/modules (IIS/Print/SQL). Fouten worden afgehandeld en in het rapport getoond.
#>

[CmdletBinding()]
param([string]$OutputPath)

#region === Helpers ===
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
  param([Parameter(Mandatory)][object]$InputObject,[string[]]$Properties,[string]$Title)
  try {
    $pre  = $( if ($Title) { "<h3>$Title</h3>" } else { $null } )
    $frag = @($InputObject) | Select-Object -Property $Properties | ConvertTo-Html -Fragment -PreContent $pre
    ($frag -join "`n") -replace "<table>", "<table class=""compact"">"
  } catch {
    New-Alert -Text "Kon tabel '$Title' niet renderen: $($_.Exception.Message)" -Type error
  }
}

function ConvertTo-NameValueTable {
  [CmdletBinding()]
  param([Parameter(Mandatory)][object]$Object,[string[]]$Properties,[string]$Title="Overzicht")
  $props = if ($Properties) { $Properties } else { $Object.PSObject.Properties.Name }
  $rows = foreach ($p in $props) { [PSCustomObject]@{ Naam=$p; Waarde=(($Object.$p | Out-String).Trim()) } }
  ConvertTo-HtmlTable -InputObject $rows -Title $Title -Properties "Naam","Waarde"
}

function Format-Preformatted {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Text,[string]$Title)
  $enc = [System.Net.WebUtility]::HtmlEncode($Text)
  $pre = $( if ($Title) { "<h3>$Title</h3>" } else { "" } )
  "$pre<pre>$enc</pre>"
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

# Robuuste conversie: accepteert DateTime en DMTF/string
function ConvertTo-DateTimeSafe {
  [CmdletBinding()] param([Parameter(Mandatory)][object]$Value)
  if ($null -eq $Value) { return $null }
  if ($Value -is [datetime]) { return $Value }
  $s = [string]$Value
  # DMTF formaat?
  if ($s -match "^\d{14}\.\d{6}(\+|-)\d{3}$") {
    try { return [Management.ManagementDateTimeConverter]::ToDateTime($s) } catch {}
  }
  # Probeer als normale datum
  try { return [DateTime]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture) } catch {
    try { return [DateTime]::Parse($s) } catch { return $null }
  }
}
#endregion Helpers

#region === Output target ===
if (-not $OutputPath) {
  $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
  $desktop = [Environment]::GetFolderPath("Desktop")
  $OutputPath = Join-Path $desktop "Server-Inventory_$env:COMPUTERNAME_$stamp.html"
}
$reportSections = New-Object System.Collections.Generic.List[string]
#endregion

#region === Omschrijvingen per tab ===
$sectionDescriptions = @{
  system   = "Samenvatting van systeemkenmerken (OS, CPU, geheugen) en ruwe systeminfo uitvoer."
  network  = "Netwerkconfiguratie van adapters, inclusief IPs, gateways, DNS en ipconfig uitvoer."
  firewall = "Firewall profielen, aangepaste regels en luisterende TCP poorten."
  storage  = "Overzicht van vaste volumes: formaat, totale capaciteit en vrije ruimte."
  apps     = "Geinstalleerde software uit Uninstall registratiesleutels (64/32 bit en per gebruiker)."
  roles    = "Geinstalleerde rollen en features, SQL (.MDF) en IIS (sites, bindings, applications)."
  services = "Alle Windows services met status, starttype, account en pad."
  shares   = "SMB shares, share permissies en NTFS ACLs (excl. administratieve shares)."
  printers = "Geinstalleerde printers, poorten, drivers en indien bekend IP adressen."
}
#endregion

#region === System Info ===
try {
  $cs   = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
  $os   = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
  $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
  $proc = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1

  $installDt = ConvertTo-DateTimeSafe $os.InstallDate
  $bootDt    = ConvertTo-DateTimeSafe $os.LastBootUpTime
  $uptimeDays = if($bootDt){ [Math]::Round((New-TimeSpan -Start $bootDt -End (Get-Date)).TotalDays,1) } else { $null }

  $sysSummary = [PSCustomObject]@{
    ComputerName   = $env:COMPUTERNAME
    Domain         = $cs.Domain
    Manufacturer   = $cs.Manufacturer
    Model          = $cs.Model
    SerialNumber   = ($bios.SerialNumber | Out-String).Trim()
    OSName         = $os.Caption
    OSVersion      = $os.Version
    InstallDate    = if($installDt){ $installDt.ToString("yyyy-MM-dd HH:mm") } else { "(onbekend)" }
    LastBoot       = if($bootDt){ $bootDt.ToString("yyyy-MM-dd HH:mm") } else { "(onbekend)" }
    UptimeDays     = if($uptimeDays){ $uptimeDays } else { "(onbekend)" }
    CPU            = $proc.Name
    Cores          = $proc.NumberOfCores
    LogicalCPU     = $proc.NumberOfLogicalProcessors
    MemoryGB_Total = [Math]::Round($cs.TotalPhysicalMemory/1GB,2)
    MemoryGB_Free  = [Math]::Round($os.FreePhysicalMemory*1KB/1GB,2)
  }

  $systeminfoRaw = try { (cmd /c systeminfo) -join "`r`n" } catch { "" }

  $topCards = @"
  <div class='grid'>
    <div class='card'><h4>Computer</h4><p>$env:COMPUTERNAME</p></div>
    <div class='card'><h4>OS</h4><p>$($os.Caption)</p></div>
    <div class='card'><h4>Versie</h4><p>$($os.Version)</p></div>
    <div class='card'><h4>Uptime (dagen)</h4><p>$($sysSummary.UptimeDays)</p></div>
    <div class='card'><h4>CPU</h4><p>$($proc.Name)</p></div>
    <div class='card'><h4>RAM (GB)</h4><p>$([Math]::Round($cs.TotalPhysicalMemory/1GB,2)) totaal / $([Math]::Round($os.FreePhysicalMemory*1KB/1GB,2)) vrij</p></div>
  </div>
"@

  $sysHtml  = $topCards
  $sysHtml += ConvertTo-NameValueTable -Object $sysSummary -Title "Overzicht"
  if ($systeminfoRaw) { $sysHtml += Format-Preformatted -Text $systeminfoRaw -Title "systeminfo (ruwe output)" }

  $reportSections.Add((Add-Section -Id "system" -Title "System Info" -BodyHtml $sysHtml -Description $sectionDescriptions.system))
} catch {
  $reportSections.Add((Add-Section -Id "system" -Title "System Info" -BodyHtml (New-Alert -Text "Kon systeeminfo niet ophalen: $($_.Exception.Message)") -Description $sectionDescriptions.system))
}
#endregion

#region === Netwerk ===
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
      $dhcp = $null
      try { $iface = Get-NetIPInterface -InterfaceIndex $c.InterfaceIndex -AddressFamily IPv4 -ErrorAction Stop; $dhcp=$iface.Dhcp } catch {}
      $adapterRows += [PSCustomObject]@{
        Interface   = $c.InterfaceAlias; Index=$c.InterfaceIndex; Description=$c.NetAdapter.Description
        Status      = $c.NetAdapter.Status; MAC=$c.NetAdapter.MacAddress; IPv4=$ipv4; IPv6=$ipv6
        Gateway=$gw; DNS=$dns; DHCP=$dhcp; IPv6Enabled= if ($bind) { $bind.Enabled } else { $null }
      }
    }
  }
  $netHtml = ""
  if ($adapterRows) { $netHtml += ConvertTo-HtmlTable -InputObject $adapterRows -Title "Netwerkadapters" -Properties "Interface","Index","Description","Status","MAC","IPv4","IPv6","Gateway","DNS","DHCP","IPv6Enabled" }
  if ($ipconfigRaw) { $netHtml += Format-Preformatted -Text $ipconfigRaw -Title "ipconfig /all (ruwe output)" }
  $reportSections.Add((Add-Section -Id "network" -Title "Netwerkconfiguratie" -BodyHtml $netHtml -Description $sectionDescriptions.network))
} catch {
  $reportSections.Add((Add-Section -Id "network" -Title "Netwerkconfiguratie" -BodyHtml (New-Alert -Text "Kon netwerkinfo niet ophalen: $($_.Exception.Message)") -Description $sectionDescriptions.network))
}
#endregion

#region === Firewall en Poorten ===
try {
  $fwHtml = ""
  if (Test-CommandExists Get-NetFirewallProfile) {
    $profiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, NotifyOnListen, AllowInboundRules
    if ($profiles) { $fwHtml += ConvertTo-HtmlTable -InputObject $profiles -Title "Firewall-profielen" -Properties * }
  }
  if (Test-CommandExists Get-NetFirewallRule) {
    $customRules = Get-NetFirewallRule -PolicyStore ActiveStore -ErrorAction SilentlyContinue | Where-Object { -not $_.Group -and $_.PolicyStoreSourceType -eq "PersistentStore" }
    if ($customRules) {
      $portFilters = $customRules | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue | Select-Object Name, Protocol, LocalPort, RemotePort, DynamicTarget, Program
      if ($portFilters) { $fwHtml += ConvertTo-HtmlTable -InputObject $portFilters -Title "Aangepaste firewallregels (PersistentStore, zonder Group)" -Properties * }
    }
  }
  $netstatRaw = try { (netstat -a -n -o) -join "`r`n" } catch { "" }
  $tcpListen = @()
  if (Test-CommandExists Get-NetTCPConnection) {
    $tcpListen = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Select-Object LocalAddress, LocalPort, OwningProcess
  }
  if ($tcpListen) { $fwHtml += ConvertTo-HtmlTable -InputObject $tcpListen -Title "Luisterende TCP-poorten (Get-NetTCPConnection)" -Properties * }
  if ($netstatRaw) { $fwHtml += Format-Preformatted -Text $netstatRaw -Title "netstat -a -n -o (ruwe output)" }
  $reportSections.Add((Add-Section -Id "firewall" -Title "Firewall en Poorten" -BodyHtml $fwHtml -Description $sectionDescriptions.firewall))
} catch {
  $reportSections.Add((Add-Section -Id "firewall" -Title "Firewall en Poorten" -BodyHtml (New-Alert -Text "Kon firewall/poorten niet ophalen: $($_.Exception.Message)") -Description $sectionDescriptions.firewall))
}
#endregion

#region === Storage ===
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
  $reportSections.Add((Add-Section -Id "storage" -Title "Storage" -BodyHtml $stHtml -Description $sectionDescriptions.storage))
} catch {
  $reportSections.Add((Add-Section -Id "storage" -Title "Storage" -BodyHtml (New-Alert -Text "Kon storage-informatie niet ophalen: $($_.Exception.Message)") -Description $sectionDescriptions.storage))
}
#endregion

#region === Applications (geinstalleerde software) ===
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
                UninstallString = $p.UninstallString
                Wow6432     = ($k -like "*WOW6432Node*")
              }
            }
          } catch {}
        }
      }
    }
  )
  $apps = @($apps) | Sort-Object Name, Version
  $appHtml = ConvertTo-HtmlTable -InputObject $apps -Title "Geinstalleerde software" -Properties "Name","Version","Publisher","InstallDate","Wow6432","UninstallString"
  $reportSections.Add((Add-Section -Id "apps" -Title "Applications" -BodyHtml $appHtml -Description $sectionDescriptions.apps))
} catch {
  $reportSections.Add((Add-Section -Id "apps" -Title "Applications" -BodyHtml (New-Alert -Text "Kon applicatielijst niet ophalen: $($_.Exception.Message)") -Description $sectionDescriptions.apps))
}
#endregion

#region === Server Roles, SQL, IIS ===
$rolesHtml = ""
# Rollen
try {
  if (Test-CommandExists Get-WindowsFeature) {
    $roles = Get-WindowsFeature -ErrorAction SilentlyContinue | Where-Object Installed | Select-Object Name, DisplayName, Installed
    if ($roles) { $rolesHtml += ConvertTo-HtmlTable -InputObject $roles -Title "Geinstalleerde rollen en features" -Properties "Name","DisplayName","Installed" }
  } else {
    $rolesHtml += New-Alert -Text "Get-WindowsFeature niet beschikbaar. Serverrollen kunnen niet bepaald worden." -Type warn
  }
} catch { $rolesHtml += New-Alert -Text "Fout bij ophalen rollen: $($_.Exception.Message)" }

# SQL: alleen MDF per database
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
      if ($_.Name -eq "MSSQLSERVER") { $instances += $env:COMPUTERNAME } else { $instances += "$env:COMPUTERNAME\$($_.Name -replace '^MSSQL\$','')" }
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
          } catch { $out += [PSCustomObject]@{ Instance=$inst; Database=$db.Name; MdfPath="(kon MDF-pad niet ophalen)" } }
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
          if ($line -match '\|') {
            $parts = $line -split '\|'
            $dbn=$parts[0].Trim(); $path=$parts[1].Trim()
            if ($dbn -and $path -and $path -match '\.mdf$') { $out += [PSCustomObject]@{ Instance=$inst; Database=$dbn; MdfPath=$path } }
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
    $rolesHtml += ConvertTo-HtmlTable -InputObject $sqlData -Title "SQL Server databases (.MDF)" -Properties "Instance","Database","MdfPath"
  } else {
    $rolesHtml += New-Alert -Text "SQL Server lijkt niet aanwezig of toegankelijk op deze host." -Type error
  }
} catch { $rolesHtml += New-Alert -Text "Fout bij SQL-detectie: $($_.Exception.Message)" }

# IIS (sites + bindings + applications)
try {
  $iisInstalled = $false
  if (Test-CommandExists Get-WindowsFeature) { $feat = Get-WindowsFeature -Name Web-Server -ErrorAction SilentlyContinue; $iisInstalled = [bool]($feat -and $feat.Installed) }
  if ($iisInstalled) {
    Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null
    $sites = Get-Website -ErrorAction SilentlyContinue

    # Sites + bindings
    $siteBindRows = @()
    foreach ($s in $sites) {
      $bindings = Get-WebBinding -Name $s.Name -ErrorAction SilentlyContinue
      foreach ($b in $bindings) {
        $proto = $b.protocol
        $info  = $b.bindingInformation   # ip:port:host
        $ip,$port,$hostHeader = $info -split ':'
        $siteBindRows += [PSCustomObject]@{
          Site        = $s.Name
          State       = $s.State
          AppPool     = $s.applicationPool
          Protocol    = $proto
          IP          = $ip
          Port        = $port
          HostHeader  = $hostHeader
          PhysicalRoot= $s.physicalPath
        }
      }
    }
    if ($siteBindRows) {
      $rolesHtml += ConvertTo-HtmlTable -InputObject $siteBindRows -Title "IIS Sites en bindings" -Properties "Site","State","AppPool","Protocol","IP","Port","HostHeader","PhysicalRoot"
    }

    # Applications per site
    $appRows = @()
    foreach ($s in $sites) {
      $apps = Get-WebApplication -Site $s.Name -ErrorAction SilentlyContinue
      foreach ($a in $apps) {
        $appRows += [PSCustomObject]@{
          Site        = $s.Name
          Application = ($a.Path.TrimStart("/"))
          AppPool     = $a.ApplicationPool
          PhysicalPath= $a.PhysicalPath
        }
      }
    }
    if ($appRows) { $rolesHtml += ConvertTo-HtmlTable -InputObject $appRows -Title "IIS Applications" -Properties "Site","Application","AppPool","PhysicalPath" }

  } else {
    $rolesHtml += New-Alert -Text "IIS (Web-Server) is niet geinstalleerd op deze server." -Type error
  }
} catch { $rolesHtml += New-Alert -Text "Fout bij IIS-informatie: $($_.Exception.Message)" }

$reportSections.Add((Add-Section -Id "roles" -Title "Server Roles / SQL / IIS" -BodyHtml $rolesHtml -Description $sectionDescriptions.roles))
#endregion

#region === Services ===
try {
  $svcs = Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | Select-Object Name, DisplayName, State, StartMode, StartName, PathName
  $svcHtml = ConvertTo-HtmlTable -InputObject $svcs -Title "Alle services" -Properties "Name","DisplayName","State","StartMode","StartName","PathName"
  $reportSections.Add((Add-Section -Id "services" -Title "Services" -BodyHtml $svcHtml -Description $sectionDescriptions.services))
} catch {
  $reportSections.Add((Add-Section -Id "services" -Title "Services" -BodyHtml (New-Alert -Text "Kon services niet ophalen: $($_.Exception.Message)") -Description $sectionDescriptions.services))
}
#endregion

#region === Shares ===
try {
  $sharesHtml = ""
  if (Test-CommandExists Get-SmbShare) {
    $shares = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { -not $_.Special }
    if ($shares) {
      $sharesHtml += ConvertTo-HtmlTable -InputObject $shares -Title "Shares (excl. administratieve shares)" -Properties "Name","Path","Description","CachingMode","EncryptData"
      $sp = foreach ($s in $shares) { Get-SmbShareAccess -Name $s.Name -ErrorAction SilentlyContinue | Select-Object @{n="Share";e={$s.Name}}, AccountName, AccessControlType, AccessRight }
      if ($sp) { $sharesHtml += ConvertTo-HtmlTable -InputObject $sp -Title "Share-permissies" -Properties "Share","AccountName","AccessControlType","AccessRight" }
      $ntfs = foreach ($s in $shares) {
        try { $acl = Get-Acl -Path $s.Path -ErrorAction Stop } catch { $acl = $null }
        if ($acl) {
          foreach ($ace in $acl.Access) {
            [PSCustomObject]@{ Path=$s.Path; Identity=$ace.IdentityReference; Rights=$ace.FileSystemRights; Inherited=$ace.IsInherited; Type=$ace.AccessControlType }
          }
        } else {
          [PSCustomObject]@{ Path=$s.Path; Identity="(geen toegang)"; Rights="n/a"; Inherited="n/a"; Type="n/a" }
        }
      }
      if ($ntfs) { $sharesHtml += ConvertTo-HtmlTable -InputObject $ntfs -Title "NTFS-permissies" -Properties "Path","Identity","Rights","Inherited","Type" }
    } else {
      $sharesHtml += New-Alert -Text "Geen niet-administratieve shares gevonden." -Type info
    }
  } else {
    $sharesHtml += New-Alert -Text "Get-SmbShare niet beschikbaar op dit systeem." -Type warn
  }
  $reportSections.Add((Add-Section -Id "shares" -Title "Shares" -BodyHtml $sharesHtml -Description $sectionDescriptions.shares))
} catch {
  $reportSections.Add((Add-Section -Id "shares" -Title "Shares" -BodyHtml (New-Alert -Text "Kon share-informatie niet ophalen: $($_.Exception.Message)") -Description $sectionDescriptions.shares))
}
#endregion

#region === Printers ===
try {
  $prtHtml = ""
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
    $prtHtml += New-Alert -Text "Printer-cmdlets niet beschikbaar (PrintManagement module ontbreekt?)." -Type warn
  }
  $reportSections.Add((Add-Section -Id "printers" -Title "Printers" -BodyHtml $prtHtml -Description $sectionDescriptions.printers))
} catch {
  $reportSections.Add((Add-Section -Id "printers" -Title "Printers" -BodyHtml (New-Alert -Text "Kon printerinformatie niet ophalen: $($_.Exception.Message)") -Description $sectionDescriptions.printers))
}
#endregion

#region === HTML skeleton ===
$css = @"
:root{--hdrH:64px}
*{box-sizing:border-box}html{font-family:Segoe UI,Arial;line-height:1.35}body{margin:0;background:#0f172a;color:#e5e7eb}

/* Sticky header (blauwe balk) */
header{position:sticky;top:0;z-index:30;background:linear-gradient(90deg,#0ea5e9,#6366f1);padding:16px 24px;color:white}
header h1{margin:0;font-size:20px}
header .meta{opacity:.9;font-size:12px}

/* Sticky tabs onder header — geen interne scrollbalk */
nav.tabs{
  position:sticky; top:64px; z-index:25;
  display:flex; flex-wrap:wrap; gap:6px; align-content:flex-start;
  padding:10px 12px; background:#111827; border-bottom:1px solid #1f2937;
  overflow:visible !important; max-height:none;
}
nav.tabs::-webkit-scrollbar{display:none} /* voor de zekerheid, Edge/Chromium */
nav.tabs a{padding:8px 12px;border-radius:10px;background:#1f2937;color:#e5e7eb;text-decoration:none;font-size:13px;transition:.15s}
nav.tabs a:hover{background:#374151}
nav.tabs a.active{background:#2563eb;color:#fff}

main{padding:18px}
.section{background:#0b1220;border:1px solid #1e293b;border-radius:14px;padding:16px;margin-bottom:16px;box-shadow:0 1px 0 rgba(255,255,255,.03) inset}
.section h2{margin-top:0;font-size:18px;color:#93c5fd}

/* Omschrijving onder titel */
.desc{margin:8px 0 14px;font-size:13px;color:#cbd5e1;background:#0a1324;border:1px solid #1e293b;border-radius:10px;padding:10px 12px}

/* Alerts */
.alert{display:flex;gap:8px;align-items:flex-start;border-radius:10px;padding:10px 12px;margin:8px 0}
.alert .ico{font-size:12px}
.alert.error{background:#7f1d1d;color:#fecaca}
.alert.warn{background:#78350f;color:#fde68a}
.alert.info{background:#1e293b;color:#93c5fd}
.alert.ok{background:#064e3b;color:#a7f3d0}

/* Preformatted blok */
pre{background:#0a0f1a;border:1px solid #1e293b;border-radius:10px;padding:12px;overflow:auto;max-height:15em;white-space:pre}

/* Tabel */
.tablewrap{overflow:auto}
table{border-collapse:collapse;width:100%;margin:8px 0}
th,td{border-bottom:1px solid #1f2937;padding:8px 10px;text-align:left}
th{position:sticky;top:0;background:#111827}
tr:hover{background:#0f1a2c}
.compact th, .compact td{font-size:12px}

/* Grid & cards */
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:8px}
.card{background:#0b1220;border:1px solid #1e293b;border-radius:12px;padding:12px}
.card h4{margin:0 0 6px 0;font-size:12px;color:#9ca3af;text-transform:uppercase;letter-spacing:.06em}
.card p{margin:0;font-size:14px;color:#e5e7eb}

/* Tabs: alles verbergen; alleen .active tonen */
.tab-content{display:none}
.tab-content.active{display:block}

footer{opacity:.7;font-size:12px;padding:12px 18px}
"@

$js = @"
(function(){
  var tabs = document.querySelectorAll('nav.tabs a');
  var secs = document.querySelectorAll('.tab-content');

  function activate(id){
    for (var i=0;i<secs.length;i++){ secs[i].classList.toggle('active', secs[i].id===id); }
    for (var j=0;j<tabs.length;j++){ tabs[j].classList.toggle('active', (tabs[j].getAttribute('href')==='#'+id)); }
    try{ history.replaceState(null,'','#'+id); }catch(e){}
  }

  for (var k=0;k<tabs.length;k++){
    tabs[k].addEventListener('click', function(e){ e.preventDefault(); activate(this.getAttribute('href').substring(1)); });
  }

  activate('system'); // default
})();
"@

$idsAndTitles = @(
  @{Id="system";Title="System Info"},
  @{Id="network";Title="Netwerk"},
  @{Id="firewall";Title="Firewall en Poorten"},
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
<html lang="nl">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Server Inventory | $env:COMPUTERNAME</title>
  <style>$css</style>
</head>
<body>
  <header>
    <h1>Server Inventory  $env:COMPUTERNAME</h1>
    <div class="meta">Gegenereerd: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")  User: $env:USERNAME  Domein: $env:USERDOMAIN</div>
  </header>
  <nav class="tabs">
    $($nav -join "`n")
  </nav>
  <main>
    $($reportSections -join "`n")
  </main>
  <footer>Rapport gegenereerd door ServerInventory-Report.ps1</footer>
  <script>$js</script>
</body>
</html>
"@
#endregion

#region === Write file ===
try {
  $null = New-Item -Path (Split-Path $OutputPath) -ItemType Directory -Force -ErrorAction SilentlyContinue
  $html | Out-File -FilePath $OutputPath -Encoding UTF8
  Write-Host "Rapport geschreven naar: $OutputPath" -ForegroundColor Cyan
} catch {
  Write-Warning "Kon rapport niet wegschrijven: $($_.Exception.Message)"
}
#endregion
