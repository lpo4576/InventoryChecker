$pcs = Import-Excel -Path "\\ctsfile\groups\CTS IT Desktop Support\Inventory\St Louis Master IT Inventory Hardware.xlsx" | where -filterscript {$_."Area/Assignee" -ne "IT Office" -and $_.Type -ne "Thinclient" -and $_."Operating System" -ne "Win 7" -and $_.'Vitalant Asset' -ne $env:COMPUTERNAME}
[System.Collections.Generic.List[pscustomobject]]$finalresults = @()
[System.Collections.Generic.List[pscustomobject]]$buildresults = @()


foreach ($item in $pcs) {
    $dns = Resolve-DnsName -Name $item."Vitalant Asset" -ErrorAction SilentlyContinue | where -FilterScript {$_.type -eq "A"}
    $pingbool = Test-Connection -ComputerName $dns.IPAddress -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($pingbool -eq $true) {
        if ($(Test-WSMan -ComputerName $item.'Vitalant Asset' -ErrorAction SilentlyContinue).ProductVersion -ne $null) {
            $winrmstatus = 'Yes'
            }
            else {$winrmstatus = 'Not found'}
        $results = "" | select Asset, IP, Pingable?, 'WinRM Enabled?', Area/Assignee, Type
        $results.Asset = $item.'Vitalant Asset'
        $results.IP = $dns.IPAddress
        $results.'Pingable?' = $pingbool
        $results.'WinRM Enabled?' = $winrmstatus
        $results.'Area/Assignee' = $item.'Area/Assignee'
        $results.Type = $item.Type
        $null = $finalresults.Add($results)
        }
    else {Write-Host "$($item.'Vitalant Asset') ($($item.'Area/Assignee')) is currently unreachable" -foregroundcolor yellow}
    }

foreach ($item in $finalresults) {
    $pcbuild = Invoke-Command -ComputerName $item.Asset -ScriptBlock {
        (Get-ItemProperty -Path "HKLM:\software\Microsoft\windows NT\currentversion" -Name Releaseid).releaseid
        } -ErrorAction SilentlyContinue
    $CurrentReaderVersion = Invoke-Command -ComputerName $item.Asset -ScriptBlock {
        Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
        Select-Object DisplayName, DisplayVersion | Where-Object{$_.DisplayName -like "*Adobe*" -and $_.DisplayName -like "*Reader*"}
        } -ErrorAction SilentlyContinue
    $o365version = Invoke-Command -ComputerName $item.Asset -ScriptBlock {
        (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us').DisplayVersion
        } -ErrorAction SilentlyContinue
    $chromeversion = Invoke-Command -ComputerName $item.asset -ScriptBlock {
        ((Get-Item (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)').VersionInfo).ProductVersion
        } -ErrorAction SilentlyContinue
    $hotfix = Invoke-Command -ComputerName $item.asset -ScriptBlock {
        get-hotfix | sort -Property hotfixid -Descending
        } -ErrorAction SilentlyContinue
    $javav = Invoke-Command -ComputerName $item.asset -ScriptBlock {
        (Get-Command java).Version.ToString()} -ErrorAction SilentlyContinue
    $results = '' | select Hostname, 'Area/Assignee', Pingable?, WinRM?, 'Win10 Build', 'Adobe Version', 'O365 Version', 'Chrome Version', 'Latest Windows Hotfix', 'Java Version'
    $results.Hostname = $item.Asset
    $results.'Pingable?' = $item.'Pingable?'
    $results.'WinRM?'= $item.'WinRM Enabled?'
    $results.'Win10 Build' = $pcbuild
    $results.'Area/Assignee' = $item.'Area/Assignee'
    $results.'Adobe Version' = $CurrentReaderVersion.DisplayVersion
    $results.'O365 Version' = $o365version
    $results.'Chrome Version' = $chromeversion
    if ($hotfix -ne $null) {
        $results.'Latest Windows Hotfix' = $($hotfix[0]).hotfixid
        }
    $results.'Java Version' = $javav
    $null = $buildresults.Add($results)
    }

$buildresults | Export-Csv -Path "C:\Users\312127\Desktop\Testing\Inventory$((get-date).ToString('MMddyyyy')).csv" -NoTypeInformation