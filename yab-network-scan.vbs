Dim ps, tmp, fso, f, sh
Set fso = CreateObject("Scripting.FileSystemObject")
tmp = fso.GetTempName() & ".ps1"
tmp = fso.BuildPath(fso.GetSpecialFolder(2), tmp)

ps = ""
ps = ps & "#Requires -Version 5.1" & vbCrLf
ps = ps & "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf
ps = ps & "Add-Type -AssemblyName System.Drawing" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Config --------------------------------------------------------------------" & vbCrLf
ps = ps & "$PingTimeoutMs      = 500" & vbCrLf
ps = ps & "$MaxConcurrentPings = 20" & vbCrLf
ps = ps & "$MaxHistoryPoints   = 3600   # ~2 hrs at 2s intervals" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Subnet / interface detection ----------------------------------------------" & vbCrLf
ps = ps & "function Get-SubnetInterfaceMap {" & vbCrLf
ps = ps & "    $map = @{}" & vbCrLf
ps = ps & "    foreach ($iface in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {" & vbCrLf
ps = ps & "        if ($iface.OperationalStatus -ne 'Up') { continue }" & vbCrLf
ps = ps & "        if ($iface.NetworkInterfaceType -eq 'Loopback') { continue }" & vbCrLf
ps = ps & "        if ($iface.Description -match 'Teredo|6TO4|isatap|Pseudo|WAN Miniport|Bluetooth') { continue }" & vbCrLf
ps = ps & "        $ifType = switch ($iface.NetworkInterfaceType) {" & vbCrLf
ps = ps & "            'Wireless80211' { 'WiFi' }" & vbCrLf
ps = ps & "            'Ethernet'      { 'Ethernet' }" & vbCrLf
ps = ps & "            default         { 'Other' }" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        foreach ($ua in $iface.GetIPProperties().UnicastAddresses) {" & vbCrLf
ps = ps & "            if ($ua.Address.AddressFamily -ne 'InterNetwork') { continue }" & vbCrLf
ps = ps & "            $ip = $ua.Address.ToString()" & vbCrLf
ps = ps & "            if ($ip -eq '127.0.0.1') { continue }" & vbCrLf
ps = ps & "            $parts = $ip -split '\.'" & vbCrLf
ps = ps & "            $sub   = ""$($parts[0]).$($parts[1]).$($parts[2])""" & vbCrLf
ps = ps & "            if (-not $map.ContainsKey($sub)) { $map[$sub] = $ifType }" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    return $map" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$SubnetIfaceMap  = Get-SubnetInterfaceMap" & vbCrLf
ps = ps & "$DetectedSubnets = @($SubnetIfaceMap.Keys)" & vbCrLf
ps = ps & "$DetectedSubnet  = if ($DetectedSubnets.Count -gt 0) {" & vbCrLf
ps = ps & "    [string]($DetectedSubnets | Sort-Object { if ($_ -notmatch '^192\.168') {0} else {1} } | Select-Object -First 1)" & vbCrLf
ps = ps & "} else { ""192.168.1"" }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Global state --------------------------------------------------------------" & vbCrLf
ps = ps & "$Global:NetworkStats    = @{}     # IP -> { History; TimestampedHistory; Timeouts; IfType; DnsName }" & vbCrLf
ps = ps & "$Global:DeviceLabels    = @{}     # IP -> custom label string" & vbCrLf
ps = ps & "$Global:DeviceAlerts    = @{}     # IP -> { WasUp; AlertOnDown; AlertOnUp }" & vbCrLf
ps = ps & "$Global:UIReferences    = @{}     # IP -> { listview sub-item refs }" & vbCrLf
ps = ps & "$Global:AllItems        = [System.Collections.Generic.List[System.Windows.Forms.ListViewItem]]::new()" & vbCrLf
ps = ps & "$Global:ScanComplete    = $false" & vbCrLf
ps = ps & "$Global:StopRequested   = $false" & vbCrLf
ps = ps & "$Global:StartTime       = [DateTime]::Now" & vbCrLf
ps = ps & "$Global:AsyncJobs       = [System.Collections.Generic.List[hashtable]]::new()" & vbCrLf
ps = ps & "$Global:PingQueue       = [System.Collections.Generic.Queue[string]]::new()" & vbCrLf
ps = ps & "$Global:ActivePings     = 0" & vbCrLf
ps = ps & "$Global:TotalQueued     = 0" & vbCrLf
ps = ps & "$Global:FilterIfType    = ""All""" & vbCrLf
ps = ps & "$Global:RefreshMs       = 2000    # controlled by interval picker" & vbCrLf
ps = ps & "$Global:Paused          = $false" & vbCrLf
ps = ps & "$Global:SortCol         = -1" & vbCrLf
ps = ps & "$Global:SortAsc         = $true" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Helper functions ----------------------------------------------------------" & vbCrLf
ps = ps & "function Get-LatencyColor ($ms) {" & vbCrLf
ps = ps & "    if ($ms -ge $PingTimeoutMs) { return [System.Drawing.Color]::DarkRed }" & vbCrLf
ps = ps & "    if ($ms -gt 10)             { return [System.Drawing.Color]::Red }" & vbCrLf
ps = ps & "    if ($ms -gt 3)              { return [System.Drawing.Color]::DarkOrange }" & vbCrLf
ps = ps & "    return [System.Drawing.Color]::Green" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Get-IfTypeIcon  ($t) { switch($t){ ""WiFi""{""[~]""} ""Ethernet""{""[=]""} default{""[?]""} } }" & vbCrLf
ps = ps & "function Get-IfTypeColor ($t) {" & vbCrLf
ps = ps & "    switch($t){" & vbCrLf
ps = ps & "        ""WiFi""     { return [System.Drawing.Color]::FromArgb(0,100,200) }" & vbCrLf
ps = ps & "        ""Ethernet"" { return [System.Drawing.Color]::FromArgb(0,140,60) }" & vbCrLf
ps = ps & "        default    { return [System.Drawing.Color]::Gray }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Get-DisplayName ($ip) {" & vbCrLf
ps = ps & "    if ($Global:DeviceLabels.ContainsKey($ip) -and $Global:DeviceLabels[$ip]) {" & vbCrLf
ps = ps & "        return $Global:DeviceLabels[$ip]" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    if ($Global:NetworkStats.ContainsKey($ip) -and $Global:NetworkStats[$ip].DnsName) {" & vbCrLf
ps = ps & "        return $Global:NetworkStats[$ip].DnsName" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    return $ip" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Get-Sparkline ($history) {" & vbCrLf
ps = ps & "    if (-not $history -or $history.Count -eq 0) { return """" }" & vbCrLf
ps = ps & "    $bars  = @(""."", "":"", ""i"", ""I"", ""X"")" & vbCrLf
ps = ps & "    $pts   = @($history | Select-Object -Last 14)" & vbCrLf
ps = ps & "    $realPts = @($pts | Where-Object { $_ -lt $PingTimeoutMs })" & vbCrLf
ps = ps & "    $minV  = if ($realPts.Count -gt 0) { ($realPts | Measure-Object -Minimum).Minimum } else { 1 }" & vbCrLf
ps = ps & "    $maxV  = if ($realPts.Count -gt 0) { ($realPts | Measure-Object -Maximum).Maximum } else { 1 }" & vbCrLf
ps = ps & "    $range = $maxV - $minV" & vbCrLf
ps = ps & "    $spark = """"" & vbCrLf
ps = ps & "    foreach ($v in $pts) {" & vbCrLf
ps = ps & "        if ($v -ge $PingTimeoutMs) { $spark += ""!"" }" & vbCrLf
ps = ps & "        elseif ($range -lt 2)     { $spark += ""."" }" & vbCrLf
ps = ps & "        else {" & vbCrLf
ps = ps & "            $idx   = [int]([Math]::Min(4, ($v - $minV) / $range * 5))" & vbCrLf
ps = ps & "            $spark += $bars[$idx]" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    return $spark" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Parse-Range ($s, $e) {" & vbCrLf
ps = ps & "    $si = 0; $ei = 0" & vbCrLf
ps = ps & "    if ([string]::IsNullOrWhiteSpace($s) -or [string]::IsNullOrWhiteSpace($e)) { return @() }" & vbCrLf
ps = ps & "    if (-not [int]::TryParse($s.Trim(),[ref]$si) -or -not [int]::TryParse($e.Trim(),[ref]$ei)) { return @() }" & vbCrLf
ps = ps & "    $si = [Math]::Max(1,[Math]::Min(254,$si)); $ei = [Math]::Max(1,[Math]::Min(254,$ei))" & vbCrLf
ps = ps & "    if ($si -gt $ei) { return @() }" & vbCrLf
ps = ps & "    return ($si..$ei)" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Get-SubnetForIP ($ip) {" & vbCrLf
ps = ps & "    $p = $ip -split '\.'; return ""$($p[0]).$($p[1]).$($p[2])""" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Form ----------------------------------------------------------------------" & vbCrLf
ps = ps & "$Form               = New-Object System.Windows.Forms.Form" & vbCrLf
ps = ps & "$Form.Text          = ""YAB Network Monitor""" & vbCrLf
ps = ps & "$Form.Size          = New-Object System.Drawing.Size(1120, 900)" & vbCrLf
ps = ps & "$Form.StartPosition = ""CenterScreen""" & vbCrLf
ps = ps & "$Form.BackColor     = [System.Drawing.Color]::FromArgb(245,245,248)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Top Panel -----------------------------------------------------------------" & vbCrLf
ps = ps & "$TopPanel           = New-Object System.Windows.Forms.Panel" & vbCrLf
ps = ps & "$TopPanel.Dock      = [System.Windows.Forms.DockStyle]::Top" & vbCrLf
ps = ps & "$TopPanel.Height    = 115" & vbCrLf
ps = ps & "$TopPanel.BackColor = [System.Drawing.Color]::FromArgb(235,235,242)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Add-TopLabel ($text, $x, $y) {" & vbCrLf
ps = ps & "    $l = New-Object System.Windows.Forms.Label" & vbCrLf
ps = ps & "    $l.Text = $text; $l.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "    $l.Location = New-Object System.Drawing.Point($x, $y); $l.AutoSize = $true" & vbCrLf
ps = ps & "    [void]$TopPanel.Controls.Add($l); return $l" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "function Add-TopTextBox ($text, $x, $y, $w) {" & vbCrLf
ps = ps & "    $t = New-Object System.Windows.Forms.TextBox" & vbCrLf
ps = ps & "    $t.Text = $text; $t.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "    $t.Location = New-Object System.Drawing.Point($x, $y); $t.Size = New-Object System.Drawing.Size($w, 24)" & vbCrLf
ps = ps & "    $t.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center" & vbCrLf
ps = ps & "    [void]$TopPanel.Controls.Add($t); return $t" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "[void](Add-TopLabel ""Subnet:"" 12 14)" & vbCrLf
ps = ps & "$TxtSubnet  = Add-TopTextBox $DetectedSubnet 68 11 115" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "if ($DetectedSubnets.Count -gt 1) {" & vbCrLf
ps = ps & "    $parts = $DetectedSubnets | ForEach-Object { ""$_ ($($SubnetIfaceMap[$_]))"" }" & vbCrLf
ps = ps & "    $ld = Add-TopLabel (""Detected: "" + ($parts -join "", "")) 68 36" & vbCrLf
ps = ps & "    $ld.ForeColor = [System.Drawing.Color]::DimGray" & vbCrLf
ps = ps & "    $ld.Font = New-Object System.Drawing.Font(""Segoe UI"", 8)" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "[void](Add-TopLabel ""Range 1:"" 210 14)" & vbCrLf
ps = ps & "$TxtR1Start = Add-TopTextBox ""1""   270 11 46" & vbCrLf
ps = ps & "[void](Add-TopLabel ""-"" 319 11)" & vbCrLf
ps = ps & "$TxtR1End   = Add-TopTextBox ""254"" 333 11 46" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "[void](Add-TopLabel ""Range 2:"" 210 50)" & vbCrLf
ps = ps & "$TxtR2Start = Add-TopTextBox """" 270 47 46" & vbCrLf
ps = ps & "[void](Add-TopLabel ""-"" 319 47)" & vbCrLf
ps = ps & "$TxtR2End   = Add-TopTextBox """" 333 47 46" & vbCrLf
ps = ps & "$lh = Add-TopLabel ""(optional)"" 384 51" & vbCrLf
ps = ps & "$lh.ForeColor = [System.Drawing.Color]::DimGray" & vbCrLf
ps = ps & "$lh.Font = New-Object System.Drawing.Font(""Segoe UI"", 8)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Filter buttons ------------------------------------------------------------" & vbCrLf
ps = ps & "[void](Add-TopLabel ""Show:"" 450 14)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function New-TopBtn ($text, $x, $y, $w, $h) {" & vbCrLf
ps = ps & "    $b = New-Object System.Windows.Forms.Button" & vbCrLf
ps = ps & "    $b.Text = $text; $b.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "    $b.Size = New-Object System.Drawing.Size($w, $h)" & vbCrLf
ps = ps & "    $b.Location = New-Object System.Drawing.Point($x, $y)" & vbCrLf
ps = ps & "    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat" & vbCrLf
ps = ps & "    $b.FlatAppearance.BorderSize = 1" & vbCrLf
ps = ps & "    $b.BackColor = [System.Drawing.Color]::FromArgb(220,220,230)" & vbCrLf
ps = ps & "    $b.ForeColor = [System.Drawing.Color]::Black" & vbCrLf
ps = ps & "    [void]$TopPanel.Controls.Add($b); return $b" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$BtnFilterAll  = New-TopBtn ""All""      450 35 88 28" & vbCrLf
ps = ps & "$BtnFilterEth  = New-TopBtn ""Ethernet"" 542 35 88 28" & vbCrLf
ps = ps & "$BtnFilterWifi = New-TopBtn ""WiFi""     634 35 88 28" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Set-FilterHighlight ($active) {" & vbCrLf
ps = ps & "    foreach ($b in @($BtnFilterAll, $BtnFilterEth, $BtnFilterWifi)) {" & vbCrLf
ps = ps & "        $b.BackColor = [System.Drawing.Color]::FromArgb(220,220,230)" & vbCrLf
ps = ps & "        $b.ForeColor = [System.Drawing.Color]::Black" & vbCrLf
ps = ps & "        $b.Font      = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $active.BackColor = [System.Drawing.Color]::FromArgb(0,120,212)" & vbCrLf
ps = ps & "    $active.ForeColor = [System.Drawing.Color]::White" & vbCrLf
ps = ps & "    $active.Font      = New-Object System.Drawing.Font(""Segoe UI"", 9, [System.Drawing.FontStyle]::Bold)" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Apply-Filter {" & vbCrLf
ps = ps & "    $ListView.BeginUpdate()" & vbCrLf
ps = ps & "    $ListView.Items.Clear()" & vbCrLf
ps = ps & "    foreach ($lvi in $Global:AllItems) {" & vbCrLf
ps = ps & "        $ip = $lvi.Text" & vbCrLf
ps = ps & "        $ift = if ($Global:NetworkStats.ContainsKey($ip)) { $Global:NetworkStats[$ip].IfType } else { ""Other"" }" & vbCrLf
ps = ps & "        if ($Global:FilterIfType -eq ""All"" -or $ift -eq $Global:FilterIfType) {" & vbCrLf
ps = ps & "            [void]$ListView.Items.Add($lvi)" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $ListView.EndUpdate()" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "Set-FilterHighlight $BtnFilterAll" & vbCrLf
ps = ps & "$BtnFilterAll.Add_Click({  $Global:FilterIfType = ""All"";      Set-FilterHighlight $BtnFilterAll;  Apply-Filter })" & vbCrLf
ps = ps & "$BtnFilterEth.Add_Click({  $Global:FilterIfType = ""Ethernet""; Set-FilterHighlight $BtnFilterEth;  Apply-Filter })" & vbCrLf
ps = ps & "$BtnFilterWifi.Add_Click({ $Global:FilterIfType = ""WiFi"";     Set-FilterHighlight $BtnFilterWifi; Apply-Filter })" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Interval picker -----------------------------------------------------------" & vbCrLf
ps = ps & "[void](Add-TopLabel ""Interval:"" 730 14)" & vbCrLf
ps = ps & "$CmbInterval = New-Object System.Windows.Forms.ComboBox" & vbCrLf
ps = ps & "$CmbInterval.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "$CmbInterval.Location = New-Object System.Drawing.Point(790, 11)" & vbCrLf
ps = ps & "$CmbInterval.Size = New-Object System.Drawing.Size(85, 24)" & vbCrLf
ps = ps & "$CmbInterval.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList" & vbCrLf
ps = ps & "@(""1s"",""2s"",""5s"",""10s"",""30s"") | ForEach-Object { [void]$CmbInterval.Items.Add($_) }" & vbCrLf
ps = ps & "$CmbInterval.SelectedIndex = 1" & vbCrLf
ps = ps & "[void]$TopPanel.Controls.Add($CmbInterval)" & vbCrLf
ps = ps & "$CmbInterval.Add_SelectedIndexChanged({" & vbCrLf
ps = ps & "    $map = @{ ""1s""=1000; ""2s""=2000; ""5s""=5000; ""10s""=10000; ""30s""=30000 }" & vbCrLf
ps = ps & "    $Global:RefreshMs = $map[$CmbInterval.SelectedItem]" & vbCrLf
ps = ps & "    if ($Global:ScanComplete) { $Timer.Interval = $Global:RefreshMs }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Alert checkbox ------------------------------------------------------------" & vbCrLf
ps = ps & "$ChkAlerts = New-Object System.Windows.Forms.CheckBox" & vbCrLf
ps = ps & "$ChkAlerts.Text = ""Down/Up Alerts""" & vbCrLf
ps = ps & "$ChkAlerts.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "$ChkAlerts.Location = New-Object System.Drawing.Point(730, 40)" & vbCrLf
ps = ps & "$ChkAlerts.AutoSize = $true" & vbCrLf
ps = ps & "$ChkAlerts.Checked = $true" & vbCrLf
ps = ps & "[void]$TopPanel.Controls.Add($ChkAlerts)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$BtnPause = New-Object System.Windows.Forms.Button" & vbCrLf
ps = ps & "$BtnPause.Text = ""Pause""" & vbCrLf
ps = ps & "$BtnPause.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "$BtnPause.Size = New-Object System.Drawing.Size(80, 28)" & vbCrLf
ps = ps & "$BtnPause.Location = New-Object System.Drawing.Point(870, 40)" & vbCrLf
ps = ps & "$BtnPause.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat" & vbCrLf
ps = ps & "$BtnPause.FlatAppearance.BorderSize = 1" & vbCrLf
ps = ps & "$BtnPause.BackColor = [System.Drawing.Color]::FromArgb(220,220,230)" & vbCrLf
ps = ps & "$BtnPause.ForeColor = [System.Drawing.Color]::Black" & vbCrLf
ps = ps & "$BtnPause.Enabled = $false" & vbCrLf
ps = ps & "[void]$TopPanel.Controls.Add($BtnPause)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$BtnPause.Add_Click({" & vbCrLf
ps = ps & "    $Global:Paused = -not $Global:Paused" & vbCrLf
ps = ps & "    if ($Global:Paused) {" & vbCrLf
ps = ps & "        $BtnPause.Text = ""Resume""" & vbCrLf
ps = ps & "        $BtnPause.BackColor = [System.Drawing.Color]::FromArgb(196,43,28)" & vbCrLf
ps = ps & "        $BtnPause.ForeColor = [System.Drawing.Color]::White" & vbCrLf
ps = ps & "        $RuntimeLabel.ForeColor = [System.Drawing.Color]::DarkOrange" & vbCrLf
ps = ps & "        $RuntimeLabel.Text = ""PAUSED  |  $($Global:UIReferences.Count) devices  (monitoring halted)""" & vbCrLf
ps = ps & "    } else {" & vbCrLf
ps = ps & "        $BtnPause.Text = ""Pause""" & vbCrLf
ps = ps & "        $BtnPause.BackColor = [System.Drawing.Color]::FromArgb(220,220,230)" & vbCrLf
ps = ps & "        $BtnPause.ForeColor = [System.Drawing.Color]::Black" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Reset / Save / Load / Export buttons --------------------------------------" & vbCrLf
ps = ps & "$BtnReset  = New-TopBtn ""Reset Counters""  450 75 120 28" & vbCrLf
ps = ps & "$BtnSave   = New-TopBtn ""Save Session""    578 75 105 28" & vbCrLf
ps = ps & "$BtnLoad   = New-TopBtn ""Load Session""    688 75 105 28" & vbCrLf
ps = ps & "$BtnExpCSV = New-TopBtn ""Export CSV""      799 75  90 28" & vbCrLf
ps = ps & "$BtnExpHTML= New-TopBtn ""Export HTML""     893 75  95 28" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$BtnReset.Add_Click({" & vbCrLf
ps = ps & "    foreach ($ip in @($Global:NetworkStats.Keys)) {" & vbCrLf
ps = ps & "        $s = $Global:NetworkStats[$ip]; $ui = $Global:UIReferences[$ip]" & vbCrLf
ps = ps & "        $s.Timeouts = 0; $s.History.Clear(); $s.TimestampedHistory.Clear()" & vbCrLf
ps = ps & "        if ($ui) {" & vbCrLf
ps = ps & "            $ui.Current.Text = ""-""; $ui.Min.Text = ""-""; $ui.Max.Text = ""-""" & vbCrLf
ps = ps & "            $ui.Avg.Text = ""-""; $ui.Timeout.Text = ""0""" & vbCrLf
ps = ps & "            $ui.Timeout.ForeColor = [System.Drawing.Color]::Black" & vbCrLf
ps = ps & "            $ui.Spark.Text = """"" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $Global:StartTime = [DateTime]::Now" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Start button --------------------------------------------------------------" & vbCrLf
ps = ps & "$StartButton = New-Object System.Windows.Forms.Button" & vbCrLf
ps = ps & "$StartButton.Text = ""Start Scan""" & vbCrLf
ps = ps & "$StartButton.Font = New-Object System.Drawing.Font(""Segoe UI"", 10, [System.Drawing.FontStyle]::Bold)" & vbCrLf
ps = ps & "$StartButton.Size = New-Object System.Drawing.Size(148, 90)" & vbCrLf
ps = ps & "$StartButton.Location = New-Object System.Drawing.Point(958, 12)" & vbCrLf
ps = ps & "$StartButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,212)" & vbCrLf
ps = ps & "$StartButton.ForeColor = [System.Drawing.Color]::White" & vbCrLf
ps = ps & "$StartButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat" & vbCrLf
ps = ps & "$StartButton.FlatAppearance.BorderSize = 0" & vbCrLf
ps = ps & "[void]$TopPanel.Controls.Add($StartButton)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Bottom status bar ---------------------------------------------------------" & vbCrLf
ps = ps & "$BottomPanel = New-Object System.Windows.Forms.Panel" & vbCrLf
ps = ps & "$BottomPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom" & vbCrLf
ps = ps & "$BottomPanel.Height = 38" & vbCrLf
ps = ps & "$BottomPanel.BackColor = [System.Drawing.Color]::FromArgb(228,228,238)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$RuntimeLabel = New-Object System.Windows.Forms.Label" & vbCrLf
ps = ps & "$RuntimeLabel.Text = ""Configure ranges above and click Start Scan.""" & vbCrLf
ps = ps & "$RuntimeLabel.Font = New-Object System.Drawing.Font(""Segoe UI"", 10, [System.Drawing.FontStyle]::Bold)" & vbCrLf
ps = ps & "$RuntimeLabel.ForeColor = [System.Drawing.Color]::DarkCyan" & vbCrLf
ps = ps & "$RuntimeLabel.Location = New-Object System.Drawing.Point(12, 10)" & vbCrLf
ps = ps & "$RuntimeLabel.AutoSize = $true" & vbCrLf
ps = ps & "[void]$BottomPanel.Controls.Add($RuntimeLabel)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$BtnCopy = New-Object System.Windows.Forms.Button" & vbCrLf
ps = ps & "$BtnCopy.Text = ""Copy Grid""" & vbCrLf
ps = ps & "$BtnCopy.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "$BtnCopy.Size = New-Object System.Drawing.Size(90, 26)" & vbCrLf
ps = ps & "$BtnCopy.Anchor = [System.Windows.Forms.AnchorStyles]::Right" & vbCrLf
ps = ps & "$BtnCopy.Location = New-Object System.Drawing.Point(1010, 6)" & vbCrLf
ps = ps & "[void]$BottomPanel.Controls.Add($BtnCopy)" & vbCrLf
ps = ps & "$BtnCopy.Add_Click({" & vbCrLf
ps = ps & "    $sb = New-Object System.Text.StringBuilder" & vbCrLf
ps = ps & "    [void]$sb.AppendLine(""IP`tLabel`tDNS Name`tType`tCurrent`tMin`tMax`tAvg`tTimeouts"")" & vbCrLf
ps = ps & "    foreach ($lvi in $Global:AllItems) {" & vbCrLf
ps = ps & "        $ip = $lvi.Text" & vbCrLf
ps = ps & "        $label = if ($Global:DeviceLabels.ContainsKey($ip)) { $Global:DeviceLabels[$ip] } else { """" }" & vbCrLf
ps = ps & "        $dns   = if ($Global:NetworkStats.ContainsKey($ip))  { $Global:NetworkStats[$ip].DnsName } else { """" }" & vbCrLf
ps = ps & "        $row = @($ip, $label, $dns)" & vbCrLf
ps = ps & "        for ($i = 2; $i -lt $lvi.SubItems.Count; $i++) { $row += $lvi.SubItems[$i].Text }" & vbCrLf
ps = ps & "        [void]$sb.AppendLine(($row -join ""`t""))" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    [System.Windows.Forms.Clipboard]::SetText($sb.ToString())" & vbCrLf
ps = ps & "    [System.Windows.Forms.MessageBox]::Show(""Copied to clipboard."", ""OK"", 0, 64)" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- ListView ------------------------------------------------------------------" & vbCrLf
ps = ps & "$ListView = New-Object System.Windows.Forms.ListView" & vbCrLf
ps = ps & "$ListView.View = [System.Windows.Forms.View]::Details" & vbCrLf
ps = ps & "$ListView.FullRowSelect = $true" & vbCrLf
ps = ps & "$ListView.GridLines = $true" & vbCrLf
ps = ps & "$ListView.Dock = [System.Windows.Forms.DockStyle]::Fill" & vbCrLf
ps = ps & "$ListView.Font = New-Object System.Drawing.Font(""Segoe UI"", 10)" & vbCrLf
ps = ps & "# Col indices:  0=IP  1=Label  2=DNS  3=Type  4=Current  5=Min  6=Max  7=Avg  8=Timeouts  9=Trend" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""IP Address"",  110)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Label"",       120)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Device Name"", 190)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Type"",         65)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Current"",      80)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Min"",          70)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Max"",          70)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Avg"",          70)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Timeouts"",     70)" & vbCrLf
ps = ps & "[void]$ListView.Columns.Add(""Trend"",       110)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Live sort -----------------------------------------------------------------" & vbCrLf
ps = ps & "function Apply-Sort {" & vbCrLf
ps = ps & "    if ($Global:SortCol -lt 0) { return }   # no sort selected yet" & vbCrLf
ps = ps & "    $all = @($Global:AllItems | ForEach-Object { $_ })" & vbCrLf
ps = ps & "    $col = $Global:SortCol" & vbCrLf
ps = ps & "    $sorted = $all | Sort-Object {" & vbCrLf
ps = ps & "        $v = $_.SubItems[$col].Text -replace ' ms','' -replace 'TIMEOUT','99999'" & vbCrLf
ps = ps & "        $n = 0.0" & vbCrLf
ps = ps & "        if ([double]::TryParse($v, [ref]$n)) { $n } else { $v }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    if (-not $Global:SortAsc) { $sorted = [Linq.Enumerable]::Reverse([object[]]$sorted) }" & vbCrLf
ps = ps & "    $Global:AllItems.Clear()" & vbCrLf
ps = ps & "    foreach ($i in $sorted) { [void]$Global:AllItems.Add($i) }" & vbCrLf
ps = ps & "    Apply-Filter" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$ListView.Add_ColumnClick({" & vbCrLf
ps = ps & "    param($s, $e)" & vbCrLf
ps = ps & "    if ($Global:SortCol -eq $e.Column) { $Global:SortAsc = -not $Global:SortAsc }" & vbCrLf
ps = ps & "    else { $Global:SortCol = $e.Column; $Global:SortAsc = $true }" & vbCrLf
ps = ps & "    # Update header indicator in column text" & vbCrLf
ps = ps & "    for ($ci = 0; $ci -lt $ListView.Columns.Count; $ci++) {" & vbCrLf
ps = ps & "        $hdr = $ListView.Columns[$ci].Text -replace ' [av]$',''" & vbCrLf
ps = ps & "        if ($ci -eq $Global:SortCol) {" & vbCrLf
ps = ps & "            $ListView.Columns[$ci].Text = $hdr + (if ($Global:SortAsc) { ' v' } else { ' ^' })" & vbCrLf
ps = ps & "        } else {" & vbCrLf
ps = ps & "            $ListView.Columns[$ci].Text = $hdr" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    Apply-Sort" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# Dock order matters" & vbCrLf
ps = ps & "[void]$Form.Controls.Add($BottomPanel)" & vbCrLf
ps = ps & "[void]$Form.Controls.Add($ListView)" & vbCrLf
ps = ps & "[void]$Form.Controls.Add($TopPanel)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Chart popup ---------------------------------------------------------------" & vbCrLf
ps = ps & "function Show-LatencyChart ($ip) {" & vbCrLf
ps = ps & "    $stats = $Global:NetworkStats[$ip]" & vbCrLf
ps = ps & "    if (-not $stats -or $stats.TimestampedHistory.Count -lt 2) {" & vbCrLf
ps = ps & "        [System.Windows.Forms.MessageBox]::Show(""Not enough data yet for $ip."", ""Chart"", 0, 64); return" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $displayName = Get-DisplayName $ip" & vbCrLf
ps = ps & "    $dnsName     = if ($stats.DnsName) { $stats.DnsName } else { ""Unknown"" }" & vbCrLf
ps = ps & "    $title = if ($displayName -ne $ip) { ""$displayName  ($dnsName)  -  $ip"" } else { ""$dnsName  -  $ip"" }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $cf = New-Object System.Windows.Forms.Form" & vbCrLf
ps = ps & "    $cf.Text = ""Latency: $title""" & vbCrLf
ps = ps & "    $cf.Size = New-Object System.Drawing.Size(820, 440)" & vbCrLf
ps = ps & "    $cf.StartPosition = ""CenterParent""" & vbCrLf
ps = ps & "    $cf.BackColor = [System.Drawing.Color]::White" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $pTop = New-Object System.Windows.Forms.Panel" & vbCrLf
ps = ps & "    $pTop.Dock = [System.Windows.Forms.DockStyle]::Top; $pTop.Height = 38" & vbCrLf
ps = ps & "    $pTop.BackColor = [System.Drawing.Color]::FromArgb(240,240,245)" & vbCrLf
ps = ps & "    [void]$cf.Controls.Add($pTop)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $lw = New-Object System.Windows.Forms.Label" & vbCrLf
ps = ps & "    $lw.Text = ""Show last:""; $lw.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "    $lw.Location = New-Object System.Drawing.Point(10,10); $lw.AutoSize = $true" & vbCrLf
ps = ps & "    [void]$pTop.Controls.Add($lw)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $cmb = New-Object System.Windows.Forms.ComboBox" & vbCrLf
ps = ps & "    $cmb.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "    $cmb.Location = New-Object System.Drawing.Point(82,7); $cmb.Size = New-Object System.Drawing.Size(120,24)" & vbCrLf
ps = ps & "    $cmb.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList" & vbCrLf
ps = ps & "    @(""5 minutes"",""15 minutes"",""30 minutes"",""1 hour"",""All data"") | ForEach-Object { [void]$cmb.Items.Add($_) }" & vbCrLf
ps = ps & "    $cmb.SelectedIndex = 1" & vbCrLf
ps = ps & "    [void]$pTop.Controls.Add($cmb)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $lsum = New-Object System.Windows.Forms.Label" & vbCrLf
ps = ps & "    $lsum.Font = New-Object System.Drawing.Font(""Segoe UI"", 9)" & vbCrLf
ps = ps & "    $lsum.Location = New-Object System.Drawing.Point(215, 10); $lsum.AutoSize = $true" & vbCrLf
ps = ps & "    [void]$pTop.Controls.Add($lsum)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $canvas = New-Object System.Windows.Forms.PictureBox" & vbCrLf
ps = ps & "    $canvas.Dock = [System.Windows.Forms.DockStyle]::Fill" & vbCrLf
ps = ps & "    $canvas.BackColor = [System.Drawing.Color]::White" & vbCrLf
ps = ps & "    [void]$cf.Controls.Add($canvas)" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    function Draw-Chart {" & vbCrLf
ps = ps & "        $winMin = switch ($cmb.SelectedItem) {" & vbCrLf
ps = ps & "            ""5 minutes""  { 5 } ""15 minutes"" { 15 } ""30 minutes"" { 30 } ""1 hour"" { 60 } default { [int]::MaxValue }" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        $cutoff = [DateTime]::Now.AddMinutes(-$winMin)" & vbCrLf
ps = ps & "        $pts = @($stats.TimestampedHistory | Where-Object { $_.Time -ge $cutoff })" & vbCrLf
ps = ps & "        if ($pts.Count -lt 2) { $lsum.Text = ""Not enough data in this window.""; return }" & vbCrLf
ps = ps & "        $vals = $pts | ForEach-Object { $_.Ms }" & vbCrLf
ps = ps & "        $minV = ($vals | Measure-Object -Minimum).Minimum" & vbCrLf
ps = ps & "        $maxV = ($vals | Measure-Object -Maximum).Maximum" & vbCrLf
ps = ps & "        $avgV = [Math]::Round(($vals | Measure-Object -Average).Average, 1)" & vbCrLf
ps = ps & "        $toCount = @($vals | Where-Object { $_ -ge $PingTimeoutMs }).Count" & vbCrLf
ps = ps & "        $lsum.Text = ""  Points: $($pts.Count)   Min: ${minV}ms   Max: ${maxV}ms   Avg: ${avgV}ms   Timeouts: $toCount""" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "        $W = [Math]::Max(200, $canvas.Width); $H = [Math]::Max(100, $canvas.Height)" & vbCrLf
ps = ps & "        $bmp = New-Object System.Drawing.Bitmap($W, $H)" & vbCrLf
ps = ps & "        $g   = [System.Drawing.Graphics]::FromImage($bmp)" & vbCrLf
ps = ps & "        $g.Clear([System.Drawing.Color]::White)" & vbCrLf
ps = ps & "        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias" & vbCrLf
ps = ps & "        $padL=58;$padR=15;$padT=18;$padB=32" & vbCrLf
ps = ps & "        $cW=$W-$padL-$padR; $cH=$H-$padT-$padB" & vbCrLf
ps = ps & "        $gridPen  = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(220,220,220))" & vbCrLf
ps = ps & "        $txtFont  = New-Object System.Drawing.Font(""Segoe UI"", 7)" & vbCrLf
ps = ps & "        $txtBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Gray)" & vbCrLf
ps = ps & "        $range    = [Math]::Max(10, $maxV - $minV)" & vbCrLf
ps = ps & "        for ($i=0; $i -le 5; $i++) {" & vbCrLf
ps = ps & "            $pct=$i/5.0; $yVal=$minV+$pct*$range" & vbCrLf
ps = ps & "            $yPx=$padT+$cH-[int]($pct*$cH)" & vbCrLf
ps = ps & "            $g.DrawLine($gridPen,$padL,$yPx,$padL+$cW,$yPx)" & vbCrLf
ps = ps & "            $g.DrawString([Math]::Round($yVal,0).ToString()+""ms"",$txtFont,$txtBrush,2,$yPx-7)" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        $t0=$pts[0].Time; $tN=$pts[-1].Time; $tSpan=($tN-$t0).TotalSeconds" & vbCrLf
ps = ps & "        if ($tSpan -gt 0) {" & vbCrLf
ps = ps & "            for ($i=0; $i -le 4; $i++) {" & vbCrLf
ps = ps & "                $frac=$i/4.0; $xPx=$padL+[int]($frac*$cW)" & vbCrLf
ps = ps & "                $tLbl=$t0.AddSeconds($frac*$tSpan).ToString(""HH:mm:ss"")" & vbCrLf
ps = ps & "                $g.DrawString($tLbl,$txtFont,$txtBrush,$xPx-18,$H-$padB+4)" & vbCrLf
ps = ps & "            }" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        $toBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(35,220,0,0))" & vbCrLf
ps = ps & "        foreach ($pt in $pts) {" & vbCrLf
ps = ps & "            if ($pt.Ms -ge $PingTimeoutMs) {" & vbCrLf
ps = ps & "                $frac=if($tSpan -gt 0){($pt.Time-$t0).TotalSeconds/$tSpan}else{0}" & vbCrLf
ps = ps & "                $xPx=$padL+[int]($frac*$cW)" & vbCrLf
ps = ps & "                $g.FillRectangle($toBrush,$xPx-1,$padT,3,$cH)" & vbCrLf
ps = ps & "            }" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0,120,212), 1.8)" & vbCrLf
ps = ps & "        $prevX=-1; $prevY=-1" & vbCrLf
ps = ps & "        for ($i=0; $i -lt $pts.Count; $i++) {" & vbCrLf
ps = ps & "            $frac=if($tSpan -gt 0){($pts[$i].Time-$t0).TotalSeconds/$tSpan}else{0.5}" & vbCrLf
ps = ps & "            $xPx=$padL+[int]($frac*$cW)" & vbCrLf
ps = ps & "            $yFrac=if($range -gt 0){($pts[$i].Ms-$minV)/$range}else{0.5}" & vbCrLf
ps = ps & "            $yPx=$padT+$cH-[int]($yFrac*$cH)" & vbCrLf
ps = ps & "            if ($prevX -ge 0){ $g.DrawLine($linePen,$prevX,$prevY,$xPx,$yPx) }" & vbCrLf
ps = ps & "            $prevX=$xPx; $prevY=$yPx" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        $borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(180,180,180))" & vbCrLf
ps = ps & "        $g.DrawRectangle($borderPen,$padL,$padT,$cW,$cH)" & vbCrLf
ps = ps & "        # Title" & vbCrLf
ps = ps & "        $titleFont = New-Object System.Drawing.Font(""Segoe UI"", 8, [System.Drawing.FontStyle]::Bold)" & vbCrLf
ps = ps & "        $g.DrawString($title, $titleFont, [System.Drawing.Brushes]::DimGray, $padL, 2)" & vbCrLf
ps = ps & "        $g.Dispose()" & vbCrLf
ps = ps & "        if ($canvas.Image) { $canvas.Image.Dispose() }" & vbCrLf
ps = ps & "        $canvas.Image = $bmp" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $cmb.Add_SelectedIndexChanged({ Draw-Chart })" & vbCrLf
ps = ps & "    $canvas.Add_Resize({ Draw-Chart })" & vbCrLf
ps = ps & "    $cf.Add_Shown({ Draw-Chart })" & vbCrLf
ps = ps & "    [void]$cf.ShowDialog(); $cf.Dispose()" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Right-click context menu --------------------------------------------------" & vbCrLf
ps = ps & "$CtxMenu      = New-Object System.Windows.Forms.ContextMenuStrip" & vbCrLf
ps = ps & "$CtxTagEth    = New-Object System.Windows.Forms.ToolStripMenuItem; $CtxTagEth.Text  = ""[=]  Tag as Ethernet""" & vbCrLf
ps = ps & "$CtxTagWifi   = New-Object System.Windows.Forms.ToolStripMenuItem; $CtxTagWifi.Text = ""[~]  Tag as WiFi""" & vbCrLf
ps = ps & "$CtxSep1      = New-Object System.Windows.Forms.ToolStripSeparator" & vbCrLf
ps = ps & "$CtxLabel     = New-Object System.Windows.Forms.ToolStripMenuItem; $CtxLabel.Text   = ""     Set Label...""" & vbCrLf
ps = ps & "$CtxClearLabel= New-Object System.Windows.Forms.ToolStripMenuItem; $CtxClearLabel.Text = ""     Clear Label""" & vbCrLf
ps = ps & "$CtxSep2      = New-Object System.Windows.Forms.ToolStripSeparator" & vbCrLf
ps = ps & "$CtxChart     = New-Object System.Windows.Forms.ToolStripMenuItem; $CtxChart.Text   = ""     View Latency Chart""" & vbCrLf
ps = ps & "$CtxSep3      = New-Object System.Windows.Forms.ToolStripSeparator" & vbCrLf
ps = ps & "$CtxRemove    = New-Object System.Windows.Forms.ToolStripMenuItem; $CtxRemove.Text  = ""     Remove from Monitor""" & vbCrLf
ps = ps & "$CtxRemove.ForeColor = [System.Drawing.Color]::Firebrick" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "foreach ($item in @($CtxTagEth,$CtxTagWifi,$CtxSep1,$CtxLabel,$CtxClearLabel,$CtxSep2,$CtxChart,$CtxSep3,$CtxRemove)) {" & vbCrLf
ps = ps & "    [void]$CtxMenu.Items.Add($item)" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Apply-IfTypeTag ($ip, $type) {" & vbCrLf
ps = ps & "    if ($Global:NetworkStats.ContainsKey($ip))  { $Global:NetworkStats[$ip].IfType = $type }" & vbCrLf
ps = ps & "    if ($Global:UIReferences.ContainsKey($ip)) {" & vbCrLf
ps = ps & "        $ui = $Global:UIReferences[$ip]" & vbCrLf
ps = ps & "        $ui.Type.Text = Get-IfTypeIcon $type; $ui.Type.ForeColor = Get-IfTypeColor $type" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    Apply-Filter" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$CtxTagEth.Add_Click({" & vbCrLf
ps = ps & "    if ($ListView.SelectedItems.Count -gt 0) { Apply-IfTypeTag $ListView.SelectedItems[0].Text ""Ethernet"" }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "$CtxTagWifi.Add_Click({" & vbCrLf
ps = ps & "    if ($ListView.SelectedItems.Count -gt 0) { Apply-IfTypeTag $ListView.SelectedItems[0].Text ""WiFi"" }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$CtxLabel.Add_Click({" & vbCrLf
ps = ps & "    if ($ListView.SelectedItems.Count -eq 0) { return }" & vbCrLf
ps = ps & "    $ip  = $ListView.SelectedItems[0].Text" & vbCrLf
ps = ps & "    $cur = if ($Global:DeviceLabels.ContainsKey($ip)) { $Global:DeviceLabels[$ip] } else { """" }" & vbCrLf
ps = ps & "    # Simple input dialog" & vbCrLf
ps = ps & "    $dlg = New-Object System.Windows.Forms.Form" & vbCrLf
ps = ps & "    $dlg.Text = ""Set Label for $ip""; $dlg.Size = New-Object System.Drawing.Size(360, 130)" & vbCrLf
ps = ps & "    $dlg.StartPosition = ""CenterParent""; $dlg.FormBorderStyle = ""FixedDialog""" & vbCrLf
ps = ps & "    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false" & vbCrLf
ps = ps & "    $tb = New-Object System.Windows.Forms.TextBox" & vbCrLf
ps = ps & "    $tb.Text = $cur; $tb.Font = New-Object System.Drawing.Font(""Segoe UI"", 10)" & vbCrLf
ps = ps & "    $tb.Location = New-Object System.Drawing.Point(12,12); $tb.Size = New-Object System.Drawing.Size(320,24)" & vbCrLf
ps = ps & "    [void]$dlg.Controls.Add($tb)" & vbCrLf
ps = ps & "    $ok = New-Object System.Windows.Forms.Button; $ok.Text = ""OK""" & vbCrLf
ps = ps & "    $ok.Location = New-Object System.Drawing.Point(175,50); $ok.Size = New-Object System.Drawing.Size(75,28)" & vbCrLf
ps = ps & "    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK; [void]$dlg.Controls.Add($ok)" & vbCrLf
ps = ps & "    $cn = New-Object System.Windows.Forms.Button; $cn.Text = ""Cancel""" & vbCrLf
ps = ps & "    $cn.Location = New-Object System.Drawing.Point(257,50); $cn.Size = New-Object System.Drawing.Size(75,28)" & vbCrLf
ps = ps & "    $cn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; [void]$dlg.Controls.Add($cn)" & vbCrLf
ps = ps & "    $dlg.AcceptButton = $ok; $dlg.CancelButton = $cn" & vbCrLf
ps = ps & "    if ($dlg.ShowDialog() -eq ""OK"") {" & vbCrLf
ps = ps & "        $lbl = $tb.Text.Trim()" & vbCrLf
ps = ps & "        $Global:DeviceLabels[$ip] = $lbl" & vbCrLf
ps = ps & "        if ($Global:UIReferences.ContainsKey($ip)) {" & vbCrLf
ps = ps & "            $Global:UIReferences[$ip].Label.Text = $lbl" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $dlg.Dispose()" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$CtxClearLabel.Add_Click({" & vbCrLf
ps = ps & "    if ($ListView.SelectedItems.Count -eq 0) { return }" & vbCrLf
ps = ps & "    $ip = $ListView.SelectedItems[0].Text" & vbCrLf
ps = ps & "    $Global:DeviceLabels[$ip] = """"" & vbCrLf
ps = ps & "    if ($Global:UIReferences.ContainsKey($ip)) { $Global:UIReferences[$ip].Label.Text = """" }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$CtxChart.Add_Click({" & vbCrLf
ps = ps & "    if ($ListView.SelectedItems.Count -gt 0) { Show-LatencyChart $ListView.SelectedItems[0].Text }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$CtxRemove.Add_Click({" & vbCrLf
ps = ps & "    if ($ListView.SelectedItems.Count -eq 0) { return }" & vbCrLf
ps = ps & "    $ip  = $ListView.SelectedItems[0].Text" & vbCrLf
ps = ps & "    $lvi = $ListView.SelectedItems[0]" & vbCrLf
ps = ps & "    [void]$ListView.Items.Remove($lvi)" & vbCrLf
ps = ps & "    [void]$Global:AllItems.Remove($lvi)" & vbCrLf
ps = ps & "    if ($Global:NetworkStats.ContainsKey($ip))  { $Global:NetworkStats.Remove($ip) }" & vbCrLf
ps = ps & "    if ($Global:UIReferences.ContainsKey($ip))  { $Global:UIReferences.Remove($ip) }" & vbCrLf
ps = ps & "    if ($Global:DeviceLabels.ContainsKey($ip))  { $Global:DeviceLabels.Remove($ip) }" & vbCrLf
ps = ps & "    if ($Global:DeviceAlerts.ContainsKey($ip))  { $Global:DeviceAlerts.Remove($ip) }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$ListView.Add_MouseDown({" & vbCrLf
ps = ps & "    param($s,$e)" & vbCrLf
ps = ps & "    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {" & vbCrLf
ps = ps & "        $hit = $ListView.HitTest($e.Location)" & vbCrLf
ps = ps & "        if ($hit.Item -ne $null) {" & vbCrLf
ps = ps & "            $hit.Item.Selected = $true" & vbCrLf
ps = ps & "            $ip = $hit.Item.Text" & vbCrLf
ps = ps & "            $hasData = $Global:NetworkStats.ContainsKey($ip) -and $Global:NetworkStats[$ip].TimestampedHistory.Count -ge 2" & vbCrLf
ps = ps & "            $CtxChart.Enabled = $hasData" & vbCrLf
ps = ps & "            $hasLabel = $Global:DeviceLabels.ContainsKey($ip) -and $Global:DeviceLabels[$ip]" & vbCrLf
ps = ps & "            $CtxClearLabel.Enabled = $hasLabel" & vbCrLf
ps = ps & "            $CtxMenu.Show($ListView, $e.Location)" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$ListView.Add_DoubleClick({" & vbCrLf
ps = ps & "    if ($ListView.SelectedItems.Count -gt 0) { Show-LatencyChart $ListView.SelectedItems[0].Text }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Save / Load session -------------------------------------------------------" & vbCrLf
ps = ps & "$BtnSave.Add_Click({" & vbCrLf
ps = ps & "    $dlg = New-Object System.Windows.Forms.SaveFileDialog" & vbCrLf
ps = ps & "    $dlg.Filter = ""YAB Session (*.yabsession)|*.yabsession|All Files (*.*)|*.*""" & vbCrLf
ps = ps & "    $dlg.Title  = ""Save Session""" & vbCrLf
ps = ps & "    $dlg.FileName = ""NetworkSession_"" + [DateTime]::Now.ToString(""yyyyMMdd_HHmmss"")" & vbCrLf
ps = ps & "    if ($dlg.ShowDialog() -ne ""OK"") { return }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $export = @{}" & vbCrLf
ps = ps & "    foreach ($ip in $Global:NetworkStats.Keys) {" & vbCrLf
ps = ps & "        $s = $Global:NetworkStats[$ip]" & vbCrLf
ps = ps & "        $pts = @($s.TimestampedHistory | ForEach-Object {" & vbCrLf
ps = ps & "            @{ T = $_.Time.ToString(""o""); M = $_.Ms }" & vbCrLf
ps = ps & "        })" & vbCrLf
ps = ps & "        $export[$ip] = @{" & vbCrLf
ps = ps & "            IfType   = $s.IfType" & vbCrLf
ps = ps & "            DnsName  = $s.DnsName" & vbCrLf
ps = ps & "            Timeouts = $s.Timeouts" & vbCrLf
ps = ps & "            Label    = if ($Global:DeviceLabels.ContainsKey($ip)) { $Global:DeviceLabels[$ip] } else { """" }" & vbCrLf
ps = ps & "            History  = $pts" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $export | ConvertTo-Json -Depth 5 | Set-Content -Path $dlg.FileName -Encoding UTF8" & vbCrLf
ps = ps & "    [System.Windows.Forms.MessageBox]::Show(""Session saved: $($dlg.FileName)"", ""Saved"", 0, 64)" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Load-DeviceFromData ($ip, $data) {" & vbCrLf
ps = ps & "    $ifType  = if ($data.IfType)  { $data.IfType }  else { ""Other"" }" & vbCrLf
ps = ps & "    $dnsName = if ($data.DnsName) { $data.DnsName } else { ""Unknown"" }" & vbCrLf
ps = ps & "    $label   = if ($data.Label)   { $data.Label }   else { """" }" & vbCrLf
ps = ps & "    $toCount = if ($data.Timeouts) { [int]$data.Timeouts } else { 0 }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $histList  = [System.Collections.Generic.List[long]]::new()" & vbCrLf
ps = ps & "    $tsHistory = [System.Collections.Generic.List[hashtable]]::new()" & vbCrLf
ps = ps & "    foreach ($pt in $data.History) {" & vbCrLf
ps = ps & "        try {" & vbCrLf
ps = ps & "            $t = [DateTime]::Parse($pt.T)" & vbCrLf
ps = ps & "            $m = [long]$pt.M" & vbCrLf
ps = ps & "            [void]$histList.Add($m)" & vbCrLf
ps = ps & "            [void]$tsHistory.Add(@{ Time=$t; Ms=$m })" & vbCrLf
ps = ps & "        } catch {}" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $Global:NetworkStats[$ip] = @{" & vbCrLf
ps = ps & "        History            = $histList" & vbCrLf
ps = ps & "        TimestampedHistory = $tsHistory" & vbCrLf
ps = ps & "        Timeouts           = $toCount" & vbCrLf
ps = ps & "        IfType             = $ifType" & vbCrLf
ps = ps & "        DnsName            = $dnsName" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $Global:DeviceLabels[$ip] = $label" & vbCrLf
ps = ps & "    $Global:DeviceAlerts[$ip] = @{ WasUp=$true }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    # Build list view row" & vbCrLf
ps = ps & "    $lvi      = New-Object System.Windows.Forms.ListViewItem($ip)" & vbCrLf
ps = ps & "    $labelCell= $lvi.SubItems.Add($label)" & vbCrLf
ps = ps & "    $nameCell = $lvi.SubItems.Add($dnsName)" & vbCrLf
ps = ps & "    $typeCell = $lvi.SubItems.Add((Get-IfTypeIcon $ifType))" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $vals = @($histList); $lastMs = if ($vals.Count -gt 0) { $vals[-1] } else { 0 }" & vbCrLf
ps = ps & "    $realVals = @($vals | Where-Object { $_ -lt $PingTimeoutMs })" & vbCrLf
ps = ps & "    $minMs = if ($realVals.Count -gt 0) { ($realVals | Measure-Object -Minimum).Minimum } else { 0 }" & vbCrLf
ps = ps & "    $maxMs = if ($realVals.Count -gt 0) { ($realVals | Measure-Object -Maximum).Maximum } else { 0 }" & vbCrLf
ps = ps & "    $avgMs = if ($realVals.Count -gt 0) { [Math]::Round(($realVals | Measure-Object -Average).Average,1) } else { 0 }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $curCell  = $lvi.SubItems.Add(""$lastMs ms"")" & vbCrLf
ps = ps & "    $minCell  = $lvi.SubItems.Add(""$minMs ms"")" & vbCrLf
ps = ps & "    $maxCell  = $lvi.SubItems.Add(""$maxMs ms"")" & vbCrLf
ps = ps & "    $avgCell  = $lvi.SubItems.Add(""$avgMs ms"")" & vbCrLf
ps = ps & "    $toCell   = $lvi.SubItems.Add($toCount)" & vbCrLf
ps = ps & "    $sparkCell= $lvi.SubItems.Add((Get-Sparkline $histList))" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $lvi.UseItemStyleForSubItems = $false" & vbCrLf
ps = ps & "    $curCell.ForeColor  = Get-LatencyColor $lastMs" & vbCrLf
ps = ps & "    $typeCell.ForeColor = Get-IfTypeColor $ifType" & vbCrLf
ps = ps & "    $nameCell.ForeColor = [System.Drawing.Color]::Black" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    [void]$Global:AllItems.Add($lvi)" & vbCrLf
ps = ps & "    $Global:UIReferences[$ip] = @{" & vbCrLf
ps = ps & "        Current=$curCell; Min=$minCell; Max=$maxCell; Avg=$avgCell" & vbCrLf
ps = ps & "        Timeout=$toCell; Name=$nameCell; Type=$typeCell; Spark=$sparkCell" & vbCrLf
ps = ps & "        Label=$labelCell; Item=$lvi" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$BtnLoad.Add_Click({" & vbCrLf
ps = ps & "    $dlg = New-Object System.Windows.Forms.OpenFileDialog" & vbCrLf
ps = ps & "    $dlg.Filter = ""YAB Session (*.yabsession)|*.yabsession|All Files (*.*)|*.*""" & vbCrLf
ps = ps & "    $dlg.Title  = ""Load Session""" & vbCrLf
ps = ps & "    if ($dlg.ShowDialog() -ne ""OK"") { return }" & vbCrLf
ps = ps & "    try {" & vbCrLf
ps = ps & "        $json = Get-Content -Path $dlg.FileName -Encoding UTF8 -Raw | ConvertFrom-Json" & vbCrLf
ps = ps & "        # Clear existing" & vbCrLf
ps = ps & "        $ListView.Items.Clear(); $Global:AllItems.Clear()" & vbCrLf
ps = ps & "        $Global:NetworkStats.Clear(); $Global:UIReferences.Clear()" & vbCrLf
ps = ps & "        $Global:DeviceLabels.Clear(); $Global:DeviceAlerts.Clear()" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "        foreach ($prop in $json.PSObject.Properties) {" & vbCrLf
ps = ps & "            $ip   = $prop.Name" & vbCrLf
ps = ps & "            $data = $prop.Value" & vbCrLf
ps = ps & "            Load-DeviceFromData $ip $data" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        Apply-Filter" & vbCrLf
ps = ps & "        $Global:ScanComplete = $true" & vbCrLf
ps = ps & "        $StartButton.Text = ""Rescan""" & vbCrLf
ps = ps & "        $StartButton.BackColor = [System.Drawing.Color]::FromArgb(16,124,16)" & vbCrLf
ps = ps & "        $RuntimeLabel.Text = ""Session loaded: $($Global:UIReferences.Count) devices. Monitoring active.""" & vbCrLf
ps = ps & "        $RuntimeLabel.ForeColor = [System.Drawing.Color]::DarkGreen" & vbCrLf
ps = ps & "        $Timer.Interval = $Global:RefreshMs; $Timer.Start()" & vbCrLf
ps = ps & "    } catch {" & vbCrLf
ps = ps & "        [System.Windows.Forms.MessageBox]::Show(""Failed to load session:`n$_"", ""Error"", 0, 16)" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Export CSV ----------------------------------------------------------------" & vbCrLf
ps = ps & "$BtnExpCSV.Add_Click({" & vbCrLf
ps = ps & "    $dlg = New-Object System.Windows.Forms.SaveFileDialog" & vbCrLf
ps = ps & "    $dlg.Filter = ""CSV (*.csv)|*.csv|All Files (*.*)|*.*""" & vbCrLf
ps = ps & "    $dlg.FileName = ""NetworkExport_"" + [DateTime]::Now.ToString(""yyyyMMdd_HHmmss"")" & vbCrLf
ps = ps & "    if ($dlg.ShowDialog() -ne ""OK"") { return }" & vbCrLf
ps = ps & "    $sb = New-Object System.Text.StringBuilder" & vbCrLf
ps = ps & "    [void]$sb.AppendLine(""IP,Label,DNS Name,Type,Min(ms),Max(ms),Avg(ms),Timeouts,Samples,First Seen,Last Seen"")" & vbCrLf
ps = ps & "    foreach ($ip in ($Global:NetworkStats.Keys | Sort-Object)) {" & vbCrLf
ps = ps & "        $s = $Global:NetworkStats[$ip]" & vbCrLf
ps = ps & "        $label = if ($Global:DeviceLabels.ContainsKey($ip)) { $Global:DeviceLabels[$ip] } else { """" }" & vbCrLf
ps = ps & "        $realH = @($s.History | Where-Object { $_ -lt $PingTimeoutMs })" & vbCrLf
ps = ps & "        $minV = if ($realH.Count -gt 0) { ($realH | Measure-Object -Minimum).Minimum } else { ""N/A"" }" & vbCrLf
ps = ps & "        $maxV = if ($realH.Count -gt 0) { ($realH | Measure-Object -Maximum).Maximum } else { ""N/A"" }" & vbCrLf
ps = ps & "        $avgV = if ($realH.Count -gt 0) { [Math]::Round(($realH | Measure-Object -Average).Average,1) } else { ""N/A"" }" & vbCrLf
ps = ps & "        $first = if ($s.TimestampedHistory.Count -gt 0) { $s.TimestampedHistory[0].Time.ToString(""yyyy-MM-dd HH:mm:ss"") } else { """" }" & vbCrLf
ps = ps & "        $last  = if ($s.TimestampedHistory.Count -gt 0) { $s.TimestampedHistory[-1].Time.ToString(""yyyy-MM-dd HH:mm:ss"") } else { """" }" & vbCrLf
ps = ps & "        $row = ""$ip,`""$label`"",`""$($s.DnsName)`"",$($s.IfType),$minV,$maxV,$avgV,$($s.Timeouts),$($s.History.Count),`""$first`"",`""$last`""""" & vbCrLf
ps = ps & "        [void]$sb.AppendLine($row)" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $sb.ToString() | Set-Content -Path $dlg.FileName -Encoding UTF8" & vbCrLf
ps = ps & "    [System.Windows.Forms.MessageBox]::Show(""CSV exported: $($dlg.FileName)"", ""Done"", 0, 64)" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Export HTML ---------------------------------------------------------------" & vbCrLf
ps = ps & "$BtnExpHTML.Add_Click({" & vbCrLf
ps = ps & "    $dlg = New-Object System.Windows.Forms.SaveFileDialog" & vbCrLf
ps = ps & "    $dlg.Filter = ""HTML Report (*.html)|*.html|All Files (*.*)|*.*""" & vbCrLf
ps = ps & "    $dlg.FileName = ""NetworkReport_"" + [DateTime]::Now.ToString(""yyyyMMdd_HHmmss"")" & vbCrLf
ps = ps & "    if ($dlg.ShowDialog() -ne ""OK"") { return }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    # Build per-device JSON for Chart.js" & vbCrLf
ps = ps & "    $deviceJsonList = [System.Collections.Generic.List[string]]::new()" & vbCrLf
ps = ps & "    foreach ($ip in ($Global:NetworkStats.Keys | Sort-Object)) {" & vbCrLf
ps = ps & "        $s     = $Global:NetworkStats[$ip]" & vbCrLf
ps = ps & "        $label = if ($Global:DeviceLabels.ContainsKey($ip) -and $Global:DeviceLabels[$ip]) { $Global:DeviceLabels[$ip] } else { $s.DnsName }" & vbCrLf
ps = ps & "        $realH = @($s.History | Where-Object { $_ -lt $PingTimeoutMs })" & vbCrLf
ps = ps & "        $minV  = if ($realH.Count -gt 0) { ($realH | Measure-Object -Minimum).Minimum } else { 0 }" & vbCrLf
ps = ps & "        $maxV  = if ($realH.Count -gt 0) { ($realH | Measure-Object -Maximum).Maximum } else { 0 }" & vbCrLf
ps = ps & "        $avgV  = if ($realH.Count -gt 0) { [Math]::Round(($realH | Measure-Object -Average).Average,1) } else { 0 }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "        # Downsample to max 500 points for HTML" & vbCrLf
ps = ps & "        $tsPts = @($s.TimestampedHistory)" & vbCrLf
ps = ps & "        $step  = [Math]::Max(1, [int]($tsPts.Count / 500))" & vbCrLf
ps = ps & "        $tLabels = [System.Collections.Generic.List[string]]::new()" & vbCrLf
ps = ps & "        $tVals   = [System.Collections.Generic.List[string]]::new()" & vbCrLf
ps = ps & "        for ($i=0; $i -lt $tsPts.Count; $i += $step) {" & vbCrLf
ps = ps & "            [void]$tLabels.Add('""' + $tsPts[$i].Time.ToString(""HH:mm:ss"") + '""')" & vbCrLf
ps = ps & "            $mv = if ($tsPts[$i].Ms -ge $PingTimeoutMs) { ""null"" } else { $tsPts[$i].Ms.ToString() }" & vbCrLf
ps = ps & "            [void]$tVals.Add($mv)" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "        $safeName = $ip -replace '\.','_'" & vbCrLf
ps = ps & "        $dispName = ($label -replace '""','') + "" ("" + $ip + "")""" & vbCrLf
ps = ps & "        $deviceJsonList.Add(@""" & vbCrLf
ps = ps & "  {" & vbCrLf
ps = ps & "    ""id"": ""$safeName""," & vbCrLf
ps = ps & "    ""ip"": ""$ip""," & vbCrLf
ps = ps & "    ""name"": ""$dispName""," & vbCrLf
ps = ps & "    ""dns"": ""$($s.DnsName)""," & vbCrLf
ps = ps & "    ""type"": ""$($s.IfType)""," & vbCrLf
ps = ps & "    ""min"": $minV," & vbCrLf
ps = ps & "    ""max"": $maxV," & vbCrLf
ps = ps & "    ""avg"": $avgV," & vbCrLf
ps = ps & "    ""timeouts"": $($s.Timeouts)," & vbCrLf
ps = ps & "    ""samples"": $($s.History.Count)," & vbCrLf
ps = ps & "    ""labels"": [$($tLabels -join ',')]," & vbCrLf
ps = ps & "    ""data"": [$($tVals -join ',')]" & vbCrLf
ps = ps & "  }" & vbCrLf
ps = ps & """@)" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    $devicesJson = ""["" + ($deviceJsonList -join "",`n"") + ""]""" & vbCrLf
ps = ps & "    $generatedAt = [DateTime]::Now.ToString(""yyyy-MM-dd HH:mm:ss"")" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $html = @""" & vbCrLf
ps = ps & "<!DOCTYPE html>" & vbCrLf
ps = ps & "<html lang=""en"">" & vbCrLf
ps = ps & "<head>" & vbCrLf
ps = ps & "<meta charset=""UTF-8"">" & vbCrLf
ps = ps & "<title>YAB Network Report - $generatedAt</title>" & vbCrLf
ps = ps & "<script src=""https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js""></script>" & vbCrLf
ps = ps & "<style>" & vbCrLf
ps = ps & "  *{box-sizing:border-box;margin:0;padding:0}" & vbCrLf
ps = ps & "  body{font-family:'Segoe UI',sans-serif;background:#f4f4f8;color:#222}" & vbCrLf
ps = ps & "  header{background:#1a237e;color:#fff;padding:18px 28px}" & vbCrLf
ps = ps & "  header h1{font-size:1.4rem;font-weight:600}" & vbCrLf
ps = ps & "  header p{font-size:.85rem;opacity:.8;margin-top:4px}" & vbCrLf
ps = ps & "  .wrap{max-width:1200px;margin:0 auto;padding:20px}" & vbCrLf
ps = ps & "  h2{font-size:1.1rem;margin:20px 0 10px;color:#1a237e}" & vbCrLf
ps = ps & "  table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.1)}" & vbCrLf
ps = ps & "  th{background:#3949ab;color:#fff;padding:9px 12px;text-align:left;font-size:.82rem}" & vbCrLf
ps = ps & "  td{padding:8px 12px;font-size:.85rem;border-bottom:1px solid #eee}" & vbCrLf
ps = ps & "  tr:last-child td{border-bottom:none}" & vbCrLf
ps = ps & "  tr:hover td{background:#f0f4ff}" & vbCrLf
ps = ps & "  .badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:.78rem;font-weight:600}" & vbCrLf
ps = ps & "  .eth{background:#e8f5e9;color:#2e7d32}" & vbCrLf
ps = ps & "  .wifi{background:#e3f2fd;color:#1565c0}" & vbCrLf
ps = ps & "  .good{color:#2e7d32;font-weight:600}" & vbCrLf
ps = ps & "  .warn{color:#e65100;font-weight:600}" & vbCrLf
ps = ps & "  .bad{color:#c62828;font-weight:600}" & vbCrLf
ps = ps & "  .chart-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(480px,1fr));gap:18px;margin-top:12px}" & vbCrLf
ps = ps & "  .chart-card{background:#fff;border-radius:8px;padding:14px;box-shadow:0 1px 4px rgba(0,0,0,.1)}" & vbCrLf
ps = ps & "  .chart-card h3{font-size:.9rem;color:#333;margin-bottom:8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}" & vbCrLf
ps = ps & "  .chart-card .meta{font-size:.78rem;color:#777;margin-bottom:8px}" & vbCrLf
ps = ps & "  canvas{width:100%!important;height:180px!important}" & vbCrLf
ps = ps & "  .filter-bar{margin:14px 0 8px;display:flex;gap:8px;flex-wrap:wrap}" & vbCrLf
ps = ps & "  .filter-btn{padding:5px 14px;border:1px solid #3949ab;border-radius:14px;background:#fff;color:#3949ab;cursor:pointer;font-size:.83rem}" & vbCrLf
ps = ps & "  .filter-btn.active{background:#3949ab;color:#fff}" & vbCrLf
ps = ps & "</style>" & vbCrLf
ps = ps & "</head>" & vbCrLf
ps = ps & "<body>" & vbCrLf
ps = ps & "<header>" & vbCrLf
ps = ps & "  <h1>YAB Network Monitor - Session Report</h1>" & vbCrLf
ps = ps & "  <p>Generated: $generatedAt</p>" & vbCrLf
ps = ps & "</header>" & vbCrLf
ps = ps & "<div class=""wrap"">" & vbCrLf
ps = ps & "  <h2>Summary</h2>" & vbCrLf
ps = ps & "  <table id=""summaryTable"">" & vbCrLf
ps = ps & "    <thead><tr><th>IP</th><th>Label / Name</th><th>Type</th><th>Min</th><th>Max</th><th>Avg</th><th>Timeouts</th><th>Samples</th></tr></thead>" & vbCrLf
ps = ps & "    <tbody id=""summaryBody""></tbody>" & vbCrLf
ps = ps & "  </table>" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "  <h2>Latency Timelines</h2>" & vbCrLf
ps = ps & "  <div class=""filter-bar"">" & vbCrLf
ps = ps & "    <button class=""filter-btn active"" onclick=""filterCharts('all',this)"">All</button>" & vbCrLf
ps = ps & "    <button class=""filter-btn"" onclick=""filterCharts('Ethernet',this)"">Ethernet</button>" & vbCrLf
ps = ps & "    <button class=""filter-btn"" onclick=""filterCharts('WiFi',this)"">WiFi</button>" & vbCrLf
ps = ps & "  </div>" & vbCrLf
ps = ps & "  <div class=""chart-grid"" id=""chartGrid""></div>" & vbCrLf
ps = ps & "</div>" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "<script>" & vbCrLf
ps = ps & "const devices = $devicesJson;" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function latencyClass(v){" & vbCrLf
ps = ps & "  if(v===null||v===undefined) return 'bad';" & vbCrLf
ps = ps & "  if(v>10) return 'warn';" & vbCrLf
ps = ps & "  return 'good';" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "// Summary table" & vbCrLf
ps = ps & "const tbody = document.getElementById('summaryBody');" & vbCrLf
ps = ps & "devices.forEach(d => {" & vbCrLf
ps = ps & "  const tr = document.createElement('tr');" & vbCrLf
ps = ps & "  tr.dataset.type = d.type;" & vbCrLf
ps = ps & "  const badge = d.type==='WiFi'" & vbCrLf
ps = ps & "    ? '<span class=""badge wifi"">[~] WiFi</span>'" & vbCrLf
ps = ps & "    : '<span class=""badge eth"">[=] Ethernet</span>';" & vbCrLf
ps = ps & "  tr.innerHTML = [" & vbCrLf
ps = ps & "    '<td>'+d.ip+'</td>'," & vbCrLf
ps = ps & "    '<td>'+d.name+'</td>'," & vbCrLf
ps = ps & "    '<td>'+badge+'</td>'," & vbCrLf
ps = ps & "    '<td class=""'+latencyClass(d.min)+'"">'+d.min+'ms</td>'," & vbCrLf
ps = ps & "    '<td class=""'+latencyClass(d.max)+'"">'+d.max+'ms</td>'," & vbCrLf
ps = ps & "    '<td class=""'+latencyClass(d.avg)+'"">'+d.avg+'ms</td>'," & vbCrLf
ps = ps & "    '<td class=""'+(d.timeouts>0?'bad':'good')+'"">'+d.timeouts+'</td>'," & vbCrLf
ps = ps & "    '<td>'+d.samples+'</td>'" & vbCrLf
ps = ps & "  ].join('');" & vbCrLf
ps = ps & "  tbody.appendChild(tr);" & vbCrLf
ps = ps & "});" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "// Charts" & vbCrLf
ps = ps & "const chartGrid = document.getElementById('chartGrid');" & vbCrLf
ps = ps & "const charts = {};" & vbCrLf
ps = ps & "devices.forEach(d => {" & vbCrLf
ps = ps & "  const card = document.createElement('div');" & vbCrLf
ps = ps & "  card.className = 'chart-card';" & vbCrLf
ps = ps & "  card.dataset.type = d.type;" & vbCrLf
ps = ps & "  card.id = 'card_'+d.id;" & vbCrLf
ps = ps & "  const toRate = d.samples>0 ? ((d.timeouts/d.samples)*100).toFixed(1) : 0;" & vbCrLf
ps = ps & "  card.innerHTML =" & vbCrLf
ps = ps & "    '<h3>'+d.name+'</h3>'+" & vbCrLf
ps = ps & "    '<div class=""meta"">Min: '+d.min+'ms &nbsp; Max: '+d.max+'ms &nbsp; Avg: '+d.avg+'ms &nbsp; Timeouts: '+d.timeouts+' ('+toRate+'%)</div>'+" & vbCrLf
ps = ps & "    '<canvas id=""c_'+d.id+'""></canvas>';" & vbCrLf
ps = ps & "  chartGrid.appendChild(card);" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "  const ctx = document.getElementById('c_'+d.id).getContext('2d');" & vbCrLf
ps = ps & "  charts[d.id] = new Chart(ctx, {" & vbCrLf
ps = ps & "    type: 'line'," & vbCrLf
ps = ps & "    data: {" & vbCrLf
ps = ps & "      labels: d.labels," & vbCrLf
ps = ps & "      datasets: [{" & vbCrLf
ps = ps & "        data: d.data," & vbCrLf
ps = ps & "        borderColor: 'rgba(0,120,212,0.85)'," & vbCrLf
ps = ps & "        backgroundColor: 'rgba(0,120,212,0.08)'," & vbCrLf
ps = ps & "        borderWidth: 1.5," & vbCrLf
ps = ps & "        pointRadius: 0," & vbCrLf
ps = ps & "        spanGaps: false," & vbCrLf
ps = ps & "        tension: 0.2" & vbCrLf
ps = ps & "      }]" & vbCrLf
ps = ps & "    }," & vbCrLf
ps = ps & "    options: {" & vbCrLf
ps = ps & "      animation: false," & vbCrLf
ps = ps & "      responsive: true," & vbCrLf
ps = ps & "      maintainAspectRatio: false," & vbCrLf
ps = ps & "      plugins: { legend: { display: false } }," & vbCrLf
ps = ps & "      scales: {" & vbCrLf
ps = ps & "        x: { display: true, ticks: { maxTicksLimit: 6, font:{size:10} }, grid:{color:'#eee'} }," & vbCrLf
ps = ps & "        y: { display: true, ticks: { font:{size:10}, callback: v => v===null?'T/O':v+'ms' }," & vbCrLf
ps = ps & "             grid:{color:'#eee'}, min: 0 }" & vbCrLf
ps = ps & "      }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "  });" & vbCrLf
ps = ps & "});" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function filterCharts(type, btn) {" & vbCrLf
ps = ps & "  document.querySelectorAll('.filter-btn').forEach(b=>b.classList.remove('active'));" & vbCrLf
ps = ps & "  btn.classList.add('active');" & vbCrLf
ps = ps & "  document.querySelectorAll('#chartGrid .chart-card, #summaryBody tr').forEach(el=>{" & vbCrLf
ps = ps & "    el.style.display = (type==='all'||el.dataset.type===type) ? '' : 'none';" & vbCrLf
ps = ps & "  });" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "</script>" & vbCrLf
ps = ps & "</body>" & vbCrLf
ps = ps & "</html>" & vbCrLf
ps = ps & """@" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $html | Set-Content -Path $dlg.FileName -Encoding UTF8" & vbCrLf
ps = ps & "    [System.Windows.Forms.MessageBox]::Show(""HTML report saved:`n$($dlg.FileName)`n`nOpen in any browser."", ""Done"", 0, 64)" & vbCrLf
ps = ps & "    Start-Process $dlg.FileName" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Async helpers -------------------------------------------------------------" & vbCrLf
ps = ps & "function Start-AsyncPing ($IP) {" & vbCrLf
ps = ps & "    $rs = [runspacefactory]::CreateRunspace()" & vbCrLf
ps = ps & "    $rs.ApartmentState = ""STA""; $rs.ThreadOptions = ""ReuseThread""; $rs.Open()" & vbCrLf
ps = ps & "    $ps = [PowerShell]::Create(); $ps.Runspace = $rs" & vbCrLf
ps = ps & "    [void]$ps.AddScript({" & vbCrLf
ps = ps & "        param($ip,$tms)" & vbCrLf
ps = ps & "        try {" & vbCrLf
ps = ps & "            $p = New-Object System.Net.NetworkInformation.Ping" & vbCrLf
ps = ps & "            $r = $p.Send($ip,$tms); $p.Dispose()" & vbCrLf
ps = ps & "            [pscustomobject]@{ IP=$ip; Status=$r.Status; RTT=$r.RoundtripTime }" & vbCrLf
ps = ps & "        } catch {" & vbCrLf
ps = ps & "            [pscustomobject]@{ IP=$ip; Status=""Error""; RTT=0 }" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "    }).AddArgument($IP).AddArgument($PingTimeoutMs)" & vbCrLf
ps = ps & "    $handle = $ps.BeginInvoke()" & vbCrLf
ps = ps & "    [void]$Global:AsyncJobs.Add(@{ Type=""Ping""; PS=$ps; Handle=$handle; IP=$IP })" & vbCrLf
ps = ps & "    $Global:ActivePings++" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Start-AsyncDns ($IP, $Cell) {" & vbCrLf
ps = ps & "    $rs = [runspacefactory]::CreateRunspace()" & vbCrLf
ps = ps & "    $rs.ApartmentState = ""STA""; $rs.ThreadOptions = ""ReuseThread""; $rs.Open()" & vbCrLf
ps = ps & "    $ps = [PowerShell]::Create(); $ps.Runspace = $rs" & vbCrLf
ps = ps & "    [void]$ps.AddScript({" & vbCrLf
ps = ps & "        param($ip)" & vbCrLf
ps = ps & "        try   { [System.Net.Dns]::GetHostEntry($ip).HostName }" & vbCrLf
ps = ps & "        catch { ""Unknown"" }" & vbCrLf
ps = ps & "    }).AddArgument($IP)" & vbCrLf
ps = ps & "    $handle = $ps.BeginInvoke()" & vbCrLf
ps = ps & "    [void]$Global:AsyncJobs.Add(@{ Type=""DNS""; PS=$ps; Handle=$handle; IP=$IP; Cell=$Cell })" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "function Update-StatusBar {" & vbCrLf
ps = ps & "    $found = $Global:UIReferences.Count" & vbCrLf
ps = ps & "    if (-not $Global:ScanComplete) {" & vbCrLf
ps = ps & "        $done = $Global:TotalQueued - $Global:PingQueue.Count - $Global:ActivePings" & vbCrLf
ps = ps & "        $pct  = if ($Global:TotalQueued -gt 0) { [int](($done/$Global:TotalQueued)*100) } else { 0 }" & vbCrLf
ps = ps & "        $RuntimeLabel.ForeColor = [System.Drawing.Color]::DarkCyan" & vbCrLf
ps = ps & "        $RuntimeLabel.Text = ""Scanning...  $done / $Global:TotalQueued  ($($Global:ActivePings) active)  |  $found found  [$pct%]""" & vbCrLf
ps = ps & "    } else {" & vbCrLf
ps = ps & "        $span = [DateTime]::Now - $Global:StartTime" & vbCrLf
ps = ps & "        $RuntimeLabel.ForeColor = [System.Drawing.Color]::Black" & vbCrLf
ps = ps & "        $RuntimeLabel.Text = ""Runtime: {0:d2}:{1:d2}:{2:d2}  |  {3} devices  |  interval: $($CmbInterval.SelectedItem)"" -f $span.Hours,$span.Minutes,$span.Seconds,$found" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "}" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Main Timer ----------------------------------------------------------------" & vbCrLf
ps = ps & "$Timer = New-Object System.Windows.Forms.Timer" & vbCrLf
ps = ps & "$Timer.Interval = 50" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$Timer.Add_Tick({" & vbCrLf
ps = ps & "    $Timer.Stop()" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    if ($Global:StopRequested) {" & vbCrLf
ps = ps & "        $Global:StopRequested = $false; $Global:ScanComplete = $true" & vbCrLf
ps = ps & "        $Global:PingQueue.Clear()" & vbCrLf
ps = ps & "        $RuntimeLabel.ForeColor = [System.Drawing.Color]::DarkOrange" & vbCrLf
ps = ps & "        $RuntimeLabel.Text = ""Scan stopped  |  $($Global:UIReferences.Count) devices found""" & vbCrLf
ps = ps & "        $StartButton.Text = ""Rescan""; $StartButton.BackColor = [System.Drawing.Color]::FromArgb(16,124,16)" & vbCrLf
ps = ps & "        $StartButton.Enabled = $true" & vbCrLf
ps = ps & "        $BtnPause.Enabled = $true" & vbCrLf
ps = ps & "        $Timer.Interval = $Global:RefreshMs; $Timer.Start(); return" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    while ($Global:PingQueue.Count -gt 0 -and $Global:ActivePings -lt $MaxConcurrentPings) {" & vbCrLf
ps = ps & "        Start-AsyncPing ($Global:PingQueue.Dequeue())" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $done = [System.Collections.Generic.List[hashtable]]::new()" & vbCrLf
ps = ps & "    foreach ($job in @($Global:AsyncJobs)) {" & vbCrLf
ps = ps & "        if (-not $job.Handle.IsCompleted) { continue }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "        if ($job.Type -eq ""Ping"") {" & vbCrLf
ps = ps & "            $Global:ActivePings--" & vbCrLf
ps = ps & "            try {" & vbCrLf
ps = ps & "                $res = $job.PS.EndInvoke($job.Handle)[0]" & vbCrLf
ps = ps & "                if ($res.Status -eq ""Success"") {" & vbCrLf
ps = ps & "                    $rtt    = [Math]::Max(1, $res.RTT)" & vbCrLf
ps = ps & "                    $subnet = Get-SubnetForIP $res.IP" & vbCrLf
ps = ps & "                    $ifType = if ($SubnetIfaceMap.ContainsKey($subnet)) { $SubnetIfaceMap[$subnet] } else { ""Other"" }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "                    $Global:NetworkStats[$res.IP] = @{" & vbCrLf
ps = ps & "                        History            = [System.Collections.Generic.List[long]]::new()" & vbCrLf
ps = ps & "                        TimestampedHistory = [System.Collections.Generic.List[hashtable]]::new()" & vbCrLf
ps = ps & "                        Timeouts           = 0; IfType = $ifType; DnsName = """"" & vbCrLf
ps = ps & "                    }" & vbCrLf
ps = ps & "                    [void]$Global:NetworkStats[$res.IP].History.Add($rtt)" & vbCrLf
ps = ps & "                    [void]$Global:NetworkStats[$res.IP].TimestampedHistory.Add(@{ Time=[DateTime]::Now; Ms=$rtt })" & vbCrLf
ps = ps & "                    $Global:DeviceAlerts[$res.IP] = @{ WasUp=$true }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "                    $lbl      = if ($Global:DeviceLabels.ContainsKey($res.IP)) { $Global:DeviceLabels[$res.IP] } else { """" }" & vbCrLf
ps = ps & "                    $lvi      = New-Object System.Windows.Forms.ListViewItem($res.IP)" & vbCrLf
ps = ps & "                    $lblCell  = $lvi.SubItems.Add($lbl)" & vbCrLf
ps = ps & "                    $nameCell = $lvi.SubItems.Add(""Resolving..."")" & vbCrLf
ps = ps & "                    $typeCell = $lvi.SubItems.Add((Get-IfTypeIcon $ifType))" & vbCrLf
ps = ps & "                    $curCell  = $lvi.SubItems.Add(""$rtt ms"")" & vbCrLf
ps = ps & "                    $minCell  = $lvi.SubItems.Add(""$rtt ms"")" & vbCrLf
ps = ps & "                    $maxCell  = $lvi.SubItems.Add(""$rtt ms"")" & vbCrLf
ps = ps & "                    $avgCell  = $lvi.SubItems.Add(""$rtt ms"")" & vbCrLf
ps = ps & "                    $toCell   = $lvi.SubItems.Add(""0"")" & vbCrLf
ps = ps & "                    $sparkCell= $lvi.SubItems.Add("""")" & vbCrLf
ps = ps & "                    $lvi.UseItemStyleForSubItems = $false" & vbCrLf
ps = ps & "                    $curCell.ForeColor  = Get-LatencyColor $rtt" & vbCrLf
ps = ps & "                    $nameCell.ForeColor = [System.Drawing.Color]::Gray" & vbCrLf
ps = ps & "                    $typeCell.ForeColor = Get-IfTypeColor $ifType" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "                    [void]$Global:AllItems.Add($lvi)" & vbCrLf
ps = ps & "                    if ($Global:FilterIfType -eq ""All"" -or $ifType -eq $Global:FilterIfType) {" & vbCrLf
ps = ps & "                        [void]$ListView.Items.Add($lvi)" & vbCrLf
ps = ps & "                    }" & vbCrLf
ps = ps & "                    $Global:UIReferences[$res.IP] = @{" & vbCrLf
ps = ps & "                        Current=$curCell; Min=$minCell; Max=$maxCell; Avg=$avgCell" & vbCrLf
ps = ps & "                        Timeout=$toCell; Name=$nameCell; Type=$typeCell" & vbCrLf
ps = ps & "                        Spark=$sparkCell; Label=$lblCell; Item=$lvi" & vbCrLf
ps = ps & "                    }" & vbCrLf
ps = ps & "                    Start-AsyncDns $res.IP $nameCell" & vbCrLf
ps = ps & "                }" & vbCrLf
ps = ps & "            } catch {}" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        elseif ($job.Type -eq ""DNS"") {" & vbCrLf
ps = ps & "            try {" & vbCrLf
ps = ps & "                $name = $job.PS.EndInvoke($job.Handle)[0]" & vbCrLf
ps = ps & "                $resolvedName = if ($name) { $name } else { ""Unknown"" }" & vbCrLf
ps = ps & "                $job.Cell.Text = $resolvedName" & vbCrLf
ps = ps & "                $job.Cell.ForeColor = [System.Drawing.Color]::Black" & vbCrLf
ps = ps & "                if ($Global:NetworkStats.ContainsKey($job.IP)) {" & vbCrLf
ps = ps & "                    $Global:NetworkStats[$job.IP].DnsName = $resolvedName" & vbCrLf
ps = ps & "                }" & vbCrLf
ps = ps & "            } catch {}" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "        $job.PS.Dispose(); [void]$done.Add($job)" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "    foreach ($j in $done) { [void]$Global:AsyncJobs.Remove($j) }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    if (-not $Global:ScanComplete -and $Global:PingQueue.Count -eq 0 -and $Global:ActivePings -eq 0) {" & vbCrLf
ps = ps & "        $Global:ScanComplete = $true" & vbCrLf
ps = ps & "        $StartButton.Text = ""Rescan""; $StartButton.BackColor = [System.Drawing.Color]::FromArgb(16,124,16)" & vbCrLf
ps = ps & "        $StartButton.Enabled = $true; $Timer.Interval = $Global:RefreshMs" & vbCrLf
ps = ps & "        $BtnPause.Enabled = $true" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    Update-StatusBar" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    # -- Monitoring phase -------------------------------------------------------" & vbCrLf
ps = ps & "    if ($Global:ScanComplete -and -not $Global:Paused -and $Global:UIReferences.Count -gt 0 -and" & vbCrLf
ps = ps & "        $Global:PingQueue.Count -eq 0 -and $Global:ActivePings -eq 0) {" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "        foreach ($IP in @($Global:UIReferences.Keys)) {" & vbCrLf
ps = ps & "            try {" & vbCrLf
ps = ps & "                $p = New-Object System.Net.NetworkInformation.Ping" & vbCrLf
ps = ps & "                $r = $p.Send($IP, $PingTimeoutMs); $p.Dispose()" & vbCrLf
ps = ps & "                $stats = $Global:NetworkStats[$IP]; $ui = $Global:UIReferences[$IP]" & vbCrLf
ps = ps & "                $wasUp = if ($Global:DeviceAlerts.ContainsKey($IP)) { $Global:DeviceAlerts[$IP].WasUp } else { $true }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "                if ($r.Status -eq ""Success"") {" & vbCrLf
ps = ps & "                    $ms = [Math]::Max(1, $r.RoundtripTime)" & vbCrLf
ps = ps & "                    [void]$stats.History.Add($ms)" & vbCrLf
ps = ps & "                    [void]$stats.TimestampedHistory.Add(@{ Time=[DateTime]::Now; Ms=$ms })" & vbCrLf
ps = ps & "                    while ($stats.TimestampedHistory.Count -gt $MaxHistoryPoints) { $stats.TimestampedHistory.RemoveAt(0) }" & vbCrLf
ps = ps & "                    while ($stats.History.Count -gt $MaxHistoryPoints)            { $stats.History.RemoveAt(0) }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "                    $realH = @($stats.History | Where-Object { $_ -lt $PingTimeoutMs })" & vbCrLf
ps = ps & "                    $min = if ($realH.Count -gt 0) { ($realH | Measure-Object -Minimum).Minimum } else { $ms }" & vbCrLf
ps = ps & "                    $max = if ($realH.Count -gt 0) { ($realH | Measure-Object -Maximum).Maximum } else { $ms }" & vbCrLf
ps = ps & "                    $avg = if ($realH.Count -gt 0) { [Math]::Round(($realH | Measure-Object -Average).Average,1) } else { $ms }" & vbCrLf
ps = ps & "                    $ui.Current.Text = ""$ms ms""; $ui.Current.ForeColor = Get-LatencyColor $ms" & vbCrLf
ps = ps & "                    $ui.Min.Text = ""$min ms""; $ui.Max.Text = ""$max ms""; $ui.Avg.Text = ""$avg ms""" & vbCrLf
ps = ps & "                    $ui.Spark.Text = Get-Sparkline $stats.History" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "                    # Alert: device came back up" & vbCrLf
ps = ps & "                    if (-not $wasUp -and $ChkAlerts.Checked) {" & vbCrLf
ps = ps & "                        $dispName = Get-DisplayName $IP" & vbCrLf
ps = ps & "                        [System.Windows.Forms.MessageBox]::Show(""$dispName ($IP) is back online."", ""Device Up"", 0, 64)" & vbCrLf
ps = ps & "                    }" & vbCrLf
ps = ps & "                    $Global:DeviceAlerts[$IP].WasUp = $true" & vbCrLf
ps = ps & "                } else {" & vbCrLf
ps = ps & "                    $stats.Timeouts++" & vbCrLf
ps = ps & "                    [void]$stats.TimestampedHistory.Add(@{ Time=[DateTime]::Now; Ms=$PingTimeoutMs })" & vbCrLf
ps = ps & "                    while ($stats.TimestampedHistory.Count -gt $MaxHistoryPoints) { $stats.TimestampedHistory.RemoveAt(0) }" & vbCrLf
ps = ps & "                    $ui.Current.Text = ""TIMEOUT""; $ui.Current.ForeColor = [System.Drawing.Color]::DarkRed" & vbCrLf
ps = ps & "                    $ui.Timeout.Text = $stats.Timeouts; $ui.Timeout.ForeColor = [System.Drawing.Color]::Red" & vbCrLf
ps = ps & "                    $ui.Spark.Text = Get-Sparkline $stats.History" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "                    # Alert: device went down" & vbCrLf
ps = ps & "                    if ($wasUp -and $ChkAlerts.Checked) {" & vbCrLf
ps = ps & "                        $dispName = Get-DisplayName $IP" & vbCrLf
ps = ps & "                        [System.Windows.Forms.MessageBox]::Show(""$dispName ($IP) is NOT responding!"", ""Device Down"", 0, 48)" & vbCrLf
ps = ps & "                    }" & vbCrLf
ps = ps & "                    $Global:DeviceAlerts[$IP].WasUp = $false" & vbCrLf
ps = ps & "                }" & vbCrLf
ps = ps & "            } catch {}" & vbCrLf
ps = ps & "        }" & vbCrLf
ps = ps & "        # Re-apply live sort after all cells updated" & vbCrLf
ps = ps & "        if ($Global:SortCol -ge 0) { Apply-Sort }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $Timer.Start()" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "# -- Start / Stop / Rescan -----------------------------------------------------" & vbCrLf
ps = ps & "$StartButton.Add_Click({" & vbCrLf
ps = ps & "    if (-not $Global:ScanComplete -and ($Global:PingQueue.Count -gt 0 -or $Global:ActivePings -gt 0)) {" & vbCrLf
ps = ps & "        $Global:StopRequested = $true; $StartButton.Enabled = $false; $StartButton.Text = ""Stopping...""; return" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $subnet = $TxtSubnet.Text.Trim()" & vbCrLf
ps = ps & "    if ([string]::IsNullOrWhiteSpace($subnet)) { $subnet = $DetectedSubnet }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $r1 = Parse-Range $TxtR1Start.Text $TxtR1End.Text" & vbCrLf
ps = ps & "    $r2 = Parse-Range $TxtR2Start.Text $TxtR2End.Text" & vbCrLf
ps = ps & "    if ($r1.Count -eq 0 -and $r2.Count -eq 0) {" & vbCrLf
ps = ps & "        [System.Windows.Forms.MessageBox]::Show(""Enter at least one valid range (1-254)."",""Invalid"",0,48); return" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $existing = @($Global:UIReferences.Keys)" & vbCrLf
ps = ps & "    $Global:PingQueue.Clear()" & vbCrLf
ps = ps & "    foreach ($oct in ($r1 + $r2 | Sort-Object -Unique)) {" & vbCrLf
ps = ps & "        $ip = ""$subnet.$oct""" & vbCrLf
ps = ps & "        if ($existing -notcontains $ip) { [void]$Global:PingQueue.Enqueue($ip) }" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    if ($Global:PingQueue.Count -eq 0) {" & vbCrLf
ps = ps & "        $RuntimeLabel.Text = ""All IPs already known. Monitoring active.""; return" & vbCrLf
ps = ps & "    }" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "    $Global:ScanComplete = $false; $Global:StopRequested = $false; $Global:Paused = $false" & vbCrLf
ps = ps & "    $BtnPause.Text = ""Pause""; $BtnPause.BackColor = [System.Drawing.Color]::FromArgb(220,220,230)" & vbCrLf
ps = ps & "    $BtnPause.ForeColor = [System.Drawing.Color]::Black; $BtnPause.Enabled = $false" & vbCrLf
ps = ps & "    $Global:TotalQueued = $Global:PingQueue.Count; $Global:ActivePings = 0" & vbCrLf
ps = ps & "    $Global:StartTime = [DateTime]::Now" & vbCrLf
ps = ps & "    $StartButton.Text = ""Stop""; $StartButton.BackColor = [System.Drawing.Color]::FromArgb(196,43,28)" & vbCrLf
ps = ps & "    $StartButton.Enabled = $true" & vbCrLf
ps = ps & "    Update-StatusBar; $Timer.Interval = 50; $Timer.Start()" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "$Form.Add_FormClosing({" & vbCrLf
ps = ps & "    $Timer.Stop()" & vbCrLf
ps = ps & "    foreach ($j in $Global:AsyncJobs) { try { $j.PS.Dispose() } catch {} }" & vbCrLf
ps = ps & "})" & vbCrLf
ps = ps & "" & vbCrLf
ps = ps & "[void]$Form.ShowDialog()" & vbCrLf
ps = ps & "" & vbCrLf

Set f = fso.OpenTextFile(tmp, 2, True, -1)  ' -1 = Unicode
f.Write ps
f.Close

Set sh = CreateObject("WScript.Shell")
sh.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File """ & tmp & """", 0, False

' Wait for PS to load the file, then clean up
WScript.Sleep 3000
If fso.FileExists(tmp) Then fso.DeleteFile tmp