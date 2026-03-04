# 1. 設定輸出路徑、電腦名稱與時間戳記
$CurrentDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($CurrentDir)) { $CurrentDir = Get-Location }

$CompName = $env:COMPUTERNAME
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"

$FileNormal = "System_Report_${CompName}_${TimeStamp}.html"
$FileAI     = "System_Report_forAI_${CompName}_${TimeStamp}.html"

$OutNormal = Join-Path $CurrentDir $FileNormal
$OutAI     = Join-Path $CurrentDir $FileAI

Write-Host "正在對 [$CompName] 進行深度系統與資安掃描，請稍候..." -ForegroundColor Cyan

# ==========================================
# === 模組 1: 系統硬體與作業系統資訊 ===
# ==========================================
Write-Host " [1/8] 收集系統硬體資訊..." -ForegroundColor Yellow
$os = Get-CimInstance Win32_OperatingSystem
$cpu = (Get-CimInstance Win32_Processor | Select-Object -ExpandProperty Name -First 1).Trim()
$ram = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
$gpu = (Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name) -join ", "
$disks = (Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
    "$($_.DeviceID) (剩餘: $([math]::Round($_.FreeSpace/1GB, 2))GB / 總共: $([math]::Round($_.Size/1GB, 2))GB)"
}) -join " | "

$sysinfo = @([PSCustomObject]@{
    OS_Version = "$($os.Caption) $($os.OSArchitecture)"
    CPU = $cpu
    RAM = "$ram GB"
    GPU = $gpu
    Disk_Usage = $disks
})

# ==========================================
# === 模組 2: 執行中程序 ===
# ==========================================
Write-Host " [2/8] 掃描執行中程序..." -ForegroundColor Yellow
$procs = Get-Process -IncludeUserName -ErrorAction SilentlyContinue | Where-Object Path | Select-Object Name, UserName, @{N="Status";E={"Running"}}, Path

# ==========================================
# === 模組 3: 對外網路連線 (Established) ===
# ==========================================
Write-Host " [3/8] 掃描對外網路連線..." -ForegroundColor Yellow
$procCache = @{}
Get-Process -ErrorAction SilentlyContinue | ForEach-Object { $procCache[$_.Id] = $_.Name }

$net = Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue | Select-Object `
    @{N="ProcessName";E={if($procCache[$_.OwningProcess]){$procCache[$_.OwningProcess]}else{"PID: $($_.OwningProcess)"}}},
    LocalAddress, LocalPort, RemoteAddress, RemotePort, State

if (-not $net) { $net = @([PSCustomObject]@{ProcessName="無";LocalAddress="-";LocalPort="-";RemoteAddress="-";RemotePort="-";State="-"}) }

# ==========================================
# === 模組 4: 啟動執行項 ===
# ==========================================
Write-Host " [4/8] 破解啟動項狀態..." -ForegroundColor Yellow
$startupStatus = @{}
$approvedPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run32",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\StartupFolder"
)
foreach ($p in $approvedPaths) {
    if (Test-Path $p) {
        $item = Get-Item $p
        foreach ($valName in $item.GetValueNames()) {
            $bytes = $item.GetValue($valName)
            if ($bytes -is [byte[]] -and $bytes.Length -gt 0) {
                $startupStatus[$valName] = if ($bytes[0] -eq 2) { "Enabled" } else { "Disabled" }
            }
        }
    }
}

$starts = Get-CimInstance Win32_StartupCommand | Select-Object Name, @{
    N="Status";
    E={ if ($startupStatus.ContainsKey($_.Name)) { $startupStatus[$_.Name] } else { "Enabled" } }
}, Command, Location

foreach ($st in $starts) {
    if (![string]::IsNullOrEmpty($st.Command) -and $st.Command -match '~\d') {
        if ($st.Command -match '^(?<path>"[^"]+"|[a-zA-Z]:\\[^\s]+)(?<args>.*)$') {
            $pPath = $matches['path'].Trim('"')
            $pArgs = $matches['args']
            try { $st.Command = "`"$((Get-Item -LiteralPath $pPath -ErrorAction Stop).FullName)`"$pArgs" } catch {}
        }
    }
}

# ==========================================
# === 模組 5: 排程任務 ===
# ==========================================
Write-Host " [5/8] 提取排程任務與引數..." -ForegroundColor Yellow
$tasks = Get-ScheduledTask | Select-Object TaskName, @{N="State";E={[string]$_.State}}, @{
    N="Path"; 
    E={
        ($_.Actions | Where-Object Execute | ForEach-Object {
            $exe = $_.Execute; $arg = $_.Arguments
            if ([string]::IsNullOrWhiteSpace($arg)) { $exe } else { "$exe $arg" }
        }) -join " | "
    }
} | Where-Object { ![string]::IsNullOrWhiteSpace($_.Path) }

# ==========================================
# === 模組 6: 系統服務 ===
# ==========================================
Write-Host " [6/8] 匯出系統服務清單..." -ForegroundColor Yellow
$srvs = Get-CimInstance Win32_Service | Select-Object Name, DisplayName, State, StartMode, PathName, StartName

# ==========================================
# === 模組 7: 已安裝軟體清單 ===
# ==========================================
Write-Host " [7/8] 撈取已安裝軟體紀錄..." -ForegroundColor Yellow
$uninstallPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$software = Get-ItemProperty $uninstallPaths -ErrorAction SilentlyContinue | 
    Where-Object DisplayName | 
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | 
    Sort-Object DisplayName -Unique
if (-not $software) { $software = @([PSCustomObject]@{DisplayName="無資料";DisplayVersion="";Publisher="";InstallDate=""}) }

# ==========================================
# === 模組 8: WMI 永久事件訂閱 ===
# ==========================================
Write-Host " [8/8] 偵測 WMI 無檔案惡意訂閱..." -ForegroundColor Yellow
$wmiB = Get-CimInstance -Namespace root\subscription -ClassName __FilterToConsumerBinding -ErrorAction SilentlyContinue
$wmi = @()
if ($wmiB) {
    foreach ($b in $wmiB) {
        $wmi += [PSCustomObject]@{ Filter = [string]$b.Filter; Consumer = [string]$b.Consumer }
    }
}
if ($wmi.Count -eq 0) { $wmi = @([PSCustomObject]@{Filter="無可疑訂閱"; Consumer="安全"}) }

# ==========================================
# === 隱私遮罩函數 (For AI) ===
# ==========================================
function Mask-User {
    param([string]$u)
    if ([string]::IsNullOrWhiteSpace($u)) { return "" }
    $parts = $u -split '\\', 2
    if ($parts.Length -eq 2) { $dom = $parts[0]; $name = $parts[1] } else { $dom = ""; $name = $u }
    if ($name.Length -gt 2) {
        $mid = "o" * ($name.Length - 2)
        $name = $name.Substring(0,1) + $mid + $name.Substring($name.Length - 1, 1)
    }
    if ($dom) { return "$dom\$name" }
    return $name
}

function Build-SimpleTable {
    param($data, $props)
    if (-not $data) { return "<table><tr><td>無資料</td></tr></table>" }
    $arr = @($data)
    $sb = [System.Text.StringBuilder]::new()
    $sb.Append("<table border='1'><tr>") | Out-Null
    foreach ($p in $props) { $sb.Append("<th>$p</th>") | Out-Null }
    $sb.Append("</tr>") | Out-Null
    foreach ($item in $arr) {
        $sb.Append("<tr>") | Out-Null
        foreach ($p in $props) {
            $val = $item.$p
            if ($null -eq $val) { $val = "" }
            $val = $val.ToString().Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;")
            $sb.Append("<td>$val</td>") | Out-Null
        }
        $sb.Append("</tr>") | Out-Null
    }
    $sb.Append("</table>") | Out-Null
    return $sb.ToString()
}

# === JSON 轉換 (正常版使用) ===
$j1 = $sysinfo  | ConvertTo-Json -Compress
$j2 = $procs    | ConvertTo-Json -Compress; if(!$j2){$j2="[]"}
$j3 = $net      | ConvertTo-Json -Compress; if(!$j3){$j3="[]"}
$j4 = $starts   | ConvertTo-Json -Compress; if(!$j4){$j4="[]"}
$j5 = $tasks    | ConvertTo-Json -Compress; if(!$j5){$j5="[]"}
$j6 = $srvs     | ConvertTo-Json -Compress; if(!$j6){$j6="[]"}
$j7 = $software | ConvertTo-Json -Compress; if(!$j7){$j7="[]"}
$j8 = $wmi      | ConvertTo-Json -Compress; if(!$j8){$j8="[]"}

# === HTML 介面 (排版修復：z-index, nowrap, 精準高度) ===
$html = @(
'<!DOCTYPE html>',
'<html lang="zh-TW"><head><meta charset="UTF-8">',
'<title>' + $FileNormal.Replace('.html','') + '</title>',
'<style>',
'body{margin:0;padding-top:60px;font-family:"Segoe UI",sans-serif;background:#f5f7f9;scroll-behavior:smooth;}',
'.topbar{position:fixed;top:0;left:0;width:100%;height:50px;background:#20232a;color:#fff;padding:0 20px;z-index:100;display:flex;gap:10px;align-items:center;box-sizing:border-box;box-shadow:0 2px 5px rgba(0,0,0,0.2);overflow-x:auto;}',
'.topbar a{color:#fff;text-decoration:none;padding:6px 10px;background:#33373e;border-radius:4px;font-size:13px;white-space:nowrap;transition:all 0.3s ease}',
'.topbar a:hover{background:#4a4f58;}',
'.topbar a.active{background:#61dafb;color:#000;font-weight:bold;box-shadow:0 0 8px rgba(97,218,251,0.5);}',
'.tooltip{position:relative;display:inline-flex;align-items:center;cursor:pointer;margin-left:auto;background:#33373e;padding:6px 12px;border-radius:4px;font-size:13px;white-space:nowrap;transition:0.2s}',
'.tooltip:hover{background:#444}',
'.tooltip input{margin-right:8px;cursor:pointer;}',
'.tooltiptext{visibility:hidden;width:max-content;max-width:500px;background-color:#111;color:#fff;text-align:left;padding:12px;border-radius:6px;position:absolute;z-index:1;top:130%;right:0;font-size:13px;white-space:pre-wrap;box-shadow:0 4px 10px rgba(0,0,0,0.5);line-height:1.6;border:1px solid #555;opacity:0;transition:opacity 0.2s}',
'.tooltip:hover .tooltiptext{visibility:visible;opacity:1;}',
'.container{padding:0 25px 25px 25px;max-width:1500px;margin:auto}',
'h2{color:#333;margin-top:20px;padding-bottom:10px;border-bottom:2px solid #ccc;display:flex;justify-content:space-between;align-items:center;font-size:18px;}',
'table{width:100%;border-collapse:collapse;background:#fff;margin-bottom:30px;box-shadow:0 1px 3px rgba(0,0,0,0.1);}',
'th,td{border:1px solid #e0e0e0;padding:10px;text-align:left;font-size:13px}',
'td{word-break:break-word;} /* 讓過長的內容自然換行，而不是硬生生切斷 */',
'th{background:#e8eaed;cursor:pointer;user-select:none;position:sticky;top:50px;z-index:10;white-space:nowrap;transition:background 0.2s;box-shadow:0 1px 2px rgba(0,0,0,0.1)} /* z-index 與 nowrap 修復跑版與重疊 */',
'th:hover{background:#d5d8dc}',
'button{padding:5px 12px;cursor:pointer;background:#0078d4;color:#fff;border:none;border-radius:4px;font-size:12px;white-space:nowrap;}',
'button:hover{background:#005a9e}',
'</style></head><body>',
'<div class="topbar">',
'  <strong style="font-size:16px;margin-right:5px">💻 系統健檢</strong>',
'  <a href="#sec1" class="nav-link">1. 系統硬體</a>',
'  <a href="#sec2" class="nav-link">2. 程序</a>',
'  <a href="#sec3" class="nav-link">3. 網路連線</a>',
'  <a href="#sec4" class="nav-link">4. 啟動項</a>',
'  <a href="#sec5" class="nav-link">5. 排程</a>',
'  <a href="#sec6" class="nav-link">6. 服務</a>',
'  <a href="#sec7" class="nav-link">7. 已安裝軟體</a>',
'  <a href="#sec8" class="nav-link">8. WMI訂閱</a>',
'  <div class="tooltip">',
'    <input type="checkbox" id="hideSpecific" checked onchange="applyFilter()">',
'    <label for="hideSpecific" style="cursor:pointer">隱藏特定常駐</label>',
'    <div class="tooltiptext">隱藏包含以下路徑：<br>• C:\WINDOWS\system32\svchost.exe<br>• C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe</div>',
'  </div>',
'</div>',
'<div class="container">',
'<h2 id="sec1" class="section-title"><span>1. 系統硬體與作業系統資訊</span> <button onclick="cv(''t1'', ''SysInfo'')">匯出 CSV</button></h2>',
'<table id="t1"><thead><tr><th>系統版本</th><th>處理器 (CPU)</th><th>記憶體 (RAM)</th><th>顯示卡 (GPU)</th><th>磁碟使用量</th></tr></thead><tbody></tbody></table>',
'<h2 id="sec2" class="section-title"><span>2. 執行中的程序</span> <button onclick="cv(''t2'', ''Processes'')">匯出 CSV</button></h2>',
'<table id="t2"><thead><tr><th>名稱</th><th>使用者</th><th>狀態</th><th>完整路徑</th></tr></thead><tbody></tbody></table>',
'<h2 id="sec3" class="section-title"><span>3. 對外網路連線 (Established)</span> <button onclick="cv(''t3'', ''Network'')">匯出 CSV</button></h2>',
'<table id="t3"><thead><tr><th>程序名稱</th><th>本地 IP</th><th>本地 Port</th><th>遠端 IP</th><th>遠端 Port</th><th>狀態</th></tr></thead><tbody></tbody></table>',
'<h2 id="sec4" class="section-title"><span>4. 啟動執行項</span> <button onclick="cv(''t4'', ''Startup'')">匯出 CSV</button></h2>',
'<table id="t4"><thead><tr><th>名稱</th><th>狀態</th><th>啟動指令</th><th>位置</th></tr></thead><tbody></tbody></table>',
'<h2 id="sec5" class="section-title"><span>5. 排程任務</span> <button onclick="cv(''t5'', ''Tasks'')">匯出 CSV</button></h2>',
'<table id="t5"><thead><tr><th>任務名稱</th><th>狀態</th><th>執行路徑與引數</th></tr></thead><tbody></tbody></table>',
'<h2 id="sec6" class="section-title"><span>6. 系統服務</span> <button onclick="cv(''t6'', ''Services'')">匯出 CSV</button></h2>',
'<table id="t6"><thead><tr><th>服務名稱</th><th>顯示名稱</th><th>執行狀態</th><th>啟動類型</th><th>執行路徑</th><th>啟動身分</th></tr></thead><tbody></tbody></table>',
'<h2 id="sec7" class="section-title"><span>7. 已安裝軟體清單</span> <button onclick="cv(''t7'', ''Software'')">匯出 CSV</button></h2>',
'<table id="t7"><thead><tr><th>軟體名稱</th><th>版本</th><th>發行商</th><th>安裝日期</th></tr></thead><tbody></tbody></table>',
'<h2 id="sec8" class="section-title"><span>8. WMI 永久事件訂閱 (無檔案攻擊指標)</span> <button onclick="cv(''t8'', ''WMI'')">匯出 CSV</button></h2>',
'<table id="t8"><thead><tr><th>事件過濾器 (Filter)</th><th>消費者觸發動作 (Consumer)</th></tr></thead><tbody></tbody></table>',
'</div>',
'<script>'
)

# === JavaScript 邏輯 ===
$jsLogic = @(
'const hList = ["c:/windows/system32/svchost.exe", "c:/program files (x86)/microsoft/edge/application/msedge.exe"];',
'function fl(id, d, k) {',
'  const b = document.getElementById(id).tBodies[0];',
'  if(!d) return;',
'  const arr = Array.isArray(d) ? d : [d];',
'  const chk = document.getElementById("hideSpecific").checked;',
'  arr.forEach(i => {',
'    let r = b.insertRow();',
'    let rowPath = "";',
'    k.forEach(key => {',
'      let c = r.insertCell();',
'      let val = (i[key] === null || i[key] === undefined) ? "" : String(i[key]);',
'      if(key.toLowerCase().includes("path") || key === "Command") {',
'        rowPath = val.toLowerCase().replace(/\\/g, "/");',
'      }',
'      if(key === "Status" || key === "State" || key === "StartMode" || key === "LocalPort" || key === "RemotePort") {',
'        c.style.whiteSpace = "nowrap";',
'      }',
'      c.textContent = val;',
'    });',
'    r.setAttribute("data-path", rowPath);',
'    if (chk && rowPath && hList.some(h => rowPath.includes(h))) {',
'      r.style.display = "none";',
'    }',
'  });',
'  att(id);',
'}',
'function applyFilter() {',
'  const chk = document.getElementById("hideSpecific").checked;',
'  document.querySelectorAll("tbody tr").forEach(r => {',
'    const p = r.getAttribute("data-path");',
'    if (p && hList.some(h => p.includes(h))) {',
'      r.style.display = chk ? "none" : "";',
'    }',
'  });',
'}',
'function att(id) {',
'  const ths = document.querySelectorAll("#" + id + " th");',
'  ths.forEach((th, idx) => { th.onclick = () => st(id, idx, th); });',
'}',
'function st(tid, col, th) {',
'  const tb = document.getElementById(tid).tBodies[0];',
'  const rows = Array.from(tb.rows);',
'  const isAsc = th.getAttribute("data-dir") !== "asc";',
'  rows.sort((a, b) => {',
'    let v1 = a.cells[col].innerText.toLowerCase();',
'    let v2 = b.cells[col].innerText.toLowerCase();',
'    if (v1 < v2) return isAsc ? -1 : 1;',
'    if (v1 > v2) return isAsc ? 1 : -1;',
'    return 0;',
'  });',
'  document.querySelectorAll("#" + tid + " th").forEach(x => { x.innerText = x.innerText.replace(" ▲","").replace(" ▼",""); });',
'  th.setAttribute("data-dir", isAsc ? "asc" : "desc");',
'  th.innerText += isAsc ? " ▲" : " ▼";',
'  rows.forEach(r => tb.appendChild(r));',
'}',
'function cv(id, name) {',
'  try {',
'    let c = ["\ufeff"];',
'    let table = document.getElementById(id);',
'    if (!table) return;',
'    let rows = table.querySelectorAll("tr");',
'    rows.forEach(r => {',
'      if (r.style.display === "none") return;',
'      let d = [];',
'      r.querySelectorAll("th,td").forEach(cell => {',
'        let txt = cell.innerText.replace(" ▲","").replace(" ▼","");',
'        txt = txt.split("\x22").join("\x22\x22");',
'        d.push("\x22" + txt + "\x22");',
'      });',
'      c.push(d.join(","));',
'    });',
'    let b = new Blob([c.join("\n")], {type:"text/csv;charset=utf-8;"});',
'    let a = document.createElement("a");',
'    a.href = URL.createObjectURL(b);',
'    a.download = name + "_" + document.title + ".csv";',
'    document.body.appendChild(a);',
'    a.click();',
'    document.body.removeChild(a);',
'  } catch(e) { alert("匯出錯誤：" + e.message); }',
'}',
'window.addEventListener("scroll", () => {',
'  let current = "";',
'  const sections = document.querySelectorAll(".section-title");',
'  sections.forEach(sec => {',
'    if (pageYOffset >= sec.offsetTop - 100) { current = sec.getAttribute("id"); }',
'  });',
'  document.querySelectorAll(".nav-link").forEach(a => {',
'    a.classList.remove("active");',
'    if (current && a.getAttribute("href") === "#" + current) { a.classList.add("active"); }',
'  });',
'});',
'window.onload = () => {',
'  fl("t1", d1, ["OS_Version","CPU","RAM","GPU","Disk_Usage"]);',
'  fl("t2", d2, ["Name","UserName","Status","Path"]);',
'  fl("t3", d3, ["ProcessName","LocalAddress","LocalPort","RemoteAddress","RemotePort","State"]);',
'  fl("t4", d4, ["Name","Status","Command","Location"]);',
'  fl("t5", d5, ["TaskName","State","Path"]);',
'  fl("t6", d6, ["Name","DisplayName","State","StartMode","PathName","StartName"]);',
'  fl("t7", d7, ["DisplayName","DisplayVersion","Publisher","InstallDate"]);',
'  fl("t8", d8, ["Filter","Consumer"]);',
'  window.dispatchEvent(new Event("scroll"));',
'};',
'</script></body></html>'
)

# === 寫入 正常版 ===
Set-Content -Path $OutNormal -Value $html -Encoding UTF8
Add-Content -Path $OutNormal -Value "const d1 = $j1;" -Encoding UTF8
Add-Content -Path $OutNormal -Value "const d2 = $j2;" -Encoding UTF8
Add-Content -Path $OutNormal -Value "const d3 = $j3;" -Encoding UTF8
Add-Content -Path $OutNormal -Value "const d4 = $j4;" -Encoding UTF8
Add-Content -Path $OutNormal -Value "const d5 = $j5;" -Encoding UTF8
Add-Content -Path $OutNormal -Value "const d6 = $j6;" -Encoding UTF8
Add-Content -Path $OutNormal -Value "const d7 = $j7;" -Encoding UTF8
Add-Content -Path $OutNormal -Value "const d8 = $j8;" -Encoding UTF8
Add-Content -Path $OutNormal -Value $jsLogic -Encoding UTF8

# === 寫入 AI 版 (去識別化 & 極簡化) ===
$procs_ai = $procs | Select-Object Name, @{N="UserName";E={Mask-User $_.UserName}}, Status, Path
$srvs_ai  = $srvs | Select-Object Name, DisplayName, State, StartMode, PathName, @{N="StartName";E={Mask-User $_.StartName}}

$aiHtml = [System.Text.StringBuilder]::new()
$aiHtml.Append("<!DOCTYPE html><html><head><meta charset='UTF-8'><title>$($FileAI.Replace('.html',''))</title></head><body>") | Out-Null

$aiHtml.Append("<h2>1. System Info</h2>") | Out-Null
$aiHtml.Append((Build-SimpleTable $sysinfo @("OS_Version","CPU","RAM","GPU","Disk_Usage"))) | Out-Null
$aiHtml.Append("<h2>2. Processes</h2>") | Out-Null
$aiHtml.Append((Build-SimpleTable $procs_ai @("Name","UserName","Status","Path"))) | Out-Null
$aiHtml.Append("<h2>3. Network Connections</h2>") | Out-Null
$aiHtml.Append((Build-SimpleTable $net @("ProcessName","LocalAddress","LocalPort","RemoteAddress","RemotePort","State"))) | Out-Null
$aiHtml.Append("<h2>4. Startup</h2>") | Out-Null
$aiHtml.Append((Build-SimpleTable $starts @("Name","Status","Command","Location"))) | Out-Null
$aiHtml.Append("<h2>5. Tasks</h2>") | Out-Null
$aiHtml.Append((Build-SimpleTable $tasks @("TaskName","State","Path"))) | Out-Null
$aiHtml.Append("<h2>6. Services</h2>") | Out-Null
$aiHtml.Append((Build-SimpleTable $srvs_ai @("Name","DisplayName","State","StartMode","PathName","StartName"))) | Out-Null
$aiHtml.Append("<h2>7. Installed Software</h2>") | Out-Null
$aiHtml.Append((Build-SimpleTable $software @("DisplayName","DisplayVersion","Publisher","InstallDate"))) | Out-Null
$aiHtml.Append("<h2>8. WMI Subscriptions</h2>") | Out-Null
$aiHtml.Append((Build-SimpleTable $wmi @("Filter","Consumer"))) | Out-Null

$aiHtml.Append("</body></html>") | Out-Null
Set-Content -Path $OutAI -Value $aiHtml.ToString() -Encoding UTF8

Write-Host "成功！排版修正版報告已生成：" -ForegroundColor Green
Write-Host "1. $FileNormal" -ForegroundColor Yellow
Write-Host "2. $FileAI" -ForegroundColor Cyan