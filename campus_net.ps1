param(
    [switch]$Test,
    [switch]$Status
)

# ============================================================
# 校园网助手 - Dr.COM 网页认证自动登录工具
# 单文件、无第三方依赖、PowerShell 5.1 兼容
# ============================================================

$ErrorActionPreference = 'Stop'

if ($PSScriptRoot) { $script:Root = $PSScriptRoot }
else { $script:Root = Split-Path -Parent $MyInvocation.MyCommand.Definition }

$script:CfgPath = Join-Path $script:Root 'config.json'
$script:LogPath = Join-Path $script:Root 'campus_net.log'
$script:Version = '1.0'

$script:UA_Phone = 'Mozilla/5.0 (Linux; Android 15; Pixel 9) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Mobile Safari/537.36'
$script:UA_PC    = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36'

# 注意：这里只放“出厂默认值”，真实的账号密码请写进同目录的 config.json（已被 .gitignore 排除），
# 不要把个人账号密码提交到仓库里。门户地址需替换为你所在学校 Dr.COM 服务器的地址。
$script:DefaultCfg = [ordered]@{
    account      = ''
    password     = ''
    portal       = ''
    terminal     = 'phone'
    autoRelogin  = $true
    checkSeconds = 30
    startMinimized = $false
}

# ---------------- 配置 ----------------

function Get-Config {
    $cfg = [ordered]@{}
    foreach ($k in $script:DefaultCfg.Keys) { $cfg[$k] = $script:DefaultCfg[$k] }
    if (Test-Path -LiteralPath $script:CfgPath) {
        try {
            $raw = Get-Content -LiteralPath $script:CfgPath -Raw -Encoding UTF8
            $obj = $raw | ConvertFrom-Json
            foreach ($p in $obj.PSObject.Properties) {
                if ($cfg.Contains($p.Name)) { $cfg[$p.Name] = $p.Value }
            }
        } catch { }
    }
    return $cfg
}

function Save-Config($cfg) {
    try {
        $cfg | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $script:CfgPath -Encoding UTF8
        return $true
    } catch { return $false }
}

# ---------------- 日志 ----------------

function Write-Log([string]$msg, [string]$level = 'INFO') {
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $level, $msg
    try { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 } catch { }
    if ($script:OnLog) { & $script:OnLog $line }
    elseif ($Test) { Write-Output $line }
}

# ---------------- 网络信息 ----------------

function Get-NetInfo {
    $route = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
             Where-Object { $_.NextHop -ne '0.0.0.0' } |
             Sort-Object RouteMetric, InterfaceMetric | Select-Object -First 1
    if (-not $route) { return $null }
    $idx = $route.InterfaceIndex
    $ipObj = Get-NetIPAddress -InterfaceIndex $idx -AddressFamily IPv4 -ErrorAction SilentlyContinue |
             Where-Object { $_.IPAddress -notlike '169.254.*' } | Select-Object -First 1
    $adp = Get-NetAdapter -InterfaceIndex $idx -ErrorAction SilentlyContinue
    if (-not $ipObj) { return $null }
    return [pscustomobject]@{
        Ip       = $ipObj.IPAddress
        Mac      = ($adp.MacAddress -replace '-', '')
        Adapter  = $adp.Name
    }
}

# ---------------- HTTP ----------------

function Invoke-Portal {
    param(
        [string]$Url,
        [string]$Body = $null,
        [string]$Ua = $script:UA_Phone,
        [int]$TimeoutSec = 15
    )
    $script:LastHttpError = ''
    try {
        $req = [System.Net.HttpWebRequest]::Create($Url)
        $req.Proxy = $null
        $req.UserAgent = $Ua
        $req.Timeout = $TimeoutSec * 1000
        $req.ReadWriteTimeout = $TimeoutSec * 1000
        $req.KeepAlive = $false
        if ($Body -ne $null -and $Body -ne '') {
            $bytes = [System.Text.Encoding]::GetEncoding(936).GetBytes($Body)
            $req.Method = 'POST'
            $req.ContentType = 'application/x-www-form-urlencoded'
            $req.ContentLength = $bytes.Length
            $stream = $req.GetRequestStream()
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Close()
        } else {
            $req.Method = 'GET'
        }
        $resp = $req.GetResponse()
        $ms = New-Object System.IO.MemoryStream
        $resp.GetResponseStream().CopyTo($ms)
        $bytes = $ms.ToArray()
        $ms.Close()
        $resp.Close()
        # 门户页面是 GBK，接口返回 JSON 是 UTF-8：先按 UTF-8 严格解码，失败退回 GBK
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        try { $text = $strict.GetString($bytes) }
        catch { $text = [System.Text.Encoding]::GetEncoding(936).GetString($bytes) }
        return $text
    } catch {
        # 取最内层异常，才能看到真正的 Socket 错误（外层只是包装）
        $ex = $_.Exception
        $depth = 0
        while ($ex.InnerException -and $depth -lt 5) { $ex = $ex.InnerException; $depth++ }
        $script:LastHttpError = $ex.GetType().Name + ': ' + $ex.Message
        # 断网/换网后进程内可能残留失效的连接对象，主动丢弃，下次请求重建
        try { [System.Net.ServicePointManager]::FindServicePoint([uri]$Url).Close() } catch { }
        return $null
    }
}

function ConvertTo-FormValue([string]$s) {
    if ($null -eq $s) { return '' }
    return [System.Uri]::EscapeDataString($s)
}

# ---------------- 业务动作 ----------------

function Get-OnlineSession([string]$Portal, [string]$Account) {
    # 必须返回“数组本身”。直接 return @($o.list) 时单元素数组会被展开成单个对象，
    # 调用方拿到的 $ss.Count 变成 $null，于是永远判成“离线”——这正是反复重登的根因。
    $empty = @()
    $url = '{0}:801/eportal/portal/online_list?callback=cb&user_account={1}&jsVersion=4.1.3' -f $Portal, (ConvertTo-FormValue $Account)
    $t = Invoke-Portal -Url $url
    if (-not $t) {
        $script:LastSessionQuery = '请求失败: ' + $script:LastHttpError
        return ,$empty
    }
    $i = $t.IndexOf('(')
    if ($i -lt 0) {
        $script:LastSessionQuery = '返回非 JSONP: ' + $t.Substring(0, [Math]::Min(60, $t.Length))
        return ,$empty
    }
    $j = $t.LastIndexOf(')')
    if ($j -le $i) {
        $script:LastSessionQuery = '返回缺少右括号，长度 ' + $t.Length
        return ,$empty
    }
    $json = $t.Substring($i + 1, $j - $i - 1)
    try {
        $o = $json | ConvertFrom-Json
        $list = @($o.list | Where-Object { $null -ne $_ })
        if ($list.Count -gt 0) { return ,$list }
        $script:LastSessionQuery = '接口返回空列表（result=' + $o.result + ' msg=' + $o.msg + '）'
    } catch {
        $script:LastSessionQuery = 'JSON 解析失败: ' + $_.Exception.Message
    }
    return ,$empty
}

# 同一账号可能有多台设备同时在线（手机 + 电脑），必须挑出本机那一条，
# 否则会拿别的设备的会话判断名额，导致“名额永远对不上”。
function Select-LocalSession($sessions) {
    $arr = @($sessions | Where-Object { $null -ne $_ })
    if ($arr.Count -eq 0) { return $null }
    $ni = Get-NetInfo
    if ($ni) {
        foreach ($s in $arr) {
            if ("$($s.online_ip)" -eq $ni.Ip) { return $s }
        }
        foreach ($s in $arr) {
            if ("$($s.online_mac)" -eq $ni.Mac) { return $s }
        }
        # 本机 IP/MAC 都不在列表里：说明在线的是别的设备，本机并未认证
        return $null
    }
    return $arr[0]
}

function Test-Internet([int]$TimeoutSec = 6) {
    $t = Invoke-Portal -Url 'http://connect.rom.miui.com/generate_204' -TimeoutSec $TimeoutSec
    return ($null -ne $t)
}

# 用门户首页判断是否已认证在线：已认证会返回“注销页”，内含 uid / v4ip。
# 该请求走 80 端口，比 801 接口更稳，用于在 online_list 查询失败时兜底，避免误判掉线后反复重登。
function Get-PortalState([string]$Portal) {
    $url = '{0}/?t={1}' -f $Portal.TrimEnd('/'), [int](Get-Date -UFormat %s)
    $t = Invoke-Portal -Url $url -TimeoutSec 8
    if (-not $t) {
        return [pscustomobject]@{ Ok = $false; Online = $false; Uid = ''; Ip = ''; Error = $script:LastHttpError }
    }
    $online = ($t -match '注销页') -or ($t -match "uid='")
    $uid = ''
    if ($t -match "uid='([^']*)'") { $uid = $Matches[1] }
    $ip = ''
    if ($t -match "v4ip='([^']*)'") { $ip = $Matches[1] }
    return [pscustomobject]@{ Ok = $true; Online = $online; Uid = $uid; Ip = $ip; Error = '' }
}

function Login-Portal($cfg, [string]$Terminal = $null) {
    $portal = $cfg.portal.TrimEnd('/')
    if (-not $Terminal) { $Terminal = $cfg.terminal }
    $ua = $script:UA_Phone
    $r6 = '1'
    $termName = '手机'
    if ($Terminal -ne 'phone') { $ua = $script:UA_PC; $r6 = '0'; $termName = '电脑' }

    $body = 'DDDDD={0}&upass={1}&R1=0&R2=&R3=0&R6={2}&para=00&0MKKey=123456&v6ip=' -f `
            (ConvertTo-FormValue $cfg.account), (ConvertTo-FormValue $cfg.password), $r6

    $t = Invoke-Portal -Url ('{0}/a79.htm' -f $portal) -Body $body -Ua $ua
    if (-not $t) {
        $why = ''
        if ($script:LastHttpError) { $why = '（' + $script:LastHttpError + '）' }
        Write-Log ('登录请求失败：无法连接门户{0}，请检查是否已连上校园 Wi-Fi。' -f $why) 'ERROR'
        return [pscustomobject]@{ Ok = $false; Message = '无法连接门户' }
    }
    if ($t -match 'Dr\.COMWebLoginID_3' -or $t -match '认证成功页') {
        Write-Log ('登录成功，终端类型={0}' -f $termName) 'OK'
        $ss = @()
        for ($k = 0; $k -lt 2; $k++) {
            Start-Sleep -Milliseconds 1500
            $ss = Get-OnlineSession -Portal $portal -Account $cfg.account
            if ($ss.Count -gt 0) { break }
        }
        if ($ss.Count -gt 0) {
            $f = $ss[0].phone_flag
            $tag = '电脑名额'
            if ("$f" -eq '1') { $tag = '手机名额' }
            Write-Log ('服务端记录：ip={0} mac={1} 计费标记={2}({3})' -f $ss[0].online_ip, $ss[0].online_mac, $f, $tag) 'OK'
            return [pscustomobject]@{ Ok = $true; Message = ('登录成功 · {0}' -f $tag); PhoneFlag = $f }
        }
        return [pscustomobject]@{ Ok = $true; Message = ('登录成功 · 会话同步中'); PhoneFlag = $null }
    }
    if ($t -match 'clientip online') {
        # AC 侧仍保留该 IP 的会话：看它是不是已经是我们想要的名额
        $ss2 = Get-OnlineSession -Portal $portal -Account $cfg.account
        if ($ss2.Count -gt 0) {
            $f2 = "$($ss2[0].phone_flag)"
            $tag2 = '电脑名额'
            if ($f2 -eq '1') { $tag2 = '手机名额' }
            if ($Terminal -eq 'phone' -and $f2 -eq '1') {
                Write-Log ('该 IP 已在线且已是{0}，无需重复登录。' -f $tag2) 'OK'
                return [pscustomobject]@{ Ok = $true; Message = ('已在线 · {0}' -f $tag2); PhoneFlag = $f2 }
            }
            if ($Terminal -ne 'phone' -and $f2 -ne '1') {
                Write-Log ('该 IP 已在线且已是{0}，无需重复登录。' -f $tag2) 'OK'
                return [pscustomobject]@{ Ok = $true; Message = ('已在线 · {0}' -f $tag2); PhoneFlag = $f2 }
            }
            Write-Log ('当前会话占用的是{0}，与所选终端类型不一致，需先注销。' -f $tag2) 'WARN'
            return [pscustomobject]@{ Ok = $false; Message = ('已在线（{0}）' -f $tag2) }
        }
        $ps = Get-PortalState $portal
        if ($ps.Ok -and $ps.Online) {
            Write-Log ('门户显示本机已认证在线（账号 {0}，IP {1}），不再重复提交登录。' -f $ps.Uid, $ps.Ip) 'OK'
            return [pscustomobject]@{ Ok = $true; Message = '已在线'; PhoneFlag = $null }
        }
        $why = ''
        if ($script:LastSessionQuery) { $why = '：' + $script:LastSessionQuery }
        Write-Log ('AC 端仍保留该 IP 的会话，但在线记录暂不可得{0}，稍后重试。' -f $why) 'WARN'
        return [pscustomobject]@{ Ok = $false; Message = 'AC 端未释放旧会话' }
    }
    $msg = ''
    if ($t -match "msga='([^']*)'") { $msg = $Matches[1] }
    if ($t -match 'Msg=(\d+)') { $msg = ('Msg={0} {1}' -f $Matches[1], $msg).Trim() }
    if (-not $msg) { $msg = '未知失败原因' }
    Write-Log ('登录失败：{0}' -f $msg) 'ERROR'
    return [pscustomobject]@{ Ok = $false; Message = $msg }
}

function Logout-Portal($cfg) {
    $portal = $cfg.portal.TrimEnd('/')
    $ni = Get-NetInfo
    if (-not $ni) { return [pscustomobject]@{ Ok = $false; Message = '未取到本机地址' } }
    $url = '{0}:801/eportal/portal/logout?callback=cb&login_method=0&user_account=drcom&user_password=123&wlan_user_ip={1}&wlan_user_mac={2}&jsVersion=4.1.3' -f `
           $portal, $ni.Ip, $ni.Mac.ToUpper()
    $t = Invoke-Portal -Url $url
    if ($t -and $t -match '"result":1') {
        Write-Log '注销成功。' 'OK'
        return [pscustomobject]@{ Ok = $true; Message = '注销成功' }
    }
    Write-Log '注销失败（可能本就不在线）。' 'WARN'
    return [pscustomobject]@{ Ok = $false; Message = '注销失败' }
}

function Unbind-Mac($cfg) {
    # 清掉 AC 侧“MAC + 终端类型”绑定，否则新登录会沿用旧类型，
    # 表现为返回 clientip online、phone_flag 始终为 0。
    $portal = $cfg.portal.TrimEnd('/')
    $ni = Get-NetInfo
    if (-not $ni) { return $false }
    $ipInt = 0
    foreach ($part in $ni.Ip.Split('.')) { $ipInt = $ipInt * 256 + [int]$part }
    $url = '{0}:801/eportal/portal/mac/unbind?callback=cb&user_account={1}&wlan_user_mac={2}&wlan_user_ip={3}&jsVersion=4.1.3' -f `
           $portal, (ConvertTo-FormValue $cfg.account), $ni.Mac.ToUpper(), $ipInt
    $t = Invoke-Portal -Url $url
    if ($t -and $t -match '"result":1') {
        Write-Log '已解绑本机 MAC 的终端类型记录。' 'INFO'
        return $true
    }
    Write-Log '解绑 MAC 未成功（可能本就未绑定），继续尝试登录。' 'WARN'
    return $false
}

function Invoke-LogoutUntilOffline($cfg) {
    for ($i = 0; $i -lt 4; $i++) {
        if ((Get-OnlineSession -Portal $cfg.portal.TrimEnd('/') -Account $cfg.account).Count -eq 0) { return $true }
        Logout-Portal $cfg | Out-Null
        Start-Sleep -Seconds 2
    }
    return ((Get-OnlineSession -Portal $cfg.portal.TrimEnd('/') -Account $cfg.account).Count -eq 0)
}

function Invoke-Relogin($cfg) {
    $portal = $cfg.portal.TrimEnd('/')
    $wantPhone = ($cfg.terminal -eq 'phone')

    # 门户都连不上时不要折腾解绑/注销，那只会把日志刷满
    $ps = Get-PortalState $portal
    if (-not $ps.Ok) {
        Write-Log ('门户不可达，暂不认证：' + $ps.Error) 'WARN'
        return [pscustomobject]@{ Ok = $false; Message = '门户不可达' }
    }

    # 先看现状：已在线且名额正是所需时不做任何动作，避免把正在用的连接打断
    $ss = Get-OnlineSession -Portal $portal -Account $cfg.account
    $mine = Select-LocalSession $ss
    if ($mine) {
        $flag = "$($mine.phone_flag)"
        $tag = '电脑名额'
        if ($flag -eq '1') { $tag = '手机名额' }
        $match = (($wantPhone -and $flag -eq '1') -or ((-not $wantPhone) -and $flag -ne '1'))
        if ($match) {
            Write-Log ('当前已在线，占的是{0}，无需重复认证。' -f $tag) 'OK'
            return [pscustomobject]@{ Ok = $true; Message = ('已在线 · ' + $tag); PhoneFlag = $flag }
        }
        Write-Log ('当前在线占的是{0}，与所选终端不一致，先注销再重新认证…' -f $tag) 'WARN'
        Logout-Portal $cfg | Out-Null
        Start-Sleep -Seconds 2
    }

    # 离线或刚注销：清掉 MAC 旧绑定后以目标终端类型认证
    Unbind-Mac $cfg | Out-Null
    Start-Sleep -Milliseconds 1200
    $r = Login-Portal $cfg
    for ($i = 0; $i -lt 2 -and -not $r.Ok; $i++) {
        Write-Log '登录未成功，注销后重试…' 'WARN'
        Logout-Portal $cfg | Out-Null
        Start-Sleep -Seconds 2
        Unbind-Mac $cfg | Out-Null
        Start-Sleep -Milliseconds 1200
        $r = Login-Portal $cfg
    }
    return $r
}

# ============================================================
# 自检模式
# ============================================================

if ($Status) {
    # 只读诊断：不改动任何会话，用来确认“到底在线没有、占的是哪个名额”
    $ni = Get-NetInfo
    Write-Output '=== 网络信息 ==='
    if ($ni) { Write-Output ('  IP={0} MAC={1} 网卡={2}' -f $ni.Ip, $ni.Mac, $ni.Adapter) }
    else { Write-Output '  未取到默认路由网卡' }
    $cfg = Get-Config
    Write-Output ('=== 配置 === 账号={0} 门户={1} 终端={2}' -f $cfg.account, $cfg.portal, $cfg.terminal)
    if (-not $cfg.account -or -not $cfg.portal) { Write-Output '  配置不完整：请先填写账号与认证地址（config.json 或图形界面）。'; exit 0 }
    Write-Output '=== 在线会话（801 接口） ==='
    $ss = Get-OnlineSession -Portal $cfg.portal.TrimEnd('/') -Account $cfg.account
    if ($ss.Count -gt 0) {
        foreach ($s in $ss) {
            Write-Output ('  session={0} ip={1} mac={2} phone_flag={3} 开始={4} 时长={5}s' -f $s.online_session, $s.online_ip, $s.online_mac, $s.phone_flag, $s.online_time, $s.time_long)
        }
    } else {
        Write-Output ('  无会话记录。原因: {0}' -f $(if ($script:LastSessionQuery) { $script:LastSessionQuery } else { '接口返回空列表（可能 AC 瞬时未同步）' }))
    }
    Write-Output '=== 门户首页（80 端口） ==='
    $ps = Get-PortalState $cfg.portal
    if ($ps.Ok) { Write-Output ('  在线={0} 账号={1} IP={2}' -f $ps.Online, $ps.Uid, $ps.Ip) }
    else { Write-Output ('  访问失败: {0}' -f $ps.Error) }
    Write-Output '=== POST 认证通道探测（不会改变在线状态） ==='
    $probeBody = 'DDDDD={0}&upass={1}&R1=0&R2=&R3=0&R6=1&para=00&0MKKey=123456&v6ip=' -f (ConvertTo-FormValue $cfg.account), (ConvertTo-FormValue $cfg.password)
    $pr = Invoke-Portal -Url ($cfg.portal.TrimEnd('/') + '/a79.htm') -Body $probeBody
    if ($pr) {
        $pt = ''
        if ($pr -match '<title>([^<]*)</title>') { $pt = $Matches[1] }
        Write-Output ('  POST 成功 len={0} title={1} clientipOnline={2}' -f $pr.Length, $pt, ($pr -match 'clientip online'))
    } else {
        Write-Output ('  POST 失败: ' + $script:LastHttpError)
    }
    Write-Output ('外网可达: {0}' -f (Test-Internet))
    exit 0
}

if ($Test) {
    Write-Output '=== 网络信息 ==='
    Get-NetInfo | Format-List | Out-String | Write-Output
    $cfg = Get-Config
    Write-Output ('=== 配置 === 账号={0} 门户={1} 终端={2}' -f $cfg.account, $cfg.portal, $cfg.terminal)
    Write-Output '=== 登录前会话 ==='
    Get-OnlineSession -Portal $cfg.portal.TrimEnd('/') -Account $cfg.account | ForEach-Object {
        Write-Output ('  session={0} ip={1} mac={2} phone_flag={3}' -f $_.online_session, $_.online_ip, $_.online_mac, $_.phone_flag)
    }
    Write-Output '=== 执行注销+登录（完整链路） ==='
    $r = Invoke-Relogin $cfg
    Write-Output ('结果: Ok={0} Message={1} PhoneFlag={2}' -f $r.Ok, $r.Message, $r.PhoneFlag)
    Write-Output ('外网可达: {0}' -f (Test-Internet))
    exit 0
}

# ============================================================
# 图形界面
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# 单实例保护：多个实例同时重登会互相注销、把会话搅乱
$script:Mutex = New-Object System.Threading.Mutex($false, 'Global\DrcomCampusNetHelper')
$gotLock = $false
try { $gotLock = $script:Mutex.WaitOne(0, $false) } catch { $gotLock = $true }
if (-not $gotLock) {
    [System.Windows.Forms.MessageBox]::Show('校园网助手已经在运行了。' + [Environment]::NewLine + [Environment]::NewLine + '请看屏幕右下角托盘里的盾牌图标，双击即可打开。', '校园网助手', 'OK', 'Information') | Out-Null
    exit 0
}

$cfg = Get-Config
$script:Busy = $false
$script:Exiting = $false

$form = New-Object System.Windows.Forms.Form
$form.Text = '校园网助手'
$form.Size = New-Object System.Drawing.Size(430, 520)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

$lblAccount = New-Object System.Windows.Forms.Label
$lblAccount.Text = '账号'
$lblAccount.Location = New-Object System.Drawing.Point(18, 18)
$lblAccount.Size = New-Object System.Drawing.Size(60, 22)
$form.Controls.Add($lblAccount)

$txtAccount = New-Object System.Windows.Forms.TextBox
$txtAccount.Location = New-Object System.Drawing.Point(85, 15)
$txtAccount.Size = New-Object System.Drawing.Size(310, 25)
$txtAccount.Text = $cfg.account
$form.Controls.Add($txtAccount)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = '密码'
$lblPass.Location = New-Object System.Drawing.Point(18, 53)
$lblPass.Size = New-Object System.Drawing.Size(60, 22)
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(85, 50)
$txtPass.Size = New-Object System.Drawing.Size(310, 25)
$txtPass.UseSystemPasswordChar = $true
$txtPass.Text = $cfg.password
$form.Controls.Add($txtPass)

$lblPortal = New-Object System.Windows.Forms.Label
$lblPortal.Text = '认证地址'
$lblPortal.Location = New-Object System.Drawing.Point(18, 88)
$lblPortal.Size = New-Object System.Drawing.Size(62, 22)
$form.Controls.Add($lblPortal)

$txtPortal = New-Object System.Windows.Forms.TextBox
$txtPortal.Location = New-Object System.Drawing.Point(85, 85)
$txtPortal.Size = New-Object System.Drawing.Size(310, 25)
$txtPortal.Text = $cfg.portal
$form.Controls.Add($txtPortal)

$lblTerm = New-Object System.Windows.Forms.Label
$lblTerm.Text = '终端类型'
$lblTerm.Location = New-Object System.Drawing.Point(18, 123)
$lblTerm.Size = New-Object System.Drawing.Size(62, 22)
$form.Controls.Add($lblTerm)

$cmbTerm = New-Object System.Windows.Forms.ComboBox
$cmbTerm.Location = New-Object System.Drawing.Point(85, 120)
$cmbTerm.Size = New-Object System.Drawing.Size(150, 25)
$cmbTerm.DropDownStyle = 'DropDownList'
[void]$cmbTerm.Items.Add('手机（占手机名额）')
[void]$cmbTerm.Items.Add('电脑（占 PC 名额）')
if ($cfg.terminal -eq 'phone') { $cmbTerm.SelectedIndex = 0 } else { $cmbTerm.SelectedIndex = 1 }
$form.Controls.Add($cmbTerm)

$chkAuto = New-Object System.Windows.Forms.CheckBox
$chkAuto.Text = '掉线自动重连'
$chkAuto.Location = New-Object System.Drawing.Point(250, 122)
$chkAuto.Size = New-Object System.Drawing.Size(140, 24)
$chkAuto.Checked = [bool]$cfg.autoRelogin
$form.Controls.Add($chkAuto)

$btnLogin = New-Object System.Windows.Forms.Button
$btnLogin.Text = '登录'
$btnLogin.Location = New-Object System.Drawing.Point(18, 158)
$btnLogin.Size = New-Object System.Drawing.Size(150, 34)
$form.Controls.Add($btnLogin)

$btnLogout = New-Object System.Windows.Forms.Button
$btnLogout.Text = '注销'
$btnLogout.Location = New-Object System.Drawing.Point(178, 158)
$btnLogout.Size = New-Object System.Drawing.Size(110, 34)
$form.Controls.Add($btnLogout)

$btnHide = New-Object System.Windows.Forms.Button
$btnHide.Text = '隐藏到托盘'
$btnHide.Location = New-Object System.Drawing.Point(296, 158)
$btnHide.Size = New-Object System.Drawing.Size(100, 34)
$form.Controls.Add($btnHide)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = '状态：未检测'
$lblStatus.Location = New-Object System.Drawing.Point(18, 202)
$lblStatus.Size = New-Object System.Drawing.Size(380, 24)
$lblStatus.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblStatus)

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = ''
$lblInfo.Location = New-Object System.Drawing.Point(18, 226)
$lblInfo.Size = New-Object System.Drawing.Size(380, 22)
$lblInfo.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblInfo)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(18, 252)
$txtLog.Size = New-Object System.Drawing.Size(378, 214)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$form.Controls.Add($txtLog)

$script:OnLog = {
    param($line)
    if ($txtLog.Text.Length -gt 40000) { $txtLog.Text = '' }
    $txtLog.AppendText($line + [Environment]::NewLine)
}

function Set-Status([string]$text, [string]$color, [string]$info = '') {
    $lblStatus.Text = '状态：' + $text
    switch ($color) {
        'ok'   { $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 140, 70) }
        'bad'  { $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(190, 40, 40) }
        'busy' { $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(200, 120, 0) }
        default { $lblStatus.ForeColor = [System.Drawing.Color]::DimGray }
    }
    $lblInfo.Text = $info
    $form.Refresh()
}

function Sync-Cfg {
    $cfg.account = $txtAccount.Text.Trim()
    $cfg.password = $txtPass.Text
    $cfg.portal = $txtPortal.Text.Trim()
    if ($cmbTerm.SelectedIndex -eq 0) { $cfg.terminal = 'phone' } else { $cfg.terminal = 'pc' }
    $cfg.autoRelogin = $chkAuto.Checked
    Save-Config $cfg | Out-Null
}

function Update-State {
    $ni = Get-NetInfo
    if (-not $ni) {
        Set-Status '未连接网络' 'bad' ''
        return
    }
    $ss = Get-OnlineSession -Portal $cfg.portal.TrimEnd('/') -Account $cfg.account
    $mine = Select-LocalSession $ss
    if ($mine) {
        $f = "$($mine.phone_flag)"
        $tag = '电脑名额'
        if ($f -eq '1') { $tag = '手机名额' }
        $t = [TimeSpan]::FromSeconds([int]$mine.time_long)
        Set-Status ('在线 · ' + $tag) 'ok' ('IP {0} · MAC {1} · 已在线 {2:hh\:mm\:ss}' -f $mine.online_ip, $mine.online_mac, $t)
    } else {
        # 801 接口有时查不到会话，但门户首页仍显示已认证：以门户为准，避免误报离线
        $ps = Get-PortalState $cfg.portal
        if ($ps.Ok -and $ps.Online) {
            Set-Status '在线（名额待确认）' 'ok' ('IP {0} · 账号 {1} · 在线记录查询失败' -f $ps.Ip, $ps.Uid)
        } else {
            $extra = ''
            if (-not $ps.Ok -and $ps.Error) { $extra = ' · 门户不可达' }
            Set-Status '离线（未认证）' 'bad' ('本机 IP {0} · MAC {1}{2}' -f $ni.Ip, $ni.Mac, $extra)
        }
    }
}

$script:Working = $false

function Do-Login {
    if ($script:Working) { return }
    $script:Working = $true
    try {
        Sync-Cfg
        if (-not $cfg.account -or -not $cfg.password -or -not $cfg.portal) {
            Set-Status '配置不完整' 'bad' '请填写账号、密码与认证地址'
            [void][System.Windows.Forms.MessageBox]::Show(
                ('请先把账号、密码、认证地址填完整。' + [Environment]::NewLine + [Environment]::NewLine +
                 '认证地址就是浏览器跳转到的那个校园网登录页，例如：' + [Environment]::NewLine + 'http://172.25.251.2'),
                '校园网助手', 'OK', 'Warning')
            return
        }
        Set-Status '正在登录…' 'busy' ''
        $r = Invoke-Relogin $cfg
        if ($r.Ok) { Set-Status ('在线 · ' + $r.Message) 'ok' ''; $script:refreshTimer.Start() }
        else { Update-State }
    } catch {
        Write-Log ('登录异常：' + $_.Exception.Message) 'ERROR'
        Set-Status '登录异常' 'bad' ''
    } finally {
        $script:Working = $false
    }
}

$btnLogin.Add_Click({
    $script:Working = $false
    Do-Login
})

$btnLogout.Add_Click({
    if ($script:Working) { return }
    $script:Working = $true
    try {
        Sync-Cfg
        Set-Status '正在注销…' 'busy' ''
        Logout-Portal $cfg | Out-Null
        Start-Sleep -Milliseconds 800
        Update-State
    } finally {
        $script:Working = $false
    }
})

$btnHide.Add_Click({ $form.Hide() })

# ---- 托盘 ----
$notify = New-Object System.Windows.Forms.NotifyIcon
try {
    $notify.Icon = [System.Drawing.SystemIcons]::Shield
} catch {
    $notify.Icon = [System.Drawing.SystemIcons]::Application
}
$notify.Text = '校园网助手'
$notify.Visible = $true
$notify.Add_DoubleClick({ $form.ShowInTaskbar = $true; $form.Show(); $form.WindowState = 'Normal'; $form.Activate() })

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miShow = $menu.Items.Add('显示主界面')
$miLogin = $menu.Items.Add('立即登录')
$miLogout = $menu.Items.Add('注销')
[void]$menu.Items.Add('-')
$miExit = $menu.Items.Add('退出')
$notify.ContextMenuStrip = $menu
$miShow.Add_Click({ $form.ShowInTaskbar = $true; $form.Show(); $form.WindowState = 'Normal'; $form.Activate() })
$miLogin.Add_Click({ $script:Working = $false; Do-Login })
$miLogout.Add_Click({ Logout-Portal $cfg | Out-Null; Start-Sleep -Milliseconds 800; Update-State })
$miExit.Add_Click({
    $script:Exiting = $true
    $timer.Stop()
    $script:refreshTimer.Stop()
    $notify.Visible = $false
    $form.Close()
})

# ---- 定时检测 ----
$script:refreshTimer = New-Object System.Windows.Forms.Timer
$script:refreshTimer.Interval = 9000
$script:refreshTimer.Add_Tick({
    $script:refreshTimer.Stop()
    Update-State
})

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [int]($cfg.checkSeconds) * 1000
if ($timer.Interval -lt 15000) { $timer.Interval = 30000 }
$script:FailStreak = 0
$script:WaitNetLogged = $false
$script:BaseInterval = $timer.Interval
$timer.Add_Tick({
    if ($script:Working) { return }
    $ss = Get-OnlineSession -Portal $cfg.portal.TrimEnd('/') -Account $cfg.account
    $mine = Select-LocalSession $ss
    if ($mine) {
        $script:FailStreak = 0
        $timer.Interval = $script:BaseInterval
        # 在线但名额与所选终端不符：自动切换（会短暂断线，失败会在下轮重试）
        $wantPhone = ($cfg.terminal -eq 'phone')
        $isPhone = ("$($mine.phone_flag)" -eq '1')
        if ($chkAuto.Checked -and ($wantPhone -ne $isPhone)) {
            $cur = '电脑名额'; if ($isPhone) { $cur = '手机名额' }
            $dst = '电脑名额'; if ($wantPhone) { $dst = '手机名额' }
            Write-Log ('当前占的是{0}，目标为{1}，开始切换…' -f $cur, $dst) 'WARN'
            $script:Working = $true
            try {
                $r = Invoke-Relogin $cfg
                if ($r.Ok) { Set-Status ('在线 · ' + $r.Message) 'ok' ''; $script:refreshTimer.Start() }
            } catch {
                Write-Log ('切换名额异常：' + $_.Exception.Message) 'ERROR'
            } finally {
                $script:Working = $false
            }
            Update-State
            return
        }
        Update-State
        return
    }
    # 会话查询失败不等于掉线：门户首页显示已认证就不重登，避免反复提交把会话搅乱
    $ps = Get-PortalState $cfg.portal
    if ($ps.Ok -and $ps.Online) {
        $script:FailStreak = 0
        $timer.Interval = $script:BaseInterval
        Update-State
        return
    }
    if (-not $chkAuto.Checked) { Update-State; return }
    # 开机初期 Wi-Fi 可能尚未连上：门户不可达且外网也不通，属“网络未就绪”，不进入失败退避
    if (-not $ps.Ok) {
        if (-not (Test-Internet)) {
            if (-not $script:WaitNetLogged) {
                Write-Log '网络尚未就绪（门户不可达、外网不通），按 30 秒间隔继续等待…' 'WARN'
                $script:WaitNetLogged = $true
            }
            $timer.Interval = $script:BaseInterval
            Update-State
            return
        }
    }
    $script:WaitNetLogged = $false
    $script:FailStreak = $script:FailStreak + 1
    $wait = [int][Math]::Min(300, [Math]::Max(30, ($script:BaseInterval / 1000) * [Math]::Pow(2, $script:FailStreak - 1)))
    $timer.Interval = $wait * 1000
    Write-Log ('检测到已掉线（第 {0} 次），尝试自动重连，下次间隔 {1} 秒…' -f $script:FailStreak, $wait) 'WARN'
    $script:Working = $true
    try {
        $r = Invoke-Relogin $cfg
        if ($r.Ok) {
            $script:FailStreak = 0
            $timer.Interval = $script:BaseInterval
            Set-Status ('在线 · ' + $r.Message) 'ok' ''
            $script:refreshTimer.Start()
        }
    } finally {
        $script:Working = $false
    }
    Update-State
})
$timer.Start()

$form.Add_Shown({
    Write-Log ('校园网助手 v{0} 启动' -f $script:Version) 'INFO'
    Update-State
    # 开机自启用：直接缩进托盘，不抢焦点
    if ([bool]$cfg.startMinimized) {
        $form.ShowInTaskbar = $false
        $form.WindowState = 'Minimized'
        $form.Hide()
    }
})

$form.Add_FormClosing({
    param($sender, $e)
    if ($script:Exiting) {
        $timer.Stop()
        $script:refreshTimer.Stop()
        $notify.Visible = $false
        return
    }
    if ($e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
        $e.Cancel = $true
        $form.Hide()
        return
    }
    $notify.Visible = $false
    $timer.Stop()
})

[void]$form.ShowDialog()
$notify.Dispose()
try { $script:Mutex.ReleaseMutex() } catch { }
