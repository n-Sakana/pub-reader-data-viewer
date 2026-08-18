# e2e_notepads.ps1 -- VERIFICATION ONLY. デモに使うメモ帳の下ごしらえ。
#
# 製品（VBA）はメモ帳を起動も終了もしない。UI Automation でもう開いている窓に
# つなぐだけ。ここは検証側の道具で、やることは 3 つだけ。
#
#   -List             … いま開いているメモ帳の窓を並べる
#   -Fill  -Hwnd n    … その窓へ原稿を書く（空の無題窓にだけ書く）
#   -Focus -Hwnd n    … その窓を前に出す（製品はフォーカスした窓を掴む）
#
# 【重要】タイトルが * で始まる窓（未保存の変更あり）には絶対に書かない。
# 実測：UI Automation が返す窓の名前は古いことがあり、中身のある窓が
# 「タイトルなし」に見えた。書く前に必ず現在値を読んで、空かどうかで判断する。
[CmdletBinding()]
param(
    [switch] $List,
    [switch] $Fill,
    [switch] $Focus,
    [int] $Hwnd = 0,
    [int] $Chars = 800
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -Namespace NpW -Name W -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h);
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr h, int c);
'@

function Get-NpWindows {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $c = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty, 'Notepad')
    $seen = @{}
    $out = @()
    foreach ($w in $root.FindAll([System.Windows.Automation.TreeScope]::Children, $c)) {
        $h = $w.Current.NativeWindowHandle
        if ($seen.ContainsKey($h)) { continue }
        $seen[$h] = 1
        $out += [pscustomobject]@{ hwnd = $h; name = $w.Current.Name; el = $w }
    }
    return $out
}

function Get-NpValue($el) {
    foreach ($ct in @([System.Windows.Automation.ControlType]::Document,
                      [System.Windows.Automation.ControlType]::Edit)) {
        $ec = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ct)
        $d = $el.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $ec)
        if ($null -ne $d) {
            try { return $d.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern) } catch { }
        }
    }
    return $null
}

$wins = Get-NpWindows

if ($List) {
    foreach ($w in $wins) {
        $v = Get-NpValue $w.el
        $len = -1
        if ($v) { $len = $v.Current.Value.Length }
        "hwnd={0}  chars={1}  name=[{2}]" -f $w.hwnd, $len, $w.name
    }
    exit 0
}

if ($Hwnd -eq 0) { throw '-Hwnd が要ります（-List で調べてください）' }
$target = $wins | Where-Object { $_.hwnd -eq $Hwnd } | Select-Object -First 1
if ($null -eq $target) { throw "hwnd $Hwnd のメモ帳が見つかりません" }

if ($Focus) {
    [void][NpW.W]::ShowWindow([System.IntPtr]$Hwnd, 9)     # SW_RESTORE
    [void][NpW.W]::SetForegroundWindow([System.IntPtr]$Hwnd)
    Start-Sleep -Milliseconds 500
    Write-Host "focused hwnd=$Hwnd  name=[$($target.name)]"
    exit 0
}

if ($Fill) {
    $v = Get-NpValue $target.el
    if ($null -eq $v) { throw "hwnd $Hwnd の編集領域に届きません" }
    $cur = $v.Current.Value
    if ($cur.Trim().Length -gt 0) {
        throw "hwnd $Hwnd には既に $($cur.Length) 文字あります。中身のある窓には書きません。"
    }
    $para = @(
     'これは VBA Pixel Bridge の負荷デモ用の原稿です。',
     'この窓が「原稿」、もう一枚が「清書」になります。',
     '原稿を一文字ずつ清書へ写しますが、一文字書くたびに、',
     'そこまで書いた分を頭から読み直して原稿と突き合わせます。',
     'だから後ろへ行くほど、一文字あたりが重くなります。',
     'BE で走らせると、この画面が動いたまま清書が伸びていきます。',
     'FE で走らせると Excel は本当に固まり、清書は最後に一度だけ出ます。'
    ) -join "`r`n"
    $sb = New-Object System.Text.StringBuilder
    while ($sb.Length -lt $Chars) { [void]$sb.AppendLine($para) }
    $text = $sb.ToString().Substring(0, $Chars)
    $v.SetValue($text)
    Write-Host "wrote $($text.Length) chars into hwnd=$Hwnd"
    exit 0
}

Write-Host 'use -List / -Fill -Hwnd n / -Focus -Hwnd n'
