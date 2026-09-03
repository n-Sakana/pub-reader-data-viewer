# e2e_click.ps1 -- VERIFICATION ONLY. Presses a named cell rectangle on the
# PIXELBRIDGE sheet with a real mouse click, at its real place on screen.
#
# The showcase draws its buttons as cells, so "pressing a button" is a mouse
# click on a cell -- exactly what a person does. Nothing in the product is
# bypassed: the click lands on the sheet, the selection changes, and the
# running OnTime tick notices it.
#
# Finding where a cell is on screen is the hard part. Application.Left/Width
# and Window.PointsToScreenPixelsX are in Excel's own units, and on a 200%
# display those are NOT screen pixels -- a first attempt clicked 15 rows and
# 22 columns short of the target. So this calibrates instead of converting:
# it clicks two harmless spots, asks Excel which cells those were, and reads
# the pixels-per-cell straight off the difference. Harmless is guaranteed by
# the product itself: HandleSelection ignores every selection that is neither
# a button's top-left cell nor inside the minimap, and the probe points are
# aimed at the pi panel, which is neither.
#
# Win32 is used HERE, in a verification script, never in the shipped VBA.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Book,
    [Parameter(Mandatory=$true)][string] $Name,
    [switch] $NoClick,
    [switch] $ViaSelection
)
$ErrorActionPreference = 'Stop'

Add-Type -Namespace Win -Name In -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, System.IntPtr e);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(System.IntPtr h);
[DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr h, out RECT r);
'@

$wb = [System.Runtime.InteropServices.Marshal]::BindToMoniker($Book)
$xl = $wb.Application
$ws = $wb.Worksheets('PIXELBRIDGE')
$target = $ws.Range($Name).Cells(1, 1)
$tr = [int]$target.Row
$tc = [int]$target.Column

# -ViaSelection：マウスを使わず、選択そのものを与える。製品が見ているのは
# 「選択セルが変わったこと」なので、ここから先はクリックと同じ道を通る。
# マウスの座標合わせが要らないぶん確実で、CI 向き。
if ($ViaSelection) {
    $ws.Range($Name).Select()
    Write-Host "selected $Name ($($target.Address($false,$false))) -- same signal a click produces"
    return
}

$h = [System.IntPtr]$xl.hwnd
$r = New-Object Win.In+RECT
[void][Win.In]::GetWindowRect($h, [ref]$r)
$W = $r.R - $r.L
$H = $r.B - $r.T

[void][Win.In]::SetForegroundWindow($h)
Start-Sleep -Milliseconds 400

function Sel { $c = $xl.Selection.Cells(1, 1); return "$($c.Row),$($c.Column)" }

function Hit([int]$x, [int]$y) {
    # クリック直後に読むと前の選択が返る。FE は毎秒のティックで動いていて、
    # その最中の COM 読み取りは一拍遅れる（実測：220ms では古い値、
    # 変化を待つと 0.7-1.3 秒）。だから「変わるまで待つ」。
    # いったん遠くの無関係なセルへ退避してから叩く。そうしないと「叩いた先が
    # いまの選択と同じ」ときに変化が起きず、当たったのに外したと読んでしまう
    # （実測：離れた 2 点が同じ座標を答えて較正が落ちた）。A1 は外枠なので、
    # 製品側の HandleSelection は何もしない。
    $ws.Range('A1').Select()
    Start-Sleep -Milliseconds 250
    $before = Sel
    [void][Win.In]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 150
    [Win.In]::mouse_event(0x0002, 0, 0, 0, [System.IntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [Win.In]::mouse_event(0x0004, 0, 0, 0, [System.IntPtr]::Zero)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $after = $before
    while ($sw.Elapsed.TotalSeconds -lt 5 -and $after -eq $before) {
        Start-Sleep -Milliseconds 120
        $after = Sel
    }
    $p = $after.Split(',')
    return @{ x = $x; y = $y; row = [int]$p[0]; col = [int]$p[1] }
}

# 較正は π の盤の上で取る。ここは Fill だけで結合していないので、叩いた
# セルがそのまま返る。結合セルを叩くと左上のセルが返り、離れた 2 点が同じ
# 座標を答えてしまう（実測：カードの見出しを叩いて 96 列ぶんずれた）。
# 盤でもボタンでもミニマップでもないので、押しても何も起きない。
$a = Hit ([int]($r.L + $W * 0.50)) ([int]($r.T + $H * 0.77))
$b = Hit ([int]($r.L + $W * 0.70)) ([int]($r.T + $H * 0.83))
if ($b.col -eq $a.col -or $b.row -eq $a.row) { throw "calibration failed: $($a.row),$($a.col) then $($b.row),$($b.col)" }
$pxCol = ($b.x - $a.x) / ($b.col - $a.col)
$pxRow = ($b.y - $a.y) / ($b.row - $a.row)
Write-Host ("calibrated: {0:N3} px/col, {1:N3} px/row" -f $pxCol, $pxRow)

$x = [int]($b.x + ($tc - $b.col) * $pxCol + $pxCol / 2)
$y = [int]($b.y + ($tr - $b.row) * $pxRow + $pxRow / 2)
Write-Host "$Name is cell $($target.Address($false,$false)) -> screen $x,$y"
if ($NoClick) { return }

# ここから先は、着地直後の COM 読み取りが一時的に撥ねられる場合も受理する。
# ボタン処理は次の OnTime ティックで動くため、その最中に Selection が読めない
# ことがある。読めれば選択位置で確かめ、読めなければ処理中に入ったと判断する。
for ($i = 1; $i -le 4; $i++) {
    [void][Win.In]::SetCursorPos($x, $y)
    Start-Sleep -Milliseconds 150
    [Win.In]::mouse_event(0x0002, 0, 0, 0, [System.IntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [Win.In]::mouse_event(0x0004, 0, 0, 0, [System.IntPtr]::Zero)
    Start-Sleep -Milliseconds 2500
    $got = $null
    try { $got = Sel } catch { }
    if ($null -eq $got) {
        Write-Host "clicked $Name at $x,$y -- Excel stopped answering, so a dialog is up (took $i click(s))"
        return
    }
    $p = $got.Split(',')
    if ([int]$p[0] -eq $tr -and [int]$p[1] -eq $tc) {
        Write-Host "clicked $Name at $x,$y (took $i click(s))"
        return
    }
    # 外したぶんだけ寄せる。ボタン以外を掴んでも製品側は無視するので安全。
    $x = [int]($x + ($tc - [int]$p[1]) * $pxCol)
    $y = [int]($y + ($tr - [int]$p[0]) * $pxRow)
    Write-Host "  landed on $got, want $tr,$tc -> retry at $x,$y"
}
throw "could not land on $Name"
