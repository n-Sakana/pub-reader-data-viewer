# e2e_shotwin.ps1 -- VERIFICATION ONLY. Captures ONE window's own content with
# PrintWindow, so the shot is correct even when that window is behind others.
# Nothing is brought to the front and nobody's work is interrupted.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][int] $ProcessId,
    [Parameter(Mandatory=$true)][string] $Out,
    [int] $Scale = 100
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -Namespace Win -Name Cap -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
public struct RECT { public int Left, Top, Right, Bottom; }
'@
[void][Win.Cap]::SetProcessDPIAware()

$p = Get-Process -Id $ProcessId -ErrorAction Stop
$h = $p.MainWindowHandle
if ($h -eq [IntPtr]::Zero) { throw "process $ProcessId has no main window" }
$r = New-Object Win.Cap+RECT
[void][Win.Cap]::GetWindowRect($h, [ref]$r)
$w = $r.Right - $r.Left
$hh = $r.Bottom - $r.Top
Write-Host "window rect $($r.Left),$($r.Top) ${w}x${hh}"

$bmp = New-Object System.Drawing.Bitmap($w, $hh)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$dc = $g.GetHdc()
# 2 = PW_RENDERFULLCONTENT, needed for windows drawn with modern compositing
[void][Win.Cap]::PrintWindow($h, $dc, 2)
$g.ReleaseHdc($dc)
$g.Dispose()

if ($Scale -ne 100) {
    $nw = [int]($w * $Scale / 100); $nh = [int]($hh * $Scale / 100)
    $small = New-Object System.Drawing.Bitmap($nw, $nh)
    $g2 = [System.Drawing.Graphics]::FromImage($small)
    $g2.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g2.DrawImage($bmp, 0, 0, $nw, $nh)
    $g2.Dispose(); $bmp.Dispose(); $bmp = $small
}
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "shot: $Out ($((Get-Item $Out).Length) bytes)"
