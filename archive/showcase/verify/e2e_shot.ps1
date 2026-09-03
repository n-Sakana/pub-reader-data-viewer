# e2e_shot.ps1 -- VERIFICATION ONLY. Captures the real screen (or one window's
# rectangle) at true device pixels, so what is checked is what is on the glass.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Out,
    [int] $X = -1, [int] $Y = -1, [int] $W = -1, [int] $H = -1,
    [int] $Scale = 100
)
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing
Add-Type -Namespace Dpi -Name W -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
'@
[void][Dpi.W]::SetProcessDPIAware()

Add-Type -AssemblyName System.Windows.Forms
$all = [System.Windows.Forms.SystemInformation]::VirtualScreen
if ($X -lt 0) { $X = $all.X }
if ($Y -lt 0) { $Y = $all.Y }
if ($W -lt 0) { $W = $all.Width }
if ($H -lt 0) { $H = $all.Height }

$bmp = New-Object System.Drawing.Bitmap($W, $H)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($X, $Y, 0, 0, (New-Object System.Drawing.Size($W, $H)))
$g.Dispose()

if ($Scale -ne 100) {
    $nw = [int]($W * $Scale / 100)
    $nh = [int]($H * $Scale / 100)
    $small = New-Object System.Drawing.Bitmap($nw, $nh)
    $g2 = [System.Drawing.Graphics]::FromImage($small)
    $g2.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g2.DrawImage($bmp, 0, 0, $nw, $nh)
    $g2.Dispose()
    $bmp.Dispose()
    $bmp = $small
}
$bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()
Write-Host "shot: $Out ($((Get-Item $Out).Length) bytes)"
