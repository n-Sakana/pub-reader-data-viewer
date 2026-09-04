# Capture every standard-control dialog without putting a window on screen.
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $OutDir = "",
  [switch] $DumpDialogGeometry
)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
if ([string]::IsNullOrEmpty($OutDir)) { $OutDir = Join-Path $Root 'work\ui-v3\dialogs' }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

# This compiles the same source list and leaves its types plus PrepareControl
# in this process. Its main-screen capture is useful alongside the dialogs.
. (Join-Path $Root 'build\capture_headless.ps1') -Root $Root -Loaded -OutPath (Join-Path $OutDir 'main.png')

function Save-Offscreen([System.Windows.Forms.Form] $dialog, [string] $path) {
  if ($DumpDialogGeometry) {
    Write-Output ("before {0} client={1}x{2} autoScale={3} currentScale={4}" -f $dialog.Text, $dialog.ClientSize.Width, $dialog.ClientSize.Height, $dialog.AutoScaleDimensions, $dialog.CurrentAutoScaleDimensions)
    foreach ($rootControl in $dialog.Controls) { Write-Output ("before-control {0} bounds={1}" -f $rootControl.Name, $rootControl.Bounds) }
  }
  $dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
  $dialog.Location = New-Object Drawing.Point(-32000, -32000)
  $dialog.ShowInTaskbar = $false
  PrepareControl $dialog
  $dialog.Show()
  [System.Windows.Forms.Application]::DoEvents()
  Write-Output ("dialog {0} client={1}x{2} dpi={3}" -f $dialog.Text, $dialog.ClientSize.Width, $dialog.ClientSize.Height, [RdvHeadlessDpi]::GetDpiForWindow($dialog.Handle))
  $geometryText = $dialog.GeometryDump()
  $geometry = $geometryText | ConvertFrom-Json
  if ($DumpDialogGeometry) { Write-Output $geometryText }
  if (@($geometry.clipped).Count -gt 0) {
    $overflow = @($geometry.clipped | ForEach-Object {
      "{0} in {1} (left={2}, top={3}, right={4}, bottom={5})" -f $_.name, $_.parent, $_.left, $_.top, $_.right, $_.bottom
    })
    throw ("controls exceed their parent client bounds: " + ($overflow -join '; '))
  }
  $todo = New-Object 'System.Collections.Generic.Stack[System.Windows.Forms.Control]'
  $todo.Push($dialog)
  while ($todo.Count -gt 0) {
    $at = $todo.Pop()
    foreach ($child in $at.Controls) {
      $todo.Push($child)
      if ($child -is [System.Windows.Forms.Button] -and $child.Name -notmatch '^export\.(right|left)$') {
        $preferred = $child.GetPreferredSize((New-Object Drawing.Size(10000, 10000)))
        Write-Output ("button {0} actual={1}x{2} preferred={3}x{4} autosize={5}" -f $child.Name, $child.Width, $child.Height, $preferred.Width, $preferred.Height, $child.AutoSize)
        if (-not $child.AutoSize -or $child.Width -lt $preferred.Width -or $child.Height -lt $preferred.Height) {
          throw ("dialog button does not fit its preferred size: " + $child.Name)
        }
      }
    }
  }
  $bitmap = New-Object Drawing.Bitmap($dialog.Width, $dialog.Height)
  try {
    $dialog.DrawToBitmap($bitmap, (New-Object Drawing.Rectangle(0, 0, $bitmap.Width, $bitmap.Height)))
    $bitmap.Save($path, [Drawing.Imaging.ImageFormat]::Png)
  } finally {
    $bitmap.Dispose()
    $dialog.Hide()
    $dialog.Dispose()
  }
  Write-Output $path
}

$ledgerPath = Join-Path $Root 'dist\app-csharp\ReaderDataViewer-Ledger.xlsx'
$baseDir = Split-Path -Parent $ledgerPath

$dialogs = New-Object System.Collections.ArrayList
[void]$dialogs.Add(@('record-update.png', [Rdv3ProcessForm]::ForCheck($cfg.Data, 'merge-ledger', $dataDir, $ledgerPath)))
[void]$dialogs.Add(@('record-delete.png', [Rdv3ProcessForm]::ForCheck($cfg.Data, 'delete-listed-records', $dataDir, $ledgerPath)))
[void]$dialogs.Add(@('table-export.png', [Rdv3ExportForm]::ForCheck($cfg.Data, $cfg.Screen, $baseDir)))
[void]$dialogs.Add(@('settings.png', [Rdv3SettingsForm]::ForCheck($cfg.Clone())))
$sendBody = [Rdv3Text]::ConfirmSendBody.Replace('{n}', '3')
[void]$dialogs.Add(@('send-confirm.png', [Rdv3ConfirmForm]::ForCheck([Rdv3Text]::SendTitle, $sendBody)))

$merged = [Rdv3Ledger]::BuildFromCsv($cfg.Data, $dataDir)
$candidateRows = New-Object 'System.Collections.Generic.List[Rdv3CandRow]'
for ($i = 0; $i -lt 5; $i++) {
  $row = New-Object Rdv3CandRow
  $row.Line = $merged.Lines[$i]
  $row.Stored = $cfg.Screen.Work.InitialStored
  $candidateRows.Add($row)
}
$candidateFields = New-Object Rdv3Fields (,$cfg.Data.ColumnRefs)
[void]$dialogs.Add(@('candidates.png', [Rdv3CandidatesForm]::ForCheck($cfg.Screen.Candidates, $candidateFields, $cfg.Screen.Work, $candidateRows, 5, 0)))
[void]$dialogs.Add(@('merge-reset.png', [Rdv3LedgerUpdateForm]::ForCheck($form, $candidateRows)))

foreach ($entry in $dialogs) {
  Save-Offscreen $entry[1] (Join-Path $OutDir $entry[0])
}
