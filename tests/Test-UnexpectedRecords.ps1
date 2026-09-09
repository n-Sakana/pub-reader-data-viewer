param([string]$Manifest, [string]$Output)
$ErrorActionPreference = 'Stop'
$unexpectedRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $unexpectedRoot 'build/test_support.ps1')
Import-RdvProduct -Root $unexpectedRoot
$cases = [IO.File]::ReadAllText($Manifest) | ConvertFrom-Json
$observed = @()
foreach ($case in $cases) {
    $validate = [Rdv3Headless]::Run($case.directory, $case.config, $case.directory, $false, '', $case.baseline)
    $execute = [Rdv3Headless]::Run($case.directory, $case.config, $case.directory, $true, $case.report, $case.baseline)
    $item = [ordered]@{ name=$case.name; validateExit=$validate; runExit=$execute }
    if ($execute -eq 0) {
        $cfg = [Rdv3Config]::Load($case.config)
        $merge = [Rdv3Ledger]::BuildFromCsv($cfg.Data, $case.directory)
        $item.windowLines = @($merge.Lines)
        $item.windowWarnings = @($merge.Warnings)
        [object[]]$previewArgs = @($cfg.Data, $cfg.Data.UpdateJob, $case.directory, $cfg.Ledger, $false)
        $method = [Rdv3ProcessForm].GetMethod('Build', [Reflection.BindingFlags]'NonPublic,Static')
        $item.preview = $method.Invoke($null, $previewArgs) | ConvertFrom-Json
        $item.previewValid = [bool]$previewArgs[4]
    }
    $observed += [pscustomobject]$item
}
if (Test-Path -LiteralPath $Output) { throw 'Test output already exists' }
[IO.File]::WriteAllText($Output, (ConvertTo-Json -InputObject @($observed) -Depth 20), (New-Object Text.UTF8Encoding($false)))
