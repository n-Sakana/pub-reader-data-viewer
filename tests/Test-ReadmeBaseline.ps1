param([string]$Config, [string]$DataDir, [string]$Report, [string]$Output, [switch]$MarkReviewed)
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'build/test_support.ps1')
Import-RdvProduct -Root $root
$cfg=[Rdv3Config]::Load($Config)
[Rdv3Headless]::Evaluate($cfg,$root,$DataDir,$false,'') | Out-Null
$result=[IO.File]::ReadAllText($Report) | ConvertFrom-Json
[string[]]$lines=@($result.rows | ForEach-Object {$_ -join "`t"})
[string[]]$states=$result.states
if($MarkReviewed) {
    for($i=0;$i -lt $states.Length;$i++) {
        if($result.rows[$i][0] -in @('R01','R02')) { $states[$i]='TRUE' }
    }
}
if(Test-Path -LiteralPath $Output) { throw 'The test must not overwrite an existing ledger' }
[Rdv3Xlsx]::Write($Output,$cfg.Data.Head,$cfg.Screen.Work.Column,$lines,$states,'readme-example',[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work),[Rdv3LedgerProtection]::Create($cfg.Data))
Write-Output ('WROTE '+$Output)
