param([string]$AppRoot='', [string]$Evidence='')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.Encoding]::UTF8
$sourceRoot=Split-Path -Parent $PSScriptRoot
if(-not $AppRoot){$AppRoot=$sourceRoot}
if(-not $Evidence){$Evidence=Join-Path $sourceRoot ('work/excel-roundtrip-'+[Guid]::NewGuid().ToString('N'))}
$Evidence=[IO.Path]::GetFullPath($Evidence)
if(Test-Path -LiteralPath $Evidence){throw 'Use a new evidence directory'}
[IO.Directory]::CreateDirectory($Evidence)|Out-Null
. (Join-Path $sourceRoot 'build/test_support.ps1')
Import-RdvProduct -Root $AppRoot
$inputRoot=if(Test-Path -LiteralPath (Join-Path $AppRoot 'configs/sample/settings.json')){Join-Path $AppRoot 'samples/current'}else{$AppRoot}
$configPath=if($inputRoot -eq $AppRoot){Join-Path $AppRoot 'settings.json'}else{Join-Path $AppRoot 'configs/sample/settings.json'}
$cfg=[Rdv3Config]::Load($configPath)
$inputDir=Join-Path $inputRoot 'data'
[Rdv3Ledger]::BuildFromCsv($cfg.Data,$inputDir)|Out-Null
$start=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$inputDir,[string[]]@(),[string[]]@(),$cfg.Screen.Work.InitialStored)
for($i=0;$i -lt $start.Lines.Length;$i++){if($start.Lines[$i].Contains('AB10000001CD')){$start.States[$i]='TRUE'}}
$deleted=[Rdv3Process]::Run($cfg.Data,$cfg.Data.JobOf('delete-processed-records'),$inputDir,$start.Lines,$start.States,$cfg.Screen.Work.InitialStored)
$protection=[Rdv3LedgerProtection]::Create($cfg.Data)
$protection.ArchiveRemoved($cfg.Data,$start.Lines,$start.States,$deleted.Lines)
$definition=$protection.Definition
$deletedLine=$protection.Deleted[0].Line
$deletedAt=$protection.Deleted[0].DeletedAt
$file=Join-Path $Evidence 'ledger.xlsx'
[Rdv3Xlsx]::Write($file,$cfg.Data.Head,$cfg.Screen.Work.Column,$deleted.Lines,$deleted.States,'roundtrip',[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work),$protection)
$archive=@(Get-ChildItem -LiteralPath (Join-Path $Evidence 'archived') -Filter '*.xlsx')[0].FullName
$archiveHash=(Get-FileHash -LiteralPath $archive).Hash
$excel=$null;$book=$null
$results=New-Object Collections.Generic.List[object]
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$true;$excel.DisplayAlerts=$false
    Write-Output ('Excel '+$excel.Version+'; owned HWND '+$excel.Hwnd)
    for($round=1;$round -le 2;$round++){
        $book=$excel.Workbooks.Open($file)
        if($book.ReadOnly){throw 'Excel opened read-only'}
        $saved=Join-Path $Evidence ('excel-save-'+$round+'.xlsx')
        $book.SaveAs($saved,51)
        if(-not $book.Saved){throw 'Excel did not finish saving'}
        $book.Close($false);$book=$null
        [string[]]$lines=@();[string[]]$states=@();$warning='';$p=$null
        [Rdv3Xlsx]::ReadProtected($saved,$cfg.Data.Head,$cfg.Screen.Work.Column,[ref]$lines,[ref]$states,[ref]$warning,[Rdv3Files]::StorageContract($cfg.Data,$cfg.Screen.Work),[Rdv3Files]::LegacyStorageContract($cfg.Data,$cfg.Screen.Work),$definition,[ref]$p,$true)
        $p.RequireWritable($cfg.Data);$p.Validate($cfg.Data,$cfg.Screen.Work,$lines)
        if($warning -or $p.Legacy -or $p.Definition -ne $definition -or $p.Deleted.Count -ne 1 -or $p.Deleted[0].Line -cne $deletedLine -or $p.Deleted[0].DeletedAt -cne $deletedAt){throw 'Protection metadata did not survive Excel save'}
        if(($lines -join "`0") -cne ($deleted.Lines -join "`0") -or ($states -join "`0") -cne ($deleted.States -join "`0")){throw 'Excel changed row content or confirmation states'}
        $reimport=[Rdv3Process]::Run($cfg.Data,$cfg.Data.UpdateJob,$inputDir,$lines,$states,$cfg.Screen.Work.InitialStored)
        $p.ProtectUpdate($cfg.Data,$cfg.Screen.Work.InitialStored,$lines,$states,$start.Lines,$reimport.Update)
        if($reimport.Update.Lines.Length -ne 2 -or $reimport.Update.SkippedDeleted -ne 1){throw 'Excel save allowed reimport of the deleted record'}
        if((Get-FileHash -LiteralPath $archive).Hash -ne $archiveHash){throw 'Ledger save rewrote the separate archive'}
        $results.Add(@{round=$round;rows=$lines.Length;deleted=$p.Deleted.Count;reimportBlocked=$true;writable=$true;allContentEqual=$true;archiveBytesEqual=$true})
        Write-Output ('PASS Excel save '+$round+': all content/states, definition, archive reference, reimport exclusion and writable state')
        $file=$saved
    }
} finally {
    if($null -ne $book){$book.Close($false)}
    if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [IO.File]::WriteAllText((Join-Path $Evidence 'results.json'),($results.ToArray()|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
}
