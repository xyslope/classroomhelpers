<#
.SYNOPSIS
    カレントディレクトリのPPTXファイルを一括変換
#>

# スクリプトの絶対パスを取得
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

# convertppt.ps1のパス
$converterScript = Join-Path $scriptRoot "convertppt.ps1"

# カレントディレクトリのPPTXファイルを検索
$pptFiles = @(Get-ChildItem -Path . -Filter *.pptx -File)

if ($pptFiles.Count -eq 0) {
    Write-Host "現在のフォルダに.pptxファイルが見つかりませんでした" -ForegroundColor Yellow
    Write-Host "カレントディレクトリ: $(Get-Location)" -ForegroundColor Cyan
    exit 1
}

# 処理開始
Write-Host "`n==== PPTXファイル一括変換 ====" -ForegroundColor Green
Write-Host "検出されたファイル:"
$pptFiles | ForEach-Object { Write-Host " - $($_.Name)" }

foreach ($file in $pptFiles) {
    Write-Host "`n処理中: $($file.Name)" -ForegroundColor Cyan
    try {
        & $converterScript -PPT_FILE $file.FullName
        Write-Host "[成功] $($file.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "[失敗] $($file.Name)" -ForegroundColor Red
        Write-Host "エラー詳細: $_" -ForegroundColor Red
    }
}

Write-Host "`n==== 処理完了 ====" -ForegroundColor Green