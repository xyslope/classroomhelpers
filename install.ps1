# install.ps1
$binPath = "$HOME\bin"
$scriptNames = @(
    "revealhugo/starthugo.ps1", # ローカルhugoサーバを開始
    "revealhugo/deployClass.bat", # いろいろ処理してデプロイ
    "revealhugo/postEditContents.py", # サブディレクトリ
    "revealhugo/organize_images.py", # サブディレクトリ
    "revealhugo/convertppt.ps1", # pptファイルを変換
    "revealhugo/batchconvertppt.ps1" # フォルダ内のpptファイルをまとめて変換
)

# ~/bin がなければ作成
if (!(Test-Path $binPath)) {
    New-Item -ItemType Directory -Path $binPath
}

foreach ($name in $scriptNames) {
    $source = "$PSScriptRoot\$name"
    $baseName = Split-Path -Leaf $name  # ディレクトリを除去してファイル名だけ取得
    $target = "$binPath\$baseName"
    
    if (Test-Path $source) {
        # 既存のファイル/シンボリックリンクがあれば削除
        if (Test-Path $target) {
            Remove-Item $target -Force
            Write-Host "  Removed existing: $baseName" -ForegroundColor Yellow
        }
        
        New-Item -ItemType SymbolicLink -Path $target -Target $source -Force
        Write-Host "✓ Installed: $baseName" -ForegroundColor Green
    } else {
        Write-Host "✗ Not found: $source" -ForegroundColor Red
    }
}

Write-Host "`nDone! Installed $($scriptNames.Count) scripts to $binPath"
