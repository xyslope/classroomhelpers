#=======================================================================
# PowerPoint to Markdown Converter
# ファイルをMarkdownに変換し、アルファベット数字部分+ユニーク文字でフォルダ名を作成
# Author: Claude Code
# Date: $(Get-Date -Format "yyyy-MM-dd")
#=======================================================================

param (
    [string]$PPT_FILE
)

if (-not $PPT_FILE) {
    Write-Host "パワーポイントファイルを指定してください"
    exit 1
}

# 絶対パスに変換
$pptFullPath = (Get-Item $PPT_FILE).FullName

if (-not (Test-Path -Path $pptFullPath -PathType Leaf)) {
    Write-Host "指定されたパワーポイントファイルが存在しません: $pptFullPath"
    exit 1
}

# 拡張子を除いたファイル名を取得
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($pptFullPath)

# ファイル名からアルファベットと数字部分のみを抽出
$alphanumericName = $baseName -replace '[^a-zA-Z0-9]', ''
if (-not $alphanumericName) {
    $alphanumericName = "converted"
}

# ユニークな3文字を生成
$uniqueChars = -join ((1..3) | ForEach-Object { [char]((Get-Random -Minimum 97 -Maximum 123)) })
$finalFolderName = $alphanumericName + $uniqueChars

# 作成するフォルダパス（PPTファイルと同じディレクトリに作成）
$parentDir = Split-Path -Parent $pptFullPath
$outputFolder = Join-Path -Path $parentDir -ChildPath $finalFolderName

# フォルダが既に存在する場合はスキップ
if (Test-Path -Path $outputFolder -PathType Container) {
    Write-Host "フォルダは既に存在するためスキップします: $outputFolder"
    exit 0
}

# 新しいフォルダを作成
New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
Write-Host "新しいフォルダを作成しました: $outputFolder"

# imagesフォルダを作成
$imagesFolderName = "images"
$imagesFolder = Join-Path -Path $outputFolder -ChildPath $imagesFolderName
New-Item -Path $imagesFolder -ItemType Directory -Force | Out-Null
Write-Host "imagesフォルダを作成しました: $imagesFolder"

# pptx2mdを実行（絶対パスを使用）
try {
    pptx2md.exe $pptFullPath -i "$imagesFolder" -o "$outputFolder\_index.md"
    Write-Host "pptx2mdの変換が完了しました"
    
    # 生成されたMarkdownファイルにヘッダと見出しレベル変更を適用
    $markdownFile = "$outputFolder\_index.md"
    if (Test-Path $markdownFile) {
        $content = Get-Content $markdownFile -Raw
        
        # 現在の日時を取得
        $currentDate = Get-Date -Format "yyyy-MM-ddTHH:mm:sszzz"
        
        # タイトル用のクリーンな文字列を作成（特殊文字を除去）
        $cleanTitle = $baseName -replace '[^\w\s]', '' -replace '\s+', ' '
        $cleanTitle = $cleanTitle.Trim()
        if (-not $cleanTitle) {
            $cleanTitle = "Presentation"
        }
        
        # ヘッダを作成
        $header = @"
+++
title = '$cleanTitle'
date = $currentDate
outputs = ['Reveal']
draft = true
+++

"@
        
        # # を ---改行# に変更  
        $content = $content -replace '(?m)^(#+\s)', "---`r`n`$1"
        
        # ヘッダを先頭に追加
        $finalContent = $header + $content
        
        Set-Content $markdownFile -Value $finalContent -Encoding UTF8
        Write-Host "ヘッダを追加し、見出しレベルを変更しました"
    }
}
catch {
    Write-Host "pptx2mdの実行中にエラーが発生しました:" -ForegroundColor Red
    Write-Host $_ -ForegroundColor Red
    exit 1
}