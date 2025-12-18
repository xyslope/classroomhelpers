$port = 1313

# 1313ポートを使用しているプロセスを取得
$process = Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue | 
           Select-Object -ExpandProperty OwningProcess -First 1

if ($process) {
    Write-Host "ポート $port はプロセスID $process によって使用されています。プロセスを終了します。"
    Stop-Process -Id $process -Force
    Start-Sleep -Seconds 1 # プロセス終了を待つ
}

# Hugoサーバーを起動
hugo server -D --bind 0.0.0.0 --port $port --disableFastRender --ignoreCache
