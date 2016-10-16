# PSR実行ファイル
[string] $psrExe = 'C:\Windows\System32\psr.exe'

# スクショ上限
[int] $maxsc = 10

# 出力場所＆ファイル名
#[string] $output = 'C:\20161016_tmp\evi_' + (Get-Date).ToString("yyyy-MM-dd_HHmmss") + '.zip'

# PSRの起動
$proc = Start-Process -FilePath $psrExe -ArgumentList ('/maxsc ' + $maxsc.ToString()) -PassThru

# 「スクショ一時保存フォルダ」が格納されるフォルダ
[string] $tmpParentDir = ('C:\Users\' + ([Environment]::UserName) + '\AppData\Local\Microsoft\UIR')

# 記録開始の監視
while (1)
{
    if ($proc.HasExited) {
        break
    }

    # スクショ一時保存フォルダの取得（3秒以内に作成されたフォルダ）
    $tmpDirs = @(Get-ChildItem  $tmpParentDir | Where-Object {$_.LastWriteTime -gt (Get-Date).AddSeconds(-3)})
    if ($tmpDirs.length -eq 1) {
        break
    }
    if ($tmpDirs.length -gt 1) {
        Write-Host '実行に失敗しました。（一時フォルダの重複）'
        Write-Host '実行しなおしてください。'
        Start-Process -FilePath $psrExe -ArgumentList '/stop'
        pause
        exit
    }

    Start-Sleep -s 1
}

# スクショ数の監視
[int] $prevLength = 0
while (1)
{
    if ($proc.HasExited) {
        break
    }

    $jpegs = @(Get-ChildItem ($tmpDirs[0].FullName.ToString() + '\*.JPEG'))

    if ($jpegs.Length -ge $maxsc)
    {
        [string] $msgTitle  = 'PsrLauncher'
        [string] $msgPrompt = '記録できる画像数が上限に達します。下の[OK]を押す前に「ステップ記録ツール」ウィンドウの[記録の停止]からPSRを終了してください。その後、[OK]を押してから、PsrLauncherを再実行してください。'
        [Void][Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
        [Microsoft.VisualBasic.Interaction]::MsgBox($msgPrompt, ([Microsoft.VisualBasic.MsgBoxStyle]::Yes -bor [Microsoft.VisualBasic.MsgBoxStyle]::SystemModal),$msgTitle)
    }

    if (($jpegs.Length -ne $prevLength) -and ($jpegs.Length -ne 0)) {
        Write-Host ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + ': 記録した画像数: ' + $jpegs.Length + '／' + $maxsc.ToString() + '')
    }

    $prevLength = $jpegs.Length

    #Start-Sleep -s 1
}

exit 0
