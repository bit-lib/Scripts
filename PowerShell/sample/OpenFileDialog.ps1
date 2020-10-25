
Function OpenFileDialog($filter)
{
    # アセンブリロード
    Add-Type -AssemblyName System.Windows.Forms

    # ダイアログインスタンス製紙絵
    $dlog = New-Object System.Windows.Forms.OpenFileDialog

    # 設定
    $dlog.Title = "開く"
    $dlog.Filter = $filter
    $dlog.FilterIndex = 1
    $dlog.InitialDirectory = ""
    $dlog.RestoreDirectory = $true
    $dlog.Multiselect = $false

    # ダイアログを開く
    $result = $dlog.ShowDialog()

    If($result -eq "OK"){
        Return $dlog.FileName 
    }Else{
        Break
    }
}

# テスト
Write-Output (OpenFileDialog("テキスト (*.txt,*.csv)|*.txt;*.csv|すべてのファイル(*.*)|*.*"))