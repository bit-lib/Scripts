
Function OpenFileDialog($filter)
{
    # �A�Z���u�����[�h
    Add-Type -AssemblyName System.Windows.Forms

    # �_�C�A���O�C���X�^���X�����G
    $dlog = New-Object System.Windows.Forms.OpenFileDialog

    # �ݒ�
    $dlog.Title = "�J��"
    $dlog.Filter = $filter
    $dlog.FilterIndex = 1
    $dlog.InitialDirectory = ""
    $dlog.RestoreDirectory = $true
    $dlog.Multiselect = $false

    # �_�C�A���O���J��
    $result = $dlog.ShowDialog()

    If($result -eq "OK"){
        Return $dlog.FileName 
    }Else{
        Break
    }
}

# �e�X�g
Write-Output (OpenFileDialog("�e�L�X�g (*.txt,*.csv)|*.txt;*.csv|���ׂẴt�@�C��(*.*)|*.*"))