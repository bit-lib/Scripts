# クリップボードの内容を表示します

Function print_clipboard
{
    Add-Type -AssemblyName System.Windows.Forms
    $lst = @("Text", "UnicodeText", "Rtf", "Html", "CommaSeparatedValue")

    $cb = [Windows.Forms.Clipboard]
    
    for($i=0; $i -lt 5; $i++){
        Write-Output "-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-"
        Write-Output ("    ClipBoard: " + $lst[$i])
        $res = $cb::GetText($i)
        Write-Output "----------------------------------------"
        Write-Output $res
        Write-Output "-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-`r`n"
    }
}

# 表示
print_clipboard
