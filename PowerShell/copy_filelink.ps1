
function Set-ClipboardText($text, $html){
    Add-Type -AssemblyName System.Windows.Forms
    $data = New-Object System.Windows.Forms.DataObject
    $data.SetData([System.Windows.Forms.DataFormats]::Text, $text)
    $data.SetData([System.Windows.Forms.DataFormats]::Html, $html)
    [System.Windows.Forms.Clipboard]::SetDataObject($data, $true)
}

function Get-HtmlStr_forClipboard($fragment){
    # Clipboard�p�}�[�J�[���
    $marker_block = "Version:0.9`r`n" `
                  + "StartHTML:0000000000`r`n" `
                  + "EndHTML:0000000000`r`n" `
                  + "StartFragment:0000000000`r`n" `
                  + "EndFragment:0000000000`r`n"
    # HTML�p�[�c
    $html_top = "<html>`r`n" `
              + "<body>`r`n" `
              + "<!--StartFragment-->`r`n"
    $html_end = "<!--EndFragment-->`r`n" `
              + "</body>`r`n" `
              + "</html>"
    

    $html_str = $marker_block + $html_top + $fragment + $html_end
    
    # �I�t�Z�b�g��񏑂�����
    $html_str = $html_str.Replace("StartHTML:0000000000", "StartHTML:{0:0000000000}" -f $marker_block.Length)
    $html_str = $html_str.Replace("EndHTML:0000000000", "EndHTML:{0:0000000000}" -f $html_str.IndexOf("</html>"))
    $html_str = $html_str.Replace("StartFragment:0000000000", "StartFragment:{0:0000000000}" -f ($marker_block.Length + $html_top.Length))
    $html_str = $html_str.Replace("EndFragment:0000000000", "EndFragment:{0:0000000000}" -f ($marker_block.Length + $html_top.Length + $fragment.Length))

    return $html_str
}

function Get-ClipboardLinkText_Html($links, $link_header){
    Add-Type -AssemblyName System.Web
    # Fragment
    $fragment = ""
    $flg_http_link = ($link_header.Substring(0,4) -eq "http")
    foreach($link in $links){
        if ($flg_http_link){
            $hyperlink = $link_header + [System.Web.HttpUtility]::UrlEncode($link)
        }else{
            $hyperlink = $link_header + $link
        }
        $fragment += "<a href=""$hyperlink"">$link</a><br>`r`n"
    }
    return Get-HtmlStr_forClipboard($fragment)
}
function Get-ClipboardLinkText_Text($links){
    $file_head = "file://"
    $txt_str = ""
    foreach($link in $links){
        $txt_str += "<"+$file_head+$link+">`r`n"
    }
    return $txt_str
}

#�e�X�g�R�[�h
function main($links){

    # $links = @("c:\\PowerShell", "c:\\PowerShell2")
    
    $html_link = Get-ClipboardLinkText_Html $links "http://localhost:8000/lancher?path="
    $text_link = Get-ClipboardLinkText_Text $links
    
    # Write-Output $text_link
    Write-Output $html_link
    
    # �N���b�v�{�[�h�ɓ\��t��
    Set-ClipboardText $text_link $html_link
    
    # pause�̑���
    # Read-Host "please input.."
}

# D&D�����t�@�C���p�X��$args�Ɋi�[����Ă���
main $args