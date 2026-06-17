Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead('C:\Users\lenovo\Desktop\常用运算符和数学函数_含示例解释.docx')
$stream = $zip.GetEntry('word/document.xml').Open()
$reader = New-Object System.IO.StreamReader($stream)
$content = $reader.ReadToEnd()
$reader.Close()
$zip.Dispose()

# Find all paragraphs with their text
$paragraphs = [regex]::Matches($content, '<w:p[ >].*?</w:p>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
foreach ($p in $paragraphs) {
    $texts = [regex]::Matches($p.Value, '<w:t[^>]*>([^<]*)</w:t>')
    $line = ""
    foreach ($t in $texts) {
        $line += $t.Groups[1].Value
    }
    if ($line.Trim() -ne "") {
        Write-Output $line
    }
}
