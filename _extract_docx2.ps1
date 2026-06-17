Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead('/tmp/target.docx')
$stream = $zip.GetEntry('word/document.xml').Open()
$reader = New-Object System.IO.StreamReader($stream)
$content = $reader.ReadToEnd()
$reader.Close()
$zip.Dispose()

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
