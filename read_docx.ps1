Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead("C:\Users\LENOVO\Documents\PharmaFinder_Analyse_Concurrentielle.docx")
$entry = $zip.GetEntry("word/document.xml")
$stream = $entry.Open()
$reader = New-Object System.IO.StreamReader($stream)
$xml = $reader.ReadToEnd()
$reader.Close()
$stream.Close()
$zip.Dispose()

[xml]$doc = $xml
$ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

$paragraphs = $doc.SelectNodes("//w:p", $ns)
foreach ($p in $paragraphs) {
    if ($p.InnerText.Trim()) {
        Write-Output "P: $($p.InnerText.Trim())"
    }
}

Write-Output "=== TABLES ==="
$rows = $doc.SelectNodes("//w:tr", $ns)
foreach ($r in $rows) {
    $cells = $r.SelectNodes(".//w:tc", $ns)
    $rowText = ($cells | ForEach-Object { $_.InnerText.Trim() }) -join " | "
    Write-Output $rowText
}
