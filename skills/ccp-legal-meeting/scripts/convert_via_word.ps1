# ASCII-only. Word COM batch-convert docx files to PDF.
# Usage: 1) Python writes a UTF-8-BOM list file (one absolute path per line)
#        2) Edit $LIST and $OUT below (ASCII paths only)
#        3) Run from an elevated/outside-sandbox PowerShell session
# Word COM requires escaping the sandbox on Windows (sandbox blocks COM instantiation).
$ErrorActionPreference = "Continue"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$LIST = "C:\temp\ccp_list.txt"      # UTF-8 BOM, one .docx path per line
$OUT  = "C:\temp\ccp_out"           # output folder for PDFs
$MODE = "pdf"                       # "pdf" = SaveAs 17 ; "docx" = SaveAs 16 (for .doc sources, saves next to source)

if ($MODE -eq "pdf") { New-Item -ItemType Directory -Force -Path $OUT | Out-Null }
$lines = [System.IO.File]::ReadAllLines($LIST, [System.Text.Encoding]::UTF8)
$log = @()
$i = 0
foreach($src in $lines){
  if([string]::IsNullOrWhiteSpace($src)){ continue }
  $i++
  $src = [string]$src
  try{
    $doc = $word.Documents.Open($src, $false, $true)
    if ($MODE -eq "pdf") {
      $dst = [string](Join-Path $OUT ("chk{0:D2}.pdf" -f $i))
      $doc.SaveAs([ref]$dst, [ref]17)
    } else {
      $dst = [string]([System.IO.Path]::ChangeExtension($src, ".docx"))
      $doc.SaveAs([ref]$dst, [ref]16)
    }
    $doc.Close($false)
    $log += "OK $i"
  } catch { $log += "FAIL $i : $($_.Exception.Message)" }
}
$word.Quit()
$log | Out-File (Join-Path $OUT "convert_log.txt") -Encoding unicode
# NOTE: convert_log.txt is UTF-16; read it back with Python: open(p, encoding='utf-16')
