$out = "C:\Users\Andalus\Downloads\coding\extracted"
New-Item -ItemType Directory -Force -Path $out | Out-Null
$files = @("Lecture 1 & 2.pdf","Lecture 3 _Computers.pdf","Lecture 4 _Computers.pdf","Lecture 5 _Computers.pdf","Lecture 6 _Computers.pdf")
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
foreach($f in $files){
  $src = Join-Path "C:\Users\Andalus\Downloads" $f
  $dst = Join-Path $out ($f -replace '\.pdf$','.txt')
  try {
    $doc = $word.Documents.Open($src, $false, $false)
    Start-Sleep -Milliseconds 1000
    $doc.SaveAs2($dst, 7)
    $doc.Close(0)
    Write-Output "OK: $f"
  } catch {
    Write-Output ("FAIL: $f -> " + $_.Exception.Message)
  }
}
$word.Quit()
Start-Sleep -Milliseconds 500
Get-Process winword -ErrorAction SilentlyContinue | Stop-Process -Force
