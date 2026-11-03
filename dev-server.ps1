# שרת פיתוח מקומי עם רענון אוטומטי
# הרצה: powershell -ExecutionPolicy Bypass -File dev-server.ps1
# ואז לפתוח בדפדפן: http://localhost:5500
param([int]$Port = 5500)

$root = $PSScriptRoot
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Serving $root at http://localhost:$Port  (Ctrl+C to stop)"

$types = @{
  ".html"="text/html; charset=utf-8"; ".htm"="text/html; charset=utf-8"; ".css"="text/css"; ".js"="text/javascript"
  ".json"="application/json"; ".png"="image/png"; ".jpg"="image/jpeg"; ".jpeg"="image/jpeg"; ".gif"="image/gif"
  ".svg"="image/svg+xml"; ".webp"="image/webp"; ".avif"="image/avif"; ".ico"="image/x-icon"
}
$reload = "<script>(()=>{let t;setInterval(async()=>{try{const v=await (await fetch('/__reload',{cache:'no-store'})).text();if(t&&v!==t)location.reload();t=v}catch(e){}},700)})()</script>"

function Get-Stamp {
  $max = 0
  foreach ($f in [IO.Directory]::GetFiles($root, "*", "AllDirectories")) {
    if ($f -like "*\.git\*") { continue }
    $t = [IO.File]::GetLastWriteTimeUtc($f).Ticks
    if ($t -gt $max) { $max = $t }
  }
  "$max"
}

while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  $req = $ctx.Request; $res = $ctx.Response
  try {
    $path = [Uri]::UnescapeDataString($req.Url.AbsolutePath)
    $res.Headers.Add("Cache-Control", "no-store")
    if ($path -eq "/__reload") {
      $bytes = [Text.Encoding]::UTF8.GetBytes((Get-Stamp)); $res.ContentType = "text/plain"
    } else {
      if ($path -eq "/") { $path = "/index.html" }
      $file = [IO.Path]::GetFullPath((Join-Path $root $path.TrimStart("/")))
      if (-not $file.StartsWith($root) -or -not (Test-Path $file -PathType Leaf)) {
        $res.StatusCode = 404; $bytes = [Text.Encoding]::UTF8.GetBytes("Not found")
      } else {
        $ext = [IO.Path]::GetExtension($file).ToLower()
        $res.ContentType = $(if ($types.ContainsKey($ext)) { $types[$ext] } else { "application/octet-stream" })
        if ($ext -eq ".html") {
          $html = [IO.File]::ReadAllText($file, [Text.Encoding]::UTF8) -replace "</body>", "$reload</body>"
          $bytes = [Text.Encoding]::UTF8.GetBytes($html)
        } else { $bytes = [IO.File]::ReadAllBytes($file) }
      }
    }
    $res.ContentLength64 = $bytes.Length
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
  } catch { $res.StatusCode = 500 } finally { $res.OutputStream.Close() }
}
