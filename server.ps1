$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$port = 18888

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()

Start-Process "http://localhost:$port/"

Write-Host ""
Write-Host "  DCS Converter is running"
Write-Host "  Open: http://localhost:$port/"
Write-Host "  Press Ctrl+C to stop"
Write-Host ""

while ($listener.IsListening) {
    try {
        $ctx = $listener.GetContext()
        $path = $ctx.Request.Url.LocalPath
        if ($path -eq '/') { $path = '/dcs_converter.html' }
        $file = Join-Path $baseDir $path.TrimStart('/')
        if (Test-Path $file) {
            $bytes = [System.IO.File]::ReadAllBytes($file)
            $ext = [System.IO.Path]::GetExtension($file).ToLower()
            $ct = 'text/plain'
            if ($ext -eq '.html') { $ct = 'text/html; charset=utf-8' }
            elseif ($ext -eq '.js') { $ct = 'application/javascript; charset=utf-8' }
            $ctx.Response.ContentType = $ct
            $ctx.Response.ContentLength64 = $bytes.Length
            $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $ctx.Response.StatusCode = 404
        }
        $ctx.Response.Close()
    } catch { break }
}

$listener.Stop()
$listener.Close()
