param(
  [string]$Repo = "Chengyi7888/Acacia-flow",
  [string]$Tag = "v0.1.0",
  [string]$AssetPath = "Acacia Flow Setup 0.1.0.exe"
)

$ErrorActionPreference = "Stop"

$resolvedAsset = (Resolve-Path -LiteralPath $AssetPath).Path
$assetName = [IO.Path]::GetFileName($resolvedAsset)
$credentialInput = "protocol=https`nhost=github.com`n`n"
$credentialText = $credentialInput | git credential fill
$tokenLine = ($credentialText -split "`n") | Where-Object { $_ -like "password=*" } | Select-Object -First 1
if (-not $tokenLine) {
  throw "GitHub token not available from Git Credential Manager."
}

$token = $tokenLine.Substring(9).Trim()
$headers = @{
  Authorization = "Bearer $token"
  Accept = "application/vnd.github+json"
  "X-GitHub-Api-Version" = "2022-11-28"
  "User-Agent" = "Acacia-Flow-Release-Uploader"
}

try {
  $release = Invoke-RestMethod -Method Get -Uri "https://api.github.com/repos/$Repo/releases/tags/$Tag" -Headers $headers
} catch {
  $body = @{
    tag_name = $Tag
    name = "Acacia Flow 0.1.0"
    body = "Initial Windows installer release for Acacia Flow.`n`nDownload the setup executable below and run it on Windows."
    draft = $false
    prerelease = $false
  } | ConvertTo-Json
  $release = Invoke-RestMethod -Method Post -Uri "https://api.github.com/repos/$Repo/releases" -Headers $headers -Body $body -ContentType "application/json"
}

foreach ($asset in @($release.assets)) {
  if ($asset.name -eq $assetName) {
    Invoke-RestMethod -Method Delete -Uri $asset.url -Headers $headers | Out-Null
  }
}

$uploadUrl = $release.upload_url -replace "\{\?name,label\}", "?name=$([uri]::EscapeDataString($assetName))"
$uploaded = Invoke-RestMethod -Method Post -Uri $uploadUrl -Headers $headers -InFile $resolvedAsset -ContentType "application/octet-stream"

[PSCustomObject]@{
  tag = $Tag
  releaseUrl = $release.html_url
  asset = $uploaded.name
  size = $uploaded.size
  downloadUrl = $uploaded.browser_download_url
} | ConvertTo-Json
