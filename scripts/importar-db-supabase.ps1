$ErrorActionPreference = "Stop"

$backendRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent $backendRoot
$dbFile = Join-Path $workspaceRoot "sistema-asistencia-web\db.json"
$nodeScript = Join-Path $PSScriptRoot "migrar-db-local-a-supabase.js"

if (-not (Test-Path -LiteralPath $dbFile)) {
  throw "No encontré el archivo local de datos: $dbFile"
}
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js no está disponible en esta terminal."
}

$previousUrl = $env:SUPABASE_URL
$previousSecretKey = $env:SUPABASE_SECRET_KEY
$previousServiceRoleKey = $env:SUPABASE_SERVICE_ROLE_KEY
$secureKey = $null
$pointer = [IntPtr]::Zero

try {
  $env:SUPABASE_URL = "https://hswszojtwbuejxmbklfn.supabase.co"
  $secureKey = Read-Host "Pega la clave secret de Supabase (no se mostrará)" -AsSecureString
  $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
  $env:SUPABASE_SECRET_KEY = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)

  & node $nodeScript $dbFile "--apply"
  if ($LASTEXITCODE -ne 0) {
    throw "El importador terminó con error $LASTEXITCODE. Revisa el mensaje anterior; no compartas la clave."
  }
}
finally {
  if ($pointer -ne [IntPtr]::Zero) {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
  }
  if ($secureKey) { $secureKey.Dispose() }

  if ($null -eq $previousUrl) { Remove-Item Env:SUPABASE_URL -ErrorAction SilentlyContinue }
  else { $env:SUPABASE_URL = $previousUrl }

  if ($null -eq $previousSecretKey) { Remove-Item Env:SUPABASE_SECRET_KEY -ErrorAction SilentlyContinue }
  else { $env:SUPABASE_SECRET_KEY = $previousSecretKey }

  if ($null -eq $previousServiceRoleKey) { Remove-Item Env:SUPABASE_SERVICE_ROLE_KEY -ErrorAction SilentlyContinue }
  else { $env:SUPABASE_SERVICE_ROLE_KEY = $previousServiceRoleKey }
}
