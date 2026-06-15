# Build the decompiler (handling TLB flakiness) and optionally decompile Dungeon.
# Usage: powershell -File build_and_test.ps1 [-OutDir <dir>] [-NoBuild]
param(
  [string]$OutDir = "C:\Users\Owner\Desktop\websites\dungeondecomipler\ourdecompiler_v29",
  [switch]$NoBuild
)
$ErrorActionPreference = "Stop"
$exe = ".\Install Folder\SemiVBDecompiler.exe"
$vb6 = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
$target = "C:\Users\Owner\Desktop\forummods\rpgwo\DungeonFateSource\Dungeon.exe"

if (-not $NoBuild) {
  $before = (Get-Item $exe).LastWriteTime.Ticks
  $ok = $false
  for ($i=1; $i -le 12; $i++) {
    Get-Process VB6 -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Milliseconds 1500
    & $vb6 /make "VBDecompiler.vbp" /out "build_check2.log" | Out-Null
    Start-Sleep -Milliseconds 800
    $after = (Get-Item $exe).LastWriteTime.Ticks
    if ($after -ne $before) { $ok = $true; "BUILD OK (attempt $i, $([datetime]$after))"; break }
  }
  if (-not $ok) {
    # Distinguish a real compile error from TLB flakiness
    $tail = Get-Content build_check2.log -Tail 4
    if ($tail -match "failed") { "BUILD FAILED (compile error):"; $tail; exit 1 }
    "BUILD: exe timestamp unchanged after 12 tries (TLB flakiness or no source change)"; exit 2
  }
}

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force $OutDir | Out-Null }
& $exe $target /vbp /out $OutDir | Out-Null
"DECOMPILED -> $OutDir"
