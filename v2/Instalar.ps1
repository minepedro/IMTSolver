# IMTSolver 2.0 - instalador
#
# Copia o suplemento e o motor para a pasta de suplementos do Excel,
# tira o bloqueio do Windows ("arquivo baixado da internet") e
# registra o suplemento para carregar sozinho quando o Excel abrir.
#
# Nao precisa de administrador: tudo fica no perfil do usuario.

param([switch]$Desinstalar)

$ErrorActionPreference = "Stop"
$aqui    = Split-Path -Parent $MyInvocation.MyCommand.Path
$destino = Join-Path $env:APPDATA "Microsoft\AddIns"
# A chave do registro depende da versao do Office:
#   16.0 = 2016/2019/2021/365    15.0 = 2013    14.0 = 2010
# Isto estava fixo em 16.0. Num Office 2013 a instalacao dizia "pronto"
# e o suplemento nunca carregava, sem dizer por que.
$chaves = @()
foreach ($v in @("16.0", "15.0", "14.0")) {
    if (Test-Path "HKCU:\Software\Microsoft\Office\$v\Excel") {
        $chaves += "HKCU:\Software\Microsoft\Office\$v\Excel\Options"
    }
}
if ($chaves.Count -eq 0) { $chaves = @("HKCU:\Software\Microsoft\Office\16.0\Excel\Options") }
$valor   = '/R "IMTSolver.xlam"'
$soltos  = @("IMTSolver.xlam", "IMTSolver_Exemplos.xlsx")
$pasta   = "imtsolver"      # o motor: exe + bibliotecas

function Excel-Aberto { return (Get-Process EXCEL -ErrorAction SilentlyContinue) -ne $null }

Write-Host ""
Write-Host "  IMTSolver 2.0" -ForegroundColor Cyan
Write-Host "  ============="
Write-Host ""

if (Excel-Aberto) {
    Write-Host "  O Excel esta aberto. Feche-o e rode de novo." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "  Enter para sair"
    exit 1
}

if ($Desinstalar) {
    foreach ($a in $soltos) {
        $p = Join-Path $destino $a
        if (Test-Path $p) { Remove-Item $p -Force; Write-Host "  removido  $a" }
    }
    $pm = Join-Path $destino $pasta
    if (Test-Path $pm) { Remove-Item $pm -Recurse -Force; Write-Host "  removida  pasta $pasta" }
    # sobra de versoes antigas, quando o motor era um exe solto
    $velho = Join-Path $destino "imtsolver.exe"
    if (Test-Path $velho) { Remove-Item $velho -Force; Write-Host "  removido  imtsolver.exe (versao antiga)" }

    foreach ($chave in $chaves) {
        if (-not (Test-Path $chave)) { continue }
        Get-ItemProperty $chave | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -like "OPEN*" } | ForEach-Object {
            $v = (Get-ItemProperty $chave -Name $_.Name).($_.Name)
            if ($v -like "*IMTSolver.xlam*") { Remove-ItemProperty $chave -Name $_.Name; Write-Host "  registro  $($_.Name) apagado" }
        }
    }
    Write-Host ""
    Write-Host "  IMTSolver desinstalado." -ForegroundColor Green
    Write-Host ""
    Read-Host "  Enter para sair"
    exit 0
}

# --- conferir o que veio no pacote ---
foreach ($a in $soltos) {
    if (-not (Test-Path (Join-Path $aqui $a))) {
        if ($a -eq "IMTSolver_Exemplos.xlsx") { continue }
        Write-Host "  Nao achei $a nesta pasta." -ForegroundColor Red
        Read-Host "  Enter para sair"
        exit 1
    }
}
if (-not (Test-Path (Join-Path $aqui "$pasta\imtsolver.exe"))) {
    Write-Host "  Nao achei a pasta '$pasta' com o motor." -ForegroundColor Red
    Write-Host "  Descompacte o pacote inteiro antes de instalar." -ForegroundColor Red
    Read-Host "  Enter para sair"
    exit 1
}

# --- copiar ---
if (-not (Test-Path $destino)) { New-Item -ItemType Directory -Path $destino -Force | Out-Null }

foreach ($a in $soltos) {
    $origem = Join-Path $aqui $a
    if (-not (Test-Path $origem)) { continue }
    Copy-Item $origem (Join-Path $destino $a) -Force
    try { Unblock-File (Join-Path $destino $a) -ErrorAction Stop } catch {}
    Write-Host ("  copiado   {0,-26} {1,8:n0} KB" -f $a, ((Get-Item $origem).Length / 1KB))
}

# o motor: pasta inteira, substituindo a anterior
$pm = Join-Path $destino $pasta
if (Test-Path $pm) { Remove-Item $pm -Recurse -Force }
Copy-Item (Join-Path $aqui $pasta) $pm -Recurse -Force
Get-ChildItem $pm -Recurse -File | ForEach-Object { try { Unblock-File $_.FullName -ErrorAction Stop } catch {} }
$n = (Get-ChildItem $pm -Recurse -File | Measure-Object).Count
$mb = (Get-ChildItem $pm -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB
Write-Host ("  copiado   {0,-26} {1,5:n0} arquivos, {2:n0} MB" -f "$pasta\", $n, $mb)

# sobra da versao antiga: o motor era um exe solto de 80 MB
$velho = Join-Path $destino "imtsolver.exe"
if (Test-Path $velho) { Remove-Item $velho -Force; Write-Host "  removido  imtsolver.exe da versao anterior" }

# --- registrar em cada Excel encontrado ---
# Dentro de cada versao, ocupamos a primeira chave OPEN / OPEN1 / OPEN2
# livre, para nao atropelar o Solver do Excel nem o OpenSolver.
foreach ($chave in $chaves) {
    if (-not (Test-Path $chave)) { New-Item -Path $chave -Force | Out-Null }
    $ver = ($chave -split "\\")[5]
    $props = Get-ItemProperty $chave
    $jaTem = $false
    $usadas = @()
    foreach ($n in ($props | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -match "^OPEN\d*$" } | ForEach-Object { $_.Name })) {
        $usadas += $n
        if ($props.$n -like "*IMTSolver.xlam*") { $jaTem = $true }
    }
    if ($jaTem) {
        Write-Host "  registro  Office $ver - ja estava registrado"
    } else {
        $nome = "OPEN"; $i = 0
        while ($usadas -contains $nome) { $i++; $nome = "OPEN$i" }
        New-ItemProperty -Path $chave -Name $nome -Value $valor -PropertyType String -Force | Out-Null
        Write-Host "  registro  Office $ver - $nome = $valor"
    }
}

Write-Host ""
Write-Host "  Pronto. Abra o Excel: a aba IMTSolver aparece na faixa de opcoes." -ForegroundColor Green
Write-Host "  Pasta: $destino"
Write-Host ""
Read-Host "  Enter para sair"
