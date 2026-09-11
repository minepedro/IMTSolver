# Um imageMso que nao existe nesta versao do Office aparece como botao
# em branco na faixa - sem erro nenhum. Aqui conferimos todos.
#
# Truque: GetImageMso devolve um IPictureDisp que o PowerShell nao sabe
# converter, entao um id VALIDO estoura "Falha catastrofica" na conversao,
# e um id INVALIDO estoura "valor fora do intervalo" antes disso.
param([string]$xml = "$PSScriptRoot\customUI14.xml")
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") }
catch { $xl = New-Object -ComObject Excel.Application; $novo = $true }
$ids = Select-String -Path $xml -Pattern 'imageMso="([^"]+)"' -AllMatches |
       ForEach-Object { $_.Matches } | ForEach-Object { $_.Groups[1].Value }
$ruins = 0
foreach ($c in $ids) {
  try { $null = $xl.CommandBars.GetImageMso($c, 32, 32) }
  catch {
    if ($_.Exception.Message -like "*catastr*" -or $_.Exception.Message -like "*catastrophic*") { "OK       $c" }
    else { "INVALIDO $c"; $ruins++ }
  }
}
if ($novo) { $xl.Quit() }
if ($ruins -gt 0) { "`n$ruins icone(s) apareceriam em branco na faixa"; exit 1 } else { "`ntodos os icones existem" }
