# atualizar.ps1 — Atualização incremental: CSV de log do Windows → Supabase
# Uso: .\atualizar.ps1 -Csv "caminho\para\log.csv"
# O script insere apenas registros com data POSTERIOR ao último registro no banco.

param(
    [Parameter(Mandatory)][string]$Csv
)

# Mapeamento impressora → departamento (espelho do IMP_DEPT do index.html)
$IMP_DEPT = @{
    'Brother MFC-8480DN - Secretaria da Presidência'     = 'Presidência'
    'EPSON L1455 - Secretaria da Presidência - Colorida' = 'Presidência'
    'Brother MFC-8480DN - Financeiro'                    = 'Financeiro'
    'Brother DCP-8157DN - Editora'                       = 'Editora'
    'KONICA MINOLTA bizhub C284e - Editora'              = 'Editora'
    'Brother DCP-8080DN - RH'                            = 'RH'
    'Brother MFC-8480DN - Livraria'                      = 'Livraria'
    'Brother MFC-8480DN - DAS'                           = 'DAS'
    'Brother DCP-L5652DN - Patrimônio do Livro'          = 'Patrimônio do Livro'
    'Brother DCP-8485DN - Biblioteca'                    = 'Biblioteca'
    'EPSON L605 Series - Juridico'                       = 'Jurídico'
    'RICOH SP 3710DN - Jurídico'                         = 'Jurídico'
    'Brother DCP-L5652DN Printer - Expedição 1'          = 'Expedição'
    'Brother DCP-L5652DN Printer - Expedição 2'          = 'Expedição'
    # Aliases históricos
    'Brother Impressora Presidencia'                     = 'Presidência'
    'EPSON L1455 Series Presidencia'                     = 'Presidência'
}

$COLORIDAS = @(
    'EPSON L1455 - Secretaria da Presidência - Colorida'
    'KONICA MINOLTA bizhub C284e - Editora'
)

# 1. Ler credenciais do .env
$envFile = Join-Path $PSScriptRoot '.env'
if (-not (Test-Path $envFile)) { Write-Error ".env não encontrado"; exit 1 }

$creds = @{}
foreach ($line in (Get-Content $envFile -Encoding UTF8)) {
    if ($line -match '^\s*([^#=\s][^=]*)=(.*)$') {
        $creds[$Matches[1].Trim()] = $Matches[2].Trim()
    }
}

$SUPABASE_URL = $creds['SUPABASE_URL']
$SUPABASE_KEY = $creds['SUPABASE_SERVICE_KEY']
if (-not $SUPABASE_URL -or -not $SUPABASE_KEY) {
    Write-Error "Credenciais ausentes no .env"; exit 1
}

$getHeaders  = @{ 'apikey' = $SUPABASE_KEY; 'Authorization' = "Bearer $SUPABASE_KEY" }
$postHeaders = @{
    'apikey'        = $SUPABASE_KEY
    'Authorization' = "Bearer $SUPABASE_KEY"
    'Content-Type'  = 'application/json'
    'Prefer'        = 'return=minimal'
}

# 2. Buscar última data registrada no Supabase
Write-Host "Consultando último registro no Supabase..."
$lastResp = Invoke-RestMethod `
    -Uri "$SUPABASE_URL/rest/v1/impressoes?select=dt&order=dt.desc&limit=1" `
    -Headers $getHeaders

$lastDt = if ($lastResp -and $lastResp.Count -gt 0) { $lastResp[0].dt } else { '2000-01-01' }
Write-Host "Última data no banco  : $lastDt"

# 3. Ler e parsear o CSV do Windows Event Viewer
# Formato: Nível,Data e Hora,Fonte,Identificação do Evento,Categoria da Tarefa,<Descrição>
if (-not (Test-Path $Csv)) { Write-Error "Arquivo CSV não encontrado: $Csv"; exit 1 }

$lines = Get-Content -Path $Csv -Encoding UTF8 | Select-Object -Skip 1
$rows  = $lines | ConvertFrom-Csv -Header 'Nivel','DataHora','Fonte','EventoId','Categoria','Descricao'

$parsed  = [System.Collections.Generic.List[object]]::new()
$skipped = 0
$unknown = [System.Collections.Generic.List[string]]::new()

foreach ($row in $rows) {
    if ([string]::IsNullOrWhiteSpace($row.Descricao)) { $skipped++; continue }

    # Extrair nome da impressora
    if ($row.Descricao -notmatch 'foi impresso em (.+?) pela porta') { $skipped++; continue }
    $impressora = $Matches[1].Trim()

    # Extrair número de páginas
    if ($row.Descricao -notmatch 'P[aá]ginas impressas: (\d+)') { $skipped++; continue }
    $pages = [int]$Matches[1]

    # Parsear data: DD/MM/YYYY HH:MM:SS → YYYY-MM-DD
    $dh = $row.DataHora.Trim()
    if ($dh -notmatch '^(\d{2})/(\d{2})/(\d{4})') { $skipped++; continue }
    $dt  = "$($Matches[3])-$($Matches[2])-$($Matches[1])"
    $mes = "$($Matches[2])/$($Matches[3])"

    # Ignorar datas já presentes no banco
    if ($dt -le $lastDt) { continue }

    # Mapear departamento
    $dept = $IMP_DEPT[$impressora]
    if (-not $dept) {
        if (-not $unknown.Contains($impressora)) { $unknown.Add($impressora) }
        $skipped++
        continue
    }

    $colorida = $COLORIDAS -contains $impressora

    $parsed.Add([PSCustomObject]@{
        dt         = $dt
        mes        = $mes
        pages      = $pages
        impressora = $impressora
        dept       = $dept
        colorida   = $colorida
    })
}

Write-Host "Registros novos (após $lastDt) : $($parsed.Count)"
Write-Host "Pulados (antes da data / sem match): $skipped"

if ($unknown.Count -gt 0) {
    Write-Warning "Impressoras não mapeadas (registros ignorados):"
    $unknown | ForEach-Object { Write-Warning "  - $_" }
}

if ($parsed.Count -eq 0) {
    Write-Host "Nenhum registro novo. Nada a inserir."; exit 0
}

# 4. Remover duplicatas internas do CSV
$before  = $parsed.Count
$grouped = $parsed | Group-Object { "$($_.dt)|$($_.impressora)|$($_.pages)|$($_.colorida)" }
$parsed  = [System.Collections.Generic.List[object]]($grouped | ForEach-Object { $_.Group[0] })
if ($parsed.Count -lt $before) {
    Write-Warning "Removidas $($before - $parsed.Count) duplicatas internas no CSV."
}

# 5. Inserir em lotes de 500
$BATCH    = 500
$inserted = 0

for ($i = 0; $i -lt $parsed.Count; $i += $BATCH) {
    $end   = [Math]::Min($i + $BATCH - 1, $parsed.Count - 1)
    $slice = @($parsed[$i..$end])

    $body = ConvertTo-Json -InputObject $slice -Depth 3 -Compress
    if ($slice.Count -eq 1 -and -not $body.StartsWith('[')) { $body = "[$body]" }

    try {
        Invoke-RestMethod -Uri "$SUPABASE_URL/rest/v1/impressoes" `
            -Method POST -Headers $postHeaders -Body $body | Out-Null
        $inserted += $slice.Count
        Write-Host "  Inseridos: $inserted / $($parsed.Count)"
    } catch {
        Write-Error "Erro no lote $i–$end : $_"
        exit 1
    }
}

$totalPages = ($parsed | Measure-Object -Property pages -Sum).Sum
Write-Host ""
Write-Host "Atualização concluída!"
Write-Host "  Registros inseridos : $inserted"
Write-Host "  Páginas adicionadas : $totalPages"
Write-Host "  Novo período máximo : $($parsed[-1].dt)"
