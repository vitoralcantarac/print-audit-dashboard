# sincronizar.ps1 -- Sincronizacao automatica: Event Log -> Supabase
# Agendado diariamente via Task Scheduler no servidor de impressao.
# Nao requer exportacao manual de CSV.

param()

$BASE = $PSScriptRoot
$LOG  = Join-Path $BASE 'sincronizar.log'

function Write-Log {
    param([string]$msg, [string]$level = 'INFO')
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts][$level] $msg"
    Add-Content -Path $LOG -Value $line -Encoding UTF8
    Write-Host $line
}

$IMP_DEPT = @{
    # Presidencia
    'Brother MFC-8480DN - Secretaria da Presidencia'     = 'Presidencia'
    'Brother MFC-8480DN - Secretaria da Presidência'     = 'Presidência'
    'EPSON L1455 - Secretaria da Presidencia - Colorida' = 'Presidencia'
    'EPSON L1455 - Secretaria da Presidência - Colorida' = 'Presidência'
    'Brother Impressora Presidencia'                     = 'Presidência'
    'Brother Impressora Presidência'                     = 'Presidência'
    'EPSON L1455 Series Presidencia'                     = 'Presidência'
    'EPSON L1455 Series Presidência'                     = 'Presidência'
    # Editora (inclui Konica Minolta nova)
    'KONICA MINOLTA Universal V4 PCL - Editora - Colorida' = 'Editora'
    # Financeiro
    'Brother MFC-8480DN - Financeiro'                    = 'Financeiro'
    # Editora
    'Brother DCP-8157DN - Editora'                       = 'Editora'
    'KONICA MINOLTA bizhub C284e - Editora'              = 'Editora'
    # RH
    'Brother DCP-8080DN - RH'                            = 'RH'
    # Livraria
    'Brother MFC-8480DN - Livraria'                      = 'Livraria'
    # DAS
    'Brother MFC-8480DN - DAS'                           = 'DAS'
    # Patrimonio do Livro
    'Brother DCP-L5652DN - Patrimonio do Livro'          = 'Patrimônio do Livro'
    'Brother DCP-L5652DN - Patrimônio do Livro'          = 'Patrimônio do Livro'
    # Biblioteca
    'Brother DCP-8485DN - Biblioteca'                    = 'Biblioteca'
    # Juridico
    'EPSON L605 Series - Juridico'                       = 'Jurídico'
    'RICOH SP 3710DN - Juridico'                         = 'Jurídico'
    'RICOH SP 3710DN - Jurídico'                         = 'Jurídico'
    # Expedicao
    'Brother DCP-L5652DN Printer - Expedicao 1'          = 'Expedição'
    'Brother DCP-L5652DN Printer - Expedicao 2'          = 'Expedição'
    'Brother DCP-L5652DN Printer - Expedição 1'          = 'Expedição'
    'Brother DCP-L5652DN Printer - Expedição 2'          = 'Expedição'
}

$COLORIDAS = @(
    'EPSON L1455 - Secretaria da Presidência - Colorida'
    'EPSON L1455 - Secretaria da Presidencia - Colorida'
    'EPSON L1455 Series Presidencia'
    'EPSON L1455 Series Presidência'
    'KONICA MINOLTA bizhub C284e - Editora'
    'KONICA MINOLTA Universal V4 PCL - Editora - Colorida'
)

Write-Log "=== Iniciando sincronizacao ==="

# 1. Ler credenciais do .env
$envFile = Join-Path $BASE '.env'
if (-not (Test-Path $envFile)) {
    Write-Log ".env nao encontrado em $BASE" 'ERROR'
    exit 1
}

$creds = @{}
foreach ($line in (Get-Content $envFile -Encoding UTF8)) {
    if ($line -match '^\s*([^#=\s][^=]*)=(.*)$') {
        $creds[$Matches[1].Trim()] = $Matches[2].Trim()
    }
}

$SUPABASE_URL = $creds['SUPABASE_URL']
$SUPABASE_KEY = $creds['SUPABASE_SERVICE_KEY']
if (-not $SUPABASE_URL -or -not $SUPABASE_KEY) {
    Write-Log "Credenciais ausentes no .env" 'ERROR'
    exit 1
}

$getHeaders  = @{ 'apikey' = $SUPABASE_KEY; 'Authorization' = "Bearer $SUPABASE_KEY" }
$postHeaders = @{
    'apikey'        = $SUPABASE_KEY
    'Authorization' = "Bearer $SUPABASE_KEY"
    'Content-Type'  = 'application/json'
    'Prefer'        = 'return=minimal'
}

# 2. Buscar ultima data registrada no Supabase
try {
    $lastResp = Invoke-RestMethod `
        -Uri "$SUPABASE_URL/rest/v1/impressoes?select=dt&order=dt.desc&limit=1" `
        -Headers $getHeaders
    $lastDt = if ($lastResp -and $lastResp.Count -gt 0) { $lastResp[0].dt } else { '2000-01-01' }
} catch {
    Write-Log "Erro ao consultar Supabase: $($_.Exception.Message)" 'ERROR'
    exit 1
}

Write-Log "Ultima data no banco: $lastDt"

# 3. Ler eventos do Event Log a partir do dia seguinte ao ultimo registrado
$startTime = [DateTime]::Parse($lastDt).AddDays(1)
Write-Log "Buscando eventos a partir de: $($startTime.ToString('yyyy-MM-dd'))"

try {
    $events = Get-WinEvent -FilterHashtable @{
        LogName   = 'Microsoft-Windows-PrintService/Operational'
        Id        = 307
        StartTime = $startTime
    } -ErrorAction Stop
} catch {
    if ($_.Exception.Message -match 'No events were found|nenhum evento') {
        Write-Log "Nenhum evento novo no Event Log."
        Write-Log "=== Sincronizacao concluida: 0 registros inseridos ==="
        exit 0
    }
    Write-Log "Erro ao ler Event Log: $($_.Exception.Message)" 'ERROR'
    exit 1
}

Write-Log "Eventos encontrados: $($events.Count)"

# 4. Parsear eventos
$parsed  = [System.Collections.Generic.List[object]]::new()
$skipped = 0
$unknown = [System.Collections.Generic.List[string]]::new()

foreach ($ev in $events) {
    $msg = $ev.Message

    if ($msg -notmatch 'foi impresso em (.+?) pela porta') { $skipped++; continue }
    $impressora = $Matches[1].Trim()

    if ($msg -notmatch 'P.ginas impressas: (\d+)') { $skipped++; continue }
    $pages = [int]$Matches[1]

    $dt   = $ev.TimeCreated.ToString('yyyy-MM-dd')
    $hora = $ev.TimeCreated.ToString('HH:mm:ss')
    $mes  = $ev.TimeCreated.ToString('MM/yyyy')

    $dept = $IMP_DEPT[$impressora]
    if (-not $dept) {
        # Fallback: normaliza via NFD e remove diacriticos (robusto a diferenca de encoding)
        $norm = [System.Text.RegularExpressions.Regex]::Replace(
            $impressora.ToLower().Normalize([System.Text.NormalizationForm]::FormD), '\p{M}', '')
        foreach ($k in $IMP_DEPT.Keys) {
            $kNorm = [System.Text.RegularExpressions.Regex]::Replace(
                $k.ToLower().Normalize([System.Text.NormalizationForm]::FormD), '\p{M}', '')
            if ($kNorm -eq $norm) { $dept = $IMP_DEPT[$k]; break }
        }
    }
    if (-not $dept) {
        if (-not $unknown.Contains($impressora)) { $unknown.Add($impressora) }
        $skipped++
        continue
    }

    $impNorm = [System.Text.RegularExpressions.Regex]::Replace(
        $impressora.ToLower().Normalize([System.Text.NormalizationForm]::FormD), '\p{M}', '')
    $colorida = ($COLORIDAS | ForEach-Object {
        [System.Text.RegularExpressions.Regex]::Replace(
            $_.ToLower().Normalize([System.Text.NormalizationForm]::FormD), '\p{M}', '')
    }) -contains $impNorm

    $parsed.Add([PSCustomObject]@{
        dt         = $dt
        hora       = $hora
        mes        = $mes
        pages      = $pages
        impressora = $impressora
        dept       = $dept
        colorida   = $colorida
    })
}

Write-Log "Registros validos: $($parsed.Count)  |  Pulados: $skipped"

if ($unknown.Count -gt 0) {
    Write-Log "Impressoras nao mapeadas (ignoradas): $($unknown -join ' | ')" 'WARN'
}

if ($parsed.Count -eq 0) {
    Write-Log "Nenhum registro novo a inserir."
    Write-Log "=== Sincronizacao concluida: 0 registros inseridos ==="
    exit 0
}

# 5. Remover duplicatas internas
$before  = $parsed.Count
$grouped = $parsed | Group-Object { "$($_.dt)|$($_.hora)|$($_.impressora)|$($_.pages)" }
$parsed  = [System.Collections.Generic.List[object]]($grouped | ForEach-Object { $_.Group[0] })
if ($parsed.Count -lt $before) {
    Write-Log "Duplicatas removidas: $($before - $parsed.Count)" 'WARN'
}

# 6. Inserir em lotes de 500
$BATCH    = 500
$inserted = 0

for ($i = 0; $i -lt $parsed.Count; $i += $BATCH) {
    $endIdx = [Math]::Min($i + $BATCH - 1, $parsed.Count - 1)
    $slice  = @($parsed[$i..$endIdx])

    $sliceClean = $slice | Select-Object dt, mes, pages, impressora, dept, colorida
    $body = ConvertTo-Json -InputObject $sliceClean -Depth 3 -Compress
    if ($sliceClean.Count -eq 1 -and -not $body.StartsWith('[')) { $body = "[$body]" }

    try {
        Invoke-RestMethod -Uri "$SUPABASE_URL/rest/v1/impressoes" `
            -Method POST -Headers $postHeaders -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) | Out-Null
        $inserted += $slice.Count
        Write-Log "  Inseridos: $inserted / $($parsed.Count)"
    } catch {
        Write-Log "Erro ao inserir lote $i-$endIdx : $($_.Exception.Message)" 'ERROR'
        exit 1
    }
}

$totalPages = ($parsed | Measure-Object -Property pages -Sum).Sum
Write-Log "=== Sincronizacao concluida: $inserted registros | $totalPages paginas ==="
