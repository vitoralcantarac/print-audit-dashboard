# recuperar.ps1 -- Insere no Supabase os registros historicos das impressoras
# que foram silenciosamente descartadas pelo bug de encoding do sincronizar.ps1
# Fontes: registros_impressao.xlsx (15/04-22/05) + 28-05-26.csv (23/05-28/05)

param()
$BASE = $PSScriptRoot

function Write-Log {
    param([string]$msg, [string]$level = 'INFO')
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Write-Host "[$ts][$level] $msg"
}

# Impressoras a recuperar — comparacao normalizada (sem acentos) para evitar encoding mismatch
$RECUPERAR_NORM = @(
    'epson l1455 - secretaria da presidencia - colorida'
    'brother impressora presidencia'
    'epson l1455 series presidencia'
    'konica minolta universal v4 pcl - editora - colorida'
)

function Norm([string]$s) {
    # Decompoe caracteres acentuados (NFD) e remove os diacriticos
    $nfd = $s.ToLower().Normalize([System.Text.NormalizationForm]::FormD)
    return [System.Text.RegularExpressions.Regex]::Replace($nfd, '\p{M}', '')
}

# Chaves normalizadas para evitar encoding mismatch
$IMP_DEPT_NORM = @{
    'epson l1455 - secretaria da presidencia - colorida'   = 'Presidência'
    'brother impressora presidencia'                        = 'Presidência'
    'epson l1455 series presidencia'                        = 'Presidência'
    'konica minolta universal v4 pcl - editora - colorida'  = 'Editora'
}

$COLORIDAS_NORM = @(
    'epson l1455 - secretaria da presidencia - colorida'
    'epson l1455 series presidencia'
    'konica minolta universal v4 pcl - editora - colorida'
)

# Ler credenciais
$envFile = Join-Path $BASE '.env'
if (-not (Test-Path $envFile)) { Write-Log ".env nao encontrado" 'ERROR'; exit 1 }
$creds = @{}
foreach ($line in (Get-Content $envFile -Encoding UTF8)) {
    if ($line -match '^\s*([^#=\s][^=]*)=(.*)$') { $creds[$Matches[1].Trim()] = $Matches[2].Trim() }
}
$SUPABASE_URL = $creds['SUPABASE_URL']
$SUPABASE_KEY = $creds['SUPABASE_SERVICE_KEY']
if (-not $SUPABASE_URL -or -not $SUPABASE_KEY) { Write-Log "Credenciais ausentes" 'ERROR'; exit 1 }

$postHeaders = @{
    'apikey'        = $SUPABASE_KEY
    'Authorization' = "Bearer $SUPABASE_KEY"
    'Content-Type'  = 'application/json'
    'Prefer'        = 'return=minimal'
}

$parsed = [System.Collections.Generic.List[object]]::new()

# ── FONTE 1: registros_impressao.xlsx (15/04–22/05) ──────────────────────────
Write-Log "Lendo registros_impressao.xlsx..."
$xlsxPath = Join-Path $BASE 'registros_impressao.xlsx'
$tmpDir   = Join-Path $env:TEMP 'recuperar_xlsx'
if (Test-Path $tmpDir) { Remove-Item $tmpDir -Recurse -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($xlsxPath, $tmpDir)

$sheet = [xml](Get-Content "$tmpDir\xl\worksheets\sheet1.xml" -Encoding UTF8)
$rows  = $sheet.worksheet.sheetData.row
$xlsxCount = 0

foreach ($row in $rows[1..($rows.Count - 1)]) {
    # Mapear celulas pela referencia de coluna (ex: C5 -> coluna C) para nao
    # depender de posicao no array quando celulas vazias sao omitidas no XML
    $cm = @{}
    foreach ($c in $row.c) { $cm[($c.r -replace '\d+','')] = $c.v }
    $dt    = $cm['A']
    $mes   = $cm['B']
    $imp   = $cm['C']
    $pages = [int]($cm['E'])
    $col   = $cm['F']   # Colorida (Sim/Nao)

    if (-not $imp) { continue }
    if ((Norm $imp) -notin $RECUPERAR_NORM) { continue }
    if ($dt -gt '2026-05-22') { continue }  # xlsx cobre ate 22/05

    $dept     = $IMP_DEPT_NORM[(Norm $imp)]
    $colorida = ($col -eq 'Sim')

    $parsed.Add([PSCustomObject]@{
        dt         = $dt
        mes        = $mes
        pages      = $pages
        impressora = $imp
        dept       = $dept
        colorida   = $colorida
    })
    $xlsxCount++
}
Write-Log "xlsx: $xlsxCount registros extraidos das impressoras ausentes"

# ── FONTE 2: 28-05-26.csv (23/05–28/05) ──────────────────────────────────────
Write-Log "Lendo 28-05-26.csv (23/05-28/05)..."
$csvPath  = Join-Path $BASE '28-05-26.csv'
$csvLines = [System.IO.File]::ReadAllLines($csvPath, [System.Text.Encoding]::UTF8)
$csvCount = 0

foreach ($line in $csvLines[1..($csvLines.Count - 1)]) {
    $mImp   = [regex]::Match($line, 'foi impresso em (.+?) pela porta')
    $mPages = [regex]::Match($line, 'P.ginas impressas: (\d+)')
    $mDt    = [regex]::Match($line, '(\d{2}/\d{2}/\d{4}) \d{2}:\d{2}:\d{2}')
    if (-not $mImp.Success -or -not $mPages.Success -or -not $mDt.Success) { continue }

    $imp = $mImp.Groups[1].Value.Trim()
    if ((Norm $imp) -notin $RECUPERAR_NORM) { continue }

    $dtParsed = [DateTime]::ParseExact($mDt.Groups[1].Value, 'dd/MM/yyyy', $null)
    if ($dtParsed -lt [DateTime]'2026-05-23') { continue }  # xlsx ja cobre ate 22/05

    $dt    = $dtParsed.ToString('yyyy-MM-dd')
    $mes   = $dtParsed.ToString('MM/yyyy')
    $pages = [int]$mPages.Groups[1].Value
    $dept  = $IMP_DEPT_NORM[(Norm $imp)]
    if (-not $dept) { continue }
    $colorida = $COLORIDAS_NORM -contains (Norm $imp)

    $parsed.Add([PSCustomObject]@{
        dt         = $dt
        mes        = $mes
        pages      = $pages
        impressora = $imp
        dept       = $dept
        colorida   = $colorida
    })
    $csvCount++
}
Write-Log "csv: $csvCount registros extraidos das impressoras ausentes (23/05-28/05)"

# Sem dedup adicional: xlsx e csv já cobrem períodos distintos (≤22/05 e ≥23/05)
# e cada linha representa um trabalho de impressão único na origem

if ($parsed.Count -eq 0) {
    Write-Log "Nenhum registro a inserir."
    exit 0
}

$totalPages = ($parsed | Measure-Object -Property pages -Sum).Sum
Write-Log "Total de paginas a inserir: $totalPages"

# ── Inserir em lotes de 500 ───────────────────────────────────────────────────
$BATCH    = 500
$inserted = 0

for ($i = 0; $i -lt $parsed.Count; $i += $BATCH) {
    $endIdx     = [Math]::Min($i + $BATCH - 1, $parsed.Count - 1)
    $slice      = @($parsed[$i..$endIdx])
    $body       = ConvertTo-Json -InputObject $slice -Depth 3 -Compress
    if ($slice.Count -eq 1 -and -not $body.StartsWith('[')) { $body = "[$body]" }

    try {
        Invoke-RestMethod -Uri "$SUPABASE_URL/rest/v1/impressoes" `
            -Method POST -Headers $postHeaders `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) | Out-Null
        $inserted += $slice.Count
        Write-Log "  Inseridos: $inserted / $($parsed.Count)"
    } catch {
        Write-Log "Erro ao inserir lote: $($_.Exception.Message)" 'ERROR'
        exit 1
    }
}

Write-Log "=== Recuperacao concluida: $inserted registros | $totalPages paginas ==="
