# migrar.ps1 -- Migracao unica: data.json -> Supabase
# Execute uma unica vez com a tabela impressoes vazia.

param()

# 1. Ler credenciais do .env
$envFile = Join-Path $PSScriptRoot '.env'
if (-not (Test-Path $envFile)) { Write-Error ".env nao encontrado em $PSScriptRoot"; exit 1 }

$creds = @{}
foreach ($line in (Get-Content $envFile -Encoding UTF8)) {
    if ($line -match '^\s*([^#=\s][^=]*)=(.*)$') {
        $creds[$Matches[1].Trim()] = $Matches[2].Trim()
    }
}

$SUPABASE_URL = $creds['SUPABASE_URL']
$SUPABASE_KEY = $creds['SUPABASE_SERVICE_KEY']

if (-not $SUPABASE_URL -or -not $SUPABASE_KEY) {
    Write-Error "SUPABASE_URL ou SUPABASE_SERVICE_KEY ausentes no .env"; exit 1
}

# 2. Ler data.json
$dataFile = Join-Path $PSScriptRoot 'data.json'
if (-not (Test-Path $dataFile)) { Write-Error "data.json nao encontrado"; exit 1 }

$json    = Get-Content $dataFile -Raw -Encoding UTF8 | ConvertFrom-Json
$records = $json.data
Write-Host "Registros a migrar: $($records.Count)"

# 3. Inserir em lotes de 500
$BATCH    = 500
$inserted = 0
$headers  = @{
    'apikey'        = $SUPABASE_KEY
    'Authorization' = "Bearer $SUPABASE_KEY"
    'Content-Type'  = 'application/json'
    'Prefer'        = 'return=minimal'
}

for ($i = 0; $i -lt $records.Count; $i += $BATCH) {
    $end   = [Math]::Min($i + $BATCH - 1, $records.Count - 1)
    $slice = $records[$i..$end]

    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $slice) {
        $items.Add([PSCustomObject]@{
            dt         = [string]$r.dt
            mes        = [string]$r.mes
            pages      = [int]$r.pages
            impressora = [string]$r.impressora
            dept       = [string]$r.dept
            colorida   = [bool]$r.colorida
        })
    }

    $body  = ConvertTo-Json -InputObject $items.ToArray() -Depth 3 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)

    try {
        Invoke-RestMethod -Uri "$SUPABASE_URL/rest/v1/impressoes" `
            -Method POST -Headers $headers -Body $bytes | Out-Null
        $inserted += $slice.Count
        Write-Host "  Inseridos: $inserted / $($records.Count)"
    } catch {
        Write-Error "Erro no lote $i-$end : $_"
        exit 1
    }
}

Write-Host ""
Write-Host "Migracao concluida: $inserted registros inseridos no Supabase."
