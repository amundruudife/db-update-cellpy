[CmdletBinding()]
param(
    [ValidateSet('update', 'full', 'cluster-only', 'query', 'path', 'explain')]
    [string]$Mode = 'update',
    [string]$Question,
    [string[]]$ExtraArgs = @()
)

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$pythonFile = Join-Path $projectRoot 'graphify-out\.graphify_python'

if (-not (Test-Path -LiteralPath $pythonFile)) {
    throw "Graphify interpreter record is missing: $pythonFile. Install requirements.txt and run Graphify once from the project root."
}

$python = (Get-Content -LiteralPath $pythonFile -Raw).Trim()
if (-not (Test-Path -LiteralPath $python)) {
    throw "Recorded Graphify interpreter does not exist: $python"
}

switch ($Mode) {
    'update' {
        $cliArgs = @($projectRoot, '--update')
    }
    'full' {
        $cliArgs = @($projectRoot)
    }
    'cluster-only' {
        $cliArgs = @($projectRoot, '--cluster-only')
    }
    'query' {
        if ([string]::IsNullOrWhiteSpace($Question)) {
            throw '-Question is required for -Mode query.'
        }
        $cliArgs = @('query', $Question) + $ExtraArgs
    }
    'path' {
        if ($ExtraArgs.Count -lt 2) {
            throw '-ExtraArgs must contain two node names for -Mode path.'
        }
        $cliArgs = @('path') + $ExtraArgs
    }
    'explain' {
        if ([string]::IsNullOrWhiteSpace($Question)) {
            throw '-Question is required for -Mode explain.'
        }
        $cliArgs = @('explain', $Question) + $ExtraArgs
    }
}

Push-Location $projectRoot
try {
    & $python -m graphify @cliArgs
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}
finally {
    Pop-Location
}
