<#
  lint.ps1 — Verificação de erros do projeto ProjectMoney
  Equivalente ao `npm run lint` / ESLint do mundo React+TypeScript.

  Uso no terminal:
      .\lint.ps1
      .\lint.ps1 -Project projectMoney.lpi

  O que faz:
    1. Roda "lazbuild -B" (compilação limpa de todas as unidades)
    2. Converte as mensagens do FPC no formato
         <arquivo>(<linha>,<coluna>) <severidade>: <mensagem>
       que é o formato que o problem matcher do VS Code entende
    3. Sai com código 1 se houver erro (igual ao ESLint)

  Saída das severidades:
      Fatal|Error -> error
      Warning     -> warning
      Note|Hint   -> info
#>

param(
    [string]$Project = 'projectMoney.lpi'
)

$ErrorActionPreference = 'Continue'

Push-Location $PSScriptRoot
try {
    $out = & lazbuild -B $Project 2>&1 | ForEach-Object { [string]$_ }
    $buildFailed = ($LASTEXITCODE -ne 0)
}
finally {
    Pop-Location
}

# file(line,col) Fatal|Error|Warning|Note|Hint: message
$msgPattern = '^(?<file>[^\s].*?\.(pas|pp|p|lpr|dpr|inc))\((?<line>\d+),(?<col>\d+)\)\s+(?<kind>Fatal|Error|Warning|Note|Hint):\s*(?<msg>.*)$'

$nErr = 0
$nWarn = 0
$nInfo = 0

$fileMsgs = New-Object System.Collections.Generic.List[string]
$plainMsgs = New-Object System.Collections.Generic.List[string]

foreach ($line in $out) {
    if ($line -match $msgPattern) {
        $file = $Matches['file']
        $ln   = $Matches['line']
        $col  = $Matches['col']
        $kind = $Matches['kind']
        # remove o código "(5024) " que o FPC coloca no começo das dicas
        $msg  = ($Matches['msg'] -replace '^\(\d+\)\s*', '')

        switch ($kind) {
            { $_ -in 'Fatal', 'Error' } { $sev = 'error';  $nErr++  }
            'Warning'                  { $sev = 'warning'; $nWarn++ }
            default                    { $sev = 'info';    $nInfo++ }
        }

        $fileMsgs.Add(("{0}({1},{2}) {3}: {4}" -f $file, $ln, $col, $sev, $msg))
    }
    elseif ($line -match '^\s*(Fatal|Error):\s*(.*)$') {
        # erro fatal sem arquivo (ex.: falha de link); guardado para depois
        $plainMsgs.Add($Matches[2].Trim())
    }
}

# mensagens com arquivo: sempre mostradas (são as apontadas pelo editor)
$fileMsgs | ForEach-Object { Write-Output $_ }

# sem arquivo: só entram se nenhuma mensagem com arquivo apareceu,
# evitando duplicar o mesmo problema no painel de Problemas
if ($fileMsgs.Count -eq 0) {
    $plainMsgs | Select-Object -Unique | ForEach-Object {
        $nErr++
        Write-Output ("{0}(1,1) error: {1}" -f $Project, $_)
    }
}

Write-Output ("lint: {0} erro(s), {1} aviso(s), {2} dica(s)" -f $nErr, $nWarn, $nInfo)

if ($nErr -gt 0 -or $buildFailed) {
    exit 1
}
exit 0
