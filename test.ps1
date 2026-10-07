<#
  test.ps1 — Suite de verificação do ProjectMoney
  Ponto de entrada único (equivalente a "npm test").

  Uso:
      .\test.ps1              # roda tudo (compilação/lint + testes de execução)
      .\test.ps1 -SkipLint    # só os testes de execução (usa o .exe atual)
      .\test.ps1 -SkipDllTest # pula o teste de DLL ausente

  O que verifica:
    [1] Compilação limpa + mensagens do FPC (usa o lint.ps1)
    [2] Sem banks.db  -> aplicação cria o arquivo com a estrutura esperada
    [3] Com banks.db  -> aplicação não recria nem altera o arquivo
    [4] Sem sqlite3.dll -> erro tratado com diálogo (sem crash, sem arquivo parcial)
    [5] Novo Database (miNew) -> cria um .db com as tabelas contas/extratos/saldos
    [6] Abrir Database (miOpen) -> abre o .db valido e rejeita arquivo invalido
    [7] Fechar Database (miClose) -> encerra a conexao com o arquivo de contas
    [8] Gerenciar Contas (miGerCon) -> abre tbContas e o "Voltar" devolve o estado
    [9] Nenhum crash registrado no log de eventos do Windows (WER)

  Sai com código 0 quando tudo passa, 1 quando há alguma falha.
  As mensagens "arquivo(linha,col) severidade: mensagem" alimentam o
  problem matcher do .vscode/tasks.json (painel Problems + Error Lens).
#>

[CmdletBinding()]
param(
    [switch]$SkipLint,
    [switch]$SkipDllTest,
    [int]$RunSeconds = 4
)

$ErrorActionPreference = 'Continue'
$root   = $PSScriptRoot
$exe    = Join-Path $root 'projectMoney.exe'
$dll    = Join-Path $root 'sqlite3.dll'
$db     = Join-Path $root 'banks.db'
$minewDb = Join-Path $root 'miNew-test.db'
$lixoDb  = Join-Path $root 'lixo-test.db'
$passed = 0
$failed = 0

Set-Location $root

# P/Invoke mínimo para dirigir a interface: clique no menu -> dialogo de
# arquivo -> caminho digitado -> confirmar (usado por miNew e miOpen).
if (-not ('UiTest' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class UiTest {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindowEx(IntPtr p, IntPtr a, string c, string t);
    // Unicode (SendMessageW) de proposito: o buffer de SetText e UTF-16 e a
    // versao A cortaria o texto no primeiro byte nulo ("C:\..." viraria "C").
    [DllImport("user32.dll", CharSet=CharSet.Unicode, EntryPoint="SendMessageW")]
    static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, EntryPoint="SendMessageW")]
    static extern IntPtr SendMessageBuf(IntPtr h, uint m, IntPtr w, [Out] StringBuilder s);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, IntPtr l, uint f, uint t, out IntPtr r);
    [DllImport("user32.dll")] static extern IntPtr GetMenu(IntPtr h);
    [DllImport("user32.dll")] static extern IntPtr GetSubMenu(IntPtr h, int pos);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetMenuString(IntPtr h, uint id, StringBuilder s, int max, bool byPos);
    [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);

    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int left, top, right, bottom; }

    public static IntPtr Msg(IntPtr h, uint m, IntPtr w, IntPtr l) {
        IntPtr r; SendMessageTimeout(h, m, w, l, 2, 3000, out r); return r;
    }
    // Primeiro dialogo (#32770) que pertence ao processo indicado.
    public static IntPtr FindDialog(uint procId) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((h, l) => {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p != procId) return true;
            var c = new StringBuilder(64); GetClassName(h, c, 64);
            if (c.ToString() == "#32770") { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
    public static IntPtr FindChild(IntPtr parent, string cls) {
        return FindWindowEx(parent, IntPtr.Zero, cls, null);
    }
    // Botao de comando 1 (OK/Salvar): nao depende do idioma do Windows.
    public static IntPtr FindOkButton(IntPtr parent) {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(parent, (h, l) => {
            if (GetDlgCtrlID(h) == 1) { found = h; return false; }
            return true;
        }, IntPtr.Zero);
        return found;
    }
    // WM_SETTEXT sincrono (seguro liberar a memoria em seguida).
    public static bool SetText(IntPtr h, string s) {
        IntPtr p = Marshal.StringToHGlobalUni(s);
        try { return SendMessage(h, 0x000C, IntPtr.Zero, p) != IntPtr.Zero; }
        finally { Marshal.FreeHGlobal(p); }
    }
    // WM_GETTEXT sincrono (confirma que gravamos no controle certo).
    public static string GetText(IntPtr h) {
        var sb = new StringBuilder(1024);
        SendMessageBuf(h, 0x000D, (IntPtr)sb.Capacity, sb);
        return sb.ToString();
    }
    // Campo de nome do arquivo: no dialogo "Salvar" ele e o Edit de id 1001,
    // mas no "Abrir" o campo e outro (id 1148) e o Edit da barra de endereco
    // (id 41477) esta invisivel - por isso a reserva e o primeiro Edit VISIVEL
    // (sem isso o caminho ia para a barra de endereco e o OK nao faria nada).
    public static IntPtr FindFileNameEdit(IntPtr parent) {
        IntPtr byId = IntPtr.Zero, firstVis = IntPtr.Zero, first = IntPtr.Zero;
        EnumChildWindows(parent, (h, l) => {
            var c = new StringBuilder(128); GetClassName(h, c, 128);
            if (c.ToString() != "Edit") return true;
            if (GetDlgCtrlID(h) == 1001) { byId = h; return false; }
            if (first == IntPtr.Zero) first = h;
            if (firstVis == IntPtr.Zero && IsWindowVisible(h)) firstVis = h;
            return true;
        }, IntPtr.Zero);
        if (byId != IntPtr.Zero) return byId;
        if (firstVis != IntPtr.Zero) return firstVis;
        return first;
    }
    // Id do comando de menu procurado pelo texto. Varre todos os submenus:
    // o id e unico no menu principal, entao basta achar em qualquer um
    // (necessario para itens fora do "Arquivo", como o miGerCon).
    public static int MenuId(IntPtr main, string text) {
        IntPtr menu = GetMenu(main);
        if (menu == IntPtr.Zero) return -1;
        for (int sub = 0; sub < 16; sub++) {
            IntPtr sm = GetSubMenu(menu, sub);
            if (sm == IntPtr.Zero) continue;
            var sb = new StringBuilder(256);
            for (uint id = 1; id <= 120; id++) {
                sb.Length = 0;
                GetMenuString(sm, id, sb, 256, false);
                if (sb.ToString() == text) return (int)id;
            }
        }
        return -1;
    }
    // Janelas filhas visiveis, no formato "classe | id | texto | x,y WxH".
    // Comparar antes/depois prova que a aba abriu (novas janelas) e que o
    // "Voltar" devolveu o estado (volta ao conjunto inicial).
    public static string[] Visible(IntPtr parent) {
        var res = new System.Collections.Generic.List<string>();
        EnumChildWindows(parent, (h, l) => {
            if (!IsWindowVisible(h)) return true;
            var c = new StringBuilder(128); GetClassName(h, c, 128);
            var t = new StringBuilder(256); GetWindowText(h, t, 256);
            RECT r; GetWindowRect(h, out r);
            res.Add(c.ToString() + " | id=" + h + " | \"" + t.ToString() + "\" | " +
                    r.left + "," + r.top + " " + (r.right - r.left) + "x" + (r.bottom - r.top));
            return true;
        }, IntPtr.Zero);
        return res.ToArray();
    }
    // Clique de verdade: WM_LBUTTONDOWN/UP com coordenadas relativas ao
    // proprio HWND, SEM WindowFromPoint - assim a janela nao precisa estar
    // em primeiro plano (na suíte o form fica atras da IDE e WindowFromPoint
    // devolveria a janela errada). TSpeedButton nao tem HWND proprio, entao
    // o alvo e o painel pai (o LCL roteia o mouse para o botao grafico).
    public static string ClickOn(IntPtr h, int rx, int ry) {
        IntPtr lp = (IntPtr)((ry << 16) | (rx & 0xFFFF));
        PostMessage(h, 0x0201, (IntPtr)1, lp);   // WM_LBUTTONDOWN
        PostMessage(h, 0x0202, (IntPtr)0, lp);   // WM_LBUTTONUP
        return "hwnd=" + h + " rel=" + rx + "," + ry;
    }
}
"@
}

function Write-Banner([string]$Text) {
    Write-Output ''
    Write-Output ('== ' + $Text)
}

function Check([string]$Name, [bool]$Ok, [string]$Detail = '') {
    if ($Ok) {
        $script:passed++
        Write-Output ('   [OK]      ' + $Name)
    }
    else {
        $script:failed++
        $suffix = if ($Detail) { ' -> ' + $Detail } else { '' }
        Write-Output ('   [FALHA]   ' + $Name + $suffix)
    }
}

# Inicia a aplicação, espera, mata se ainda estiver viva e devolve o resultado.
function Invoke-App([int]$Seconds) {
    $p = Start-Process -FilePath $exe -WorkingDirectory $root -PassThru
    Start-Sleep -Seconds $Seconds
    $p.Refresh()
    $exited = $p.HasExited
    $code = $null
    if ($exited) {
        $code = $p.ExitCode
    }
    else {
        try { $p.Kill(); $p.WaitForExit() } catch { }
    }
    [pscustomobject]@{ Exited = $exited; ExitCode = $code }
}

# Verifica se o processo morreu sozinho (crash) ou se saiu com código != 0.
function Test-NoCrash($Result, [string]$Label) {
    if (-not $Result.Exited) {
        Check $Label $true
    }
    elseif ($Result.ExitCode -eq 0) {
        Check $Label $false 'a aplicação saiu sozinha (exit 0)'
    }
    else {
        Check $Label $false ('crash, exit code ' + $Result.ExitCode)
    }
}

function Get-DbInfo([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text  = [Text.Encoding]::ASCII.GetString($bytes)
    $head  = if ($bytes.Length -ge 15) { [Text.Encoding]::ASCII.GetString($bytes, 0, 15) } else { '' }
    [pscustomobject]@{
        IsSqlite     = ($head -eq 'SQLite format 3')
        HasTable     = $text.Contains('CREATE TABLE "banks"')
        HasPrimaryKey= $text.Contains('PRIMARY KEY("id")')
        HasColumns   = $text.Contains('"id"') -and $text.Contains('"name"') -and $text.Contains('"alias"')
        Size         = $bytes.Length
    }
}

function Get-Hash([string]$Path) {
    (Get-FileHash -Path $Path -Algorithm SHA256).Hash
}

# Le o arquivo: so' funciona com a conexao ENCERRADA, porque o SQLite mantem
# o arquivo travado enquanto a query da tbContas estiver aberta. E a prova
# observavel (sem tocar na interface) de que o miClose fez o trabalho.
function Test-FileReadable([string]$Path) {
    try {
        $null = [IO.File]::ReadAllBytes($Path)
        $true
    }
    catch { $false }
}

function Get-AppCrashes([datetime]$Since) {
    try {
        @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Id = 1000; StartTime = $Since } -ErrorAction Stop |
            Where-Object { $_.Message -like '*projectMoney*' })
    }
    catch { @() }
}

# ------------------------------------------------------------------
# Helpers do dialogo de arquivo (compartilhados pelos passos do miNew/miOpen)

# Clica num item de menu e espera o dialogo #32770 do proprio processo.
# Devolve o HWND do dialogo (Zero quando o item nao existe ou o dialogo nao abriu).
function Invoke-MenuFileDialog($Process, [string]$MenuText) {
    $Process.Refresh()
    $main = $Process.MainWindowHandle
    if ($main -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
    $id = [UiTest]::MenuId($main, $MenuText)
    if ($id -le 0) { return [IntPtr]::Zero }
    # PostMessage e nao SendMessage: o handler abre um dialogo MODAL e so
    # responderia quando ele fechar (a chamada esperaria os 3s do timeout).
    [void][UiTest]::PostMessage($main, 0x0111, [IntPtr]$id, [IntPtr]::Zero)
    for ($t = 0; $t -lt 20; $t++) {
        Start-Sleep -Milliseconds 250
        $w = [UiTest]::FindDialog([uint32]$Process.Id)
        if ($w -ne [IntPtr]::Zero) { return $w }
    }
    [IntPtr]::Zero
}

# Grava o caminho no campo de nome e confirma (botao OK / IDOK).
# Devolve o campo usado (Zero quando o dialogo nao tem campo de nome).
function Set-FileDialogName([IntPtr]$Dlg, [string]$Path) {
    $edit = [UiTest]::FindFileNameEdit($Dlg)
    if ($edit -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
    if (-not [UiTest]::SetText($edit, $Path)) { return [IntPtr]::Zero }
    $btnOk = [UiTest]::FindOkButton($Dlg)
    if ($btnOk -ne [IntPtr]::Zero) {
        [void][UiTest]::PostMessage($btnOk, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)  # BM_CLICK
    }
    else {
        [void][UiTest]::PostMessage($Dlg, 0x0111, [IntPtr]1, [IntPtr]::Zero)          # WM_COMMAND/IDOK
    }
    $edit
}

# Espera o dialogo informado ser destruido (20 x 250ms = 5s).
function Wait-DialogClosed([IntPtr]$Dlg) {
    for ($t = 0; $t -lt 20; $t++) {
        if (-not [UiTest]::IsWindow($Dlg)) { return $true }
        Start-Sleep -Milliseconds 250
    }
    -not [UiTest]::IsWindow($Dlg)
}

# Espera um #32770 aparecer depois de o dialogo de arquivo fechar: e o aviso
# do MessageDlg (SoErros/miOpen). Chamar sempre apos Wait-DialogClosed, senao
# pode devolver o proprio dialogo de arquivo. Devolve o HWND ou Zero.
function Wait-Warning([uint32]$ProcId) {
    for ($t = 0; $t -lt 20; $t++) {
        Start-Sleep -Milliseconds 250
        $w = [UiTest]::FindDialog($ProcId)
        if ($w -ne [IntPtr]::Zero) { return $w }
    }
    [IntPtr]::Zero
}

# ------------------------------------------------------------------
Write-Output '============================================================'
Write-Output ' ProjectMoney - verificacao de funcionamento'
Write-Output (' Executado em: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Write-Output '============================================================'

$startedAt   = Get-Date
$dbExisted   = Test-Path $db
$dbBackup    = Join-Path ([IO.Path]::GetTempPath()) ('projectMoney.banks.' + $PID + '.bak')
$dllRenamed  = $false

if ($dbExisted) {
    Copy-Item -Path $db -Destination $dbBackup -Force
}

try {
    # ------------------------------------------------ [1] compilação/lint
    Write-Banner '[1/9] Compilacao e lint (lazbuild -B)'
    if ($SkipLint) {
        Write-Output '   (pulado por -SkipLint)'
    }
    elseif (-not (Test-Path (Join-Path $root 'lint.ps1'))) {
        Check 'lint.ps1 encontrado' $false (Join-Path $root 'lint.ps1')
    }
    else {
        $lintLines = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'lint.ps1') 2>&1
        $lintExit  = $LASTEXITCODE
        # as linhas sao repassadas sem indentacao de proposito: o problem
        # matcher do VS Code exige que a linha comece com o caminho do arquivo
        foreach ($l in $lintLines) { Write-Output ([string]$l) }
        Check 'compilacao sem erros' ($lintExit -eq 0) ('exit code ' + $lintExit)
        if ($lintExit -ne 0) {
            Write-Output '   (aviso: a compilação falhou; os testes de execução abaixo usam o .exe da última compilação válida)'
        }
    }

    Check 'executavel projectMoney.exe existe' (Test-Path $exe)
    Check 'sqlite3.dll existe' (Test-Path $dll)

    # ------------------------------------------------ [2] cria o banco
    Write-Banner '[2/9] Sem banks.db -> deve criar o arquivo'
    if (Test-Path $db) { Remove-Item $db -Force }
    $r = Invoke-App $RunSeconds
    Test-NoCrash $r 'aplicacao permanece viva (sem crash)'
    Check 'banks.db foi criado' (Test-Path $db)
    if (Test-Path $db) {
        $info = Get-DbInfo $db
        Check 'cabecalho SQLite valido' $info.IsSqlite
        Check 'tabela "banks" presente' $info.HasTable
        Check 'colunas id/name/alias presentes' $info.HasColumns
        Check 'PRIMARY KEY(id) presente' $info.HasPrimaryKey
    }

    # ------------------------------------------------ [3] já existe
    Write-Banner '[3/9] Com banks.db -> nao deve recriar nem alterar'
    if (Test-Path $db) {
        $hashBefore = Get-Hash $db
        $r = Invoke-App $RunSeconds
        Test-NoCrash $r 'aplicacao permanece viva (sem crash)'
        Check 'banks.db nao foi alterado' ((Get-Hash $db) -eq $hashBefore)
    }
    else {
        Check 'banks.db disponivel para o teste' $false 'arquivo ausente'
    }

    # ------------------------------------------------ [4] DLL ausente
    Write-Banner '[4/9] Sem sqlite3.dll -> erro deve ser tratado'
    if ($SkipDllTest) {
        Write-Output '   (pulado por -SkipDllTest)'
    }
    else {
        try {
            if (Test-Path $dll) {
                Rename-Item $dll ($dll + '.bak')
                $dllRenamed = $true
            }
            if (Test-Path $db) { Remove-Item $db -Force }

            $r = Invoke-App $RunSeconds
            # esperado: processo vivo mostrando o dialogo de erro (ou saiu com 0)
            Check 'aplicacao nao crashou com a DLL ausente' (
                (-not $r.Exited) -or ($r.ExitCode -eq 0)
            ) ('exit code ' + $r.ExitCode)
            Check 'nenhum banks.db parcial criado' (-not (Test-Path $db))
        }
        finally {
            if ($dllRenamed -and (Test-Path ($dll + '.bak'))) {
                Rename-Item ($dll + '.bak') 'sqlite3.dll'
                $dllRenamed = $false
            }
        }
        Check 'sqlite3.dll restaurada' (Test-Path $dll)
    }

    # ------------------------------------------------ [5] miNew -> database
    Write-Banner '[5/9] Novo Database (miNew) -> cria .db com as 3 tabelas'
    if (Test-Path $minewDb) { Remove-Item $minewDb -Force -ErrorAction SilentlyContinue }
    $pNovo = $null
    # Qualquer excecao no fluxo tem de virar FALHA: se escapasse em silencio,
    # os Checks seguintes seriam pulados e a suite sairia verde sem testar nada.
    $ErrorActionPreference = 'Stop'
    try {
        $pNovo = Start-Process -FilePath $exe -WorkingDirectory $root -PassThru
        Start-Sleep -Seconds 3
        $pNovo.Refresh()
        $mainNovo = $pNovo.MainWindowHandle
        Check 'aplicacao abriu (janela principal)' ($mainNovo -ne [IntPtr]::Zero)

        $idNovo = -1
        if ('UiTest' -as [type]) { $idNovo = [UiTest]::MenuId($mainNovo, 'Novo Database') }
        # Um helper que nao compilou tem de virar FALHA (senao o passo inteiro
        # seria pulado e a suite sairia verde sem ter testado nada).
        Check 'helper de interface (UiTest) carregado' ([bool]('UiTest' -as [type]))
        Check 'item de menu "Novo Database" encontrado' ($idNovo -gt 0) ('id=' + $idNovo)

        if ($idNovo -gt 0) {
            # Clique no menu: mesmo caminho de um clique real (WM_COMMAND)
            [void][UiTest]::Msg($mainNovo, 0x0111, [IntPtr]$idNovo, [IntPtr]::Zero)

            # Espera o dialogo "Salvar" (modal, mesmo processo)
            $dlg = [IntPtr]::Zero
            for ($t = 0; ($t -lt 20) -and ($dlg -eq [IntPtr]::Zero); $t++) {
                Start-Sleep -Milliseconds 250
                $dlg = [UiTest]::FindDialog([uint32]$pNovo.Id)
            }
            Check 'dialogo "Salvar" abriu' ($dlg -ne [IntPtr]::Zero)

            if ($dlg -ne [IntPtr]::Zero) {
                $edit = [UiTest]::FindFileNameEdit($dlg)
                Check 'campo de nome do arquivo encontrado' ($edit -ne [IntPtr]::Zero)
                if ($edit -ne [IntPtr]::Zero) {
                    Check 'nome do arquivo gravado no dialogo' ([UiTest]::SetText($edit, $minewDb))
                    Check 'campo de nome confere o caminho' ([UiTest]::GetText($edit) -eq $minewDb) (
                        [UiTest]::GetText($edit))
                    $btnOk = [UiTest]::FindOkButton($dlg)
                    if ($btnOk -ne [IntPtr]::Zero) {
                        [void][UiTest]::PostMessage($btnOk, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)  # BM_CLICK
                    }
                    else {
                        [void][UiTest]::PostMessage($dlg, 0x0111, [IntPtr]1, [IntPtr]::Zero)         # WM_COMMAND/IDOK
                    }
                    for ($t = 0; ($t -lt 20) -and (-not (Test-Path $minewDb)); $t++) {
                        Start-Sleep -Milliseconds 250
                    }
                    Check 'miNew-test.db foi criado' (Test-Path $minewDb)
                }

                # Sem dialogo pendente = salvou E a query da aba tbContas abriu:
                # qualquer erro em NewDatabase/Open deixa um MessageDlg na tela.
                Check 'nenhum dialogo pendente (criacao/conexao)' (
                    [UiTest]::FindDialog([uint32]$pNovo.Id) -eq [IntPtr]::Zero)
            }

            $pNovo.Refresh()
            Check 'aplicacao permanece viva (sem crash)' (-not $pNovo.HasExited)

            # O SQLite mantem o database aberto enquanto a query da tbContas
            # estiver ativa: fecha a aplicacao antes de ler os bytes do arquivo.
            try { $pNovo.Kill(); $pNovo.WaitForExit() } catch { }
        }

        if (Test-Path $minewDb) {
            $bytesNovo = [IO.File]::ReadAllBytes($minewDb)
            $headNovo = if ($bytesNovo.Length -ge 15) { [Text.Encoding]::ASCII.GetString($bytesNovo, 0, 15) } else { '' }
            $textoNovo = [Text.Encoding]::ASCII.GetString($bytesNovo)
            Check 'cabecalho SQLite valido' ($headNovo -eq 'SQLite format 3')
            Check 'tabela "contas" presente' $textoNovo.Contains('CREATE TABLE "contas"')
            Check 'tabela "extratos" presente' $textoNovo.Contains('CREATE TABLE "extratos"')
            Check 'tabela "saldos" presente' $textoNovo.Contains('CREATE TABLE "saldos"')
        }
    }
    catch {
        Check 'fluxo do miNew sem excecao' $false $_.Exception.Message
    }
    finally {
        $ErrorActionPreference = 'Continue'
        if ($pNovo -and -not $pNovo.HasExited) {
            try { $pNovo.Kill(); $pNovo.WaitForExit() } catch { }
        }
    }

    # ------------------------------------------------ [6] miOpen -> abrir .db
    Write-Banner '[6/9] Abrir Database (miOpen) -> abre o valido e rejeita o invalido'
    $pAbrir  = $null
    $hashAbr = $null
    # O passo [4] apaga banks.db e ele so' e recriado no passo final: garante o
    # arquivo aqui, porque e' ele o alvo do teste de rejeicao.
    if (-not (Test-Path $db)) { $null = Invoke-App $RunSeconds }
    if (Test-Path $db) { $hashAbr = Get-Hash $db }
    Check 'banks.db disponivel para o teste do miOpen' (Test-Path $db)
    if (-not (Test-Path $minewDb)) {
        Check 'miNew-test.db criado no passo anterior' $false 'arquivo ausente'
    }
    # Mesma protecao dos passos anteriores: excecao tem de virar FALHA.
    $ErrorActionPreference = 'Stop'
    try {
        $pAbrir = Start-Process -FilePath $exe -WorkingDirectory $root -PassThru
        Start-Sleep -Seconds 3
        $pAbrir.Refresh()
        $mainAbrir = $pAbrir.MainWindowHandle
        Check 'aplicacao abriu (janela principal)' ($mainAbrir -ne [IntPtr]::Zero)

        # (a) arquivo valido: o dialogo fecha sozinho e nenhum aviso sobra
        $dlgOk = Invoke-MenuFileDialog $pAbrir 'Abrir Database'
        Check 'dialogo "Abrir" abriu (arquivo valido)' ($dlgOk -ne [IntPtr]::Zero)
        if ($dlgOk -ne [IntPtr]::Zero) {
            $editOk = Set-FileDialogName $dlgOk $minewDb
            Check 'campo de nome do arquivo encontrado' ($editOk -ne [IntPtr]::Zero)
            Check 'dialogo fechou ao confirmar' (Wait-DialogClosed $dlgOk)
            # A conexao e aberta DEPOIS que o dialogo some: se o arquivo nao
            # for um database de contas, o aviso aparece nesse intervalo.
            $avisoOk = [IntPtr]::Zero
            for ($t = 0; ($t -lt 12) -and ($avisoOk -eq [IntPtr]::Zero); $t++) {
                Start-Sleep -Milliseconds 250
                $avisoOk = [UiTest]::FindDialog([uint32]$pAbrir.Id)
            }
            Check 'nenhum aviso pendente (conexao no .db valido)' (
                $avisoOk -eq [IntPtr]::Zero) ('hwnd=' + $avisoOk)
            $pAbrir.Refresh()
            Check 'aplicacao permanece viva (sem crash)' (-not $pAbrir.HasExited)
        }

        # (b) arquivo que nao e' database de contas -> aviso e app continua viva
        [IO.File]::WriteAllText($lixoDb, 'isto nao e um banco sqlite')
        $dlgLixo = Invoke-MenuFileDialog $pAbrir 'Abrir Database'
        Check 'dialogo "Abrir" abriu (arquivo invalido)' ($dlgLixo -ne [IntPtr]::Zero)
        if ($dlgLixo -ne [IntPtr]::Zero) {
            $editLixo = Set-FileDialogName $dlgLixo $lixoDb
            Check 'campo de nome do arquivo encontrado (invalido)' (
                $editLixo -ne [IntPtr]::Zero)
            Check 'dialogo fechou ao confirmar (invalido)' (Wait-DialogClosed $dlgLixo)
            $avisoLixo = Wait-Warning ([uint32]$pAbrir.Id)
            Check 'aviso de arquivo invalido exibido' ($avisoLixo -ne [IntPtr]::Zero)
            if ($avisoLixo -ne [IntPtr]::Zero) {
                [void][UiTest]::PostMessage($avisoLixo, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
                Check 'aviso dispensado' (Wait-DialogClosed $avisoLixo)
            }
            $pAbrir.Refresh()
            Check 'aplicacao permanece viva apos o aviso' (-not $pAbrir.HasExited)
        }

        # (c) banks.db: arquivo real, mas sem a tabela "contas" -> rejeitado
        $dlgBanc = Invoke-MenuFileDialog $pAbrir 'Abrir Database'
        Check 'dialogo "Abrir" abriu (banks.db)' ($dlgBanc -ne [IntPtr]::Zero)
        if ($dlgBanc -ne [IntPtr]::Zero) {
            $editBanc = Set-FileDialogName $dlgBanc $db
            Check 'campo de nome do arquivo encontrado (banks.db)' (
                $editBanc -ne [IntPtr]::Zero)
            Check 'dialogo fechou ao confirmar (banks.db)' (Wait-DialogClosed $dlgBanc)
            $avisoBanc = Wait-Warning ([uint32]$pAbrir.Id)
            Check 'banks.db rejeitado com aviso' ($avisoBanc -ne [IntPtr]::Zero)
            if ($avisoBanc -ne [IntPtr]::Zero) {
                [void][UiTest]::PostMessage($avisoBanc, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
                Check 'aviso de banks.db dispensado' (Wait-DialogClosed $avisoBanc)
            }
            $pAbrir.Refresh()
            Check 'aplicacao permanece viva apos rejeitar banks.db' (-not $pAbrir.HasExited)
        }

        # O SQLite mantem o arquivo aberto enquanto a aplicacao vive: o hash
        # so' pode ser comparado depois de mata-la (bloqueio leitura/escrita).
        try { $pAbrir.Kill(); $pAbrir.WaitForExit() } catch { }
    }
    catch {
        Check 'fluxo do miOpen sem excecao' $false $_.Exception.Message
    }
    finally {
        $ErrorActionPreference = 'Continue'
        if ($pAbrir -and -not $pAbrir.HasExited) {
            try { $pAbrir.Kill(); $pAbrir.WaitForExit() } catch { }
        }
    }
    if ($hashAbr -and (Test-Path $db)) {
        Check 'banks.db nao foi alterado pelo miOpen' ((Get-Hash $db) -eq $hashAbr)
    }
    else {
        Check 'hash de banks.db comparado antes/depois' $false 'arquivo ausente'
    }

    # ------------------------------------------------ [7] miClose -> fechar
    Write-Banner '[7/9] Fechar Database (miClose) -> encerra a conexao com o arquivo'
    $pFecha = $null
    # Mesma protecao dos passos de interface: excecao tem de virar FALHA.
    $ErrorActionPreference = 'Stop'
    try {
        $pFecha = Start-Process -FilePath $exe -WorkingDirectory $root -PassThru
        Start-Sleep -Seconds 3
        $pFecha.Refresh()
        $mainFecha = $pFecha.MainWindowHandle
        Check 'aplicacao abriu (janela principal)' ($mainFecha -ne [IntPtr]::Zero)
        Check 'miNew-test.db dos passos anteriores disponivel' (Test-Path $minewDb)

        # (a) abre o database de contas: sem conexao nao ha nada para fechar
        $dlgF = Invoke-MenuFileDialog $pFecha 'Abrir Database'
        Check 'dialogo "Abrir" abriu' ($dlgF -ne [IntPtr]::Zero)
        if ($dlgF -ne [IntPtr]::Zero) {
            $editF = Set-FileDialogName $dlgF $minewDb
            Check 'campo de nome do arquivo encontrado' ($editF -ne [IntPtr]::Zero)
            Check 'dialogo fechou ao confirmar' (Wait-DialogClosed $dlgF)
            # a conexao e aberta depois que o dialogo some: espera o aviso
            $avisoF = [IntPtr]::Zero
            for ($t = 0; ($t -lt 12) -and ($avisoF -eq [IntPtr]::Zero); $t++) {
                Start-Sleep -Milliseconds 250
                $avisoF = [UiTest]::FindDialog([uint32]$pFecha.Id)
            }
            Check 'nenhum aviso pendente (conexao aberta)' (
                $avisoF -eq [IntPtr]::Zero) ('hwnd=' + $avisoF)
            Check 'arquivo travado com a conexao aberta' (-not (Test-FileReadable $minewDb))
        }

        # (b) "Fechar Database": conexao encerrada -> o arquivo volta a ser
        # legivel mesmo com a aplicacao viva (e nada de dialogo de erro).
        $idFecha = [UiTest]::MenuId($mainFecha, 'Fechar Database')
        Check 'item de menu "Fechar Database" encontrado' ($idFecha -gt 0) ('id=' + $idFecha)
        if ($idFecha -gt 0) {
            [void][UiTest]::Msg($mainFecha, 0x0111, [IntPtr]$idFecha, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 1000
            Check 'nenhum dialogo apos fechar' (
                [UiTest]::FindDialog([uint32]$pFecha.Id) -eq [IntPtr]::Zero)
            Check 'arquivo liberado (conexao encerrada)' (Test-FileReadable $minewDb)
            $pFecha.Refresh()
            Check 'aplicacao permanece viva ao fechar' (-not $pFecha.HasExited)

            # fechar de novo e' no-op: nao pode quebrar nem abrir dialogo
            [void][UiTest]::Msg($mainFecha, 0x0111, [IntPtr]$idFecha, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 600
            Check 'segundo fechar tambem sem dialogo' (
                [UiTest]::FindDialog([uint32]$pFecha.Id) -eq [IntPtr]::Zero)
            Check 'arquivo continua liberado apos o 2o fechar' (Test-FileReadable $minewDb)
            $pFecha.Refresh()
            Check 'aplicacao permanece viva (sem crash)' (-not $pFecha.HasExited)

            # (c) reabrir depois de fechar: tem de voltar a conectar
            $dlgRe = Invoke-MenuFileDialog $pFecha 'Abrir Database'
            Check 'dialogo "Abrir" reabriu' ($dlgRe -ne [IntPtr]::Zero)
            if ($dlgRe -ne [IntPtr]::Zero) {
                $editRe = Set-FileDialogName $dlgRe $minewDb
                Check 'campo de nome do arquivo encontrado (reabrir)' (
                    $editRe -ne [IntPtr]::Zero)
                Check 'dialogo fechou ao confirmar (reabrir)' (Wait-DialogClosed $dlgRe)
                Start-Sleep -Milliseconds 750
                Check 'arquivo travado de novo (reconectado)' (-not (Test-FileReadable $minewDb))
            }
            $pFecha.Refresh()
            Check 'aplicacao permanece viva ao reconectar' (-not $pFecha.HasExited)
        }

        try { $pFecha.Kill(); $pFecha.WaitForExit() } catch { }
    }
    catch {
        Check 'fluxo do miClose sem excecao' $false $_.Exception.Message
    }
    finally {
        $ErrorActionPreference = 'Continue'
        if ($pFecha -and -not $pFecha.HasExited) {
            try { $pFecha.Kill(); $pFecha.WaitForExit() } catch { }
        }
    }

    # ------------------------------------------------ [8] miGerCon -> tbContas
    Write-Banner '[8/9] Gerenciar Contas (miGerCon) -> abre tbContas e volta'
    $pNav = $null
    # Mesma protecao do passo do miNew: excecao tem de virar FALHA.
    $ErrorActionPreference = 'Stop'
    try {
        $pNav = Start-Process -FilePath $exe -WorkingDirectory $root -PassThru
        Start-Sleep -Seconds 3
        $pNav.Refresh()
        $mainNav = $pNav.MainWindowHandle
        Check 'aplicacao abriu (janela principal)' ($mainNav -ne [IntPtr]::Zero)

        $baseNav = @([UiTest]::Visible($mainNav))
        Check 'estado inicial mapeado' ($baseNav.Count -gt 0) ('janelas=' + $baseNav.Count)

        $idCon = [UiTest]::MenuId($mainNav, 'Gerenciar Contas')
        Check 'item de menu "Gerenciar Contas" encontrado' ($idCon -gt 0) ('id=' + $idCon)

        if ($idCon -gt 0) {
            # Clique no menu: mesmo caminho de um clique real (WM_COMMAND)
            [void][UiTest]::Msg($mainNav, 0x0111, [IntPtr]$idCon, [IntPtr]::Zero)

            # Espera a aba tbContas aparecer (grade, painel e navigator novos)
            $navApos = $baseNav
            for ($t = 0; ($t -lt 20) -and (@($navApos | Where-Object { $baseNav -notcontains $_ }).Count -eq 0); $t++) {
                Start-Sleep -Milliseconds 250
                $navApos = @([UiTest]::Visible($mainNav))
            }
            $novasNav = @($navApos | Where-Object { $baseNav -notcontains $_ })
            Check 'tbContas abriu (novas janelas visiveis)' ($novasNav.Count -gt 0) ('novas=' + $novasNav.Count)

            # O painel inferior (pnContas, altura 50) e onde mora o botao:
            # dele tiramos o HWND (alvo do clique) e o rect (confirma o maior).
            $painelId = [IntPtr]::Zero
            $painelW = 0
            $painelX = 0
            $painelY = 0
            foreach ($n in $novasNav) {
                if ($n -match '(\d+),(\d+) (\d+)x(\d+)$') {
                    $w = [int]$Matches[3]; $h = [int]$Matches[4]
                    if ($h -eq 50 -and $w -gt $painelW) {
                        $painelW = $w
                        $painelX = [int]$Matches[1]
                        $painelY = [int]$Matches[2]
                        if ($n -match 'id=(\d+)') { $painelId = [IntPtr][int64]$Matches[1] }
                    }
                }
            }
            Check 'painel da aba tbContas localizado' ($painelId -ne [IntPtr]::Zero) (
                'largura=' + $painelW + ' pos=' + $painelX + ',' + $painelY)

            if ($painelId -ne [IntPtr]::Zero) {
                # sbtnVoltarContas: Left=200 Top=8 79x30 dentro do painel ->
                # centro em (239,23). O clique vai DIRETO para o HWND do painel
                # (funciona mesmo com o form atras de outra janela).
                $ondeNav = [UiTest]::ClickOn($painelId, 239, 23)
                Check 'clique em sbtnVoltarContas' (-not $ondeNav.StartsWith('FALHA')) $ondeNav

                for ($t = 0; ($t -lt 20) -and (@(Compare-Object $baseNav $navApos).Count -ne 0); $t++) {
                    Start-Sleep -Milliseconds 250
                    $navApos = @([UiTest]::Visible($mainNav))
                }
                $difNav = @(Compare-Object $baseNav $navApos)
                # No falha, o detalhe diz o que ainda difere (=> surgiu, <= sumiu).
                Check 'sbtnVoltarContas devolveu o estado anterior' ($difNav.Count -eq 0) (
                    (($difNav | ForEach-Object { $_.SideIndicator + ' ' + $_.InputObject }) -join ' // '))
            }
        }

        $pNav.Refresh()
        Check 'aplicacao permanece viva (sem crash)' (-not $pNav.HasExited)
    }
    catch {
        Check 'fluxo do miGerCon sem excecao' $false $_.Exception.Message
    }
    finally {
        $ErrorActionPreference = 'Continue'
        if ($pNav -and -not $pNav.HasExited) {
            try { $pNav.Kill(); $pNav.WaitForExit() } catch { }
        }
    }

    # ------------------------------------------------ [9] estado final + WER
    Write-Banner '[9/9] Estado final e log de crashes do Windows'
    if (-not (Test-Path $db)) {
        $null = Invoke-App $RunSeconds   # recria o banco para deixar o ambiente utilizavel
    }
    Check 'banks.db disponivel ao final' (Test-Path $db)

    $crashes = @(Get-AppCrashes $startedAt)
    Check 'nenhum crash registrado no Windows (WER)' ($crashes.Count -eq 0) (
        ($crashes | ForEach-Object { $_.Message -split "`n" | Select-Object -First 1 }) -join ' | '
    )
}
finally {
    # devolve o banco original (preserva dados do usuario)
    if ($dbExisted -and (Test-Path $dbBackup)) {
        Copy-Item -Path $dbBackup -Destination $db -Force
        Remove-Item $dbBackup -Force
    }
    elseif ((-not $dbExisted) -and (Test-Path $dbBackup)) {
        Remove-Item $dbBackup -Force
    }
    if ($dllRenamed -and (Test-Path ($dll + '.bak'))) {
        Rename-Item ($dll + '.bak') 'sqlite3.dll'
    }
    # arquivo de teste do "Novo Database" (nome pertence ao teste)
    if (Test-Path $minewDb) {
        Remove-Item $minewDb -Force -ErrorAction SilentlyContinue
    }
    # arquivo invalido do teste do "Abrir Database"
    if (Test-Path $lixoDb) {
        Remove-Item $lixoDb -Force -ErrorAction SilentlyContinue
    }
}

# ------------------------------------------------------------------
Write-Output ''
Write-Output '============================================================'
$exit = 1
if ($failed -eq 0) {
    $exit = 0
    Write-Output (' RESULTADO: ' + $passed + ' OK, 0 falhas -> EXIT 0')
}
else {
    Write-Output (' RESULTADO: ' + $passed + ' OK, ' + $failed + ' FALHA(S) -> EXIT 1')
}
Write-Output '============================================================'
exit $exit
