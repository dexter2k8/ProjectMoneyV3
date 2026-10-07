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
    [5] Novo Database (miNew) -> cria um .db com as tabelas contas/extratos/saldos e revela a interface
    [6] Abrir Database (miOpen) -> abre o .db valido, rejeita arquivo invalido e revela a interface
    [7] Fechar Database (miClose) -> encerra a conexao com o arquivo e oculta a interface
    [8] Voltar SEM database -> devolve o estado vazio (so o menu); COM
        database aberto -> a interface fica; "Fechar Database" volta a esconder
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
    [DllImport("user32.dll")] static extern uint GetMenuState(IntPtr h, uint id, uint flags);
    [DllImport("user32.dll")] static extern int GetMenuItemCount(IntPtr h);
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
    // Item de menu habilitado? MF_BYCOMMAND=0; MF_GRAYED (0x01) ou
    // MF_DISABLED (0x02) ligado = desativado; item ausente = desativado.
    public static bool MenuItemEnabled(IntPtr main, int id) {
        IntPtr m = GetMenu(main);
        if (m == IntPtr.Zero || id < 0) return false;
        uint st = GetMenuState(m, (uint)id, 0x00000000);
        if (st == 0xFFFFFFFF) return false;
        return (st & 0x03) == 0;
    }
    // [itens de comando, habilitados] do submenu que contem o item "filho"
    // (mesma varredura do MenuId, entao nao depende da posicao fixa no .lfm).
    // Separadores nao sao comandos e ficam fora do total - assim da para
    // conferir que TODOS os itens de um menu mudaram de estado junto.
    public static int[] SubMenuEnabled(IntPtr main, string filho) {
        int comandos = 0, habilitados = 0;
        IntPtr menu = GetMenu(main);
        if (menu != IntPtr.Zero) {
            for (int sub = 0; sub < 16; sub++) {
                IntPtr sm = GetSubMenu(menu, sub);
                if (sm == IntPtr.Zero) continue;
                bool achou = false;
                var sb = new StringBuilder(256);
                for (uint id = 1; id <= 120 && !achou; id++) {
                    sb.Length = 0;
                    GetMenuString(sm, id, sb, 256, false);
                    if (sb.ToString() == filho) achou = true;
                }
                if (!achou) continue;
                int n = GetMenuItemCount(sm);
                for (int i = 0; i < n; i++) {
                    uint st = GetMenuState(sm, (uint)i, 0x00000400);   // MF_BYPOSITION
                    if (st == 0xFFFFFFFF) continue;
                    if ((st & 0x00000800) != 0) continue;             // separador
                    comandos++;
                    if ((st & 0x03) == 0) habilitados++;
                }
                break;
            }
        }
        return new int[] { comandos, habilitados };
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

# Espera a interface aparecer ($true) ou sumir ($false): miNew, miOpen e
# miClose mudam a visibilidade em tempo real, entao o clique precisa de um
# instante para agir. Devolve as janelas visiveis nesse momento.
function Wait-Interface([IntPtr]$Main, [bool]$Esperado) {
    $vis = @([UiTest]::Visible($Main))
    for ($t = 0; ($t -lt 12) -and (($vis.Count -gt 0) -ne $Esperado); $t++) {
        Start-Sleep -Milliseconds 250
        $vis = @([UiTest]::Visible($Main))
    }
    return $vis
}

# Acha o painel inferior com o botao "Voltar" entre as janelas NOVAS (em
# relacao a $Base) e clica nele. tbBancos (pnBancosControl) e tbContas
# (pnContas) tem a mesma geometria: 50px de altura e botao em (200,8) 79x30.
# Painel = janela de altura 50 MAIS BAIXA da tela: o pnHeader tambem tem
# 50px, mas fica no topo (e o pnFooter some nas duas abas).
# Devolve o HWND do painel (Zero quando nao achou).
function Invoke-VoltarContas([IntPtr]$Main, $Base) {
    $atual = @([UiTest]::Visible($Main))
    $novas = @($atual | Where-Object { $Base -notcontains $_ })
    $alvo = [IntPtr]::Zero
    $melhorTop = -1
    $melhorW = -1
    foreach ($n in $novas) {
        if ($n -match '(\d+),(\d+) (\d+)x(\d+)$') {
            $y = [int]$Matches[2]
            $w = [int]$Matches[3]
            $h = [int]$Matches[4]
            if ($h -eq 50 -and (($y -gt $melhorTop) -or (($y -eq $melhorTop) -and ($w -gt $melhorW)))) {
                $melhorTop = $y
                $melhorW = $w
                if ($n -match 'id=(\d+)') { $alvo = [IntPtr][int64]$Matches[1] }
            }
        }
    }
    if ($alvo -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
    # Centro do botao: Left=200 Top=8 79x30 dentro do painel -> (239,23).
    [void][UiTest]::ClickOn($alvo, 239, 23)
    return $alvo
}

# Um ciclo completo do botao "Voltar": tira o estado ATUAL, navega pelo menu
# ($MenuId = miList -> tbBancos ou miGerCon -> tbContas), espera a interface
# aparecer, clica no "Voltar" do painel inferior e espera a aba sumir. Devolve:
#   Inicio = estado capturado antes de navegar (base real desse ciclo)
#   Novas  = janelas novas que a navegacao trouxe (0 = a tela nao abriu)
#   Painel = HWND do painel onde ficou o botao (Zero = nao achou)
#   Pos    = estado estabilizado ja com o clique aplicado (o CALLER avalia)
function Invoke-CicloVoltar([IntPtr]$Main, [int]$MenuId) {
    $inicio = @([UiTest]::Visible($Main))
    [void][UiTest]::Msg($Main, 0x0111, [IntPtr]$MenuId, [IntPtr]::Zero)
    $apos = $inicio
    for ($t = 0; ($t -lt 20) -and (@($apos | Where-Object { $inicio -notcontains $_ }).Count -eq 0); $t++) {
        Start-Sleep -Milliseconds 250
        $apos = @([UiTest]::Visible($Main))
    }
    $novas = @($apos | Where-Object { $inicio -notcontains $_ })
    $painel = Invoke-VoltarContas $Main $inicio
    $pos = $apos
    if ($painel -ne [IntPtr]::Zero) {
        $idPainel = 'id=' + $painel
        for ($t = 0; ($t -lt 20) -and (@($pos | Where-Object { $_ -like ('*' + $idPainel + '*') }).Count -gt 0); $t++) {
            Start-Sleep -Milliseconds 250
            $pos = @([UiTest]::Visible($Main))
        }
    }
    Start-Sleep -Milliseconds 500   # assenta o layout antes de comparar
    [pscustomobject]@{
        Inicio = $inicio
        Novas  = $novas.Count
        Painel = $painel
        Pos    = @([UiTest]::Visible($Main))
    }
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
                # "Ao criar um database, os elementos da pagina aparecem".
                $visNovo = @(Wait-Interface $mainNovo $true)
                Check 'interface revelada apos o miNew' ($visNovo.Count -gt 0) (
                    'janelas=' + $visNovo.Count)
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
            # "Ao abrir um database existente, os elementos da pagina aparecem".
            $visAbr = @(Wait-Interface $mainAbrir $true)
            Check 'interface revelada apos o miOpen' ($visAbr.Count -gt 0) (
                'janelas=' + $visAbr.Count)
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
            # A rejeicao acontece ANTES de mexer na ligacao: a interface
            # continua exatamente como estava (revelada pelo arquivo valido).
            $visRec = @([UiTest]::Visible($mainAbrir))
            Check 'interface permanece apos rejeitar banks.db' ($visRec.Count -gt 0) (
                'janelas=' + $visRec.Count)
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

        # O item so' faz sentido com database aberto: na inicializacao nao
        # existe nada para fechar, entao ele tem de nascer desativado.
        $idFecha = [UiTest]::MenuId($mainFecha, 'Fechar Database')
        Check 'item de menu "Fechar Database" encontrado' ($idFecha -gt 0) ('id=' + $idFecha)
        Check 'miClose desativado na inicializacao (sem database)' (
            ($idFecha -gt 0) -and (-not [UiTest]::MenuItemEnabled($mainFecha, $idFecha)))
        # Mesma regra para TODOS os itens do menu "Transacoes" (do miImport
        # ate' o miGerCon): o localizador e' o primeiro item, sempre ASCII.
        $mmFecha = [UiTest]::SubMenuEnabled($mainFecha, 'Importar OFC/OFX')
        Check 'menu Transacoes desabilitado na inicializacao' (
            ($mmFecha[0] -gt 0) -and ($mmFecha[1] -eq 0)) (
            'comandos=' + $mmFecha[0] + ' habilitados=' + $mmFecha[1])

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
            $visAb = @(Wait-Interface $mainFecha $true)
            Check 'interface revelada ao abrir (miOpen)' ($visAb.Count -gt 0) (
                'janelas=' + $visAb.Count)
            Check 'miClose ativado ao abrir o database' (
                [UiTest]::MenuItemEnabled($mainFecha, $idFecha))
            $mmFecha = [UiTest]::SubMenuEnabled($mainFecha, 'Importar OFC/OFX')
            Check 'menu Transacoes habilitado ao abrir' (
                ($mmFecha[0] -gt 0) -and ($mmFecha[1] -eq $mmFecha[0])) (
                'comandos=' + $mmFecha[0] + ' habilitados=' + $mmFecha[1])
        }

        # (b) "Fechar Database": conexao encerrada -> o arquivo volta a ser
        # legivel mesmo com a aplicacao viva (e nada de dialogo de erro).
        if ($idFecha -gt 0) {
            [void][UiTest]::Msg($mainFecha, 0x0111, [IntPtr]$idFecha, [IntPtr]::Zero)
            Start-Sleep -Milliseconds 1000
            Check 'nenhum dialogo apos fechar' (
                [UiTest]::FindDialog([uint32]$pFecha.Id) -eq [IntPtr]::Zero)
            Check 'arquivo liberado (conexao encerrada)' (Test-FileReadable $minewDb)
            # "Ao fechar o database, os elementos da pagina sao ocultados".
            $visFec = @(Wait-Interface $mainFecha $false)
            Check 'interface oculta apos o Fechar Database' ($visFec.Count -eq 0) (
                'janelas=' + $visFec.Count)
            Check 'miClose desativado apos o fechar' (
                -not [UiTest]::MenuItemEnabled($mainFecha, $idFecha))
            $mmFecha = [UiTest]::SubMenuEnabled($mainFecha, 'Importar OFC/OFX')
            Check 'menu Transacoes desabilitado apos o fechar' (
                ($mmFecha[0] -gt 0) -and ($mmFecha[1] -eq 0)) (
                'comandos=' + $mmFecha[0] + ' habilitados=' + $mmFecha[1])
            $pFecha.Refresh()
            Check 'aplicacao permanece viva ao fechar' (-not $pFecha.HasExited)

            # Segundo clique em "Fechar Database": com o item ja' desativado
            # o comando nem chega ao handler; mesmo que chegue, fechar algo
            # ja' fechado e' no-op (nao pode quebrar nem abrir dialogo).
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
                $visRe = @(Wait-Interface $mainFecha $true)
                Check 'interface revelada apos reconectar' ($visRe.Count -gt 0) (
                    'janelas=' + $visRe.Count)
                Check 'miClose reativado ao reconectar' (
                    [UiTest]::MenuItemEnabled($mainFecha, $idFecha))
                $mmFecha = [UiTest]::SubMenuEnabled($mainFecha, 'Importar OFC/OFX')
                Check 'menu Transacoes habilitado ao reconectar' (
                    ($mmFecha[0] -gt 0) -and ($mmFecha[1] -eq $mmFecha[0])) (
                    'comandos=' + $mmFecha[0] + ' habilitados=' + $mmFecha[1])
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
    Write-Banner '[8/9] Voltar sem database -> estado vazio; com database -> fica'
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
        # A aplicacao abre em branco (só o MainMenu): nao ha nenhuma janela
        # filha visivel no estado inicial.
        Check 'interface oculta na inicializacao (so o menu)' ($baseNav.Count -eq 0) (
            'janelas=' + $baseNav.Count)

        $idCon = [UiTest]::MenuId($mainNav, 'Gerenciar Contas')
        Check 'item de menu "Gerenciar Contas" encontrado' ($idCon -gt 0) ('id=' + $idCon)
        # Sem database o menu "Transacoes" fica TODO desabilitado (inclusive o
        # "Gerenciar Contas"), entao a navegacao de prova passa a ser pela
        # "Lista de Bancos" (menu Arquivo), que continua sempre disponivel.
        $idLista = [UiTest]::MenuId($mainNav, 'Lista de Bancos')
        Check 'item de menu "Lista de Bancos" encontrado' ($idLista -gt 0) ('id=' + $idLista)
        $mmNav = [UiTest]::SubMenuEnabled($mainNav, 'Importar OFC/OFX')
        Check 'menu Transacoes desabilitado sem database' (
            ($mmNav[0] -gt 0) -and ($mmNav[1] -eq 0)) (
            'comandos=' + $mmNav[0] + ' habilitados=' + $mmNav[1])
        Check 'miList habilitado (navegacao disponivel sem database)' (
            [UiTest]::MenuItemEnabled($mainNav, $idLista))

        if (($idCon -gt 0) -and ($idLista -gt 0)) {
            # ---- (a) SEM database em aberto: navegar revela e o "Voltar"
            # devolve o estado vazio. Dois ciclos = da para repetir.
            for ($ciclo = 1; $ciclo -le 2; $ciclo++) {
                $c = Invoke-CicloVoltar $mainNav $idLista
                Check ('navegacao revelou a interface (ciclo ' + $ciclo + ')') (
                    $c.Novas -gt 0) ('novas=' + $c.Novas)
                Check ('painel do Voltar localizado (ciclo ' + $ciclo + ')') (
                    $c.Painel -ne [IntPtr]::Zero)
                $estado = $c.Pos
                for ($t = 0; ($t -lt 20) -and (@(Compare-Object $baseNav $estado).Count -ne 0); $t++) {
                    Start-Sleep -Milliseconds 250
                    $estado = @([UiTest]::Visible($mainNav))
                }
                $dif = @(Compare-Object $baseNav $estado)
                # No falha, o detalhe diz o que ainda difere (=> surgiu, <= sumiu).
                Check ('Voltar sem database devolveu ao estado vazio (ciclo ' + $ciclo + ')') (
                    $dif.Count -eq 0) (
                    (($dif | ForEach-Object { $_.SideIndicator + ' ' + $_.InputObject }) -join ' // '))
            }

            # ---- (b) COM database aberto: o "Voltar" NAO pode esconder nada.
            Check 'miNew-test.db do passo [5] disponivel' (Test-Path $minewDb)
            $dlgAb = Invoke-MenuFileDialog $pNav 'Abrir Database'
            Check 'dialogo "Abrir Database" abriu' ($dlgAb -ne [IntPtr]::Zero)
            if ($dlgAb -ne [IntPtr]::Zero) {
                $editAb = Set-FileDialogName $dlgAb $minewDb
                Check 'campo de nome do arquivo encontrado' ($editAb -ne [IntPtr]::Zero)
                Check 'dialogo fechou ao confirmar' (Wait-DialogClosed $dlgAb)
                # a conexao e aberta depois que o dialogo some: espera o aviso
                $avisoAb = [IntPtr]::Zero
                for ($t = 0; ($t -lt 12) -and ($avisoAb -eq [IntPtr]::Zero); $t++) {
                    Start-Sleep -Milliseconds 250
                    $avisoAb = [UiTest]::FindDialog([uint32]$pNav.Id)
                }
                Check 'database aberto sem aviso' ($avisoAb -eq [IntPtr]::Zero) ('hwnd=' + $avisoAb)
            }

            # Abrir ja levanta a interface: nao precisa navegar para ve-la.
            $visAb8 = @(Wait-Interface $mainNav $true)
            Check 'interface revelada apos o miOpen (sem navegar)' (
                $visAb8.Count -gt 0) ('janelas=' + $visAb8.Count)

            # Com database o menu "Transacoes" volta a funcionar todo - e o
            # "Gerenciar Contas" volta a ser navegavel (e' ele o proximo passo).
            $mmNav = [UiTest]::SubMenuEnabled($mainNav, 'Importar OFC/OFX')
            Check 'menu Transacoes habilitado com database' (
                ($mmNav[0] -gt 0) -and ($mmNav[1] -eq $mmNav[0])) (
                'comandos=' + $mmNav[0] + ' habilitados=' + $mmNav[1])

            $cCom = Invoke-CicloVoltar $mainNav $idCon
            Check 'navegacao revelou a interface (com database)' ($cCom.Novas -gt 0) (
                'novas=' + $cCom.Novas)
            Check 'painel da tbContas localizado (com database)' (
                $cCom.Painel -ne [IntPtr]::Zero)
            # Rodape e "Anterior:" so existem na tela de extratos com a
            # interface inteira levantada: e a prova de que NADA foi escondido.
            $estCom = $cCom.Pos
            for ($t = 0; ($t -lt 20) -and (@($estCom | Where-Object { $_ -like '*"Show Controls"*' }).Count -eq 0); $t++) {
                Start-Sleep -Milliseconds 250
                $estCom = @([UiTest]::Visible($mainNav))
            }
            Check 'Voltar com database mantem a interface visivel' (
                ($estCom.Count -gt 0) -and
                (@($estCom | Where-Object { $_ -like '*"Show Controls"*' }).Count -gt 0) -and
                (@($estCom | Where-Object { $_ -like '*"Anterior:"*' }).Count -gt 0)) (
                'janelas=' + $estCom.Count)

            # ---- (c) "Fechar Database" zera a regra: volta a esconder.
            $idFecha = [UiTest]::MenuId($mainNav, 'Fechar Database')
            Check 'item de menu "Fechar Database" encontrado' ($idFecha -gt 0) ('id=' + $idFecha)
            if ($idFecha -gt 0) {
                [void][UiTest]::Msg($mainNav, 0x0111, [IntPtr]$idFecha, [IntPtr]::Zero)
                Start-Sleep -Milliseconds 800
                Check 'nenhum dialogo ao fechar' (
                    [UiTest]::FindDialog([uint32]$pNav.Id) -eq [IntPtr]::Zero)

                # "Ao fechar o database, os elementos da pagina sao ocultados":
                # daqui pra baixo o ciclo comeca do estado vazio de novo.
                $visFec8 = @(Wait-Interface $mainNav $false)
                Check 'interface oculta apos o Fechar Database' ($visFec8.Count -eq 0) (
                    'janelas=' + $visFec8.Count)

                # E o menu "Transacoes" volta a ficar todo desabilitado.
                $mmNav = [UiTest]::SubMenuEnabled($mainNav, 'Importar OFC/OFX')
                Check 'menu Transacoes desabilitado apos o Fechar Database' (
                    ($mmNav[0] -gt 0) -and ($mmNav[1] -eq 0)) (
                    'comandos=' + $mmNav[0] + ' habilitados=' + $mmNav[1])

                $cFec = Invoke-CicloVoltar $mainNav $idLista
                Check 'navegacao revelou a interface (apos o fechar)' ($cFec.Novas -gt 0) (
                    'novas=' + $cFec.Novas)
                Check 'painel do Voltar localizado (apos o fechar)' (
                    $cFec.Painel -ne [IntPtr]::Zero)
                $estFec = $cFec.Pos
                for ($t = 0; ($t -lt 20) -and (@(Compare-Object $baseNav $estFec).Count -ne 0); $t++) {
                    Start-Sleep -Milliseconds 250
                    $estFec = @([UiTest]::Visible($mainNav))
                }
                $difFec = @(Compare-Object $baseNav $estFec)
                Check 'Voltar apos o Fechar Database volta a esconder' ($difFec.Count -eq 0) (
                    (($difFec | ForEach-Object { $_.SideIndicator + ' ' + $_.InputObject }) -join ' // '))
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
