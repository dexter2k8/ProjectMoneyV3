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
        database aberto -> a interface fica (tbContas e tbSaldos ligadas as
        tabelas); "Fechar Database" volta a esconder
    [9] Importar OFC/OFX (miImport) -> linhas novas entram em "extratos",
        a reimportacao nao duplica e extrato em ANSI vira UTF-8 no arquivo
    [10] "Voltar" da tbContas -> cbAccount e' remontado e reflete a conta
        excluida pela propria tela de gestao enquanto o app estava aberto
    [11] Nenhum crash registrado no log de eventos do Windows (WER)

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
$ofxUtf8 = Join-Path $root 'importa-test.ofx'
$ofxAnsi = Join-Path $root 'importa-test-ansi.ofx'
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
    [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool GetScrollInfo(IntPtr h, int bar, ref SCROLLINFO s);

    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int left, top, right, bottom; }
    [StructLayout(LayoutKind.Sequential)]
    public struct SCROLLINFO {
        public int cbSize; public uint fMask;
        public int nMin; public int nMax; public uint nPage; public int nPos; public int nTrackPos;
    }

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
    // Estado do menu de PRIMEIRO nivel que contem o item "filho":
    // 1 = habilitado, 0 = desabilitado, -1 = nao achou. A localizacao e' pela
    // legenda de um filho conhecido (mesmo truque do MenuId), entao nao
    // depende do acento de "Transacoes" nem da posicao fixa no .lfm.
    // No menu raiz o item e' indicado por POSICAO (MF_BYPOSITION=0x400) e o
    // GetMenuState devolve a entrada de menu - que e' o que decide se o
    // submenu abre (o estado dos filhos NAO herda o do pai).
    public static int MenuTopState(IntPtr main, string filho) {
        IntPtr menu = GetMenu(main);
        if (menu == IntPtr.Zero) return -1;
        for (int pos = 0; pos < 16; pos++) {
            IntPtr sm = GetSubMenu(menu, pos);
            if (sm == IntPtr.Zero) continue;
            bool achou = false;
            var sb = new StringBuilder(256);
            for (uint id = 1; id <= 120 && !achou; id++) {
                sb.Length = 0;
                GetMenuString(sm, id, sb, 256, false);
                if (sb.ToString() == filho) achou = true;
            }
            if (!achou) continue;
            uint st = GetMenuState(menu, (uint)pos, 0x00000400);   // MF_BYPOSITION
            if (st == 0xFFFFFFFF) return -1;
            return (st & 0x03) == 0 ? 1 : 0;
        }
        return -1;
    }
    // nMax da barra de rolagem vertical (-1 = janela sem barra). No TDBGrid o
    // LCL usa a barra NATIVA do Windows, entao o nMax e' o total de linhas
    // que a grade mostra - e' assim que a suite enxerga a tabela ligada a
    // grade (sem depender do desenho das celulas, que nao viram janela).
    public static int VScrollMax(IntPtr h) {
        SCROLLINFO s;
        s.cbSize = Marshal.SizeOf(typeof(SCROLLINFO));
        s.fMask = 0x0017;             // SIF_RANGE|SIF_PAGE|SIF_POS|SIF_TRACKPOS
        s.nMin = 0; s.nMax = 0; s.nPage = 0; s.nPos = 0; s.nTrackPos = 0;
        if (!GetScrollInfo(h, 1, ref s)) return -1;   // SB_VERT
        return s.nMax;
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
    // em primeiro plano (na suite o form fica atras da IDE e WindowFromPoint
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

# sqlite3.dll x64 (a mesma que a aplicacao carrega em runtime) para gravar
# linhas de teste em "saldos" ANTES de abrir o database: e' o que a grade da
# tbSaldos passa a mostrar, e portanto a prova observavel do vinculo.
if (-not ('Sq' -as [type])) {
    $env:PATH = $root + ';' + $env:PATH
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Sq {
    public static int Contagem = -1;
    public delegate int Cb(IntPtr arg, int ncol, IntPtr vals, IntPtr names);
    [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    static extern int sqlite3_open(string f, out IntPtr db);
    [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
    static extern int sqlite3_exec(IntPtr db, string sql, Cb cb, IntPtr arg, out IntPtr err);
    [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
    static extern int sqlite3_close(IntPtr db);
    // Primeira coluna da primeira linha de retorno (o SELECT COUNT(*)).
    static int OnRow(IntPtr a, int n, IntPtr v, IntPtr names) {
        if (n > 0) {
            int x;
            int.TryParse(Marshal.PtrToStringAnsi(Marshal.ReadIntPtr(v)), out x);
            Contagem = x;
        }
        return 0;
    }
    static int Rodar(string file, string sql) {
        IntPtr db, err;
        if (sqlite3_open(file, out db) != 0) return -1;
        try {
            Contagem = -1;
            if (sqlite3_exec(db, sql, OnRow, IntPtr.Zero, out err) != 0) return -2;
            return 0;
        } finally { sqlite3_close(db); }
    }
    // 1 = a sentenca rodou (INSERT commitado, o arquivo e' fechado no fim);
    // <= 0 = falhou. Consultar devolve o valor lido (ou -1/-2 em erro).
    public static int Executar(string file, string sql) { return Rodar(file, sql) == 0 ? 1 : -1; }
    public static int Consultar(string file, string sql) { return Rodar(file, sql) == 0 ? Contagem : -1; }
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

# Le do .lfm o bloco "object <Objeto>" ate' o "end" da MESMA indentacao (o
# .lfm nao aceita comentario, entao o fim do bloco e' o primeiro end' no nivel
# do proprio objeto). Devolve '' quando o objeto nao esta' no arquivo.
function Get-LfmBlock([string]$Texto, [string]$Objeto) {
    if ($Texto -eq '') { return '' }
    $padrao = '(?ms)^([ \t]*)object ' + [regex]::Escape($Objeto) +
        '.*?^\1end[ \t]*\r?$'
    $m = [regex]::Match($Texto, $padrao)
    if (-not $m.Success) { return '' }
    return $m.Value
}

# Grava N linhas em "saldos" (a tabela da tbSaldos) direto no arquivo: e' o
# que a grade da pagina tem de mostrar depois de abrir o database. O arquivo
# tem de estar LIVRE (conexao encerrada). Devolve as linhas gravadas (0 =
# arquivo ausente/falhou - quem reporta e' o chamador).
function Add-SaldosRows([string]$Path, [int]$Count) {
    if (-not (Test-Path $Path)) { return 0 }
    # O INSERT com CTE recursivo grava tudo em uma unica sentenca.
    $sql = 'WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x+1 FROM c' +
        ' WHERE x < ' + $Count + ') INSERT INTO "saldos"' +
        " (account_id, balance, enddate) SELECT 1, x * 1.5, '2026-01-31' FROM c;"
    if ([Sq]::Executar($Path, $sql) -le 0) { return 0 }
    # Le do proprio arquivo o numero de linhas que ficou: devolver o pedido sem
    # conferir esconderia um INSERT que nao deu certo.
    return [Sq]::Consultar($Path, 'SELECT COUNT(*) FROM saldos;')
}

# Grava N linhas em "extratos" (a tabela da grade da tela de extratos) datadas
# do ano informado: e' o que o combo de ano da tela tem de listar e o que o
# filtro do ano selecionado tem de deixar passar. O arquivo tem de estar LIVRE
# (conexao encerrada). Devolve as linhas gravadas DESSE ano (0 = arquivo
# ausente/falhou - quem reporta e' o chamador).
function Add-ExtratosRows([string]$Path, [int]$Count, [string]$Ano) {
    if (-not (Test-Path $Path)) { return 0 }
    # Mesmo truque do INSERT com CTE recursivo: uma sentenca so'.
    $sql = 'WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x+1 FROM c' +
        ' WHERE x < ' + $Count + ') INSERT INTO "extratos"' +
        " (account_id, trntype, dtposted, trnamt, memo)" +
        " SELECT 1, 'DEBIT', '" + $Ano + "-06-15', -x * 1.5, 'linha ' || x FROM c;"
    if ([Sq]::Executar($Path, $sql) -le 0) { return 0 }
    # Conta com a MESMA expressao do ano que a aplicacao usa no filtro: se o
    # texto gravado nao casar com substr(dtposted,1,4), a checagem falha aqui.
    return [Sq]::Consultar($Path,
        "SELECT COUNT(*) FROM extratos WHERE substr(dtposted, 1, 4) = '" +
        $Ano + "';")
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

# Importa um arquivo de extrato pelo menu "Importar OFC/OFX" e devolve o
# dialogo de resultado (MessageDlg) aberto, ou Zero quando algo travou no
# meio do caminho. O chamador fecha o dialogo com WM_CLOSE. A sequencia e'
# a mesma de miOpen: menu -> dialogo de arquivo -> nome -> esperar fechar ->
# esperar o aviso aparecer (nao pode ser lido junto com o dialogo de arquivo).
function Invoke-Importar($Process, [string]$Arquivo) {
    $dlg = Invoke-MenuFileDialog $Process 'Importar OFC/OFX'
    if ($dlg -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
    $edit = Set-FileDialogName $dlg $Arquivo
    if ($edit -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
    if (-not (Wait-DialogClosed $dlg)) { return [IntPtr]::Zero }
    Wait-Warning ([uint32]$Process.Id)
}

# Acha um combo da tela de extratos pela LARGURA (cbAccount = 212, cbYear =
# 100): os dois sao LCLComboBox e o id do controle muda a cada execucao, entao
# a geometria e' o unico jeito estavel de separar um do outro. A tela tem de
# estar revelada: sem database e fora da pagina de extratos os dois somem.
# Devolve o HWND (Zero quando nao achou).
function Find-ComboHwnd([IntPtr]$Main, [int]$Largura) {
    foreach ($w in @([UiTest]::Visible($Main))) {
        if (($w -match '^LCLComboBox') -and
            ($w -match (' ' + $Largura + 'x[0-9]+$')) -and
            ($w -match 'id=(\d+)')) {
            return [IntPtr][int64]$Matches[1]
        }
    }
    return [IntPtr]::Zero
}

# Quantas linhas o combo tem (CB_GETCOUNT = 0x0146). Devolve -1 quando o combo
# nao foi achado: assim "0 itens" nunca passa por falta de HWND.
function Get-ComboCount([IntPtr]$Combo) {
    if ($Combo -eq [IntPtr]::Zero) { return -1 }
    [int64][UiTest]::Msg($Combo, 0x0146, [IntPtr]::Zero, [IntPtr]::Zero)
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
# ($MenuId = miList -> tbBancos, miGerCon -> tbContas, miGerSal -> tbSaldos),
# espera a interface aparecer, clica no "Voltar" do painel inferior e espera
# a aba sumir. Devolve:
#   Inicio = estado capturado antes de navegar (base real desse ciclo)
#   Novas  = janelas novas que a navegacao trouxe (0 = a tela nao abriu)
#   Painel = HWND do painel onde ficou o botao (Zero = nao achou)
#   Scroll = nMax da barra da grade nova (0 = a tabela esta' vazia)
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
    # Linhas da grade ATIVA - lido AQUI, antes do clique no "Voltar": o
    # Invoke-VoltarContas ja clica no botao, e depois disso a tela ja' voltou
    # para o estado inicial (a grade nova nem esta' mais na tela). O TDBGrid
    # usa a barra nativa do Windows, entao o nMax e' quantas linhas a grade
    # mostra. Varre TODAS as janelas visiveis (e nao so as "novas"): so a
    # grade da pagina ativa fica visivel, e comparar por string pegaria a
    # grade nova como "velha" se o Windows reutilizasse o handle da grade da
    # pagina anterior (mesma classe/texto/geometria = mesma linha). Rele a
    # tela ate' 2s enquanto ainda for zero: ou a tabela e' vazia de verdade,
    # ou a barra ainda nao foi montada.
    $scroll = 0
    $mapa = @()
    for ($t = 0; ($t -lt 8) -and ($scroll -le 0); $t++) {
        if ($t -gt 0) { Start-Sleep -Milliseconds 250 }
        $scroll = 0
        $mapa = @()
        foreach ($n in @([UiTest]::Visible($Main))) {
            $nv = -99
            if ($n -match 'id=(\d+)') {
                $nv = [UiTest]::VScrollMax([IntPtr][int64]$Matches[1])
            }
            $mapa += ('      [' + $nv + '] ' + $n)
            if ($nv -gt $scroll) { $scroll = $nv }
        }
    }
    # Aqui sim o "Voltar": a funcao localiza o painel E clica no botao.
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
        Scroll = $scroll
        ScrollMap = $mapa
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
    Write-Banner '[1/11] Compilacao e lint (lazbuild -B)'
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

    # O vinculo da grade e do navigator com "extratos" nao tem efeito visual
    # que a suite consiga ler (gridTrans nasce oculto e o DBNavTrans esta' sem
    # botoes), entao a prova fica no .lfm: e' ali que o DataSource esta'
    # escrito - e um .lfm errado compila igual.
    $lfmTexto = ''
    $lfmPath = Join-Path $root 'unitmoney.lfm'
    if (Test-Path $lfmPath) { $lfmTexto = [IO.File]::ReadAllText($lfmPath) }
    Check 'unitmoney.lfm lido' ($lfmTexto -ne '') $lfmPath
    $blocoGrid = Get-LfmBlock $lfmTexto 'gridTrans: TDBGrid'
    $blocoNav  = Get-LfmBlock $lfmTexto 'DBNavTrans: TDBNavigator'
    $detGrid = 'DataSource ausente no bloco gridTrans'
    if ($blocoGrid -eq '') { $detGrid = 'bloco gridTrans nao encontrado no .lfm' }
    $detNav = 'DataSource ausente no bloco DBNavTrans'
    if ($blocoNav -eq '') { $detNav = 'bloco DBNavTrans nao encontrado no .lfm' }
    Check 'gridTrans ligado a DataSourceExtratos (.lfm)' (
        $blocoGrid.Contains('DataSource = DataSourceExtratos')) $detGrid
    Check 'DBNavTrans ligado a DataSourceExtratos (.lfm)' (
        $blocoNav.Contains('DataSource = DataSourceExtratos')) $detNav
    Check 'combos de filtro com OnChange (.lfm)' (
        ($lfmTexto -match 'OnChange = cbAccountChange') -and
        ($lfmTexto -match 'OnChange = cbYearChange')) (
        'cbAccountChange e/ou cbYearChange ausentes')

    # ------------------------------------------------ [2] cria o banco
    Write-Banner '[2/11] Sem banks.db -> deve criar o arquivo'
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
    Write-Banner '[3/11] Com banks.db -> nao deve recriar nem alterar'
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
    Write-Banner '[4/11] Sem sqlite3.dll -> erro deve ser tratado'
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
    Write-Banner '[5/11] Novo Database (miNew) -> cria .db com as 3 tabelas'
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
    Write-Banner '[6/11] Abrir Database (miOpen) -> abre o valido e rejeita o invalido'
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
            # Combos de filtro da tela de extratos: "contas" e "extratos" estao
            # vazias neste ponto (database recem-criado no passo [5]), entao e'
            # aqui que a regra do "ano atual" e a lista de contas aparecem.
            $cbConta6 = Find-ComboHwnd $mainAbrir 212
            $cbAno6   = Find-ComboHwnd $mainAbrir 100
            $nConta6  = Get-ComboCount $cbConta6
            $nAno6    = Get-ComboCount $cbAno6
            Check 'cbYear lista o ano atual ("extratos" vazio)' ($nAno6 -eq 1) (
                'combo=' + $cbAno6 + ' itens=' + $nAno6)
            Check 'cbAccount sem itens ("contas" vazia)' ($nConta6 -eq 0) (
                'combo=' + $cbConta6 + ' itens=' + $nConta6)
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
    Write-Banner '[7/11] Fechar Database (miClose) -> encerra a conexao com o arquivo'
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
        # Mesma regra para o menu "Transacoes" INTEIRO (e nao item a item):
        # desativando a entrada de primeiro nivel o submenu nem abre. O
        # localizador e' um filho conhecido, sempre ASCII.
        $mmFecha = [UiTest]::MenuTopState($mainFecha, 'Importar OFC/OFX')
        Check 'menu Transacoes desabilitado na inicializacao' ($mmFecha -eq 0) (
            'estado=' + $mmFecha)

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
            $mmFecha = [UiTest]::MenuTopState($mainFecha, 'Importar OFC/OFX')
            Check 'menu Transacoes habilitado ao abrir' ($mmFecha -eq 1) (
                'estado=' + $mmFecha)
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
            $mmFecha = [UiTest]::MenuTopState($mainFecha, 'Importar OFC/OFX')
            Check 'menu Transacoes desabilitado apos o fechar' ($mmFecha -eq 0) (
                'estado=' + $mmFecha)
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
                $mmFecha = [UiTest]::MenuTopState($mainFecha, 'Importar OFC/OFX')
                Check 'menu Transacoes habilitado ao reconectar' ($mmFecha -eq 1) (
                    'estado=' + $mmFecha)
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

    # ---------------- [8] miList/miGerCon/miGerSal -> tbBancos/tbContas/tbSaldos
    Write-Banner '[8/11] Voltar sem database -> estado vazio; com database -> fica'
    $pNav = $null
    # Mesma protecao do passo do miNew: excecao tem de virar FALHA.
    $ErrorActionPreference = 'Stop'
    try {
        # Linhas de teste em "saldos" ANTES de abrir (o arquivo tem de estar
        # livre): a grade da tbSaldos so' prova que esta' ligada a tabela se
        # houver linhas para mostrar - 150, para passar do que cabe na tela.
        $nLinhas = Add-SaldosRows $minewDb 150
        Check 'linhas de teste gravadas em "saldos" (passo [8])' ($nLinhas -eq 150) (
            'linhas=' + $nLinhas)
        # Mesma prova para a tela de extratos: dois anos em "extratos" (75 +
        # 75) para o combo de ano listar os dois e o filtro abrir com o ano
        # mais antigo selecionado.
        $n2025 = Add-ExtratosRows $minewDb 75 '2025'
        $n2026 = Add-ExtratosRows $minewDb 75 '2026'
        Check 'linhas de teste gravadas em "extratos" (passo [8])' (
            ($n2025 -eq 75) -and ($n2026 -eq 75)) (
            '2025=' + $n2025 + ' 2026=' + $n2026)

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
        $mmNav = [UiTest]::MenuTopState($mainNav, 'Importar OFC/OFX')
        Check 'menu Transacoes desabilitado sem database' ($mmNav -eq 0) (
            'estado=' + $mmNav)
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

            # 150 linhas em "extratos" (75 de 2025 + 75 de 2026): o combo de
            # ano tem de listar as duas e comecar na mais antiga - e' ele que
            # alimenta o filtro da query da grade. "contas" segue vazia, entao
            # o combo de conta continua sem nenhuma linha.
            $cbAno8   = Find-ComboHwnd $mainNav 100
            $cbConta8 = Find-ComboHwnd $mainNav 212
            $nAno8    = Get-ComboCount $cbAno8
            $nConta8  = Get-ComboCount $cbConta8
            $selAno8  = -1
            if ($cbAno8 -ne [IntPtr]::Zero) {
                $selAno8 = [int64][UiTest]::Msg($cbAno8, 0x0147,  # CB_GETCURSEL
                    [IntPtr]::Zero, [IntPtr]::Zero)
            }
            Check 'cbYear lista os 2 anos de "extratos"' ($nAno8 -eq 2) (
                'combo=' + $cbAno8 + ' itens=' + $nAno8)
            Check 'cbYear comeca no ano mais antigo (selecao 0)' ($selAno8 -eq 0) (
                'combo=' + $cbAno8 + ' selecao=' + $selAno8)
            Check 'cbAccount sem itens ("contas" vazia, passo [8])' (
                $nConta8 -eq 0) ('combo=' + $cbConta8 + ' itens=' + $nConta8)

            # Com database o menu "Transacoes" volta a funcionar todo - e o
            # "Gerenciar Contas" volta a ser navegavel (e' ele o proximo passo).
            $mmNav = [UiTest]::MenuTopState($mainNav, 'Importar OFC/OFX')
            Check 'menu Transacoes habilitado com database' ($mmNav -eq 1) (
                'estado=' + $mmNav)

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

            # Controle do teste de vinculo: a suite nao grava conta nenhuma,
            # entao a grade da tbContas fica com scroll 0 - prova de que o
            # numero so' fica positivo quando a tabela TEM linhas.
            Check 'grade da tbContas sem linhas (tabela vazia)' ($cCom.Scroll -eq 0) (
                'scroll=' + $cCom.Scroll)

            # ---- (c) tbSaldos: mesma navegacao/voltar da tbContas, e a grade
            # mostra as 150 linhas injetadas em "saldos" - e' o vinculo dela
            # com a tabela (painel, navigator e "Voltar" junto).
            $idSal = [UiTest]::MenuId($mainNav, 'Gerenciar Saldos')
            Check 'item de menu "Gerenciar Saldos" encontrado' ($idSal -gt 0) (
                'id=' + $idSal)
            if ($idSal -gt 0) {
                $cSal = Invoke-CicloVoltar $mainNav $idSal
                Check 'navegacao revelou a interface (tbSaldos)' ($cSal.Novas -gt 0) (
                    'novas=' + $cSal.Novas)
                Check 'painel do Voltar localizado (tbSaldos)' (
                    $cSal.Painel -ne [IntPtr]::Zero)
                Check 'grade da tbSaldos mostra as linhas de "saldos"' (
                    $cSal.Scroll -gt 0) ('scroll=' + $cSal.Scroll)
                if ($cSal.Scroll -le 0) {
                    # Falhou: mostra o que a tela mostrava NO MOMENTO da leitura
                    # (janela visivel + nMax do scroll de cada uma).
                    Write-Output '   [depuracao] varredura da pagina da tbSaldos:'
                    foreach ($w in $cSal.ScrollMap) { Write-Output $w }
                }
                # Mesma regra da tbContas: com database o "Voltar" devolve para
                # a pagina de mes e a interface continua de pe'.
                $estSal = $cSal.Pos
                for ($t = 0; ($t -lt 20) -and (@($estSal | Where-Object { $_ -like '*"Show Controls"*' }).Count -eq 0); $t++) {
                    Start-Sleep -Milliseconds 250
                    $estSal = @([UiTest]::Visible($mainNav))
                }
                Check 'Voltar da tbSaldos mantem a interface visivel' (
                    ($estSal.Count -gt 0) -and
                    (@($estSal | Where-Object { $_ -like '*"Show Controls"*' }).Count -gt 0)) (
                    'janelas=' + $estSal.Count)
            }

            # ---- (d) "Fechar Database" zera a regra: volta a esconder.
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
                $mmNav = [UiTest]::MenuTopState($mainNav, 'Importar OFC/OFX')
                Check 'menu Transacoes desabilitado apos o Fechar Database' ($mmNav -eq 0) (
                    'estado=' + $mmNav)

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

    # ------------------------------------ [9] importar OFC/OFX (miImport)
    Write-Banner '[9/11] Importar OFC/OFX -> linhas novas entram em "extratos"'
    $pImp  = $null
    $antes = 0
    # Mesma protecao dos demais passos: excecao tem de virar FALHA.
    $ErrorActionPreference = 'Stop'
    try {
        # O miImport grava na conta escolhida no combo e "extratos" exige
        # account_id NOT NULL: o passo [8] deixou "contas" vazia, entao a
        # conta destino e' criada aqui - com o arquivo LIVRE (nenhuma conexao
        # aberta em cima dele, o app ainda nem subiu).
        $sqlConta = 'INSERT INTO contas (acctid, accttype, bankid, branchid,' +
            ' description) VALUES (''4321'', ''CHECKING'', 1, ''0001'',' +
            ' ''Conta de teste'');'
        Check 'conta de destino gravada (passo [9])' (
            [Sq]::Executar($minewDb, $sqlConta) -eq 1)
        $antes = [Sq]::Consultar($minewDb, 'SELECT COUNT(*) FROM extratos;')
        Check 'extratos lido antes da importacao' ($antes -eq 150) ('qtd=' + $antes)

        # Extrato OFX 1.x em SGML (tag sem fechamento, valor terminando na
        # quebra de linha) com texto UTF-8 mas cabecalho afirmando CHARSET:1252:
        # a leitura decide pelo CONTEUDO, nao pelo que o arquivo declara. Os
        # registros cobrem o ano que ja existe (2025), um ano novo (2024), data
        # com hora+fuso (sobram os 8 primeiros digitos), um registro fechado no
        # estilo XML e um bloco incompleto (sem TRNAMT: nao pode virar linha).
        $memoAcento = 'Mercado S' + [char]0x00E3 + 'o Jo' + [char]0x00E3 + 'o'
        $ofxLinhas = @(
            'OFXHEADER:100'
            'DATA:20240101000000'
            'VERSION:102'
            'SECURITY:NONE'
            'ENCODING:USASCII'
            'CHARSET:1252'
            '<OFX>'
            '<BANKMSGSRSV1>'
            '<STMTTRNRS>'
            '<STMTRS>'
            '<CURDEF>BRL'
            '<BANKACCTFROM>'
            '<BANKID>001'
            '<ACCTID>4321'
            '<ACCTTYPE>CHECKING'
            '</BANKACCTFROM>'
            '<BANKTRANLIST>'
            '<DTSTART>20240101'
            '<DTEND>20261231'
            '<STMTTRN>'
            '<TRNTYPE>DEBIT'
            '<DTPOSTED>20240115'
            '<TRNAMT>-123.45'
            '<FITID>f1'
            '<CHECKNUM>101'
            '<MEMO>Pagamento de luz'
            '</STMTTRN>'
            '<STMTTRN>'
            '<TRNTYPE>CREDIT'
            '<DTPOSTED>20240220'
            '<TRNAMT>1500.00'
            '<FITID>f2'
            '<MEMO>Salario de fevereiro'
            '</STMTTRN>'
            '<STMTTRN>'
            '<TRNTYPE>DEBIT'
            '<DTPOSTED>20250305120000.000[-5:EST]'
            '<TRNAMT>-9.90'
            '<FITID>f3'
            '<MEMO>' + $memoAcento
            '</STMTTRN>'
            '<STMTTRN>'
            '<TRNTYPE>DEBIT</TRNTYPE>'
            '<DTPOSTED>20250410</DTPOSTED>'
            '<TRNAMT>-55.00</TRNAMT>'
            '<MEMO>Registro em formato XML</MEMO>'
            '</STMTTRN>'
            '<STMTTRN>'
            '<TRNTYPE>DEBIT'
            '<DTPOSTED>20250520'
            '<MEMO>blocos sem valor nao viram linha'
            '</STMTTRN>'
            '</BANKTRANLIST>'
            '</STMTRS>'
            '</STMTTRNRS>'
            '</BANKMSGSRSV1>'
            '</OFX>'
        )
        [IO.File]::WriteAllText($ofxUtf8, ($ofxLinhas -join "`r`n") + "`r`n")

        # O MESMO extrato, mas em Windows-1252 puro (o jeito antigo dos
        # bancos): o byte do acento quebra o UTF-8 e o parser tem de
        # converte-lo para gravar UTF-8 no banco.
        $memoAnsi = 'Padaria S' + [char]0x00E3 + 'o Jos' + [char]0x00E9
        $ofxAnsiLinhas = @(
            'OFXHEADER:100'
            'DATA:20240101000000'
            'VERSION:102'
            'SECURITY:NONE'
            'ENCODING:USASCII'
            'CHARSET:1252'
            '<OFX>'
            '<BANKMSGSRSV1>'
            '<STMTTRNRS>'
            '<STMTRS>'
            '<CURDEF>BRL'
            '<BANKACCTFROM>'
            '<BANKID>001'
            '<ACCTID>4321'
            '<ACCTTYPE>CHECKING'
            '</BANKACCTFROM>'
            '<BANKTRANLIST>'
            '<STMTTRN>'
            '<TRNTYPE>DEBIT'
            '<DTPOSTED>20240310'
            '<TRNAMT>-12.34'
            '<MEMO>' + $memoAnsi
            '</STMTTRN>'
            '</BANKTRANLIST>'
            '</STMTRS>'
            '</STMTTRNRS>'
            '</BANKMSGSRSV1>'
            '</OFX>'
        )
        $enc1252 = [Text.Encoding]::GetEncoding(1252)
        [IO.File]::WriteAllBytes($ofxAnsi,
            $enc1252.GetBytes(($ofxAnsiLinhas -join "`r`n") + "`r`n"))
        Check 'fixture OFX (UTF-8 e ANSI) gravada' (
            (Test-Path $ofxUtf8) -and (Test-Path $ofxAnsi))

        # Abre o database: e' o miOpen que carrega as combos de filtro.
        $pImp = Start-Process -FilePath $exe -WorkingDirectory $root -PassThru
        Start-Sleep -Seconds 3
        $pImp.Refresh()
        $mainImp = $pImp.MainWindowHandle
        Check 'aplicacao abriu para o passo [9]' ($mainImp -ne [IntPtr]::Zero)
        $dlgOpenImp = Invoke-MenuFileDialog $pImp 'Abrir Database'
        Check 'dialogo "Abrir" abriu (passo [9])' ($dlgOpenImp -ne [IntPtr]::Zero)
        if ($dlgOpenImp -ne [IntPtr]::Zero) {
            $editOpenImp = Set-FileDialogName $dlgOpenImp $minewDb
            Check 'campo de nome do arquivo encontrado (passo [9])' (
                $editOpenImp -ne [IntPtr]::Zero)
            Check 'dialogo "Abrir" fechou ao confirmar (passo [9])' (
                Wait-DialogClosed $dlgOpenImp)
            # A conexao abre depois que o dialogo some: um aviso aqui ficaria
            # pendente e enganaria o Invoke-Importar (mesmo #32770).
            $avisoImp = [IntPtr]::Zero
            for ($t = 0; ($t -lt 12) -and ($avisoImp -eq [IntPtr]::Zero); $t++) {
                Start-Sleep -Milliseconds 250
                $avisoImp = [UiTest]::FindDialog([uint32]$pImp.Id)
            }
            Check 'nenhum aviso ao abrir o database (passo [9])' (
                $avisoImp -eq [IntPtr]::Zero) ('hwnd=' + $avisoImp)
        }
        $visImp = @(Wait-Interface $mainImp $true)
        Check 'interface revelada para importar' ($visImp.Count -gt 0) (
            'janelas=' + $visImp.Count)

        # Estado inicial das combos: 1 conta (a criada acima) e os 2 anos que
        # ja estao em "extratos" (2025 e 2026, do passo [8]).
        $cbContaA = Find-ComboHwnd $mainImp 212
        $cbAnoA   = Find-ComboHwnd $mainImp 100
        $nContaA  = Get-ComboCount $cbContaA
        $nAnoA    = Get-ComboCount $cbAnoA
        Check 'cbAccount lista a conta de destino' ($nContaA -eq 1) (
            'combo=' + $cbContaA + ' itens=' + $nContaA)
        Check 'cbYear lista 2 anos antes da importacao' ($nAnoA -eq 2) (
            'combo=' + $cbAnoA + ' itens=' + $nAnoA)

        # 1a importacao: 4 registros validos entram, o bloco sem TRNAMT nao.
        $dlgImp1 = Invoke-Importar $pImp $ofxUtf8
        Check 'dialogo de resultado da 1a importacao' ($dlgImp1 -ne [IntPtr]::Zero)
        if ($dlgImp1 -ne [IntPtr]::Zero) {
            [void][UiTest]::PostMessage($dlgImp1, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
            Check 'resultado da 1a importacao dispensado' (
                Wait-DialogClosed $dlgImp1)
        }
        $pImp.Refresh()
        Check 'aplicacao viva apos a 1a importacao' (-not $pImp.HasExited)

        # Depois do resultado as combos ja foram recarregadas: 2024 entrou na
        # lista (4a linha da fixture) e a selecao tem de voltar para o ano que
        # o usuario tinha escolhido (2025 = indice 1 depois que 2024 entra).
        $cbContaB = Find-ComboHwnd $mainImp 212
        $cbAnoB   = Find-ComboHwnd $mainImp 100
        $nContaB  = Get-ComboCount $cbContaB
        $nAnoB    = Get-ComboCount $cbAnoB
        $selAnoB  = [int64][UiTest]::Msg($cbAnoB, 0x0147, [IntPtr]::Zero, [IntPtr]::Zero)
        Check 'cbAccount continua com a conta de destino' ($nContaB -eq 1) (
            'itens=' + $nContaB)
        Check 'cbYear passou a listar 3 anos (2024 veio do arquivo)' (
            $nAnoB -eq 3) ('itens=' + $nAnoB)
        Check 'cbYear manteve a selecao do usuario (indice 1 = 2025)' (
            $selAnoB -eq 1) ('sel=' + $selAnoB)

        # Reimportar o MESMO arquivo nao pode duplicar: as 4 linhas ja existem.
        $dlgImp2 = Invoke-Importar $pImp $ofxUtf8
        Check 'dialogo de resultado da reimportacao' ($dlgImp2 -ne [IntPtr]::Zero)
        if ($dlgImp2 -ne [IntPtr]::Zero) {
            [void][UiTest]::PostMessage($dlgImp2, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
            Check 'resultado da reimportacao dispensado' (
                Wait-DialogClosed $dlgImp2)
        }
        $cbAnoC = Find-ComboHwnd $mainImp 100
        Check 'cbYear inalterado apos a reimportacao' (
            (Get-ComboCount $cbAnoC) -eq 3) ('itens=' + (Get-ComboCount $cbAnoC))

        # Extrato em ANSI: o parser tem de converte-los para UTF-8. O ano e'
        # 2024 (ja existia), entao a lista de anos nao muda.
        $dlgImp3 = Invoke-Importar $pImp $ofxAnsi
        Check 'dialogo de resultado da importacao ANSI' ($dlgImp3 -ne [IntPtr]::Zero)
        if ($dlgImp3 -ne [IntPtr]::Zero) {
            [void][UiTest]::PostMessage($dlgImp3, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
            Check 'resultado da importacao ANSI dispensado' (
                Wait-DialogClosed $dlgImp3)
        }
        $cbAnoD = Find-ComboHwnd $mainImp 100
        Check 'cbYear segue com 3 anos (2024 ja existia no arquivo)' (
            (Get-ComboCount $cbAnoD) -eq 3) ('itens=' + (Get-ComboCount $cbAnoD))
        $pImp.Refresh()
        Check 'aplicacao viva ao fim das importacoes' (-not $pImp.HasExited)
    }
    catch {
        Check 'fluxo do miImport sem excecao' $false $_.Exception.Message
    }
    finally {
        $ErrorActionPreference = 'Continue'
        if ($pImp -and -not $pImp.HasExited) {
            try { $pImp.Kill(); $pImp.WaitForExit() } catch { }
        }
    }

    # Conferencia no arquivo: so' funciona com a conexao encerrada (app morta).
    # 5 linhas novas no total (4 do UTF-8 + 1 do ANSI) e nada duplicado.
    if (-not (Test-Path $minewDb)) {
        Check 'miNew-test.db disponivel para a conferencia final' $false (
            'arquivo ausente')
    }
    else {
        $depois = [Sq]::Consultar($minewDb, 'SELECT COUNT(*) FROM extratos;')
        Check 'importacao gravou 5 linhas e a reimportacao nao duplicou' (
            $depois -eq ($antes + 5)) ('antes=' + $antes + ' depois=' + $depois)
        Check 'data com hora+fuso virou os 8 primeiros digitos' (
            [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE dtposted = '20250305';"
            ) -eq 1)
        Check 'registro no estilo XML gravado' (
            [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE dtposted = '20250410';"
            ) -eq 1)
        Check 'CHECKNUM virou chknum' (
            [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE chknum = '101';") -eq 1)
        Check 'memo do SGML gravado' (
            [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE memo = 'Pagamento de luz';"
            ) -eq 1)
        Check 'valor numerico gravado' (
            [Sq]::Consultar($minewDb,
                'SELECT COUNT(*) FROM extratos WHERE trnamt = -123.45;') -eq 1)
        # O Sq le o .db com charset ANSI: acento so' se confere nos BYTES do
        # arquivo (em UTF-8), que e' o formato gravado pelo parser.
        $textoDb = [Text.Encoding]::UTF8.GetString(
            [IO.File]::ReadAllBytes($minewDb))
        $temUtf8 = ($null -ne $memoAcento) -and $textoDb.Contains(
            [string]$memoAcento)
        Check 'memo acentuado do arquivo UTF-8 gravado em UTF-8' $temUtf8
        $temAnsi = ($null -ne $memoAnsi) -and $textoDb.Contains(
            [string]$memoAnsi)
        Check 'memo acentuado do arquivo ANSI (1252) convertido p/ UTF-8' (
            $temAnsi)
    }

    # ---------------------------- [10] cbAccount acompanha a tbContas no Voltar
    Write-Banner '[10/11] "Voltar" da tbContas remonta cbAccount'
    $pCb = $null
    # Mesma protecao dos demais passos: excecao tem de virar FALHA.
    $ErrorActionPreference = 'Stop'
    try {
        # O passo [9] deixou miNew-test.db com 1 conta: e' o banco aberto
        # aqui. A conta tem de ser excluida COM o aplicativo rodando (a
        # gravacao externa no arquivo devolve busy: a transacao de leitura
        # da grade segura o lock) e pelo proprio fluxo do usuario - o combo
        # so' existe na tela de extratos, entao sem a remontagem feita no
        # "Voltar" a conta excluida continuaria aparecendo la.
        $pCb = Start-Process -FilePath $exe -WorkingDirectory $root -PassThru
        Start-Sleep -Seconds 3
        $pCb.Refresh()
        $mainCb = $pCb.MainWindowHandle
        Check 'aplicacao abriu para o passo [10]' ($mainCb -ne [IntPtr]::Zero)
        $dlgCb = Invoke-MenuFileDialog $pCb 'Abrir Database'
        Check 'dialogo "Abrir" abriu (passo [10])' ($dlgCb -ne [IntPtr]::Zero)
        if ($dlgCb -ne [IntPtr]::Zero) {
            $editCb = Set-FileDialogName $dlgCb $minewDb
            Check 'campo de nome do arquivo encontrado (passo [10])' (
                $editCb -ne [IntPtr]::Zero)
            Check 'dialogo "Abrir" fechou ao confirmar (passo [10])' (
                Wait-DialogClosed $dlgCb)
            # A conexao abre depois que o dialogo some: um aviso pendente
            # aqui seria confundido com o resultado de outro fluxo.
            $avisoCb = [IntPtr]::Zero
            for ($t = 0; ($t -lt 12) -and ($avisoCb -eq [IntPtr]::Zero); $t++) {
                Start-Sleep -Milliseconds 250
                $avisoCb = [UiTest]::FindDialog([uint32]$pCb.Id)
            }
            Check 'nenhum aviso ao abrir (passo [10])' ($avisoCb -eq [IntPtr]::Zero) (
                'hwnd=' + $avisoCb)
        }
        $visCb = @(Wait-Interface $mainCb $true)
        Check 'interface revelada no passo [10]' ($visCb.Count -gt 0) (
            'janelas=' + $visCb.Count)

        # Estado inicial: so' a conta do passo [9].
        $cbAntes = Find-ComboHwnd $mainCb 212
        $nAntes  = Get-ComboCount $cbAntes
        Check 'cbAccount com 1 conta antes da troca' ($nAntes -eq 1) (
            'combo=' + $cbAntes + ' itens=' + $nAntes)

        # Navega ate' a tbContas e apaga a conta PELA PROPRIA interface: o
        # DBNavigator e' botao pintado (sem janela propria para enumerar),
        # entao o teste aponta o clique pelo centro do sexto de dez botoes de
        # largura igual - a ordem e' nbFirst..nbRefresh, o "Delete" fica na
        # posicao 6 - com o tamanho lido do proprio .lfm (359x32).
        $baseCb = @([UiTest]::Visible($mainCb))
        $idConCb = [UiTest]::MenuId($mainCb, 'Gerenciar Contas')
        Check 'item de menu "Gerenciar Contas" encontrado (passo [10])' (
            $idConCb -gt 0)
        $aposCb = $baseCb
        if ($idConCb -gt 0) {
            [void][UiTest]::Msg($mainCb, 0x0111, [IntPtr]$idConCb, [IntPtr]::Zero)
            for ($t = 0; ($t -lt 20) -and (
                @($aposCb | Where-Object { $baseCb -notcontains $_ }).Count -eq 0);
                $t++) {
                Start-Sleep -Milliseconds 250
                $aposCb = @([UiTest]::Visible($mainCb))
            }
            Check 'tela da tbContas abriu (passo [10])' (
                @($aposCb | Where-Object { $baseCb -notcontains $_ }).Count -gt 0)

            $navCb = @($aposCb | Where-Object { $_ -match ' 359x32$' }) |
                Select-Object -First 1
            Check 'DBNavigator da tbContas encontrado (passo [10])' (
                [bool]$navCb) ('linha=' + $navCb)
            if ($navCb) {
                $navHwnd = [IntPtr][int64](
                    [regex]::Match($navCb, 'id=(\d+)').Groups[1].Value)
                $navBox = [regex]::Match($navCb, '(\d+)x(\d+)$')
                $xDel = [int][Math]::Floor(
                    ([int]$navBox.Groups[1].Value) * 11 / 20)   # 5.5 de 10
                $yDel = [int][Math]::Floor(([int]$navBox.Groups[2].Value) / 2)
                [void][UiTest]::ClickOn($navHwnd, $xDel, $yDel)

                # ConfirmDelete = True: o "Delete" pergunta antes. O dialogo
                # (#32770) nao tem botao com control ID 1 (nao e' o nativo
                # MessageBox) e o da ESQUERDA e' o OK de mbOKCancel - ordem
                # fixa, entao nao depende do idioma do Windows.
                $dlgConf = [IntPtr]::Zero
                for ($t = 0; ($t -lt 8) -and ($dlgConf -eq [IntPtr]::Zero); $t++) {
                    Start-Sleep -Milliseconds 250
                    $dlgConf = [UiTest]::FindDialog([uint32]$pCb.Id)
                }
                Check 'confirmacao do delete apareceu (passo [10])' (
                    $dlgConf -ne [IntPtr]::Zero) ('hwnd=' + $dlgConf)
                if ($dlgConf -ne [IntPtr]::Zero) {
                    $btnConf = $null
                    $linhasBtn = @([UiTest]::Visible($dlgConf)) |
                        Where-Object { $_ -match '^Button \|' }
                    if ($linhasBtn) {
                        $btnConf = ($linhasBtn | ForEach-Object {
                            [pscustomobject]@{
                                Hwnd = [IntPtr][int64](
                                    [regex]::Match($_, 'id=(\d+)').Groups[1].Value)
                                X    = [int]([regex]::Match($_, '\| (\d+),\d+ ').
                                    Groups[1].Value)
                            }
                        } | Sort-Object X | Select-Object -First 1).Hwnd
                    }
                    Check 'botao de confirmacao encontrado (passo [10])' (
                        $null -ne $btnConf) ('linha=' + ($linhasBtn -join ' ;; '))
                    # Duas tentativas: na primeira o clique pode chegar antes
                    # de o dialogo assentar; a segunda cobre esse caso.
                    for ($t = 0; ($t -lt 2) -and [UiTest]::IsWindow($dlgConf); $t++) {
                        if ($null -ne $btnConf) {
                            [void][UiTest]::PostMessage($btnConf, 0x00F5,
                                [IntPtr]::Zero, [IntPtr]::Zero)   # BM_CLICK
                        }
                        for ($w = 0;
                            ($w -lt 6) -and [UiTest]::IsWindow($dlgConf); $w++) {
                            Start-Sleep -Milliseconds 250
                        }
                    }
                    Check 'confirmacao do delete respondida (passo [10])' (
                        -not [UiTest]::IsWindow($dlgConf)) ('hwnd=' + $dlgConf)
                }
            }

            # O "Voltar" e' o unico caminho de volta para a tela de extratos,
            # e e' la que o combo tem de ser remontado (na tela de gestao ele
            # some junto com o cabecalho).
            $painelCb = Invoke-VoltarContas $mainCb $baseCb
            Check 'botao "Voltar" da tbContas encontrado (passo [10])' (
                $painelCb -ne [IntPtr]::Zero)
            # As combos do cabecalho voltam junto com a tela de extratos: o
            # HWND pode ser outro, entao procura de novo ate' aparecer.
            $cbDepois = [IntPtr]::Zero
            for ($t = 0; ($t -lt 20) -and ($cbDepois -eq [IntPtr]::Zero); $t++) {
                Start-Sleep -Milliseconds 250
                $cbDepois = Find-ComboHwnd $mainCb 212
            }
            $nDepois = Get-ComboCount $cbDepois
            Check 'cbAccount remontou no "Voltar" depois de excluir a conta' (
                $nDepois -eq ($nAntes - 1)) (
                'combo=' + $cbDepois + ' antes=' + $nAntes + ' depois=' + $nDepois)
            $pCb.Refresh()
            Check 'aplicacao viva apos o "Voltar" (passo [10])' (-not $pCb.HasExited)
        }
    }
    catch {
        Check 'fluxo do cbAccount no "Voltar" sem excecao' $false $_.Exception.Message
    }
    finally {
        $ErrorActionPreference = 'Continue'
        if ($pCb -and -not $pCb.HasExited) {
            try { $pCb.Kill(); $pCb.WaitForExit() } catch { }
        }
    }

    # Conferencia no arquivo: so' funciona com a conexao encerrada (app
    # morta) e prova que o delete de verdade foi aplicado - nao so' some da
    # combo em memoria.
    if (-not (Test-Path $minewDb)) {
        Check 'miNew-test.db disponivel para a conferencia do delete' $false (
            'arquivo ausente')
    }
    else {
        $restamCb = [Sq]::Consultar($minewDb, 'SELECT COUNT(*) FROM contas;')
        Check 'delete persistido no arquivo (contas ficou vazia)' (
            $restamCb -eq 0) ('qtd=' + $restamCb)
    }

    # ------------------------------------------------ [11] estado final + WER
    Write-Banner '[11/11] Estado final e log de crashes do Windows'
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
    # extratos de teste do miImport (UTF-8 e ANSI, nomes do proprio teste)
    if (Test-Path $ofxUtf8) {
        Remove-Item $ofxUtf8 -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $ofxAnsi) {
        Remove-Item $ofxAnsi -Force -ErrorAction SilentlyContinue
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
