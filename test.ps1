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
        a reimportacao REPETE as linhas (registro igual e' permitido) e
        extrato em ANSI vira UTF-8 no arquivo
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
$ofxOutraConta = Join-Path $root 'importa-test-conta.ofx'
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
    [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr h, uint cmd);
    [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr h, int i);
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
    // TDBGrid tem barra de rolagem VERTICAL NATIVA (e' por ela que a suite
    // le as linhas da tabela ligada a grade - ver VScrollMax/VScrollRows);
    // nenhuma outra janela larga do form a tem. Como a altura da grade muda conforme o rodape
    // (448 na tela de extratos, 498 numa aba de gestao), isto identifica
    // gridTrans sem depender do tamanho exato. (hsb nao foi incluido: na
    // pratica so' a vertical aparece no estilo, com AutoFillColumns.)
    public static bool HasVScrollBar(IntPtr h) {
        const int WS_VSCROLL = 0x00200000;
        int s = GetWindowLong(h, -16); // GWL_STYLE
        return (s & WS_VSCROLL) == WS_VSCROLL;
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
    // Texto de um dialogo do processo: a mensagem (excecao, aviso) fica no
    // Edit/Static filho (memo da caixa de excecao do LCL, Static de aviso).
    // Sem isto uma falha de "nenhum dialogo" so' devolve o hwnd e nao diz o
    // que o programa reclamou na hora de gravar.
    public static string DialogText(IntPtr dlg) {
        if (dlg == IntPtr.Zero) return "";
        var sb = new StringBuilder(512);
        EnumChildWindows(dlg, (h, l) => {
            var c = new StringBuilder(64); GetClassName(h, c, 64);
            string cls = c.ToString();
            if (cls == "Edit" || cls == "Static" || cls == "RichEdit20W" ||
                cls == "RichEdit50W" || cls == "TMemo" ||
                cls == "TEdit" || cls.StartsWith("Static")) {
                string s = GetText(h);
                if (s.Length > 0) {
                    if (sb.Length > 0) sb.Append(" | ");
                    sb.Append(s);
                }
            }
            return true;
        }, IntPtr.Zero);
        return sb.ToString();
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
    // LCL usa a barra NATIVA do Windows: o nMax e' o intervalo CRUDO que o
    // LCL poe na barra (linhas + visiveis - 1, com nPage = visiveis - ver
    // TCustomDBGrid.GetScrollbarParams). Serve para achar a grade ligada a
    // tabela (>0 quando a barra existe, sem depender do desenho das
    // celulas, que nao viram janela); para CONTAR linhas vale VScrollRows.
    public static int VScrollMax(IntPtr h) {
        SCROLLINFO s;
        s.cbSize = Marshal.SizeOf(typeof(SCROLLINFO));
        s.fMask = 0x0017;             // SIF_RANGE|SIF_PAGE|SIF_POS|SIF_TRACKPOS
        s.nMin = 0; s.nMax = 0; s.nPage = 0; s.nPos = 0; s.nTrackPos = 0;
        if (!GetScrollInfo(h, 1, ref s)) return -1;   // SB_VERT
        return s.nMax;
    }
    // Linhas REAIS da tabela ligada a grade: linhas = nMax - nPage + 2.
    // O +2 e' CALIBRADO, nao deduzido: o LCL poe nMax = GetRecordCount +
    // VisibleRowCount - 1 e nPage = VisibleRowCount, e o Windows devolve os
    // dois sem ajuste (medido na pratica com 3 contagens conhecidas de
    // saldos: 10 -> nMax=18/nPage=10, 30 -> 49/21, 150 -> 169/21; todas
    // casam com nMax - nPage + 2). O GetRecordCount do TSQLQuery (FPC) vem
    // com uma linha a menos que o total real da query e e' por isso que a
    // soma nao e' nMax - nPage + 1. Sem barra montada nao da para dividir
    // (nPage = 0) - devolve -1, o mesmo sinal de "sem barra".
    public static int VScrollRows(IntPtr h) {
        SCROLLINFO s;
        s.cbSize = Marshal.SizeOf(typeof(SCROLLINFO));
        s.fMask = 0x0017;             // SIF_RANGE|SIF_PAGE|SIF_POS|SIF_TRACKPOS
        s.nMin = 0; s.nMax = 0; s.nPage = 0; s.nPos = 0; s.nTrackPos = 0;
        if (!GetScrollInfo(h, 1, ref s)) return -1;   // SB_VERT
        if (s.nPage == 0) return -1;
        return (int)(s.nMax - s.nPage + 2);
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
    // Classe da primeira IRMA ACIMA de "h" na pilha de z-order do pai que
    // sobrepoe o retangulo dela (string vazia = nada cobre a janela). Serve
    // para provar que o gridTrans nao so' existe e tem dados, mas esta' por
    // cima da PageControl1: sem o BringToFront, a revelacao da pagina
    // deixava a SysTabControl32 acima da grade - o usuario via a tela de
    // extratos "sem grade nenhuma", com a grade escondida atras das guias.
    public static string AcimaSobreposta(IntPtr h) {
        RECT r;
        if (!GetWindowRect(h, out r)) return "";
        IntPtr a = h;
        for (int i = 0; (i < 32); i++) {
            a = GetWindow(a, 3);      // GW_HWNDPREV (irma acima na pilha)
            if (a == IntPtr.Zero) break;
            RECT q;
            if (!IsWindowVisible(a) || !GetWindowRect(a, out q)) continue;
            if ((q.left < r.right) && (q.right > r.left) &&
                (q.top < r.bottom) && (q.bottom > r.top)) {
                var c = new StringBuilder(128);
                GetClassName(a, c, 128);
                return c.ToString();
            }
        }
        return "";
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

    // ---- Guias (abas) do PageControl1. No Win32 elas sao um SysTabControl32
    // que so' recebe os itens COM TabVisible (TWin32WSCustomTabControl.
    // AddAllNBPages pula as ocultas), entao contar itens = contar guias
    // visiveis - e' por isso que da para provar "so' as guias do mes com
    // saldo aparecem". SO' mensagens sem ponteiro: TCM_GETITEM/GETITEMTEXT
    // recebem um buffer que o PROCESSO DO ALVO dereferencia - num envio
    // cross-process isso le memoria invalida e derruba a aplicacao (ja'
    // aconteceu); a ordem das legendas JAN..DEZ e' garantida pelo .lfm
    // (checado no passo [1]) e nao por leitura daqui.
    // TCM_GETITEMCOUNT = TCM_FIRST(0x1300) + 4. -1 = hwnd invalido. (O +3 e'
    // TCM_SETIMAGELIST, que devolve o handle da lista antiga - nao serve.)
    public static int TabCount(IntPtr tab) {
        if (tab == IntPtr.Zero) return -1;
        return (int)Msg(tab, 0x1304, IntPtr.Zero, IntPtr.Zero);
    }
    // TCM_GETCURSEL = TCM_FIRST + 11: indice da guia ativa no Win32
    // (-1 = nenhuma, que e' o estado com todas as guias escondidas).
    public static int TabSel(IntPtr tab) {
        if (tab == IntPtr.Zero) return -1;
        return (int)Msg(tab, 0x130B, IntPtr.Zero, IntPtr.Zero);
    }
    // Troca a selecao do combo pelo MESMO caminho do usuario: CB_SETCURSEL
    // (0x014E) no combo e CBN_SELCHANGE (WM_COMMAND com HiWord = 1) no pai -
    // e' esse WM_COMMAND que dispara o OnChange do LCL. Devolve falso quando
    // o indice nao existe (CB_SETCURSEL devolve -1).
    public static bool SelectCombo(IntPtr combo, int index) {
        if (combo == IntPtr.Zero) return false;
        if ((int)Msg(combo, 0x014E, (IntPtr)index, IntPtr.Zero) == -1) return false;
        int id = GetDlgCtrlID(combo);
        IntPtr pai = GetParent(combo);
        IntPtr wparam = (IntPtr)((1 << 16) | (id & 0xFFFF));  // CBN_SELCHANGE = 1
        Msg(pai, 0x0111, wparam, combo);                      // WM_COMMAND
        return true;
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

# Le as colunas declaradas num bloco de grade: cada "item" do "Columns" vira
# um par (FieldName, Title.Caption). A ordem das propriedades DENTRO do item
# nao entra na conta - o Lazarus IDE reescreve o .lfm em ordem canonica ao
# salvar o form (e ja' reescreveu uma vez); o que importa e' o par por item
# e a ordem dos itens, que e' a ordem das colunas na grade.
function Get-LfmColumns([string]$Bloco) {
    $colunas = @()
    $item = $null
    foreach ($linha in ($Bloco -split "`r?`n")) {
        if ($linha -match '^[ \t]*item[ \t]*$') { $item = @(); continue }
        if ($null -eq $item) { continue }
        if ($linha -match '^[ \t]*end\b') {
            $campo = ''
            $titulo = ''
            foreach ($prop in $item) {
                if ($prop -match "^[ \t]*FieldName = '([^']*)'") {
                    $campo = $Matches[1]
                }
                if ($prop -match "^[ \t]*Title\.Caption = '([^']*)'") {
                    $titulo = $Matches[1]
                }
            }
            $colunas += [pscustomobject]@{ Campo = $campo; Titulo = $titulo }
            $item = $null
            continue
        }
        $item += $linha
    }
    return ,$colunas
}

# Compara as colunas lidas do .lfm com o pedido: mesma quantidade, mesma
# ordem e o par (campo, rotulo) de cada coluna. Devolve $null quando bate e
# a explicacao da primeira diferenca quando nao bate.
function Test-LfmLabels($Colunas, $Esperado) {
    for ($i = 0; $i -lt $Esperado.Count; $i++) {
        if ($i -ge $Colunas.Count) {
            return 'falta a coluna ' + ($i + 1) + ' (campo ' + $Esperado[$i][0] + ')'
        }
        if (($Colunas[$i].Campo -ne $Esperado[$i][0]) -or
            ($Colunas[$i].Titulo -ne $Esperado[$i][1])) {
            return 'a coluna ' + ($i + 1) + " e '" + $Colunas[$i].Campo +
                "' com o rotulo '" + $Colunas[$i].Titulo + "' (esperado '" +
                $Esperado[$i][0] + "' com '" + $Esperado[$i][1] + "')"
        }
    }
    if ($Colunas.Count -gt $Esperado.Count) {
        return 'sobrou a coluna ' + ($Esperado.Count + 1) + " ('" +
            $Colunas[$Esperado.Count].Campo + "')"
    }
    return $null
}

# Le as guias de mes (TTabSheet) do .lfm na ordem em que aparecem, como pares
# "nome=legenda". Cada bloco vai do "object <nome>: TTabSheet" ate' o "end" do
# MESMO recuo (as propriedades ficam dentro dele), entao a ordem das
# propriedades nao importa - o Lazarus IDE reordena o .lfm ao salvar o form.
# E' assim que se sabe QUAL guia e' cada mes: a visibilidade em runtime vem
# dos saldos (AtualizarAbasMes), a identidade da guia vem do .lfm.
function Get-LfmTabSheets([string]$Texto) {
    $guias = @()
    $linhas = $Texto -split "`r?`n"
    for ($i = 0; $i -lt $linhas.Count; $i++) {
        if ($linhas[$i] -notmatch '^([ \t]*)object[ \t]+(\w+):[ \t]*TTabSheet[ \t]*$') {
            continue
        }
        $recuo = $Matches[1]
        $par = $Matches[2] + '='
        $fim = [regex]::Escape($recuo) + 'end[ \t]*$'
        for ($j = $i + 1; $j -lt $linhas.Count; $j++) {
            if ($linhas[$j] -match $fim) { break }
            if ($linhas[$j] -match "^[ \t]*Caption[ \t]*=[ \t]*'([^']*)'") {
                $par += $Matches[1]
            }
        }
        $guias += $par
    }
    # Sem a virgula: o canal de saida desmembra o array (o "@" do chamador
    # remonta), e o ",@..." dobraria a pilha - o join viraria "Object[]".
    return $guias
}

# Grava N linhas em "saldos" (a tabela da tbSaldos) direto no arquivo, todas
# datadas do ANO e do MES informados: e' de saldos.enddate que o cbYear le a
# lista de anos e que o PageControl1 le as guias de mes (mes com saldo =
# guia visivel), entao o periodo gravado aqui e' o que as duas coisas tem de
# mostrar. O arquivo tem de estar LIVRE (conexao encerrada). Devolve o TOTAL
# de linhas de "saldos" apos o INSERT (0 = arquivo ausente/falhou - quem
# reporta e' o chamador).
function Add-SaldosRows([string]$Path, [int]$Count, [string]$Ano, [string]$Mes) {
    if (-not (Test-Path $Path)) { return 0 }
    # O INSERT com CTE recursivo grava tudo em uma unica sentenca. O dia '15'
    # e' de proposito: vale para qualquer mes (o que importa e' o texto, e o
    # mes e' lido por substr(enddate, 6, 2)).
    $sql = 'WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x+1 FROM c' +
        ' WHERE x < ' + $Count + ') INSERT INTO "saldos"' +
        " (account_id, balance, enddate) SELECT 1, x * 1.5, '" + $Ano +
        "-" + $Mes + "-15' FROM c;"
    if ([Sq]::Executar($Path, $sql) -le 0) { return 0 }
    # Le do proprio arquivo o numero de linhas que ficou: devolver o pedido sem
    # conferir esconderia um INSERT que nao deu certo.
    return [Sq]::Consultar($Path, 'SELECT COUNT(*) FROM saldos;')
}

# Grava N linhas em "extratos" (a tabela da grade da tela de extratos) datadas
# do ano E do MES informados: e' o que o combo de ano da tela tem de listar e
# o que o filtro do ano selecionado + da guia ATIVA (mes da aba) tem de deixar
# passar - por isso as linhas sao gravadas em meses que tenham guia. O dia '15'
# e' de proposito (vale para qualquer mes). O arquivo tem de estar LIVRE
# (conexao encerrada). Devolve as linhas gravadas DESSE ano (0 = arquivo
# ausente/falhou - quem reporta e' o chamador).
function Add-ExtratosRows([string]$Path, [int]$Count, [string]$Ano, [string]$Mes) {
    if (-not (Test-Path $Path)) { return 0 }
    # Mesmo truque do INSERT com CTE recursivo: uma sentenca so'.
    $sql = 'WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x+1 FROM c' +
        ' WHERE x < ' + $Count + ') INSERT INTO "extratos"' +
        " (account_id, trntype, dtposted, trnamt, memo)" +
        " SELECT 1, 'DEBIT', '" + $Ano + "-" + $Mes + "-15', -x * 1.5," +
        " 'linha ' || x FROM c;"
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
# O dialogo comum do Windows ainda pode estar montando os controles quando a
# janela aparece (e' o instante em que FindDialog ja' a enxerga): nesse momento
# o primeiro Edit visivel nao e' o de nome de arquivo - e' o da barra de
# endereco (id 41477, invisivel no estado normal) - o caminho vai para o lugar
# errado e o OK nao faz nada, deixando o dialogo aberto pra sempre. O
# FECHAMENTO e' a unica prova de que o campo era o certo, entao o clique e'
# refeito num loop curto, re-achando o campo a cada volta (com a janela ja'
# montada o Edit visivel e' o certo). Devolve o campo usado (Zero quando o
# dialogo nao tem campo de nome).
function Set-FileDialogName([IntPtr]$Dlg, [string]$Path) {
    $edit = [IntPtr]::Zero
    for ($t = 0; ($t -lt 4) -and [UiTest]::IsWindow($Dlg); $t++) {
        $edit = [UiTest]::FindFileNameEdit($Dlg)
        if ($edit -ne [IntPtr]::Zero) {
            [void][UiTest]::SetText($edit, $Path)
            # Confirmacao barata de que o texto caiu no controle achado.
            if ([UiTest]::GetText($edit) -eq $Path) {
                $btnOk = [UiTest]::FindOkButton($Dlg)
                if ($btnOk -ne [IntPtr]::Zero) {
                    [void][UiTest]::PostMessage($btnOk, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)  # BM_CLICK
                }
                else {
                    [void][UiTest]::PostMessage($Dlg, 0x0111, [IntPtr]1, [IntPtr]::Zero)          # WM_COMMAND/IDOK
                }
                # O BM_CLICK e' POSTADO: o dialogo so' reage depois que a
                # propria thread dele processa o clique - e valida o arquivo,
                # o que o shell e o antivirus ja' atrasaram em ~17s aqui. Se em
                # 2,5s nada comecou a mudar, o clique nao era o campo certo:
                # tenta de novo. (Um clique a mais nao faz mal: se o dialogo
                # ja' fechou, a janela nao existe mais e a mensagem se perde.)
                for ($w = 0; ($w -lt 10) -and [UiTest]::IsWindow($Dlg); $w++) {
                    Start-Sleep -Milliseconds 250
                }
                if (-not [UiTest]::IsWindow($Dlg)) { return $edit }
            }
        }
        Start-Sleep -Milliseconds 250
    }
    $edit
}

# Espera o dialogo informado ser destruido (120 x 250ms = 30s).
# O fechamento apos o IDOK passa pela validacao do arquivo alvo no proprio
# dialogo, e software externo (antivirus/shell) ja atrasou isso em ~17s aqui
# (o passo [6] mata a aplicacao com o banco aberto, logo o [7] reabre um
# arquivo recem-modificado). 5s dava FALHA falsa; um dialogo preso de
# verdade continua falhando, so' mais tarde.
function Wait-DialogClosed([IntPtr]$Dlg) {
    for ($t = 0; $t -lt 120; $t++) {
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

# Acha o SysTabControl32 do PageControl1 (a faixa de guias JAN..DEZ) entre
# os filhos diretos do form. No Win32 a guia e' um ITEM do proprio controle e
# o LCL so' insere as paginas COM TabVisible (AddAllNBPages), entao ali se le
# quantas guias estao visiveis ([UiTest]::TabCount), quais ([UiTest]::TabText)
# e qual e' a ativa ([UiTest]::TabSel). Zero = sem faixa na tela.
function Find-TabHwnd([IntPtr]$Main) {
    [UiTest]::FindChild($Main, 'SysTabControl32')
}

# Clica na faixa de guias ate' TCM_GETCURSEL devolver o indice pedido - e'
# o caminho do usuario (WM_LBUTTONDOWN/UP na SysTabControl32, que dispara o
# OnChange da pagina e refaz o filtro da grade). A largura da guia depende da
# fonte do Windows, entao em vez de chutar a posicao varre a faixa em passos
# e confirma pelo proprio clique; nao da para medir o item, porque TCM_GETITEM
# le um buffer do PROCESSO DO ALVO (proibido em envio cross-process - ja'
# derrubou a aplicacao). Devolve a guia que ficou ativa (-1 = nao chegou la).
function Select-AbaMes([IntPtr]$Main, [int]$Alvo) {
    $tab = Find-TabHwnd $Main
    if ($tab -eq [IntPtr]::Zero) { return -1 }
    if ([UiTest]::TabSel($tab) -eq $Alvo) { return $Alvo }
    # y=10 fica dentro da fileira das guias (a faixa tem ~24px) e x anda da
    # esquerda para a direita, de onde sao as JAN..DEZ. O clique e' POSTADO,
    # entao a selecao so' muda quando o alvo bomba a fila: espera a reacao
    # antes do proximo ponto, senao a varredura avanca sem resultado.
    for ($rx = 4; ($rx -le 240) -and ([UiTest]::TabSel($tab) -ne $Alvo); $rx += 4) {
        [void][UiTest]::ClickOn($tab, $rx, 10)
        for ($t = 0; ($t -lt 6) -and ([UiTest]::TabSel($tab) -ne $Alvo); $t++) {
            Start-Sleep -Milliseconds 50
        }
    }
    return [UiTest]::TabSel($tab)
}

# A grade de extratos (gridTrans) e' controle do FORM com Align=alClient: ela
# sobrepoe a area das guias e so' aparece na tela de extratos, com database
# aberto. Identificacao: janela da largura do form (860 - as grades das abas
# de gestao tem 852) e com a barra de rolagem vertical NATIVA do TDBGrid.
# A ALTURA nao serve de pegada: ela muda com o rodape (448 na tela de
# extratos, 498 numa aba de gestao), e um matcher de altura daria falso-ok
# nos testes de "oculta".
function Test-GridVisivelEm($Janelas) {
    foreach ($w in $Janelas) {
        if ($w -notmatch ' 860x\d+$') { continue }
        if ($w -match '\| id=(\d+)') {
            if ([UiTest]::HasVScrollBar([IntPtr][int64]$Matches[1])) { return $true }
        }
    }
    return $false
}

function Test-GridVisivel([IntPtr]$Main) {
    Test-GridVisivelEm @([UiTest]::Visible($Main))
}

# O HWND da grade de extratos (mesma pegada do Test-GridVisivel: 860 de
# largura - a do form; as grades de gestao tem 852 - com a barra de rolagem
# nativa do TDBGrid). Zero quando ela nao esta' na tela.
function Find-GridTransHwnd([IntPtr]$Main) {
    foreach ($w in @([UiTest]::Visible($Main))) {
        if ($w -notmatch ' 860x\d+$') { continue }
        if ($w -match '\| id=(\d+)') {
            $h = [IntPtr][int64]$Matches[1]
            if ([UiTest]::HasVScrollBar($h)) { return $h }
        }
    }
    return [IntPtr]::Zero
}

# Linhas que a grade de extratos esta' EXIBINDO agora. A barra vertical e' a
# nativa do TDBGrid e o numero de linhas reais sai de VScrollRows (nMax -
# nPage + 2, o mesmo calculo ja' calibrado na tbSaldos). Depois de uma troca
# de guia/ano a query fecha e reabre, e a barra so' volta quando o LCL remonta
# a grade - entao rele a tela ate' dar um numero (>= 0); -1 no fim = a grade
# nao tem barra (ou nao achou a grade).
function Get-GridTransLinhas([IntPtr]$Main) {
    $linhas = -1
    for ($t = 0; ($t -lt 12) -and ($linhas -lt 0); $t++) {
        if ($t -gt 0) { Start-Sleep -Milliseconds 250 }
        $grade = Find-GridTransHwnd $Main
        if ($grade -ne [IntPtr]::Zero) {
            $linhas = [UiTest]::VScrollRows($grade)
        }
    }
    return $linhas
}

# Os controles do FORM que sobrepoeem a area das guias na tela de extratos - a
# grade de transacoes e os dois textos do saldo anterior (na faixa das guias) -
# tem de estar ACIMA da PageControl1 na pilha de janelas. Cada mexida na
# pagina (revelar a interface na abertura, mexer no TabVisible das guias de
# mes) a coloca POR CIMA deles, e ai eles ficam com WS_VISIBLE, porem escondidos
# atras das guias: a tela aparecia "sem grade nenhuma" e sem o "Anterior",
# com tudo gravado por tras. Confere que nada que sobrepoe o retangulo de cada
# um ficou acima (o GarantirSobreposicaoNaFrente do form e' o que garante).
function Find-AnteriorValorHwnd([IntPtr]$Main) {
    # tsAnterior e' o Static da MESMA fileira do rotulo "Anterior:", a direita
    # dele: nao tem texto proprio (o valor ainda nao e' preenchido pelo
    # programa), entao a pegada e' a posicao na tela.
    $lbl = [IntPtr]::Zero
    $x = -1
    $y = -1
    foreach ($w in @([UiTest]::Visible($Main))) {
        if ($w -match '\| id=(\d+) \| "Anterior:" \| (\d+),(\d+) ') {
            $lbl = [IntPtr][int64]$Matches[1]
            $x = [int]$Matches[2]
            $y = [int]$Matches[3]
        }
    }
    if ($lbl -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
    foreach ($w in @([UiTest]::Visible($Main))) {
        if (($w -match '^Static \| id=(\d+) \| "" \| (\d+),(\d+) ') -and
            ([int]$Matches[3] -eq $y) -and ([int]$Matches[2] -gt $x)) {
            return [IntPtr][int64]$Matches[1]
        }
    }
    return [IntPtr]::Zero
}

function Check-SobreposicaoExtratos([string]$Rotulo, [IntPtr]$Main) {
    $grade = Find-GridTransHwnd $Main
    $valor = Find-AnteriorValorHwnd $Main
    $lbl = [IntPtr]::Zero
    foreach ($w in @([UiTest]::Visible($Main))) {
        if ($w -match '\| id=(\d+) \| "Anterior:" \|') {
            $lbl = [IntPtr][int64]$Matches[1]
        }
    }
    $cobGrade = [UiTest]::AcimaSobreposta($grade)
    $cobLbl = [UiTest]::AcimaSobreposta($lbl)
    $cobValor = [UiTest]::AcimaSobreposta($valor)
    Check $Rotulo (($grade -ne [IntPtr]::Zero) -and ($lbl -ne [IntPtr]::Zero) -and
        ($valor -ne [IntPtr]::Zero) -and ($cobGrade -eq '') -and ($cobLbl -eq '') -and
        ($cobValor -eq '')) (
        'grade=' + $grade + ' cobrindo=' + $cobGrade +
        ' rotulo=' + $lbl + ' cobrindo=' + $cobLbl +
        ' valor=' + $valor + ' cobrindo=' + $cobValor)
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
#   Meio   = estado ja na tela de gestao, ANTES do "Voltar" (o caller avalia
#            o que so' existe na tela de extratos - ex.: gridTrans)
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
        Meio   = $apos
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
    # Sem dgDisplayMemoText o TDBGrid desenha o literal "(MEMO)" no lugar do
    # conteudo (os campos TEXT do SQLite viram ftMemo) - e o desenho nao da
    # para ler pela janela, entao a prova fica no .lfm. Vale para as 2 grades
    # que NAO formatam nada (bancos e contas); as de extratos e de saldos sao
    # a excecao de proposito, com OnGetText no .pas, logo abaixo.
    foreach ($nomeGrade in @('gridBancos', 'gridContas')) {
        $blocoGrade = Get-LfmBlock $lfmTexto ($nomeGrade + ': TDBGrid')
        $detGrade = 'bloco ' + $nomeGrade + ' nao encontrado no .lfm'
        if ($blocoGrade -ne '') {
            $detGrade = 'dgDisplayMemoText ausente no bloco ' + $nomeGrade
        }
        Check ($nomeGrade + ' exibe o texto dos campos memo (.lfm)') (
            $blocoGrade.Contains('dgDisplayMemoText')) $detGrade
    }
    # A grade de extratos formata data e valor pelos handlers de OnGetText do
    # .pas - e com dgDisplayMemoText ligado o TDBGrid le o campo memo direto e
    # ignora o handler (por isso a opcao saiu daqui). As colunas ficaram
    # declaradas na ordem pedida: id, account_id e trntype nao tem coluna.
    $detMemo = 'dgDisplayMemoText no bloco gridTrans quebra a formatacao'
    if ($blocoGrid -eq '') { $detMemo = $detGrid }
    $memoOk = ($blocoGrid -ne '') -and (-not $blocoGrid.Contains('dgDisplayMemoText'))
    Check 'gridTrans sem dgDisplayMemoText (formatacao por OnGetText) (.lfm)' (
        $memoOk) $detMemo
    # Mesma regra na grade de saldos: enddate e' texto ISO no arquivo e a
    # exibicao vem do OnGetText (DD/MM/AAAA), entao o dgDisplayMemoText tem de
    # estar fora - com ele ligado o editor receberia o ISO cru e a mascara
    # (!99/00/0000) nao casaria.
    $blocoGradeS = Get-LfmBlock $lfmTexto 'gridSaldos: TMoneyGrid'
    $detMemoS = 'bloco gridSaldos nao encontrado no .lfm'
    if ($blocoGradeS -ne '') {
        $detMemoS = 'dgDisplayMemoText no bloco gridSaldos quebra a formatacao'
    }
    Check 'gridSaldos sem dgDisplayMemoText (enddate formatado por OnGetText) (.lfm)' (
        ($blocoGradeS -ne '') -and (-not $blocoGradeS.Contains('dgDisplayMemoText'))) $detMemoS
    $ordemOk = $false
    $detColunas = $detGrid
    if ($blocoGrid -ne '') {
        $detColunas = 'coluna faltando ou fora de ordem no bloco gridTrans'
        $ordemOk = $true
        $posColuna = -1
        foreach ($nomeColuna in @('dtposted', 'memo', 'chknum', 'trnamt')) {
            $posNova = $blocoGrid.IndexOf("FieldName = '" + $nomeColuna + "'")
            if ($posNova -le $posColuna) { $ordemOk = $false; break }
            $posColuna = $posNova
        }
    }
    Check 'gridTrans mostra dtposted, memo, chknum e trnamt nessa ordem (.lfm)' (
        $ordemOk) $detColunas
    # Rotulos das colunas: o par (campo, rotulo) e' lido POR ITEM do
    # Columns - e' o vinculo dentro do item que prova que o rotulo e' do
    # campo certo. A ordem das propriedades DENTRO do item nao entra na
    # conta (o Lazarus IDE reescreve o .lfm em ordem canonica ao salvar o
    # form, o que ja' aconteceu uma vez). Os acentos vao por [char]: o .ps1
    # nao tem BOM e o PowerShell o le como ANSI, o que deixaria a
    # comparacao com o .lfm (UTF-8) fora de fase.
    $rotuloDescricao = 'Descri' + [char]0x00E7 + [char]0x00E3 + 'o'
    $esperadoTrans = @(
        @('dtposted', 'Data'), @('memo', $rotuloDescricao),
        @('chknum', 'Documento'), @('trnamt', 'Valor'))
    $detTitulos = $detGrid
    $titulosOk = $false
    if ($blocoGrid -ne '') {
        $detTitulos = Test-LfmLabels (Get-LfmColumns $blocoGrid) $esperadoTrans
        $titulosOk = ($detTitulos -eq $null)
    }
    Check 'gridTrans rotula as colunas (Data, Descricao, Documento, Valor) (.lfm)' (
        $titulosOk) $detTitulos
    $detCanvas = 'OnPrepareCanvas ausente no bloco gridTrans'
    if ($blocoGrid -eq '') { $detCanvas = $detGrid }
    $canvasOk = $blocoGrid.Contains('OnPrepareCanvas = gridTransPrepareCanvas')
    Check 'gridTrans pinta o valor por OnPrepareCanvas (.lfm)' ($canvasOk) $detCanvas
    # O texto memo da grade de extratos vem dos handlers do .pas (sem
    # dgDisplayMemoText nao ha quem escape do "(MEMO)"), e a ligacao dos
    # quatro so' existe em codigo - entao a prova e' la.
    $pasTexto = ''
    $pasPath = Join-Path $root 'unitmoney.pas'
    if (Test-Path $pasPath) { $pasTexto = [IO.File]::ReadAllText($pasPath) }
    Check 'unitmoney.pas lido' ($pasTexto -ne '') $pasPath
    # gridSaldos e' TMoneyGrid (subclasse com EditorCanAcceptKey proprio):
    # a TDBGrid de base so' entrega a tecla ao editor em campo ftMemo (enddate
    # e TEXT no SQLite) quando dgDisplayMemoText esta' ligado - e essa opcao
    # tem de ficar fora porque a celula se formata pelo OnGetText (DD/MM/AAAA).
    # Sem o override a digitacao da data some sem erro nenhum.
    $gridClasseOk = $pasTexto.Contains('gridSaldos: TMoneyGrid;') -and
        $pasTexto.Contains('TMoneyGrid = class(TDBGrid)') -and
        ($pasTexto -match 'function TMoneyGrid\.EditorCanAcceptKey')
    Check 'gridSaldos editavel: TMoneyGrid com EditorCanAcceptKey (.pas)' (
        $gridClasseOk) ('TMoneyGrid/EditorCanAcceptKey ausente no ' +
        'unitmoney.pas (digitacao em ftMemo recusada)')
    $faltaHandler = ''
    foreach ($ligacao in @(
        'dtposted=ExtratoDtpostedGetText',
        'memo=ExtratoTextoGetText',
        'chknum=ExtratoTextoGetText',
        'trnamt=ExtratoTrnAmtGetText',
        'bankid=BancoIdGetText')) {
        $campo, $handler = $ligacao -split '='
        $padrao = "FieldByName\('" + $campo + "'\)\.OnGetText\s*:=\s*@" + $handler
        if ($pasTexto -notmatch $padrao) {
            $faltaHandler = 'OnGetText de ' + $campo + ' nao ligado a ' + $handler
            break
        }
    }
    Check 'handlers de formatacao da grade ligados (.pas)' (
        $faltaHandler -eq '') $faltaHandler
    Check 'combos de filtro com OnChange (.lfm)' (
        ($lfmTexto -match 'OnChange = cbAccountChange') -and
        ($lfmTexto -match 'OnChange = cbYearChange')) (
        'cbYearChange e/ou cbAccountChange ausentes')
    # As 12 guias de mes tem de existir, em ordem e com as legendas JAN..DEZ.
    # E' o que define QUAL guia e' cada mes: a visibilidade delas em runtime
    # vem dos saldos do ano (AtualizarAbasMes mostra so' os meses com saldo),
    # mas a identidade da guia - e, com ela, "a ultima guia e' o mes mais
    # recente" - vem daqui. As 3 guias de gestao (tbBancos/tbSaldos/tbContas)
    # tambem sao TTabSheet: o filtro e' pelo nome de 3 letras (tbJan..tbDez).
    $abasMesLfm = @(Get-LfmTabSheets $lfmTexto) |
        Where-Object { $_ -match '^tb[A-Z][a-z]{2}=' }
    Check '12 guias de mes JAN..DEZ em ordem (.lfm)' (
        ($abasMesLfm -join ' ') -eq ('tbJan=JAN tbFev=FEV tbMar=MAR tbAbr=ABR ' +
        'tbMai=MAI tbJun=JUN tbJul=JUL tbAgo=AGO tbSet=SET tbOut=OUT ' +
        'tbNov=NOV tbDez=DEZ')) ('guias=' + ($abasMesLfm -join ' '))
    # Colunas declaradas nas duas grades de gestao: com Columns.Count > 0 o
    # LCL nao cria coluna automatica, entao a coluna de "id" (autoincrement,
    # sem serventia na tela) simplesmente nao existe - e a ordem de exibicao
    # e' a da propria declaracao. A ausencia do id vira checagem propria: os
    # needles de ordem so' provam PRESENCA na ordem certa. gridSaldos declara
    # so' os 3 campos e SEM Width/SizePriority: todas variaveis = larguras
    # iguais entre si, como no layout automatico (nao vale desigualar elas
    # de novo sem o id no meio).
    $blocoContas = Get-LfmBlock $lfmTexto 'gridContas: TDBGrid'
    $ordemContas = $false
    $detContas = 'bloco gridContas nao encontrado no .lfm'
    if ($blocoContas -ne '') {
        $detContas = 'coluna faltando ou fora de ordem no bloco gridContas'
        $ordemContas = $true
        $posContas = -1
        foreach ($campoContas in @('acctid', 'accttype', 'bankid', 'branchid',
                'description')) {
            $posNova = $blocoContas.IndexOf("FieldName = '" + $campoContas + "'")
            if ($posNova -le $posContas) { $ordemContas = $false; break }
            $posContas = $posNova
        }
        if ($ordemContas -and $blocoContas.Contains("FieldName = 'id'")) {
            $ordemContas = $false
            $detContas = 'coluna de "id" declarada no bloco gridContas'
        }
    }
    Check 'gridContas esconde o id e mostra acctid, accttype, bankid, branchid e description (.lfm)' (
        $ordemContas) $detContas
    # Rotulos do gridContas: mesmo par (campo, rotulo) por item - o vinculo
    # dentro do item prova que o rotulo e' do campo certo, e a ordem das
    # propriedades dentro do item nao vale (o Lazarus IDE reescreve o .lfm
    # em ordem canonica ao salvar o form). Os acentos vao por [char]: o .ps1
    # nao tem BOM e o PowerShell o le como ANSI, o que deixaria a
    # comparacao com o .lfm (UTF-8) fora de fase.
    $esperadoContas = @(
        @('acctid', 'Conta'), @('accttype', 'Tipo'), @('bankid', 'Banco'),
        @('branchid', ('Ag' + [char]0x00EA + 'ncia')),
        @('description', ('Descri' + [char]0x00E7 + [char]0x00E3 + 'o')))
    $detRotulos = $detContas
    $rotulosContas = $false
    if ($blocoContas -ne '') {
        $detRotulos = Test-LfmLabels (Get-LfmColumns $blocoContas) $esperadoContas
        $rotulosContas = ($detRotulos -eq $null)
    }
    Check 'gridContas rotula as colunas (Conta, Tipo, Banco, Agencia, Descricao) (.lfm)' (
        $rotulosContas) $detRotulos
    # Coluna "Banco" do gridContas edita por combo (cbsPickList) com os nomes
    # de tbBancos e grava o id: OnSelectEditor repoe o combo a cada edicao (a
    # lista de bancos muda e o PickList de projeto e' fixo), o OnSetText do
    # campo traduz nome->id (o DBGrid grava via Field.Text) e o AfterOpen da
    # query religa os handlers - os campos sao dinamicos, cada Open os recria.
    $blocoSQLContas = Get-LfmBlock $lfmTexto 'SQLQueryContas: TSQLQuery'
    Check 'gridContas edita o banco por combo cbsPickList (.lfm)' (
        $blocoContas.Contains('ButtonStyle = cbsPickList')) (
        'ButtonStyle = cbsPickList ausente na coluna bankid')
    Check 'gridContas repoe o combo de bancos a cada edicao (.lfm)' (
        $blocoContas.Contains('OnSelectEditor = gridContasSelectEditor')) (
        'OnSelectEditor = gridContasSelectEditor ausente no gridContas')
    Check 'SQLQueryContas religa os handlers de banco a cada Open (.lfm)' (
        $blocoSQLContas.Contains('AfterOpen = SQLQueryContasAfterOpen')) (
        'AfterOpen = SQLQueryContasAfterOpen ausente na query de contas')
    Check 'SQLQueryContas grava o banco escolhido no combo (OnSetText) (.pas)' (
        ($pasTexto -match 'procedure\s+BancoIdSetText\s*\(') -and
        ($pasTexto -match
            "FieldByName\('bankid'\)\.OnSetText\s*:=\s*@BancoIdSetText")) (
        'BancoIdSetText e/ou a ligacao do OnSetText de bankid ausentes')
    Check 'combo de bancos alimentado pela query de tbBancos (.pas)' (
        $pasTexto -match "SQLQueryBanks\.FieldByName\('name'\)") (
        'MontarListaBancos nao le a query de tbBancos')
    # O id do banco e' codigo TEXT ('001', '077'): no driver SQLite do FPC a
    # coluna TEXT vira campo memo, que nao tem AsInteger - o acesso derruba
    # "Invalid type conversion to Integer in field id" na cara do usuario. A
    # leitura tem de ser sempre AsString (a comparacao e' numerica por fora).
    # O alvo e' a query de BANCOS: contas.id e' INTEGER e o AsInteger dele
    # (AtualizarComboContas) e legitimo.
    Check 'id dos bancos lido como texto, sem AsInteger (.pas)' (
        $pasTexto -notmatch "SQLQueryBanks\.FieldByName\('id'\)\.AsInteger") (
        'AsInteger em banks.id: o TEXT do SQLite vira memo e estoura')
    # A coluna account_id SUMIU da grade: a tbSaldos mostra so' os registros
    # da conta escolhida no cbAccount, entao a conta que filtra nao aparece
    # (o registro novo recebe a conta pelo OnNewRecord da query). Seguem id
    # escondido e as 2 colunas restantes com largura igual entre si.
    $blocoSaldos = Get-LfmBlock $lfmTexto 'gridSaldos: TMoneyGrid'
    $ordemSaldos = $false
    $detSaldos = 'bloco gridSaldos nao encontrado no .lfm'
    if ($blocoSaldos -ne '') {
        $detSaldos = 'coluna faltando ou fora de ordem no bloco gridSaldos'
        $ordemSaldos = $true
        $posSaldos = -1
        foreach ($campoSaldos in @('balance', 'enddate')) {
            $posNova = $blocoSaldos.IndexOf("FieldName = '" + $campoSaldos + "'")
            if ($posNova -le $posSaldos) { $ordemSaldos = $false; break }
            $posSaldos = $posNova
        }
        if ($ordemSaldos -and
            $blocoSaldos.Contains("FieldName = 'account_id'")) {
            $ordemSaldos = $false
            $detSaldos = 'coluna de "account_id" declarada no bloco gridSaldos'
        }
        if ($ordemSaldos -and $blocoSaldos.Contains("FieldName = 'id'")) {
            $ordemSaldos = $false
            $detSaldos = 'coluna de "id" declarada no bloco gridSaldos'
        }
        if ($ordemSaldos -and $blocoSaldos.Contains('SizePriority')) {
            $ordemSaldos = $false
            $detSaldos = 'largura fixa declarada no bloco gridSaldos ' +
                '(as colunas tem de ficar iguais)'
        }
    }
    Check 'gridSaldos esconde id e account_id e mostra balance e enddate (.lfm)' (
        $ordemSaldos) $detSaldos
    # Conta escolhida no cbAccount: filtra a query de saldos (mesmo caminho
    # do filtro de extratos - Close, SQL.Text, Open) e preenche o registro
    # novo (a coluna account_id sumiu da grade, entao so' o OnNewRecord
    # escreve nela). A ligacao do OnNewRecord e' no .lfm (check acima do
    # .pas cobre declaracao + implementacao do handler).
    $blocoSQLSaldos = Get-LfmBlock $lfmTexto 'SQLQuerySaldos: TSQLQuery'
    Check 'SQLQuerySaldos preenche a conta do combo no registro novo (.lfm)' (
        $blocoSQLSaldos.Contains('OnNewRecord = SQLQuerySaldosNewRecord')) (
        'OnNewRecord = SQLQuerySaldosNewRecord ausente na query de saldos')
    Check 'saldos filtrados pela conta do cbAccount (.pas)' (
        ($pasTexto -match 'procedure\s+TFormMoney\.AplicarFiltroSaldos') -and
        ($pasTexto -match 'SELECT \* FROM saldos') -and
        ($pasTexto -match 'WHERE account_id = ')) (
        'AplicarFiltroSaldos e/ou o filtro por account_id ausentes')
    Check 'OnNewRecord de saldos declarado e implementado (.pas)' (
        ($pasTexto -match
            'procedure\s+SQLQuerySaldosNewRecord\s*\(DataSet: TDataSet\)') -and
        ($pasTexto -match
            'procedure\s+TFormMoney\.SQLQuerySaldosNewRecord')) (
        'handler SQLQuerySaldosNewRecord ausente')
    # enddate: gravado como AAAA-MM-DD (texto puro no SQLite) e exibido/editado
    # como DD/MM/AAAA. Tres pecas obrigatorias: OnGetText (formata tambem o
    # texto do editor, que e' o que a mascara espera), OnSetText (converte de
    # volta, validando a data) e EditMask (insere as barras sozinhas). Os
    # blocos sao lidos do .pas inteiros para a checagem pegar o corpo certo.
    $blocoGetS = ''
    if ($pasTexto -match '(?s)procedure\s+TFormMoney\.SaldosEnddateGetText.*?\nend;') {
        $blocoGetS = $Matches[0]
    }
    $blocoSetS = ''
    if ($pasTexto -match '(?s)procedure\s+TFormMoney\.SaldosEnddateSetText.*?\nend;') {
        $blocoSetS = $Matches[0]
    }
    Check 'enddate da tbSaldos exibido como DD/MM/AAAA (.pas)' (
        ($blocoGetS -ne '') -and
        $blocoGetS.Contains(
            "Copy(data, 9, 2) + '/' + Copy(data, 6, 2) + '/' + Copy(data, 1, 4)")) (
        'SaldosEnddateGetText ausente ou sem a conversao AAAA-MM-DD -> DD/MM/AAAA')
    Check 'enddate da tbSaldos gravado como AAAA-MM-DD (.pas)' (
        ($blocoSetS -ne '') -and
        $blocoSetS.Contains("FormatDateTime('yyyy-mm-dd'")) (
        'SaldosEnddateSetText ausente ou sem a conversao DD/MM/AAAA -> AAAA-MM-DD')
    # A mascara so' vale se for religada depois de cada Open (os campos sao
    # dinamicos e a query de saldos so' abre em AplicarFiltroSaldos).
    $blocoPrepS = ''
    if ($pasTexto -match '(?s)procedure\s+TFormMoney\.PrepararCamposSaldos.*?\nend;') {
        $blocoPrepS = $Matches[0]
    }
    $blocoFiltroS = ''
    if ($pasTexto -match '(?s)procedure\s+TFormMoney\.AplicarFiltroSaldos.*?\nend;') {
        $blocoFiltroS = $Matches[0]
    }
    Check 'mascara de data e handlers do enddate religados a cada Open (.pas)' (
        ($blocoPrepS -ne '') -and
        $blocoPrepS.Contains('OnGetText := @SaldosEnddateGetText') -and
        $blocoPrepS.Contains('OnSetText := @SaldosEnddateSetText') -and
        $blocoPrepS.Contains("EditMask := '!99/00/0000;1;_'") -and
        ($blocoFiltroS -ne '') -and
        $blocoFiltroS.Contains('PrepararCamposSaldos')) (
        'PrepararCamposSaldos ausente, incompleto ou sem chamada em AplicarFiltroSaldos')
    # A lista de anos do cbYear vem de saldos.enddate (e nao de
    # extratos.dtposted): e' o que CarregarFiltrosExtratos consulta. O
    # negativo pega um "volta para extratos" sem trocar o SELECT inteiro.
    Check 'cbYear preenchido com os anos de saldos.enddate (.pas)' (
        ($pasTexto -match 'substr\(enddate, 1, 4\) FROM saldos') -and
        (-not ($pasTexto -match 'substr\(dtposted, 1, 4\) FROM extratos'))) (
        'consulta de anos do cbYear nao vem de saldos.enddate')
    # As guias JAN..DEZ sao o cabecalho da tela de extratos e so' aparecem
    # nos meses com saldo no ano do cbYear (mes = caracteres 6..7 de
    # enddate), com a aba ativa no mes mais recente. Os blocos sao lidos do
    # .pas inteiros (e nao por texto solto) para a checagem pegar o corpo do
    # procedimento certo.
    $blocoAbasMes = ''
    if ($pasTexto -match '(?s)procedure\s+TFormMoney\.AtualizarAbasMes.*?\nend;') {
        $blocoAbasMes = $Matches[0]
    }
    Check 'guias de mes vao dos saldos do ano escolhido (.pas)' (
        ($blocoAbasMes -match 'substr\(enddate, 6, 2\) FROM saldos') -and
        ($blocoAbasMes -match 'substr\(enddate, 1, 4\) = ')) (
        'AtualizarAbasMes nao consulta os meses de saldos.enddate')
    Check 'aba ativa vai para o mes mais recente (.pas)' (
        $blocoAbasMes -match 'PageControl1\.ActivePage := AbaDoMes\(ultimoMes\)') (
        'AtualizarAbasMes nao ativa o ultimo mes com saldo')
    $blocoCbYearMes = ''
    if ($pasTexto -match '(?s)procedure\s+TFormMoney\.cbYearChange.*?\nend;') {
        $blocoCbYearMes = $Matches[0]
    }
    Check 'trocar o ano recalcula as guias de mes (.pas)' (
        $blocoCbYearMes -match 'AtualizarAbasMes') (
        'cbYearChange nao chama AtualizarAbasMes')
    $blocoFiltrosMes = ''
    if ($pasTexto -match '(?s)procedure\s+TFormMoney\.CarregarFiltrosExtratos.*?\nend;') {
        $blocoFiltrosMes = $Matches[0]
    }
    Check 'abrir o database recalcula as guias de mes (.pas)' (
        $blocoFiltrosMes -match 'AtualizarAbasMes') (
        'CarregarFiltrosExtratos nao chama AtualizarAbasMes')
    Check 'cbYear abre no ano mais recente de saldos (.pas)' (
        $pasTexto -match 'ORDER BY substr\(enddate, 1, 4\) DESC') (
        'lista de anos nao vai do mais recente para o mais antigo')

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
            # Mesma regra do miOpen (passo [6]): as telas de gestao ficam
            # disponiveis SEM database, entao o "Novo Database" de baixo roda
            # COM a Lista de Bancos na tela - criar um arquivo tem de cair na
            # tela de extratos, nao revelar a tela de gestao (e' o
            # IrParaTelaExtratos do miNew, sem o qual o gridTrans ficava
            # escondido: ele so' existe na tela de extratos).
            $idLista5 = [UiTest]::MenuId($mainNovo, 'Lista de Bancos')
            Check 'item de menu "Lista de Bancos" encontrado (passo [5])' (
                $idLista5 -gt 0) ('id=' + $idLista5)
            if ($idLista5 -gt 0) {
                [void][UiTest]::Msg($mainNovo, 0x0111, [IntPtr]$idLista5,
                    [IntPtr]::Zero)
                $visLista5 = @(Wait-Interface $mainNovo $true)
                Check 'navegacao para a Lista de Bancos revelou a interface (passo [5])' (
                    $visLista5.Count -gt 0) ('janelas=' + $visLista5.Count)
            }

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
                # gridTrans nasce oculta no .lfm e e' do form (alClient):
                # sem esta revelacao a tela de extratos ficaria sem grade.
                Check 'gridTrans exibida apos o miNew' (
                    Test-GridVisivel $mainNovo)
                # ... e ACIMA das guias, junto com os textos do saldo
                # anterior: sem isso a revelacao da pagina deixa a
                # PageControl1 por cima e a tela aparece sem grade nenhuma e
                # sem o "Anterior", com os dados escondidos la' atras.
                Check-SobreposicaoExtratos 'sobreposicao da tela de extratos sem nada por cima (apos o miNew)' $mainNovo
                # Regressao do miNew a partir de uma tela de gestao (a mesma
                # do miOpen no passo [6]): criar tem de cair na tela de
                # extratos - gridTrans visivel + o combo de ano, que so'
                # existe nessa tela. A faixa de guias NAO serve de pegada:
                # database novo nao tem saldo e guia de mes e' mes com saldo,
                # entao a tela de extratos correta vem com zero guias.
                Check 'miNew partiu da Lista de Bancos e caiu na tela de extratos' (
                    (Test-GridVisivel $mainNovo) -and
                    ((Find-ComboHwnd $mainNovo 100) -ne [IntPtr]::Zero)) (
                    'gridTrans=' + (Test-GridVisivel $mainNovo) +
                    ' cbYear=' + (Find-ComboHwnd $mainNovo 100))
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

        # As telas de gestao ficam disponiveis SEM database (o miList), entao
        # da' para abrir um arquivo por elas - e e' o caminho em que o
        # gridTrans sumia da tela: a revelacao mostrava a tela de gestao (sem
        # guias de mes e com a grade de transacoes escondida). Para provar a
        # regra, o miOpen de baixo acontece COM a Lista de Bancos na tela.
        $idLista6 = [UiTest]::MenuId($mainAbrir, 'Lista de Bancos')
        Check 'item de menu "Lista de Bancos" encontrado (passo [6])' (
            $idLista6 -gt 0) ('id=' + $idLista6)
        if ($idLista6 -gt 0) {
            [void][UiTest]::Msg($mainAbrir, 0x0111, [IntPtr]$idLista6,
                [IntPtr]::Zero)
            $visLista6 = @(Wait-Interface $mainAbrir $true)
            Check 'navegacao para a Lista de Bancos revelou a interface' (
                $visLista6.Count -gt 0) ('janelas=' + $visLista6.Count)
            Check 'gridTrans escondida na tela de Lista de Bancos' (
                -not (Test-GridVisivel $mainAbrir))
        }

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
            # Mesma revelacao do passo [5], agora pelo caminho do miOpen.
            Check 'gridTrans exibida apos o miOpen' (
                Test-GridVisivel $mainAbrir)
            # ... e ACIMA das guias (a revelacao mostra a PageControl1 e ela
            # sobe na pilha por cima da sobreposicao - e' o caminho do bug:
            # grade e "Anterior" com WS_VISIBLE, porem escondidos atras das
            # guias).
            Check-SobreposicaoExtratos 'sobreposicao da tela de extratos sem nada por cima (apos o miOpen)' $mainAbrir
            # Combos de filtro da tela de extratos: "contas" e "saldos" estao
            # vazias neste ponto (database recem-criado no passo [5]), entao e'
            # aqui que a regra do "ano atual" e a lista de contas aparecem.
            $cbConta6 = Find-ComboHwnd $mainAbrir 212
            $cbAno6   = Find-ComboHwnd $mainAbrir 100
            $nConta6  = Get-ComboCount $cbConta6
            $nAno6    = Get-ComboCount $cbAno6
            Check 'cbYear lista o ano atual ("saldos" vazio)' ($nAno6 -eq 1) (
                'combo=' + $cbAno6 + ' itens=' + $nAno6)
            # Regressao do miOpen a partir de uma tela de gestao: abrir tem de
            # CAIR NA TELA DE EXTRATOS - gridTrans visivel e o combo de ano,
            # que so' existe nessa tela (a faixa de guias nao serve de pegada:
            # database novo sem saldo vem com zero guias de mes, ja que guia
            # de mes e' mes com saldo).
            Check 'miOpen partiu da Lista de Bancos e caiu na tela de extratos' (
                (Test-GridVisivel $mainAbrir) -and ($cbAno6 -ne [IntPtr]::Zero)) (
                'gridTrans=' + (Test-GridVisivel $mainAbrir) +
                ' cbYear=' + $cbAno6)
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
            # gridTrans e' do form: esconder a PageControl1 nao a esconde, e'
            # a regra do "so o menu" que tem de desliga-la na mao.
            Check 'gridTrans oculta apos o Fechar Database' (
                -not (Test-GridVisivel $mainFecha))
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
        # Os DOIS anos (2025 e 2026) tambem sao o que o cbYear tem de listar
        # (a lista vem de saldos.enddate, do mais recente para o mais
        # antigo), e os MESES decidem quais guias JAN..DEZ aparecem: 2026
        # com saldo em JAN/FEV/MAR, 2025 so' em JAN.
        $nSal2601 = Add-SaldosRows $minewDb 25 '2026' '01'
        $nSal2602 = Add-SaldosRows $minewDb 25 '2026' '02'
        $nSal2603 = Add-SaldosRows $minewDb 25 '2026' '03'
        $nSal2501 = Add-SaldosRows $minewDb 75 '2025' '01'
        Check 'linhas de teste gravadas em "saldos" (passo [8])' (
            ($nSal2601 -eq 25) -and ($nSal2602 -eq 50) -and
            ($nSal2603 -eq 75) -and ($nSal2501 -eq 150)) (
            '2026-01=' + $nSal2601 + ' 2026-02=' + $nSal2602 +
            ' 2026-03=' + $nSal2603 + ' 2025-01=' + $nSal2501)
        # Mesma prova para a tela de extratos: linhas NOS MESES DAS GUIAS,
        # porque a grade mostra o MES da aba ativa (ano do cbYear + mes da
        # guia) - 2026 em JAN/FEV/MAR (os tres meses com saldo, que viram
        # guia) e 2025 so' em JAN (unico mes com saldo la). As contagens
        # tambem passam do que cabe na tela (~25 linhas), ja' que e' pela
        # barra de rolagem que a suite le o que a grade mostra.
        $n2601 = Add-ExtratosRows $minewDb 50 '2026' '01'
        $n2602 = Add-ExtratosRows $minewDb 60 '2026' '02'
        $n2603 = Add-ExtratosRows $minewDb 70 '2026' '03'
        $n2501 = Add-ExtratosRows $minewDb 80 '2025' '01'
        Check 'linhas de teste gravadas em "extratos" (passo [8])' (
            ($n2601 -eq 50) -and ($n2602 -eq 110) -and ($n2603 -eq 180) -and
            ($n2501 -eq 80)) (
            '2026=' + $n2601 + '/' + $n2602 + '/' + $n2603 +
            ' 2025=' + $n2501)

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

            # 150 linhas em "saldos" (75 em 2025 so' em JAN; 75 em 2026 em
            # JAN/FEV/MAR): o combo de ano - que le a lista de saldos.enddate -
            # tem de listar os dois anos, do mais RECENTE (2026) para o mais
            # antigo, e e' o primeiro que fica selecionado. "contas" segue
            # vazia, entao o combo de conta continua sem nenhuma linha.
            $cbAno8   = Find-ComboHwnd $mainNav 100
            $cbConta8 = Find-ComboHwnd $mainNav 212
            $nAno8    = Get-ComboCount $cbAno8
            $nConta8  = Get-ComboCount $cbConta8
            $selAno8  = -1
            if ($cbAno8 -ne [IntPtr]::Zero) {
                $selAno8 = [int64][UiTest]::Msg($cbAno8, 0x0147,  # CB_GETCURSEL
                    [IntPtr]::Zero, [IntPtr]::Zero)
            }
            Check 'cbYear lista os 2 anos de "saldos"' ($nAno8 -eq 2) (
                'combo=' + $cbAno8 + ' itens=' + $nAno8)
            Check 'cbYear comeca no ano mais recente (selecao 0 = 2026)' (
                $selAno8 -eq 0) ('combo=' + $cbAno8 + ' selecao=' + $selAno8)
            Check 'cbAccount sem itens ("contas" vazia, passo [8])' (
                $nConta8 -eq 0) ('combo=' + $cbConta8 + ' itens=' + $nConta8)

            # ---- Guias de mes do PageControl1: so' os meses com saldo no ano
            # selecionado aparecem (2026 tem JAN/FEV/MAR = 3 guias) e a aba
            # ativa e' o mes mais recente - a ULTIMA guia da faixa, ja' que as
            # guias JAN..DEZ vem em ordem crescente (a ordem das legendas esta
            # no .lfm, checado no passo [1]; aqui so' da para contar e ler a
            # posicao - TCM_GETITEM le ponteiro do processo do alvo).
            $guia8 = Find-TabHwnd $mainNav
            Check 'faixa de guias de mes encontrada (passo [8])' (
                $guia8 -ne [IntPtr]::Zero) ('hwnd=' + $guia8)
            $nGuias8 = [UiTest]::TabCount($guia8)
            Check 'guias de mes = meses com saldo no ano (3 em 2026)' (
                $nGuias8 -eq 3) ('guias=' + $nGuias8)
            Check 'aba ativa = mes mais recente (MAR = ultima guia)' (
                ([UiTest]::TabSel($guia8) -eq ($nGuias8 - 1)) -and
                ($nGuias8 -eq 3)) ('sel=' + [UiTest]::TabSel($guia8) +
                ' de ' + $nGuias8)
            # A grade mostra o MES da aba ativa (ano + mes, e nao mais o ano
            # inteiro): na guia MAR/2026 cabem as 70 linhas de teste desse
            # mes - as de JAN/FEV ficam escondidas atras da troca de guia.
            $linhasMar8 = Get-GridTransLinhas $mainNav
            Check 'grade mostra o mes da aba ativa (MAR 2026 = 70 linhas)' (
                $linhasMar8 -eq 70) ('linhas=' + $linhasMar8)

            # Trocar o ano refaz a faixa: 2025 so' tem saldo em janeiro, entao
            # sobra uma guia (a de janeiro, unica do ano) e ela vira a ativa.
            # E' o mesmo caminho do clique do usuario - CB_SETCURSEL +
            # CBN_SELCHANGE no pai dispara o OnChange do LCL.
            Check 'cbYear trocado para 2025 (passo [8])' (
                [UiTest]::SelectCombo($cbAno8, 1))
            $vivaAno25 = -not $pNav.HasExited
            Check 'aplicacao viva apos trocar o ano (passo [8])' ($vivaAno25) (
                'o WM_COMMAND do cbYear derrubou a aplicacao')
            $guia25 = Find-TabHwnd $mainNav
            $nGuias25 = [UiTest]::TabCount($guia25)
            Check 'ano 2025 deixa so a guia de janeiro' ($nGuias25 -eq 1) (
                'guias=' + $nGuias25 + ' hwnd=' + $guia25)
            Check 'aba ativa em janeiro no ano 2025' (
                [UiTest]::TabSel($guia25) -eq 0) (
                'sel=' + [UiTest]::TabSel($guia25))
            # Trocar o ano troca a guia E o filtro: em 2025 a guia e' so'
            # JAN, com as 80 linhas de teste desse mes (as de 2026 somem junto
            # com o ano).
            $linhasAno25 = Get-GridTransLinhas $mainNav
            Check 'grade no ano 2025 (guia JAN) = 80 linhas' (
                $linhasAno25 -eq 80) ('linhas=' + $linhasAno25)
            # Volta para o ano inicial: os passos seguintes esperam 2026 (o
            # mais recente) e as 3 guias dele.
            [void][UiTest]::SelectCombo($cbAno8, 0)
            $nGuiasVolta = [UiTest]::TabCount((Find-TabHwnd $mainNav))
            Check 'cbYear de volta ao ano com as 3 guias' ($nGuiasVolta -eq 3) (
                'guias=' + $nGuiasVolta)
            $linhasVolta = Get-GridTransLinhas $mainNav
            Check 'grade de volta em 2026 (guia MAR) = 70 linhas' (
                $linhasVolta -eq 70) ('linhas=' + $linhasVolta)

            # Clicar em OUTRA guia tem de refazer o filtro - e' o caminho do
            # usuario e o gatilho novo (o PageControl1Change chama a
            # filtragem): JAN/2026 = 50 linhas, FEV/2026 = 60, nunca as 180
            # do ano inteiro. O clique e' WM_LBUTTONDOWN/UP na propria
            # SysTabControl32 (Select-AbaMes varre a faixa, ja' que a largura
            # da guia depende da fonte do Windows).
            $selJan8 = Select-AbaMes $mainNav 0
            Check 'clique na guia JAN ativou a guia' ($selJan8 -eq 0) (
                'sel=' + $selJan8)
            Check 'aplicacao viva apos o clique na guia (passo [8])' (
                -not $pNav.HasExited)
            $linhasJan8 = Get-GridTransLinhas $mainNav
            Check 'grade refiltrada na guia JAN 2026 = 50 linhas' (
                $linhasJan8 -eq 50) ('linhas=' + $linhasJan8)

            $selFev8 = Select-AbaMes $mainNav 1
            Check 'clique na guia FEV ativou a guia' ($selFev8 -eq 1) (
                'sel=' + $selFev8)
            $linhasFev8 = Get-GridTransLinhas $mainNav
            Check 'grade refiltrada na guia FEV 2026 = 60 linhas' (
                $linhasFev8 -eq 60) ('linhas=' + $linhasFev8)

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
            # gridTrans e' da tela de extratos: nas telas de gestao ela tem de
            # sumir, senao cobriria a grade de contas.
            Check 'gridTrans oculta na tela de contas' (
                -not (Test-GridVisivelEm $cCom.Meio))
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
            # ... e a grade de extratos volta junto com a tela de extratos.
            Check 'gridTrans exibida de volta no Voltar' (
                Test-GridVisivelEm $estCom)
            # ... tambem ACIMA das guias: o "Voltar" roda o AtualizarAbasMes,
            # que mexe guia por guia no TabVisible e reordena a pilha - o
            # caminho em que a sobreposicao (grade e "Anterior") voltava
            # escondida atras da PageControl1.
            Check-SobreposicaoExtratos 'sobreposicao da tela de extratos sem nada por cima (apos o Voltar)' $mainNav

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
                    'janelas=' + $visFec8.Count + ' :: ' +
                    ($visFec8 -join ' | '))

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
        # Total do passo [8] (extratos por mes das guias: 50+60+70 em 2026 e
        # 80 em 2025 = 260) - a referencia da qual a importacao tem de SOMAR 9.
        $antes = [Sq]::Consultar($minewDb, 'SELECT COUNT(*) FROM extratos;')
        Check 'extratos lido antes da importacao' ($antes -eq 260) ('qtd=' + $antes)

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

        # Extrato valido, mas de OUTRA conta: o <ACCTID> (9999) nao e' o da
        # conta escolhida no combo (4321). A importacao tem de ser recusada
        # ANTES de gravar - o resultado visivel e' um aviso, nao o dialogo de
        # resultado. O memo do registro e' a prova no arquivo depois: se ele
        # aparecer em "extratos", a conferencia de conta nao rodou.
        $memoOutraConta = 'Lancamento da outra conta'
        $ofxOutraLinhas = @(
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
            '<ACCTID>9999'
            '<ACCTTYPE>CHECKING'
            '</BANKACCTFROM>'
            '<BANKTRANLIST>'
            '<STMTTRN>'
            '<TRNTYPE>DEBIT'
            '<DTPOSTED>20240615'
            '<TRNAMT>-77.77'
            '<MEMO>' + $memoOutraConta
            '</STMTTRN>'
            '</BANKTRANLIST>'
            '</STMTRS>'
            '</STMTTRNRS>'
            '</BANKMSGSRSV1>'
            '</OFX>'
        )
        [IO.File]::WriteAllText($ofxOutraConta,
            ($ofxOutraLinhas -join "`r`n") + "`r`n")
        Check 'fixture OFX de outra conta gravada' (Test-Path $ofxOutraConta)

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
        # ja estao em "saldos" (2025 e 2026, do passo [8]) - e' de
        # saldos.enddate que o cbYear monta a lista, do mais recente (2026)
        # para o mais antigo, e e' o primeiro que fica selecionado.
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

        # Depois do resultado as combos ja foram recarregadas. A importacao
        # grava em "extratos", entao NADA muda na lista de anos (ela vem de
        # "saldos"): continuam 2 itens, com 2026 na frente (ordem decrescente)
        # - e a selecao, devolvida pelo texto salvo, tem de seguir 2026
        # (indice 0).
        $cbContaB = Find-ComboHwnd $mainImp 212
        $cbAnoB   = Find-ComboHwnd $mainImp 100
        $nContaB  = Get-ComboCount $cbContaB
        $nAnoB    = Get-ComboCount $cbAnoB
        $selAnoB  = [int64][UiTest]::Msg($cbAnoB, 0x0147, [IntPtr]::Zero, [IntPtr]::Zero)
        Check 'cbAccount continua com a conta de destino' ($nContaB -eq 1) (
            'itens=' + $nContaB)
        Check 'cbYear segue com 2 anos (o arquivo nao muda a lista de saldos)' (
            $nAnoB -eq 2) ('itens=' + $nAnoB)
        Check 'cbYear manteve a selecao do usuario (indice 0 = 2026)' (
            $selAnoB -eq 0) ('sel=' + $selAnoB)
        # As guias de mes tambem vem de "saldos": a importacao nao mexe nelas
        # - continuam as 3 de janeiro/fevereiro/marco de 2026.
        $guiaImp = Find-TabHwnd $mainImp
        Check 'guias de mes inalteradas apos a importacao (JAN FEV MAR)' (
            [UiTest]::TabCount($guiaImp) -eq 3) (
            'guias=' + [UiTest]::TabCount($guiaImp) + ' hwnd=' + $guiaImp)

        # Reimportar o MESMO arquivo REPETE as linhas: registro igual e'
        # permitido (a conferencia disso e' no arquivo, no fim do passo). Os
        # anos continuam os mesmos, ja' que a lista vem de "saldos".
        $dlgImp2 = Invoke-Importar $pImp $ofxUtf8
        Check 'dialogo de resultado da reimportacao' ($dlgImp2 -ne [IntPtr]::Zero)
        if ($dlgImp2 -ne [IntPtr]::Zero) {
            [void][UiTest]::PostMessage($dlgImp2, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
            Check 'resultado da reimportacao dispensado' (
                Wait-DialogClosed $dlgImp2)
        }
        $cbAnoC = Find-ComboHwnd $mainImp 100
        Check 'cbYear inalterado apos a reimportacao' (
            (Get-ComboCount $cbAnoC) -eq 2) ('itens=' + (Get-ComboCount $cbAnoC))

        # Extrato em ANSI: o parser tem de converte-los para UTF-8. O arquivo
        # nao mexe em "saldos", entao a lista de anos nao muda.
        $dlgImp3 = Invoke-Importar $pImp $ofxAnsi
        Check 'dialogo de resultado da importacao ANSI' ($dlgImp3 -ne [IntPtr]::Zero)
        if ($dlgImp3 -ne [IntPtr]::Zero) {
            [void][UiTest]::PostMessage($dlgImp3, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
            Check 'resultado da importacao ANSI dispensado' (
                Wait-DialogClosed $dlgImp3)
        }
        $cbAnoD = Find-ComboHwnd $mainImp 100
        Check 'cbYear segue com 2 anos (importacao nao mexe em "saldos")' (
            (Get-ComboCount $cbAnoD) -eq 2) ('itens=' + (Get-ComboCount $cbAnoD))

        # Arquivo de outra conta: o aviso tem de vir no lugar do resultado.
        $dlgImp4 = Invoke-Importar $pImp $ofxOutraConta
        Check 'aviso de conta diferente exibido' ($dlgImp4 -ne [IntPtr]::Zero)
        if ($dlgImp4 -ne [IntPtr]::Zero) {
            [void][UiTest]::PostMessage($dlgImp4, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)  # WM_CLOSE
            Check 'aviso de conta diferente dispensado' (
                Wait-DialogClosed $dlgImp4)
        }
        # A recusa acontece antes de fechar/reabrir as queries, entao as
        # combos ficam exatamente como estavam (2 anos, 1 conta).
        $cbAnoE = Find-ComboHwnd $mainImp 100
        Check 'cbYear inalterado apos a recusa' (
            (Get-ComboCount $cbAnoE) -eq 2) ('itens=' + (Get-ComboCount $cbAnoE))
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
    # 9 linhas novas no total: 4 da 1a importacao + os MESMOS 4 repetidos na
    # reimportacao (registro igual e' permitido) + 1 do ANSI.
    if (-not (Test-Path $minewDb)) {
        Check 'miNew-test.db disponivel para a conferencia final' $false (
            'arquivo ausente')
    }
    else {
        $depois = [Sq]::Consultar($minewDb, 'SELECT COUNT(*) FROM extratos;')
        Check 'importacao gravou 9 linhas (4 + 4 repetidos + 1 do ANSI)' (
            $depois -eq ($antes + 9)) ('antes=' + $antes + ' depois=' + $depois)
        # O arquivo de outra conta (ACCTID 9999) foi recusado na tela; a
        # prova duraria no arquivo: o registro dele nao pode estar em
        # "extratos" - se estiver, a conferencia de conta nao rodou. (-1 =
        # o fixture nem foi criado, entao nao passa em silencio.)
        $qtdOutra = -1
        if ($null -ne $memoOutraConta) {
            $qtdOutra = [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE memo = '" +
                $memoOutraConta + "';")
        }
        Check 'extrato da outra conta nao foi gravado' ($qtdOutra -eq 0) (
            'linhas=' + $qtdOutra)
        # As linhas do arquivo UTF-8 estao gravadas DUAS vezes (importacao +
        # reimportacao do mesmo arquivo) - e' a prova de que registro repetido
        # entra; o ANSI foi importado uma vez so'.
        Check 'data com hora+fuso virou os 8 primeiros digitos (2x)' (
            [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE dtposted = '20250305';"
            ) -eq 2)
        Check 'registro no estilo XML gravado (2x)' (
            [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE dtposted = '20250410';"
            ) -eq 2)
        Check 'CHECKNUM virou chknum (2x)' (
            [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE chknum = '101';") -eq 2)
        Check 'memo do SGML gravado (2x)' (
            [Sq]::Consultar($minewDb,
                "SELECT COUNT(*) FROM extratos WHERE memo = 'Pagamento de luz';"
            ) -eq 2)
        Check 'valor numerico gravado (2x)' (
            [Sq]::Consultar($minewDb,
                'SELECT COUNT(*) FROM extratos WHERE trnamt = -123.45;') -eq 2)
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
        # Banco na lista de tbBancos ANTES de subir (o arquivo tem de estar
        # livre): a coluna "Banco" do gridContas traduz o bankid pelo codigo
        # do banco, e sem essa linha o laco da traducao nao roda. Foi assim
        # que o AsInteger em id TEXT (que o driver SQLite traz como campo
        # memo) passou despercebido e so' quebrou com os bancos reais.
        Check 'banco de teste gravado em banks.db (passo [10])' (
            [Sq]::Executar($db,
                "INSERT INTO banks (id, name, alias)" +
                " VALUES ('001', 'Banco do Brasil', 'BB');") -eq 1)
        # Um saldo de OUTRA conta, gravado com o arquivo livre: a tbSaldos
        # so' pode mostrar as 150 linhas da conta selecionada no combo (a
        # 151a, de outra conta, tem de ficar de fora) - e' a prova do
        # filtro da grade de saldos.
        Check 'saldo de outra conta gravado (passo [10])' (
            [Sq]::Executar($minewDb,
                'INSERT INTO saldos (account_id, balance, enddate)' +
                " VALUES (2, 999.9, '2026-12-31');") -eq 1)
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

        # ---- tbSaldos: o cbAccount continua visivel la e filtra a grade.
        $idSal10 = [UiTest]::MenuId($mainCb, 'Gerenciar Saldos')
        Check 'item de menu "Gerenciar Saldos" encontrado (passo [10])' (
            $idSal10 -gt 0) ('id=' + $idSal10)
        if ($idSal10 -gt 0) {
            [void][UiTest]::Msg($mainCb, 0x0111, [IntPtr]$idSal10,
                [IntPtr]::Zero)
            # Visivel = achado entre as janelas visiveis da tela (o mesmo
            # Find-ComboHwnd da tela de extratos, pela largura 212).
            $cbSal10 = [IntPtr]::Zero
            for ($t = 0; ($t -lt 20) -and ($cbSal10 -eq [IntPtr]::Zero); $t++) {
                Start-Sleep -Milliseconds 250
                $cbSal10 = Find-ComboHwnd $mainCb 212
            }
            Check 'cbAccount visivel na tela de Gerenciar Saldos (passo [10])' (
                $cbSal10 -ne [IntPtr]::Zero) ('combo=' + $cbSal10)
            # A grade de gestao tem 852 de largura (a de extratos, 860): a
            # barra nativa devolve nMax = GetRecordCount + visiveis - 1 e
            # nPage = visiveis, e o GetRecordCount do TSQLQuery vem com uma
            # linha a menos - logo linhas reais = nMax - nPage + 2
            # (VScrollRows, calibrado com contagens conhecidas). Tem de dar
            # 150 (saldos da conta do combo), nunca as 151 do arquivo.
            $gradeSal = [IntPtr]::Zero
            for ($t = 0; ($t -lt 20) -and ($gradeSal -eq [IntPtr]::Zero); $t++) {
                Start-Sleep -Milliseconds 250
                foreach ($wSal in @([UiTest]::Visible($mainCb))) {
                    if (($wSal -match ' 852x\d+$') -and
                        ($wSal -match '\| id=(\d+)') -and
                        ([UiTest]::HasVScrollBar([IntPtr][int64]$Matches[1]))) {
                        $gradeSal = [IntPtr][int64]$Matches[1]
                        break
                    }
                }
            }
            $linhasSal = [UiTest]::VScrollRows($gradeSal)
            Check 'tbSaldos mostra so as linhas da conta do combo (150 de 151)' (
                $linhasSal -eq 150) ('linhas=' + $linhasSal +
                ' grade=' + $gradeSal)

            # ---- enddate em EXECUCAO: o editor da coluna da data abre com o
            # valor em DD/MM/AAAA (o arquivo guarda AAAA-MM-DD) e a mascara
            # insere as barras sozinhas na digitacao. A navegacao e' TODA por
            # teclado, de proposito: TCustomGrid.MouseDown sai cedo quando a
            # grade nao esta focada (SetFocus falha com o form atras da IDE e a
            # suite nunca traz o form para frente), e tecla nao passa por esse
            # guard - e' por VK_F2 que o editor abre. HOME poe a coluna na
            # primeira (balance), VK_RIGHT anda ate' a coluna cujo editor abre
            # COM UMA DATA (a largura das colunas nao da para prever daqui:
            # AutoFillColumns estica a ultima para os 852 da grade).
            $editouData = $false
            if ($gradeSal -ne [IntPtr]::Zero) {
                $antesEd = @([UiTest]::Visible($gradeSal))
                $editorSal = [IntPtr]::Zero
                $colData = -1
                [void][UiTest]::Msg($gradeSal, 0x0100, [IntPtr]0x24,
                    [IntPtr]::Zero)                              # VK_HOME
                [void][UiTest]::Msg($gradeSal, 0x0101, [IntPtr]0x24,
                    [IntPtr]::Zero)
                for ($passo = 0; ($passo -lt 3) -and ($colData -lt 0); $passo++) {
                    if ($passo -gt 0) {
                        # ESC fecha o editor sem gravar (nada digitado ainda) e
                        # VK_RIGHT troca a coluna (MoveSel com gfEditingDone).
                        [void][UiTest]::Msg($gradeSal, 0x0100, [IntPtr]0x1B,
                            [IntPtr]::Zero)                      # VK_ESCAPE
                        [void][UiTest]::Msg($gradeSal, 0x0101, [IntPtr]0x1B,
                            [IntPtr]::Zero)
                        [void][UiTest]::Msg($gradeSal, 0x0100, [IntPtr]0x27,
                            [IntPtr]::Zero)                      # VK_RIGHT
                        [void][UiTest]::Msg($gradeSal, 0x0101, [IntPtr]0x27,
                            [IntPtr]::Zero)
                    }
                    [void][UiTest]::Msg($gradeSal, 0x0100, [IntPtr]0x71,
                        [IntPtr]::Zero)                          # VK_F2
                    [void][UiTest]::Msg($gradeSal, 0x0101, [IntPtr]0x71,
                        [IntPtr]::Zero)
                    Start-Sleep -Milliseconds 250
                    # O editor e' uma janela FILHA da grade (FEditor.Parent :=
                    # Self em grids.pas) com a classe nativa "Edit" (win32int.pp).
                    # A grade nao tem nenhum filho visivel em repouso (as barras
                    # de rolagem sao do proprio estilo da janela, sem janela
                    # propria), entao achar um "Edit" filho apos o F2 e achar o
                    # editor - e o handle NAO e' o mesmo entre aberturas (o
                    # editor e' recriado), por isso a busca e' refeita a cada
                    # passo. O texto da para ler por WM_GETTEXT (uma das 3
                    # mensagens que o Windows faz marshaling cross-process - o
                    # mesmo caminho do GetWindowText que Visible() ja' usa).
                    $editorSal = [IntPtr]::Zero
                    foreach ($wEd in @([UiTest]::Visible($gradeSal))) {
                        if (($wEd -match '^Edit \|') -and
                            ($antesEd -notcontains $wEd)) {
                            $editorSal = [IntPtr][int64](
                                [regex]::Match($wEd, 'id=(\d+)').
                                    Groups[1].Value)
                            break
                        }
                    }
                    if ($editorSal -eq [IntPtr]::Zero) { continue }
                    $txtCel = [UiTest]::GetText($editorSal)
                    Write-Host ('  enddate: coluna ' + $passo +
                        ' editor="' + $txtCel + '"')
                    if ($txtCel -match '^\d{2}/\d{2}/\d{4}$') {
                        $colData = $passo
                    }
                }
                Check 'editor de enddate abriu com a data em DD/MM/AAAA (passo [10])' (
                    $colData -ge 0) ('editor=' + $editorSal)
                if ($colData -ge 0) {
                    # ESC fecha o editor sem gravar (a data antiga continua no
                    # arquivo - nada foi digitado).
                    [void][UiTest]::Msg($gradeSal, 0x0100, [IntPtr]0x1B,
                        [IntPtr]::Zero)                          # VK_ESCAPE
                    [void][UiTest]::Msg($gradeSal, 0x0101, [IntPtr]0x1B,
                        [IntPtr]::Zero)
                    # Registro novo PELA PROPRIA interface: nbInsert e o 5o de
                    # 10 botoes de largura igual do DBNavigator. O painel do
                    # navigator nao passa pelo guard de foco do MouseDown da
                    # grade - e' o mesmo caminho do "Delete" da tbContas, que ja'
                    # funciona na suite. A coluna atual (enddate) nao muda com a
                    # insercao, so a linha.
                    $navSal = @([UiTest]::Visible($mainCb) |
                        Where-Object { $_ -match ' 359x32$' }) |
                        Select-Object -First 1
                    Check 'DBNavigator da tbSaldos encontrado (passo [10])' (
                        [bool]$navSal) ('linha=' + $navSal)
                    if ($navSal) {
                        $navSalHwnd = [IntPtr][int64](
                            [regex]::Match($navSal, 'id=(\d+)').Groups[1].Value)
                        $navSalW = [int]([regex]::Match($navSal,
                            '\| \d+,\d+ (\d+)x\d+').Groups[1].Value)
                        [void][UiTest]::ClickOn($navSalHwnd,
                            [int][Math]::Floor($navSalW * 9 / 20), 16) # nbInsert
                        Start-Sleep -Milliseconds 500
                        # Registro novo: balance e' NOT NULL e o OnNewRecord
                        # so' preenche a conta - o Post do VK_UP com o saldo
                        # vazio rebenta no SQLite e abre a caixa de excecao do
                        # LCL. O usuario preenche as duas colunas, entao o
                        # teste tambem: HOME vai para a primeira (balance) e,
                        # depois de digitar o saldo, VK_RIGHT fecha o editor e
                        # passa para a coluna de enddate.
                        [void][UiTest]::Msg($gradeSal, 0x0100, [IntPtr]0x24,
                            [IntPtr]::Zero)                      # VK_HOME
                        [void][UiTest]::Msg($gradeSal, 0x0101, [IntPtr]0x24,
                            [IntPtr]::Zero)
                        # Digita um texto numa celula como o usuario digita: a
                        # 1a tecla vai para a GRADE (e' ela que abre o editor,
                        # via EditorShowChar/EditorCanAcceptKey) e o resto vai
                        # para o EDITOR (com EditorMode a grade nao repassa
                        # WM_CHAR). As teclas vem como teclado de verdade
                        # (KEYDOWN + CHAR + KEYUP): o win32 armanda
                        # IgnoreNextCharWindow em TODO keydown (o F2/ESC/RIGHT
                        # do passo acima deixou armado na grade) e so' desarma
                        # no keydown que o LCL NAO trata - sem esse keydown o
                        # WM_CHAR era engolido antes de chegar ao KeyPress da
                        # grade. Espera o editor nascer E a 1a tecla entrar
                        # antes de mandar as seguintes, para nao misturar o
                        # POST que o SendCharToEditor fez com as mensagens da
                        # suite. Devolve o texto final ('' = editor nao abriu).
                        function Invoke-GridType([IntPtr]$grid, [string]$texto) {
                            if ($texto -eq '') { return '' }
                            $k1 = [Convert]::ToInt32($texto[0])
                            [void][UiTest]::Msg($grid, 0x0100, [IntPtr]$k1,
                                [IntPtr]::Zero)                  # WM_KEYDOWN
                            [void][UiTest]::Msg($grid, 0x0102, [IntPtr]$k1,
                                [IntPtr]::Zero)                  # WM_CHAR
                            [void][UiTest]::Msg($grid, 0x0101, [IntPtr]$k1,
                                [IntPtr]::Zero)                  # WM_KEYUP
                            # EditorShowChar manda a propria tecla para o
                            # editor (SendCharToEditor, WM_CHAR POST no
                            # editor): a janela nasce agora - e' uma FILHA da
                            # grade com a classe nativa "Edit", recriada a
                            # cada abertura.
                            $ed = [IntPtr]::Zero
                            $txt = ''
                            $alvo = [regex]::Escape($texto[0].ToString())
                            for ($i = 0; ($i -lt 8) -and ($txt -notmatch $alvo);
                                $i++) {
                                Start-Sleep -Milliseconds 250
                                if ($ed -eq [IntPtr]::Zero) {
                                    foreach ($w in @([UiTest]::Visible($grid))) {
                                        if ($w -match '^Edit \|') {
                                            $ed = [IntPtr][int64](
                                                [regex]::Match($w, 'id=(\d+)').
                                                    Groups[1].Value)
                                            break
                                        }
                                    }
                                }
                                if ($ed -ne [IntPtr]::Zero) {
                                    $txt = [UiTest]::GetText($ed)
                                }
                            }
                            if ($ed -eq [IntPtr]::Zero) { return '' }
                            for ($i = 1; $i -lt $texto.Length; $i++) {
                                [void][UiTest]::Msg($ed, 0x0102,
                                    [IntPtr][Convert]::ToInt32($texto[$i]),
                                    [IntPtr]::Zero)              # WM_CHAR
                            }
                            Start-Sleep -Milliseconds 250
                            return [UiTest]::GetText($ed)
                        }
                        $txtSaldo = Invoke-GridType $gradeSal '10'
                        Write-Host ('  balance: digitado="' + $txtSaldo + '"')
                        Check 'saldo digitado no registro novo (passo [10])' (
                            $txtSaldo -eq '10') ('texto="' + $txtSaldo + '"')
                        [void][UiTest]::Msg($gradeSal, 0x0100, [IntPtr]0x27,
                            [IntPtr]::Zero)                      # VK_RIGHT
                        [void][UiTest]::Msg($gradeSal, 0x0101, [IntPtr]0x27,
                            [IntPtr]::Zero)
                        Start-Sleep -Milliseconds 250
                        # A digitacao da data e' o caminho do usuario: a 1a
                        # tecla na GRADE abre o editor (EditorShowChar so' o
                        # abre se EditorCanAcceptKey deixar - enddate e' TEXT
                        # -> ftMemo, e ftMemo e' blob para a TDBGrid de base,
                        # que recusava a tecla e fazia a digitacao sumir sem
                        # erro nenhum; o TMoneyGrid e' o que aceita) e o resto
                        # vai para o EDITOR, onde o TCustomMaskEdit insere as
                        # barras sozinhas. Registro novo vazio: o OnNewRecord
                        # so' preenche a conta e o enddate vem NULL, que o
                        # OnGetText devolve ''.
                        $txtData = Invoke-GridType $gradeSal '31032028'
                        Write-Host ('  enddate: digitado="' + $txtData + '"')
                        Check 'digitou na grade e o editor de enddate abriu (passo [10])' (
                            $txtData -ne '') ('texto="' + $txtData + '"')
                        # A mascara roda no TCustomMaskEdit: CanInsertChar
                        # filtra o caractere e pula os separadores, entao
                        # digitar 31032028 tem de virar 31/03/2028 - sem a
                        # mascara viraria 31032028 na cara do usuario.
                        Check 'mascara de enddate insere as barras sozinhas (passo [10])' (
                            $txtData -eq '31/03/2028') (
                            'texto="' + $txtData + '"')
                        $editouData = ($txtData -eq '31/03/2028')
                        # VK_UP fecha o editor e POE o registro novo: a
                        # primeira tecla ja' pousou a query em dsEdit
                        # (EditorIsReadOnly chama FDataLink.Edit) e marcou o
                        # datalink como modified, entao o InsertCancelable do
                        # doVKUP e' False e o MoveBy(-1) da Post - e e' no
                        # Post que o FTempText vira Field.Text (OnUpdateData),
                        # onde o OnSetText grava AAAA-MM-DD. VK_UP e' de
                        # proposito: o registro novo e' o ULTIMO, e VK_DOWN
                        # cairia em EOF e pediria mais um registro (opAppend).
                        [void][UiTest]::Msg($gradeSal, 0x0100, [IntPtr]0x26,
                            [IntPtr]::Zero)                  # VK_UP
                        [void][UiTest]::Msg($gradeSal, 0x0101, [IntPtr]0x26,
                            [IntPtr]::Zero)
                        Start-Sleep -Milliseconds 500
                    }
                }
            }
            # Se voltou a rebentar na hora de gravar (ex.: balance NOT NULL
            # sem saldo), a mensagem sai no log e a caixa e' dispensada - sem
            # isso ela fica pendurada e polui os FindDialog dos passos
            # seguintes; a falha em si ja' fica registrada logo abaixo.
            $dlgEd = [UiTest]::FindDialog([uint32]$pCb.Id)
            if ($dlgEd -ne [IntPtr]::Zero) {
                Write-Host ('  dialogo ao gravar enddate: "' +
                    [UiTest]::DialogText($dlgEd) + '"')
                $okDlg = [UiTest]::FindOkButton($dlgEd)
                if ($okDlg -ne [IntPtr]::Zero) {
                    [void][UiTest]::Msg($okDlg, 0x00F5, [IntPtr]::Zero,
                        [IntPtr]::Zero)                        # BM_CLICK
                    Start-Sleep -Milliseconds 300
                }
            }
            Check 'nenhum dialogo ao editar enddate (passo [10])' (
                $dlgEd -eq [IntPtr]::Zero) ('hwnd=' + $dlgEd)
        }

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

            # A pintura da coluna "Banco" traduz o bankid pelo codigo da
            # lista de tbBancos (agora populada com o banco de teste): um
            # erro de conversao ali vira dialogo MODAL na cara do usuario -
            # e' o defeito do AsInteger em id TEXT.
            Start-Sleep -Milliseconds 500
            $dlgPint = [UiTest]::FindDialog([uint32]$pCb.Id)
            Check 'nenhum dialogo de erro ao pintar a grade de contas (passo [10])' (
                $dlgPint -eq [IntPtr]::Zero) (
                'hwnd=' + $dlgPint + ' texto="' +
                [UiTest]::DialogText($dlgPint) + '"')

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

    # A digitacao na grade do passo [10] so' da para conferir no arquivo com a
    # conexao encerrada (mesmo motivo do delete de contas): a mascara deixou
    # 31/03/2028 na tela do registro novo e o campo tem de ter gravado
    # 2028-03-31 (150 da conta + 1 da outra conta + 1 digitado).
    $qtdIsoSal = [Sq]::Consultar($minewDb,
        "SELECT COUNT(*) FROM saldos WHERE enddate = '2028-03-31';")
    Check 'enddate digitado na grade gravado em AAAA-MM-DD (passo [10])' (
        $editouData -and ($qtdIsoSal -eq 1)) (
        'editou=' + $editouData + ' qtd=' + $qtdIsoSal)
    $qtdSalTot = [Sq]::Consultar($minewDb, 'SELECT COUNT(*) FROM saldos;')
    Check 'linha digitada na grade entrou em saldos (passo [10])' (
        $qtdSalTot -eq 152) ('qtd=' + $qtdSalTot)
    # O saldo digitado na coluna balance tem de ter ido junto com a data:
    # balance e' NOT NULL, entao sem ele o Post do VK_UP nao grava nada (e'
    # por isso que o registro novo do teste leva saldo + data).
    $qtdSalSaldo = [Sq]::Consultar($minewDb,
        "SELECT COUNT(*) FROM saldos WHERE enddate = '2028-03-31'" +
        " AND balance = 10;")
    Check 'saldo digitado gravado junto com a data (passo [10])' (
        $qtdSalSaldo -eq 1) ('qtd=' + $qtdSalSaldo)
    # Nenhum enddate pode ter virado lixo no meio da edicao (o formato do
    # arquivo e' sempre AAAA-MM-DD, com ou sem edicao).
    $qtdLixoSal = [Sq]::Consultar($minewDb,
        "SELECT COUNT(*) FROM saldos WHERE enddate NOT GLOB" +
        " '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]';")
    Check 'nenhum enddate virou lixo depois da edicao (passo [10])' (
        $qtdLixoSal -eq 0) ('qtd=' + $qtdLixoSal)

    # Volta banks.db ao estado do passo: o banco de teste so' existiu para o
    # laco da traducao do "Banco" rodar com dados (arquivo livre agora, com
    # a aplicacao encerrada).
    Check 'banco de teste removido de banks.db (passo [10])' (
        [Sq]::Executar($db, "DELETE FROM banks WHERE id = '001';") -eq 1)

    # O saldo de outra conta so' existiu para provar o filtro da tbSaldos
    # (arquivo livre agora, com a aplicacao encerrada).
    Check 'saldo de outra conta removido (passo [10])' (
        [Sq]::Executar($minewDb,
            'DELETE FROM saldos WHERE account_id = 2;') -eq 1)

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
    # ... e o extrato de OUTRA conta, que a importacao tem de recusar
    if (Test-Path $ofxOutraConta) {
        Remove-Item $ofxOutraConta -Force -ErrorAction SilentlyContinue
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
