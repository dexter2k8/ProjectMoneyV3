unit unitMoney;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLite3Conn, SQLDB, Forms, Controls, Graphics, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls, Menus, DB, Grids, DBGrids, DBCtrls, Buttons,
  LCLType;

type

  { TMoneyGrid }

  // TDBGrid que também aceita digitação em campos ftMemo. Os TEXT do SQLite
  // viram ftMemo no driver (e ftMemo é blob para a grade): a TDBGrid de base
  // só entrega a tecla ao editor se dgDisplayMemoText estiver ligado - mas,
  // com a opção ligada, a célula lê o texto cru do campo e ignora o OnGetText
  // que formata (DD/MM/AAAA no fim da tbSaldos). Sem a opção a formatação
  // funciona e a digitação some; aqui vale a validação da base só sem essa
  // trava de blob, então a coluna de enddate (máscara !99/00/0000) edita e o
  // OnSetText continua gravando AAAA-MM-DD.
  TMoneyGrid = class(TDBGrid)
  protected
    function EditorCanAcceptKey(const Ch: TUTF8Char): boolean; override;
  end;

  { TFormMoney }

  TFormMoney = class(TForm)
    cbYear: TComboBox;
    cbAccount: TComboBox;
    DBNavBancos: TDBNavigator;
    DBNavContas: TDBNavigator;
    DBNavTrans: TDBNavigator;
    DBNavSaldos: TDBNavigator;
    gridContas: TDBGrid;
    gridSaldos: TMoneyGrid;
    gridTrans: TDBGrid;
    gridBancos: TDBGrid;
    lblTitle: TLabel;
    lblSaldo: TLabel;
    MainMenu: TMainMenu;
    miGerCon: TMenuItem;
    miGerSal: TMenuItem;
    miAbout: TMenuItem;
    mmHelp: TMenuItem;
    miImpSal: TMenuItem;
    miImpTrans: TMenuItem;
    miExit: TMenuItem;
    miList: TMenuItem;
    miClose: TMenuItem;
    miOpen: TMenuItem;
    miExpSal: TMenuItem;
    miExpTrans: TMenuItem;
    miImport: TMenuItem;
    mmTransactions: TMenuItem;
    miNew: TMenuItem;
    mmArquivo: TMenuItem;
    PageControl1: TPageControl;
    pnBancosControl: TPanel;
    pnContas: TPanel;
    pnSaldosControl: TPanel;
    pnHeader: TPanel;
    pnFooter: TPanel;
    sbtnVoltarContas: TSpeedButton;
    sbtnVoltarSaldos: TSpeedButton;
    Separator1: TMenuItem;
    Separator2: TMenuItem;
    Separator3: TMenuItem;
    Separator4: TMenuItem;
    sbtnVoltar: TSpeedButton;
    SQLite3ConnBancos: TSQLite3Connection;
    SQLTransactionBancos: TSQLTransaction;
    SQLQueryBanks: TSQLQuery;
    DataSourceBancos: TDataSource;
    // Database novo/aberto (menu "Novo Database" / "Abrir Database") -> aba
    // tbContas: mesma cadeia conexão -> transaction -> query -> datasource da
    // aba tbBancos; os dois menus só trocam o arquivo dessa mesma cadeia.
    // Extrato (miImport): arquivo OFC/OFX escolhido pelo usuario - o
    // conteudo vai para a tabela "extratos", na conta selecionada na tela.
    dlgImportar: TOpenDialog;
    dlgNewDatabase: TSaveDialog;
    dlgOpenDatabase: TOpenDialog;
    SQLite3ConnContas: TSQLite3Connection;
    SQLTransactionContas: TSQLTransaction;
    SQLQueryContas: TSQLQuery;
    DataSourceContas: TDataSource;
    SQLQuerySaldos: TSQLQuery;
    DataSourceSaldos: TDataSource;
    // Extratos (tela de extratos): mesma conexão das contas/saldos; a grade
    // gridTrans e o navegador DBNavTrans ligam-se a ela por este DataSource.
    SQLQueryExtratos: TSQLQuery;
    DataSourceExtratos: TDataSource;
    tbContas: TTabSheet;
    toggleShowControls: TToggleBox;
    tsAnterior: TStaticText;
    tbSaldos: TTabSheet;
    tbBancos: TTabSheet;
    tbDez: TTabSheet;
    tbNov: TTabSheet;
    tbOut: TTabSheet;
    tbSet: TTabSheet;
    tbAgo: TTabSheet;
    tbJul: TTabSheet;
    tbJun: TTabSheet;
    tbMai: TTabSheet;
    tbAbr: TTabSheet;
    tbMar: TTabSheet;
    tbFev: TTabSheet;
    tbJan: TTabSheet;
    tslblAnterior: TStaticText;
    txtSaldo: TLabel;
    // Banco da conta na grade de contas: a coluna bankid edita por combo com
    // os nomes de tbBancos (ButtonStyle = cbsPickList) e grava o id. O
    // DBGrid grava célula via Field.Text, que dispara o OnSetText — é ali que
    // o nome escolhido vira id; o OnGetText faz o caminho de volta na
    // exibição (e no texto inicial do combo, que abre no nome certo).
    procedure BancoIdGetText(Sender: TField; var AText: string;
      DisplayText: Boolean);
    procedure BancoIdSetText(Sender: TField; const AText: string);
    procedure cbAccountChange(Sender: TObject);
    procedure cbYearChange(Sender: TObject);
    // Formatação da grade de extratos: os campos TEXT do SQLite viram ftMemo,
    // então a exibição passa pelo DisplayText, que é onde o OnGetText entra.
    // (Sem isso a data aparece como 20250305 e o valor sem C/D.)
    procedure ExtratoDtpostedGetText(Sender: TField; var AText: string;
      DisplayText: Boolean);
    procedure ExtratoTextoGetText(Sender: TField; var AText: string;
      DisplayText: Boolean);
    procedure ExtratoTrnAmtGetText(Sender: TField; var AText: string;
      DisplayText: Boolean);
    procedure FormCreate(Sender: TObject);
    // O combo da coluna "Banco" é reabastecido a cada edição: a lista de
    // bancos muda em tbBancos e o PickList de projeto é fixo. OnSelectEditor
    // roda depois do preenchimento padrão do LCL, então dá para trocar os
    // itens na hora.
    procedure gridContasSelectEditor(Sender: TObject; Column: TColumn;
      var Editor: TWinControl);
    // Cor do valor na grade de extratos (crédito azul, débito vermelho): roda
    // antes do desenho padrão da célula, que é quem pinta o texto.
    procedure gridTransPrepareCanvas(Sender: TObject; DataCol: Integer;
      Column: TColumn; AState: TGridDrawState);
    procedure miCloseClick(Sender: TObject);
    procedure miGerConClick(Sender: TObject);
    procedure miGerSalClick(Sender: TObject);
    procedure miImportClick(Sender: TObject);
    procedure miListClick(Sender: TObject);
    procedure miNewClick(Sender: TObject);
    procedure miOpenClick(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure sbtnVoltarClick(Sender: TObject);
    // enddate da tbSaldos: gravado em ISO (AAAA-MM-DD, texto puro no SQLite)
    // e exibido/editado como DD/MM/AAAA - a coluna tem EditMask, que insere
    // as barras sozinhas na digitação. Ao contrário do dtposted dos extratos,
    // o texto do editor (DisplayText=False) TAMBÉM vem formatado: quem dita a
    // forma do que se digita é a máscara, e ela espera DD/MM/AAAA.
    procedure SaldosEnddateGetText(Sender: TField; var AText: string;
      DisplayText: Boolean);
    // Gravação da célula: o DBGrid faz Field.Text := texto do editor (que sai
    // da máscara como DD/MM/AAAA) e o SetEditText do campo vem parar aqui.
    // Vazio limpa a data (fica '' - enddate é NOT NULL, então NULL daria
    // erro de constraint); texto que não seja uma data válida não mexe no
    // valor, como no banco de tbBancos.
    procedure SaldosEnddateSetText(Sender: TField; const AText: string);
    // Liga os handlers de banco (OnGetText/OnSetText) a cada Open da query de
    // contas: os campos são dinâmicos e cada Open os recria sem handlers —
    // o mesmo problema, e a mesma solução, dos extratos (PrepararCamposExtratos).
    procedure SQLQueryContasAfterOpen(DataSet: TDataSet);
    // Registro novo da tbSaldos recebe a conta escolhida no cbAccount (a
    // coluna account_id está oculta na grade - é ela que filtra).
    procedure SQLQuerySaldosNewRecord(DataSet: TDataSet);
    procedure toggleShowControlsClick(Sender: TObject);
  private
    // Estado das guias antes de navegar para uma tela de gestão (para o
    // "Voltar" devolver - as guias de mês são dinâmicas e o "Voltar" as
    // recalcula pelo saldo do ano, em vez de só restaurar este estado).
    FSavedTabVisible: array of Boolean;
    FInListView: Boolean;
    // Título da tela normal (mês) antes de entrar na lista de bancos
    FSavedTitle: String;
    // Base do miList/miGerCon: guarda guias + título e ativa a aba com as
    // guias ocultas (o "Voltar" devolve o estado salvo aqui).
    procedure NavigateToTab(ATab: TTabSheet);
    // Tela de transações (a de entrada): devolve as guias que a navegação de
    // gestão escondeu, volta para tbJan e recalcula as guias de mês pelo dado
    // gravado. É o que o "Voltar" executa e o que miNew/miOpen também têm de
    // executar: as telas de gestão ficam disponíveis sem database, então dá
    // para abrir um arquivo estando nelas - e, sem este passo, a revelação
    // mostraria aquela tela, sem o gridTrans (que só existe na tela de
    // extratos).
    procedure IrParaTelaExtratos;
    // Devolve para o TOPO da pilha de janelas os controles do FORM que
    // sobrepõem a área das guias na tela de extratos: o gridTrans e os dois
    // textos do saldo anterior (tslblAnterior, o rótulo "Anterior:", e
    // tsAnterior, o campo do valor), que ficam na faixa das guias. Cada
    // mexida nas guias (mostrar a PageControl1 na revelação, mexer no
    // TabVisible das guias de mês) coloca a PageControl1 ACIMA deles na pilha
    // de irmãos - e é justamente o que cobre a área deles. Sem este empurrão
    // eles ficam com WS_VISIBLE, só que escondidos atrás das guias: a tela
    // aparece "sem grade nenhuma" e sem o "Anterior" para o usuário.
    procedure GarantirSobreposicaoNaFrente;
    // Liga/desliga os blocos de interface (cabeçalho, guias, rodapé e textos
    // "Anterior"). Revela em: navegação (miList/miGerCon), miNew e miOpen;
    // esconde em: inicialização, "Voltar" sem database e miClose. No revelar
    // a régua da página ativa é reaplicada.
    procedure SetInterfaceVisible(AVisible: Boolean);
    // "Novo Database"/"Abrir Database" ligam a query de contas e "Fechar
    // Database" a desliga: é esse estado que manda na interface.
    function DatabaseAberto: Boolean;
    // Aplica o estado acima de uma vez: visibilidade da página + enable do
    // "Fechar Database" (sem database aberto não há nada para fechar).
    procedure AtualizarEstadoDatabase;
    // Combos de filtro da tela de extratos (conta e ano): montadas por
    // miNew/miOpen a partir do database aberto, refazem a query da grade
    // quando o usuário troca de seleção e são esvaziadas ao fechar.
    // A lista de anos vem de saldos.enddate (o filtro continua em dtposted).
    procedure CarregarFiltrosExtratos;
    procedure AplicarFiltroExtratos;
    // Mês (da guia ativa, 1..12) e ano (do cbYear) que definem o período
    // que a tela de extratos está mostrando. Devolve False quando ainda não
    // há o que mostrar (sem database, sem guia de mês ativa ou sem ano
    // escolhido - é o estado do fechar, onde a tela já está escondida).
    function PeriodoExtratos(out ano, mes: Integer): Boolean;
    // Lê o balance do registro de saldos mais recente (maior enddate, com
    // empate de data desempatado pelo maior id) que satisfaz CondicaoPeriodo,
    // sempre na conta do cbAccount e com o mesmo CAST AS REAL da grade - 0
    // quando não há registro. Consulta própria na mesma conexão dos filtros:
    // não mexe no cursor de nenhuma grade.
    function ConsultarSaldo(const CondicaoPeriodo: string): Double;
    // Exibe um saldo no mesmo par C/D da coluna de valor da grade
    // (ExtratoTrnAmtGetText): a letra assume o papel do sinal e o valor é o
    // absoluto, com 2 casas e vírgula. A cor acompanha o sinal, MAS só o
    // negativo é vermelho: o positivo fica na cor do próprio campo. Usado
    // pelo tsAnterior e pelo txtSaldo (TStaticText e TLabel - os dois têm
    // Caption e Font no TControl).
    procedure ExibirSaldo(Campo: TControl; valor: Double);
    // tsAnterior = abertura do mês ativo: balance do saldo IMEDIATAMENTE
    // ANTERIOR ao 1º dia do mês da guia ativa (no ano do cbYear e na conta
    // do cbAccount - mesma regra dos filtros); zero quando não existe. Roda
    // no fim do AplicarFiltroExtratos, que é o único ponto que muda com
    // guia, ano, conta, abertura e importação.
    procedure AtualizarSaldoAnterior;
    // txtSaldo = fechamento do mês ativo ("Saldo:" no rodapé): balance do
    // saldo MAIS RECENTE gravado no próprio mês da guia ativa (mesmo ano,
    // conta e gatilho do AtualizarSaldoAnterior); zero quando não há saldo
    // lançado no mês.
    procedure AtualizarSaldoMes;
    // Guias de mês da tela de extratos (1 = tbJan ... 12 = tbDez): aparecem
    // só nos meses que têm saldo no ano do cbYear, com a aba ativa no mês
    // mais recente. Roda ao carregar os filtros, ao trocar o ano e no
    // "Voltar"; sem database a chamada não faz nada (é o estado do "Fechar
    // Database", que esconde a interface e deixa a próxima abertura
    // recalcular). É o único ponto que liga/desliga essas guias fora da
    // navegação de gestão (NavigateToTab esconde todas).
    procedure AtualizarAbasMes;
    function AbaDoMes(AMes: Integer): TTabSheet;
    // Tradutor inverso de AbaDoMes: a guia ATIVA diz qual mês a grade de
    // extratos mostra (1 = tbJan ... 12 = tbDez). Devolve 0 para o que não é
    // guia de mês (as telas de gestão), onde a grade está escondida e não há
    // mês a filtrar.
    function MesDaAba(ATab: TTabSheet): Integer;
    // Grade de saldos da tbSaldos: mesma regra do filtro de extratos - só os
    // registros da conta escolhida no cbAccount (combo vazia não restringe).
    // É o único ponto de Open de SQLQuerySaldos: quem abre/reabre é aqui.
    procedure AplicarFiltroSaldos;
    // Liga os handlers de exibição da grade de extratos (data, texto memo e
    // valor). Os campos da query são dinâmicos: cada Open os recria sem os
    // handlers, então a ligação é refeita logo depois de abrir — e abrir é
    // sempre por aqui (AplicarFiltroExtratos é o único ponto de Open).
    procedure PrepararCamposExtratos;
    // Liga a máscara de data e os handlers de enddate da tbSaldos (mesmo
    // motivo e mesmo ponto dos extratos: os campos são recriados a cada Open
    // de SQLQuerySaldos, que só acontece em AplicarFiltroSaldos).
    procedure PrepararCamposSaldos;
    procedure LimparFiltrosExtratos;
    // Remonta o combo de contas direto do database, mantendo a conta que o
    // usuário tinha escolhido (pelo id, para sobreviver à edição do texto).
    // É o caminho do "Voltar" das telas de gestão: a tbContas inclui/exclui
    // "contas" e o combo só existe na tela de extratos, então sem remontar
    // aqui a troca só aparece no próximo miNew/miOpen/miImport.
    procedure AtualizarComboContas;
    // Lista de bancos de tbBancos (SQLQueryBanks, aberta desde a abertura do
    // programa) no formato "nome=id": alimenta o combo da coluna "Banco" e
    // as traduções id<->nome. Percorre a query guardando e devolvendo a
    // posição, para a grade de bancos não mudar de linha selecionada.
    procedure MontarListaBancos(ALista: TStrings);
    function BancoNomePorId(const AId: Integer): string;
    function BancoIdPorNome(const ANome: string): Integer;

  public

  end;

var
  FormMoney: TFormMoney;

implementation

uses
  unitDatabase, unitOfx;

{$R *.lfm}

{ TMoneyGrid }

// Mesmo critério da TDBGrid.EditorCanAcceptKey, pulando só a exigência de
// blob: campo ftMemo com OnGetText (o caso do enddate) é editável como texto
// comum - a formatação da célula é do OnGetText, não do dgDisplayMemoText.
function TMoneyGrid.EditorCanAcceptKey(const Ch: TUTF8Char): boolean;
var
  campo: TField;
begin
  Result := inherited EditorCanAcceptKey(Ch);
  if Result or (Ch = '') then
    Exit;
  campo := SelectedField;
  if (campo = nil) or (campo.DataType <> ftMemo) or campo.Calculated or
     (campo.FieldKind = fkLookup) then
    Exit;
  // IsValidChar é o mesmo da base: dígito, barra ou backspace, byte a byte.
  Result := IsValidChar(campo, Ch);
end;

{ TFormMoney }

// "1 registro" / "2 registros": evita "registro(s)" na mensagem de resultado.
function Plural(AN: Integer; const ASingular, APlural: string): string;
begin
  if AN = 1 then
    Result := ASingular
  else
    Result := APlural;
end;

procedure TFormMoney.cbAccountChange(Sender: TObject);
begin
  // Trocou a conta escolhida: as duas grades que separam por conta reabrem
  // já filtradas - a de extratos e a de saldos (o cbAccount também fica
  // visível na tela de Gerenciar Saldos, que é filtrada por ele). Sem
  // database aberto não há o que fazer, e é esse o caminho do OnChange que a
  // limpeza das combos dispara.
  AplicarFiltroExtratos;
  AplicarFiltroSaldos;
end;

procedure TFormMoney.cbYearChange(Sender: TObject);
begin
  // Mesmo caminho do cbAccountChange: mudou o ano, refaz o filtro da grade.
  AplicarFiltroExtratos;
  // ... e as guias de mês: elas são os meses com saldo NO ANO escolhido
  // (saldos.enddate), então trocar o ano muda o cabeçalho da tela.
  AtualizarAbasMes;
end;

procedure TFormMoney.FormCreate(Sender: TObject);
begin
  try
    // O caminho do designer é absoluto e só vale na sua máquina: aqui o
    // arquivo é sempre o banks.db que está ao lado do executável.
    SQLite3ConnBancos.DatabaseName := DatabasePath;
    // Abrir a query dispara Prepare -> MaybeConnect + MaybeStartTransaction
    // (sqldb.pp:1295-1297), então não é preciso ligar a transação na mão.
    SQLQueryBanks.Open;
  except
    on E: Exception do
      MessageDlg('Não foi possível carregar os dados de "' + DatabaseFileName +
        '".' + LineEnding + E.Message, mtError, [mbOK], 0);
  end;
  // Captura o título ANTES de qualquer OnChange: a troca de página usa
  // FSavedTitle para restaurar lblTitle.
  FSavedTitle := lblTitle.Caption;
  // ActivePage no .lfm é apenas a página que estava selecionada no designer e
  // muda a cada salvamento (já foi tbJan, tbSaldos e tbBancos). Fixando aqui a
  // tela de entrada, a aplicação nunca inicializa na lista de bancos.
  PageControl1.ActivePage := tbJan;
  // Durante o streaming o OnChange não dispara (csLoading): aplica a regra da
  // lista de bancos aqui também (é idempotente).
  PageControl1Change(nil);
  // Por último (senão o PageControl1Change acima reexibiria): na tela inicial
  // não há database aberto, então fica só com o MainMenu e com o "Fechar
  // Database" desativado — a interface aparece na navegação e no miNew/miOpen.
  AtualizarEstadoDatabase;
end;

procedure TFormMoney.PageControl1Change(Sender: TObject);
var
  exibindoBancos: Boolean;
  exibindoContas: Boolean;
  exibindoSaldos: Boolean;
  exibindoExtrato: Boolean;
begin
  exibindoBancos := (PageControl1.ActivePage = tbBancos);
  exibindoContas := (PageControl1.ActivePage = tbContas);
  exibindoSaldos := (PageControl1.ActivePage = tbSaldos);
  // Rodapé e textos "Anterior" são da tela de extratos: somem nas três telas
  // de gestão (lista de bancos, contas e saldos) — mesma regra para todas.
  exibindoExtrato := not exibindoBancos and not exibindoContas and
    not exibindoSaldos;

  // O painel inteiro some (e não só os itens): como PageControl1 é alClient,
  // os 50px do rodapé são realinhados para a guia ativa e a grade cresce.
  pnFooter.Visible := exibindoExtrato;
  // A grade de transações é controle do FORM (também alClient, com os 24px
  // de espaço que deixam o "Anterior" à mostra): ela sobrepõe a área das
  // guias, então nasceu oculta no .lfm e é revelada aqui — na tela de
  // extratos e só com database aberto. Sem conexão não há o que mostrar, e
  // o FormCreate já entra por aqui com tbJan ativa e sem database.
  gridTrans.Visible := exibindoExtrato and DatabaseAberto;
  // ... mas o DBNavTrans mantém a regra própria (toggleShowControls) para,
  // quando o painel voltar, o navegador respeitar o toggle.
  DBNavTrans.Visible := exibindoExtrato and toggleShowControls.Checked;

  // Combos do cabeçalho: o de ano é só da tela de extratos; o de conta
  // continua visível na tbSaldos, porque é ele que filtra a grade de
  // saldos. Somem nas telas de lista de bancos e de contas, junto com os
  // textos "Anterior".
  cbYear.Visible := exibindoExtrato;
  cbAccount.Visible := exibindoExtrato or exibindoSaldos;
  tsAnterior.Visible := exibindoExtrato;
  tslblAnterior.Visible := exibindoExtrato;

  // Título principal: troca para "Lista de Bancos"/"Gerenciamento de Contas"
  // conforme a página ativa e volta ao anterior quando a página é a de entrada
  // (o "Voltar" devolve tbJan, então FSavedTitle é restaurado ali).
  if exibindoBancos then
    lblTitle.Caption := 'Lista de Bancos'
  else if exibindoContas then
    lblTitle.Caption := 'Gerenciamento de Contas'
  else if exibindoSaldos then
    lblTitle.Caption := 'Gerenciamento de Saldos'
  else
    lblTitle.Caption := FSavedTitle;

  // A grade mostra o mês da guia ativa (mais o ano do cbYear e a conta do
  // cbAccount): clicar em outra guia muda a página, então refaz o filtro
  // aqui. Só na tela de extratos - nas telas de gestão a grade está escondida
  // e MesDaAba devolve 0; o "Voltar" refaz o filtro pelo AtualizarAbasMes.
  if exibindoExtrato then
    AplicarFiltroExtratos;

  // Por último: revelar a interface (SetInterfaceVisible mostra a
  // PageControl1) é o que a coloca no TOPO da pilha de irmãos, por cima da
  // grade e dos textos "Anterior" - sem o empurrão, a página vem cobrindo
  // tudo o que o form desenha sobre a área das guias, e o usuário vê a tela
  // de extratos sem grade nenhuma e sem o saldo anterior (os dados estão lá,
  // atrás).
  GarantirSobreposicaoNaFrente;
end;

procedure TFormMoney.GarantirSobreposicaoNaFrente;
begin
  // Só quando a tela de extratos está de pé (a grade revelada é o sinal, e a
  // página tem de estar visível junto): nas telas de gestão os três estão
  // escondidos, e na inicialização a página também - não há o que empurrar.
  if not (gridTrans.Visible and PageControl1.Visible) then
    Exit;
  // Os textos ocupam a fileira de cima (Top=53, na faixa das guias) e a grade
  // vem logo abaixo (Top=76): nenhum cobre o outro, então o que importa é
  // estarem todos ACIMA da PageControl1 - a ordem entre eles é indiferente.
  tslblAnterior.BringToFront;
  tsAnterior.BringToFront;
  gridTrans.BringToFront;
end;

procedure TFormMoney.NavigateToTab(ATab: TTabSheet);
var
  i: Integer;
begin
  // Guarda como as guias e o título estavam (só na primeira entrada) para o
  // "Voltar" devolver.
  if not FInListView then
  begin
    SetLength(FSavedTabVisible, PageControl1.PageCount);
    for i := 0 to PageControl1.PageCount - 1 do
      FSavedTabVisible[i] := PageControl1.Pages[i].TabVisible;
    FSavedTitle := lblTitle.Caption;
    FInListView := True;
  end;
  // Esconde as guias ANTES de ativar a de destino: com as guias de mês
  // ligadas (elas aparecem quando há saldo), ativar primeiro faria o LCL
  // trocar de página sozinho quando a guia escondida fosse a ativa
  // (TTabSheet.SetTabVisible -> PageRemoved -> próxima guia visível).
  // Escondidas todas, ativar a página de gestão e por último revelar a
  // interface - a revelação é que roda a régua da tela pela página ativa.
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := False;
  PageControl1.ActivePage := ATab;
  // Revela SEMPRE e por último: o "Fechar Database" esconde a interface até
  // com a lista na tela, e navegar de novo tem de trazê-la de volta. Como é
  // a última linha, a régua vale a página recém-ativada (mesmo quando o
  // OnChange não dispara porque a página não mudou).
  SetInterfaceVisible(True);
end;

procedure TFormMoney.SetInterfaceVisible(AVisible: Boolean);
begin
  pnHeader.Visible := AVisible;
  PageControl1.Visible := AVisible;
  if AVisible then
    // Revelar não é só levantar os blocos: na lista de bancos e na aba de
    // contas o rodapé, os combos e o "Anterior" continuam escondidos.
    PageControl1Change(nil)
  else
  begin
    pnFooter.Visible := False;
    // gridTrans é do form, não da PageControl1: esconder a guia não a
    // esconde, então o estado "só o menu" precisa desligá-la na mão.
    gridTrans.Visible := False;
    tsAnterior.Visible := False;
    tslblAnterior.Visible := False;
  end;
end;

function TFormMoney.DatabaseAberto: Boolean;
begin
  // É a query da tbContas abrindo que define "tem database": ela só liga no
  // miNew/miOpen e desliga no miClose (o banks.db do menu não conta — ele
  // abre sozinho na inicialização).
  Result := SQLQueryContas.Active;
end;

procedure TFormMoney.AtualizarEstadoDatabase;
var
  aberto: Boolean;
begin
  aberto := DatabaseAberto;
  SetInterfaceVisible(aberto);
  // Sem database as combos de filtro ficam sem itens: não há o que escolher
  // e não pode sobrar lista do arquivo que estava aberto antes.
  if not aberto then
    LimparFiltrosExtratos;
  // Sem database aberto não há nada para fechar: o item fica desativado.
  miClose.Enabled := aberto;
  // Todo o menu "Transações" (importar, exportar, gerenciar) opera sobre o
  // database de contas: sem ligação, não há transações a fazer. Desativar o
  // menu de primeiro nível cobre os itens de uma vez - inclusive os futuros.
  mmTransactions.Enabled := aberto;
end;

// Remonta o combo de contas a partir do database. A consulta é própria (não
// percorre a query da tbContas) porque percorrê-la levaria o cursor da grade
// a cada "Voltar" e derrubaria uma linha que esteja em edição — a lista que
// importa aqui é a que está gravada, não o buffer da grade.
procedure TFormMoney.AtualizarComboContas;
var
  consulta: TSQLQuery;
  idConta, idEscolhido, indice, i: Integer;
  texto, descricao: string;
begin
  if not DatabaseAberto then
    Exit;

  // Guarda a escolha antes de esvaziar. Na montagem dos filtros a combo já
  // veio vazia do chamador, então não há o que guardar e vale a primeira
  // conta — comportamento do miNew/miOpen, do qual os testes dependem.
  idEscolhido := -1;
  if (cbAccount.ItemIndex >= 0) and (cbAccount.ItemIndex < cbAccount.Items.Count) then
    idEscolhido := Integer(PtrInt(cbAccount.Items.Objects[cbAccount.ItemIndex]));

  cbAccount.Items.Clear;
  consulta := TSQLQuery.Create(nil);
  try
    consulta.Database := SQLite3ConnContas;
    consulta.Transaction := SQLTransactionContas;
    consulta.SQL.Text := 'SELECT id, acctid, description FROM contas' +
      ' ORDER BY id;';
    consulta.Open;
    while not consulta.EOF do
    begin
      // Uma linha por registro, com o "id" em Items.Objects (é o account_id
      // dos extratos). O texto junta o número da conta e a descrição; sem
      // nada usável, fica a identificação interna.
      idConta := consulta.FieldByName('id').AsInteger;
      texto := Trim(consulta.FieldByName('acctid').AsString);
      descricao := Trim(consulta.FieldByName('description').AsString);
      if descricao <> '' then
      begin
        if texto <> '' then
          texto := texto + ' - ';
        texto := texto + descricao;
      end;
      if texto = '' then
        texto := 'Conta ' + IntToStr(idConta);
      cbAccount.Items.AddObject(texto, TObject(PtrInt(idConta)));
      consulta.Next;
    end;
    consulta.Close;
  finally
    consulta.Free;
  end;

  // De volta à conta que o usuário tinha escolhido — por id, porque é o que
  // sobrevive à edição do texto da própria conta. Sem correspondência
  // (primeira carga, conta excluída) entra a primeira da lista.
  indice := 0;
  if idEscolhido >= 0 then
    for i := 0 to cbAccount.Items.Count - 1 do
      if Integer(PtrInt(cbAccount.Items.Objects[i])) = idEscolhido then
      begin
        indice := i;
        Break;
      end;
  if cbAccount.Items.Count > 0 then
    cbAccount.ItemIndex := indice;
end;

procedure TFormMoney.CarregarFiltrosExtratos;
var
  consulta: TSQLQuery;
  ano: Integer;
  texto: string;
begin
  if not DatabaseAberto then
    Exit;

  // Esvazia as duas antes de montar: um OnChange no meio do caminho não
  // pode reabrir a query com conta ou ano herdados do database anterior.
  cbAccount.Items.Clear;
  cbYear.Items.Clear;

  // Conta: combo recém-esvaziada, então não há seleção antiga para manter e
  // entra a primeira — é o miNew/miOpen de sempre (e é isso que o "Voltar"
  // reexecuta via AtualizarComboContas, mas aí com seleção a preservar).
  AtualizarComboContas;

  // Ano: consulta direta na mesma conexão, porque a query da grade ainda não
  // abriu — e ela justamente depende do filtro que está sendo montado aqui.
  // A lista vem de "saldos" (enddate), não de "extratos": é o pedido — os
  // anos mostrados refletem o fim de período dos saldos gravados, do MAIS
  // RECENTE para o mais antigo, e é o primeiro item que fica selecionado ao
  // abrir. O filtro em si continua em dtposted (a grade é de extratos); se
  // os dois conjuntos de anos divergirem, vale o que o combo listar.
  // Só entra ano de verdade (4 dígitos): enddate é texto, qualquer outra
  // coisa nessa posição não é ano para ninguém.
  consulta := TSQLQuery.Create(nil);
  try
    consulta.Database := SQLite3ConnContas;
    consulta.Transaction := SQLTransactionContas;
    consulta.SQL.Text :=
      'SELECT DISTINCT substr(enddate, 1, 4) FROM saldos' +
      ' ORDER BY substr(enddate, 1, 4) DESC;';
    consulta.Open;
    while not consulta.EOF do
    begin
      texto := Trim(consulta.Fields[0].AsString);
      if (Length(texto) = 4) and TryStrToInt(texto, ano) then
        cbYear.Items.Add(texto);
      consulta.Next;
    end;
    consulta.Close;
  finally
    consulta.Free;
  end;

  // Sem ano nenhum (saldos vazio, logo depois do "Novo Database") o combo
  // não pode ficar em branco: entra o ano atual e o filtro segue de pé.
  if cbYear.Items.Count = 0 then
    cbYear.Items.Add(FormatDateTime('yyyy', Date));
  cbYear.ItemIndex := 0;

  // Guias de mês do ano acabado de escolher (o ItemIndex acima costuma
  // disparar o OnChange, mas a chamada explícita cobre o caso dele não
  // disparar): só os meses com saldo aparecem e a aba ativa vai para o
  // mais recente.
  AtualizarAbasMes;

  // Garante a seleção final mesmo se nenhuma das combos disparar OnChange
  // ao ganhar o primeiro ItemIndex.
  AplicarFiltroExtratos;
end;

// Guia de mês da tela de extratos: 1 = tbJan ... 12 = tbDez. Os nomes não
// seguem padrão (tbFev, tbAbr...), então o case é o tradutor - e ele também
// protege o chamador de um mês vindo fora do intervalo (devolve nil).
function TFormMoney.AbaDoMes(AMes: Integer): TTabSheet;
begin
  case AMes of
    1: Result := tbJan;
    2: Result := tbFev;
    3: Result := tbMar;
    4: Result := tbAbr;
    5: Result := tbMai;
    6: Result := tbJun;
    7: Result := tbJul;
    8: Result := tbAgo;
    9: Result := tbSet;
    10: Result := tbOut;
    11: Result := tbNov;
    12: Result := tbDez;
  else
    Result := nil;
  end;
end;

// Tradutor inverso de AbaDoMes: a aba ATIVA é quem diz qual mês a grade de
// extratos mostra (o pedido: ano do cbYear + mês da guia ativa). Devolve 0
// para o que não é guia de mês - as telas de gestão, onde a grade fica
// escondida e não há mês a filtrar. Os comparadores são objetos (TTabSheet
// não é ordinal), então não há case aqui: a cadeia é a tradução direta.
function TFormMoney.MesDaAba(ATab: TTabSheet): Integer;
begin
  if ATab = tbJan then
    Result := 1
  else if ATab = tbFev then
    Result := 2
  else if ATab = tbMar then
    Result := 3
  else if ATab = tbAbr then
    Result := 4
  else if ATab = tbMai then
    Result := 5
  else if ATab = tbJun then
    Result := 6
  else if ATab = tbJul then
    Result := 7
  else if ATab = tbAgo then
    Result := 8
  else if ATab = tbSet then
    Result := 9
  else if ATab = tbOut then
    Result := 10
  else if ATab = tbNov then
    Result := 11
  else if ATab = tbDez then
    Result := 12
  else
    Result := 0;
end;

// As guias JAN..DEZ são o cabeçalho de navegação da tela de extratos e só
// aparecem nos meses que têm saldo no ano do cbYear (mesma origem da lista
// de anos: saldos.enddate). A aba ativa vai para o mês mais recente - é o
// que define "hoje" quando o database abre e quando o ano muda.
procedure TFormMoney.AtualizarAbasMes;
var
  i, ano, mes, ultimoMes, mesLancado, mesAbertura: Integer;
  condicao: string;
  temMes: array[1..12] of Boolean;
  consulta: TSQLQuery;
begin
  // Sem database não há guia de mês para mostrar — e é justamente o estado do
  // "Fechar Database", onde a interface já foi escondida ANTES de as combos
  // esvaziarem (AtualizarEstadoDatabase). Mexer nas guias agora dispararia o
  // PageControl1Change na hora errada, e ele reergueria o rodapé e os textos
  // "Anterior" que o SetInterfaceVisible(False) acabou de esconder. As guias
  // voltam a ser calculadas na próxima abertura (CarregarFiltrosExtratos) e a
  // navegação esconde todas (NavigateToTab).
  if not DatabaseAberto then
    Exit;

  for i := 1 to 12 do
    temMes[i] := False;

  // Nas telas de gestão quem manda na guia é NavigateToTab (esconde todas e
  // ativa a página pedida, deixando o cabeçalho de mês para o "Voltar").
  // Esse ajuste é da tela de extratos - que é onde o cbYear vive.
  if (PageControl1.ActivePage = tbBancos) or
     (PageControl1.ActivePage = tbContas) or
     (PageControl1.ActivePage = tbSaldos) then
    Exit;

  // Ano do combo, com o mesmo critério de "4 dígitos" do preenchimento dele:
  // sem ano escolhido (combo vazia) não há mês para mostrar. Database já
  // conferido no início do procedimento.
  ano := -1;
  if (cbYear.ItemIndex >= 0) and (cbYear.ItemIndex < cbYear.Items.Count) then
    ano := StrToIntDef(Trim(cbYear.Items[cbYear.ItemIndex]), -1);

  mesLancado := 0;
  if ano >= 0 then
  begin
    // Consulta própria na mesma conexão (mesmo caminho do ano do cbYear):
    // meses do ano escolhido, direto de enddate. enddate é texto ISO, então
    // o mês são os caracteres 6..7 ("2026-03-31" -> "03").
    consulta := TSQLQuery.Create(nil);
    try
      consulta.Database := SQLite3ConnContas;
      consulta.Transaction := SQLTransactionContas;
      consulta.SQL.Text :=
        'SELECT DISTINCT substr(enddate, 6, 2) FROM saldos' +
        ' WHERE substr(enddate, 1, 4) = ' + QuotedStr(IntToStr(ano)) + ';';
      consulta.Open;
      while not consulta.EOF do
      begin
        mes := StrToIntDef(Trim(consulta.Fields[0].AsString), 0);
        if (mes >= 1) and (mes <= 12) then
          temMes[mes] := True;
        consulta.Next;
      end;
      consulta.Close;

      // Mês do lançamento mais recente do ano: a tela de extratos tem de
      // ABRIR no mês em que há o que ver - o último mês com saldo podia vir
      // sem nenhum lançamento e a grade abria vazia. Mesma conta do filtro
      // dos extratos (combo vazia não restringe), senão um lançamento de
      // outra conta mudaria a abertura de uma tela que não o mostra.
      // dtposted tem dois formatos no mesmo campo (ISO "YYYY-MM-DD" e OFX
      // "YYYYMMDD"), então o mês sai do MESMO CASE do filtro - e o ORDER BY
      // compara o ANO-MES já normalizado: em texto puro "2025-12" vem depois
      // de "202512", então o formatado não pode competir com o cru.
      condicao := '';
      if (cbAccount.ItemIndex >= 0) and
         (cbAccount.ItemIndex < cbAccount.Items.Count) then
        condicao := ' AND account_id = ' + IntToStr(
          Integer(PtrInt(cbAccount.Items.Objects[cbAccount.ItemIndex])));
      consulta.SQL.Text :=
        'SELECT CASE WHEN substr(dtposted, 5, 1) = ''-'' THEN' +
        ' substr(dtposted, 6, 2) ELSE substr(dtposted, 5, 2) END' +
        ' FROM extratos WHERE substr(dtposted, 1, 4) = ' +
        QuotedStr(IntToStr(ano)) + condicao +
        ' ORDER BY substr(dtposted, 1, 4) ||' +
        ' (CASE WHEN substr(dtposted, 5, 1) = ''-'' THEN substr(dtposted, 6, 2)' +
        ' ELSE substr(dtposted, 5, 2) END) DESC LIMIT 1;';
      consulta.Open;
      if not consulta.EOF then
        mesLancado := StrToIntDef(Trim(consulta.Fields[0].AsString), 0);
      consulta.Close;
    finally
      consulta.Free;
    end;
  end;

  // As guias de gestão nunca têm cabeçalho (não é por aqui que elas
  // aparecem): esconder aqui também tira do estado a dependência do .lfm,
  // que o IDE pode reescrever ao salvar.
  tbBancos.TabVisible := False;
  tbContas.TabVisible := False;
  tbSaldos.TabVisible := False;

  // Esconde/mostra as doze e anota o ÚLTIMO mês com saldo - como as guias
  // JAN..DEZ estão em ordem crescente, o último é o mais recente.
  // Ordem importa: esconder a guia ativa faz o LCL procurar outra página
  // visível (TTabSheet.SetTabVisible -> PageRemoved), então a aba final é
  // escolhida depois, com todas as guias já no lugar.
  ultimoMes := 0;
  for i := 1 to 12 do
  begin
    if temMes[i] then
      ultimoMes := i;
    AbaDoMes(i).TabVisible := temMes[i];
  end;

  // Aba de abertura: a do mês com LANÇAMENTOS quando esse mês tem guia (o
  // pedido - a grade tem de abrir com o que ver); sem lançamento no ano (ou
  // com ele num mês sem saldo, que não vira guia), fica a do último mês com
  // saldo - o comportamento de antes.
  mesAbertura := ultimoMes;
  if (mesLancado >= 1) and (mesLancado <= 12) and temMes[mesLancado] then
    mesAbertura := mesLancado;

  if mesAbertura > 0 then
    PageControl1.ActivePage := AbaDoMes(mesAbertura)
  else if PageControl1.ActivePage = nil then
    // Sem saldo no ano a última guia some e o LCL deixa a página sem guia
    // ativa (FPageIndex = -1). Volta para tbJan, que é a página de entrada.
    PageControl1.ActivePage := tbJan;

  // Esconder/mostrar guia por guia acima reorganiza a pilha de janelas e
  // pode deixar a PageControl1 por cima da sobreposição (grade + "Anterior")
  // mesmo quando a aba ativa NÃO muda (aí o OnChange não dispara e o
  // PageControl1Change não roda) - o empurrão aqui cobre esse caminho também.
  GarantirSobreposicaoNaFrente;

  // Mesmo raciocínio para o FILTRO da grade: esconder a guia ativa faz o LCL
  // procurar outra página visível lá dentro (SetTabVisible -> PageRemoved) e
  // nem esse caminho dispara OnChange, então a aba que sobrou pode ter mudado
  // sem o PageControl1Change rodar - e a grade mostra o mês da aba ativa.
  // Refazer o filtro aqui, com a guia já no lugar, cobre esse caso (é
  // redundante quando o OnChange dispara, e inofensivo: é só reabrir a query).
  AplicarFiltroExtratos;
end;

procedure TFormMoney.AplicarFiltroExtratos;
var
  sql, condicao: string;
  idConta, ano, mes: Integer;
begin
  // Sem database aberto não há o que filtrar: é o caminho do OnChange que a
  // limpeza das combos dispara ao fechar o arquivo.
  if not DatabaseAberto then
    Exit;

  condicao := '';
  // Conta: o combo carrega o "id" de "contas" em Items.Objects. Combo vazia
  // (nenhuma conta cadastrada) não restringe e valem todas as contas.
  if (cbAccount.ItemIndex >= 0) and (cbAccount.ItemIndex < cbAccount.Items.Count) then
  begin
    idConta := Integer(PtrInt(cbAccount.Items.Objects[cbAccount.ItemIndex]));
    condicao := 'account_id = ' + IntToStr(idConta);
  end;

  // Período = ano + mês: o ano vem do cbYear e o mês da aba ativa (guia
  // JAN..DEZ) - são as duas coisas que definem a tela de extratos. dtposted
  // é texto ISO (YYYY-MM-DD...), então o ano são os 4 primeiros caracteres -
  // que também batem no formato OFX (YYYYMMDD...). O valor é validado como
  // número antes de entrar na query e a comparação é contra texto, para o
  // SQLite não trocar de tipo no meio da expressão.
  if (cbYear.ItemIndex >= 0) and (cbYear.ItemIndex < cbYear.Items.Count) then
  begin
    ano := StrToIntDef(Trim(cbYear.Items[cbYear.ItemIndex]), -1);
    if ano >= 0 then
    begin
      if condicao <> '' then
        condicao := condicao + ' AND ';
      condicao := condicao + 'substr(dtposted, 1, 4) = ' +
        QuotedStr(IntToStr(ano));
    end;
  end;

  // Mês = guia ativa (MesDaAba devolve 0 fora das guias de mês: as telas de
  // gestão, onde a grade está escondida e não há mês a filtrar). O mês não
  // fica na mesma posição nos dois formatos de dtposted: em YYYY-MM-DD são
  // os caracteres 6..7 e em YYYYMMDD são os 5..6 - o '-' na posição 5 é o
  // que separa os casos (é o mesmo critério do ExtratoDtpostedGetText, que
  // formata a coluna da grade). O mês é comparado como texto de 2 dígitos
  // ('01'..'12'), que é o que substr devolve.
  mes := MesDaAba(PageControl1.ActivePage);
  if mes > 0 then
  begin
    if condicao <> '' then
      condicao := condicao + ' AND ';
    condicao := condicao +
      '(CASE WHEN substr(dtposted, 5, 1) = ''-'' THEN substr(dtposted, 6, 2)' +
      ' ELSE substr(dtposted, 5, 2) END) = ' +
      QuotedStr(Format('%.2d', [mes]));
  end;

  // O * não serve por causa do "trnamt": a coluna é NUMERIC e o driver
  // SQLite a tipa como TLargeintField, que LÊ por sqlite3_column_int64 - os
  // centavos de um -198,50 apareciam como "198,00 D" e a GRAVAÇÃO de uma
  // edição destruía a fração (o bind também era int64). O CAST faz o campo
  // nascer TFloatField (leitura por sqlite3_column_double, gravação por
  // sqlite3_bind_double) e o AS mantém o NOME da coluna - que é o que o sqldb
  // usa para gerar o INSERT/UPDATE ("trnamt"=:"trnamt"), então a edição
  // continua valendo para o arquivo. As demais colunas vêm nomeadas e sem
  // CAST, iguais ao *.
  sql := 'SELECT id, account_id, trntype, dtposted,' +
    ' CAST(trnamt AS REAL) AS trnamt, memo, chknum FROM extratos';
  if condicao <> '' then
    sql := sql + ' WHERE ' + condicao;
  sql := sql + ' ORDER BY dtposted, id;';

  // Mudou o filtro, muda a query: só dá para trocar o SQL com a query
  // fechada, então fecha, reescreve e reabre na sequência.
  SQLQueryExtratos.Close;
  SQLQueryExtratos.SQL.Text := sql;
  SQLQueryExtratos.Open;
  PrepararCamposExtratos;

  // Os DOIS campos de saldo são da MESMA tela (ano + mês + conta que
  // acabaram de ser aplicados), então são calculados aqui: é o único ponto
  // que roda em todas as mudanças de guia, de ano, de conta, na abertura e
  // na importação.
  AtualizarSaldoAnterior;
  AtualizarSaldoMes;
end;

// O mês vem da guia ativa (é a guia que diz qual mês a grade filtra) e o
// ano do cbYear (é ele que define a lista de anos e as guias). Sem
// database/mês/ano válido não há o que consultar - é também o estado do
// fechar, onde a tela inteira já está escondida.
function TFormMoney.PeriodoExtratos(out ano, mes: Integer): Boolean;
begin
  mes := MesDaAba(PageControl1.ActivePage);
  ano := -1;
  if (cbYear.ItemIndex >= 0) and (cbYear.ItemIndex < cbYear.Items.Count) then
    ano := StrToIntDef(Trim(cbYear.Items[cbYear.ItemIndex]), -1);
  Result := DatabaseAberto and (mes > 0) and (ano >= 0);
end;

// O CondicaoPeriodo é só a parte da data (quem chama escolhe entre "antes do
// mês" e "dentro do mês"); a CONTA entra sempre aqui, mesma regra dos
// filtros (o combo carrega o id em Items.Objects e vazio não restringe) -
// senão os campos misturariam contas. A consulta é própria, na mesma
// conexão (mesmo caminho do ano do cbYear e das guias): não mexe no cursor
// de nenhuma grade. Em empate de data (vários saldos no mesmo dia) o
// desempate é pelo MAIOR id, que é AUTOINCREMENT = ordem de gravação - senão
// o resultado dependeria da ordem em que o SQLite devolvesse as linhas. E o
// CAST AS REAL não é cosmético: o driver SQLite tipa "NUMERIC" (sem
// precisão) como TLargeintField, que LÊ por sqlite3_column_int64 - sem o CAST
// um balance 37,5 chegava aqui como 37 (o CAST cai em TFloatField, lido por
// sqlite3_column_double).
function TFormMoney.ConsultarSaldo(const CondicaoPeriodo: string): Double;
var
  consulta: TSQLQuery;
  condConta: string;
begin
  Result := 0;

  condConta := '';
  if (cbAccount.ItemIndex >= 0) and (cbAccount.ItemIndex < cbAccount.Items.Count)
  then
    condConta := ' AND account_id = ' + IntToStr(
      Integer(PtrInt(cbAccount.Items.Objects[cbAccount.ItemIndex])));

  consulta := TSQLQuery.Create(nil);
  try
    consulta.Database := SQLite3ConnContas;
    consulta.Transaction := SQLTransactionContas;
    consulta.SQL.Text := 'SELECT CAST(balance AS REAL) FROM saldos WHERE ' +
      CondicaoPeriodo + condConta + ' ORDER BY enddate DESC, id DESC LIMIT 1;';
    consulta.Open;
    if not consulta.EOF then
      Result := consulta.Fields[0].AsFloat;
    consulta.Close;
  finally
    consulta.Free;
  end;
end;

// Mesmo par C/D da coluna de valor da grade (ExtratoTrnAmtGetText): a letra
// assume o papel do sinal e o valor é o absoluto, com 2 casas, vírgula e sem
// separador de milhar - zero conta como crédito (>= 0), a mesma regra do
// trnamt. A cor acompanha o sinal como na grade, MAS só o negativo é
// vermelho: o positivo fica na cor do próprio campo (o pedido foi "na cor
// atual", não o azul da grade). Os DOIS caminhos religam a cor de propósito
// - sem isso, depois de um saldo negativo o vermelho ficaria grudado no
// campo. Campo é TControl porque os dois exibidores são disso para baixo: o
// tsAnterior (TStaticText) e o txtSaldo (TLabel).
procedure TFormMoney.ExibirSaldo(Campo: TControl; valor: Double);
var
  formatos: TFormatSettings;
begin
  formatos := DefaultFormatSettings;
  formatos.DecimalSeparator := ',';
  if valor >= 0 then
  begin
    Campo.Font.Color := clDefault;
    Campo.Caption := FormatFloat('0.00', valor, formatos) + ' C';
  end
  else
  begin
    Campo.Font.Color := clRed;
    Campo.Caption := FormatFloat('0.00', -valor, formatos) + ' D';
  end;
end;

// O tsAnterior mostra o balance do saldo imediatamente anterior ao mês ativo
// - o saldo com que o mês começa. "Imediatamente anterior" = o registro mais
// recente de saldos ANTERIOR ao primeiro dia do mês da guia ativa (no ano do
// cbYear, que é o que define o ano desse limite). Não havendo nenhum, o
// pedido é exibir zero - que aparece como crédito ("0,00 C").
procedure TFormMoney.AtualizarSaldoAnterior;
var
  ano, mes: Integer;
begin
  if PeriodoExtratos(ano, mes) then
    // enddate é texto ISO (AAAA-MM-DD - formato que o próprio form grava),
    // então a comparação de texto ordena por data: o limite é o 1º dia do
    // mês ATIVO e o que vale é o registro mais recente antes dele.
    ExibirSaldo(tsAnterior, ConsultarSaldo('enddate < ' +
      QuotedStr(Format('%.4d-%.2d-01', [ano, mes]))))
  else
    ExibirSaldo(tsAnterior, 0);
end;

// O txtSaldo (rótulo "Saldo:" no rodapé) mostra o balance do saldo DO mês
// ativo - o saldo com que o mês fecha, o par do "Anterior" (que é a
// abertura). "Do mês" = o registro mais recente gravado com enddate dentro
// do mês da guia ativa (mesmo ano do cbYear); como as guias só existem para
// meses com saldo, na prática sempre há registro - mas o zero cobre o caso
// de a conta filtrada não ter saldo no mês.
procedure TFormMoney.AtualizarSaldoMes;
var
  ano, mes: Integer;
begin
  if PeriodoExtratos(ano, mes) then
    // substr(enddate, 1, 4) e substr(enddate, 6, 2) são ano e mês, o mesmo
    // recorte que o AtualizarAbasMes usa para decidir quais guias aparecem.
    ExibirSaldo(txtSaldo, ConsultarSaldo('substr(enddate, 1, 4) = ' +
      QuotedStr(Format('%.4d', [ano])) + ' AND substr(enddate, 6, 2) = ' +
      QuotedStr(Format('%.2d', [mes]))))
  else
    ExibirSaldo(txtSaldo, 0);
end;

// Grade de saldos da tbSaldos: só os registros da conta escolhida no
// cbAccount (que fica visível nessa tela justamente por isso). Combo vazia
// (nenhuma conta cadastrada) não restringe - mesma regra dos extratos, e o
// caminho enquanto "contas" não tem linhas. account_id não aparece na grade:
// o registro novo nasce com a conta selecionada (OnNewRecord).
procedure TFormMoney.AplicarFiltroSaldos;
var
  sql, condicao: string;
  idConta: Integer;
begin
  if not DatabaseAberto then
    Exit;

  condicao := '';
  if (cbAccount.ItemIndex >= 0) and (cbAccount.ItemIndex < cbAccount.Items.Count) then
  begin
    idConta := Integer(PtrInt(cbAccount.Items.Objects[cbAccount.ItemIndex]));
    condicao := ' WHERE account_id = ' + IntToStr(idConta);
  end;

  sql := 'SELECT * FROM saldos' + condicao + ';';

  // Mesmo caminho do filtro de extratos: só dá para trocar o SQL com a query
  // fechada, então fecha, reescreve e reabre.
  SQLQuerySaldos.Close;
  SQLQuerySaldos.SQL.Text := sql;
  SQLQuerySaldos.Open;
  // Campos recriados no Open: máscara de data e formatação de enddate têm de
  // ser religados (é o que PrepararCamposExtratos faz nos extratos).
  PrepararCamposSaldos;
end;

// dtposted é texto: o importador grava AAAAMMDD (o formato ISO AAAA-MM-DD
// do arquivo passa direto). A grade mostra DD/MM/AAAA nos dois casos; com
// DisplayText=False volta o valor cru, que é o que o editor da grade usa
// para editar sem transformar a data em lixo.
procedure TFormMoney.ExtratoDtpostedGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
var
  data: string;
begin
  data := Sender.AsString;
  if DisplayText then
  begin
    if (Length(data) = 8) and (StrToIntDef(data, -1) >= 0) then
      data := Copy(data, 7, 2) + '/' + Copy(data, 5, 2) + '/' + Copy(data, 1, 4)
    else if (Length(data) = 10) and (data[5] = '-') and (data[8] = '-') then
      data := Copy(data, 9, 2) + '/' + Copy(data, 6, 2) + '/' + Copy(data, 1, 4);
  end;
  AText := data;
end;

// memo e chknum são TEXT, que o SQLite entrega como ftMemo: a grade sem
// dgDisplayMemoText desenha via DisplayText, e aí o TDBGrid cairia no
// "(MEMO)" do campo memo — este handler devolve o texto de verdade (o mesmo
// que AsString, que é o que as outras grades mostram com a opção ligada).
procedure TFormMoney.ExtratoTextoGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
begin
  if Sender.IsNull then
    AText := ''
  else
    AText := Sender.AsString;
end;

// trnamt vira "20,00 C" (crédito) ou "20,00 D" (débito): a letra assume o
// papel do sinal e o valor é o absoluto, com 2 casas e vírgula. O editor da
// grade lê o valor cru (DisplayText=False), então o gravado não muda.
procedure TFormMoney.ExtratoTrnAmtGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
var
  formatos: TFormatSettings;
  valor: Double;
begin
  if not DisplayText then
  begin
    AText := Sender.AsString;
    Exit;
  end;
  if Sender.IsNull then
  begin
    AText := '';
    Exit;
  end;
  valor := Sender.AsFloat;
  formatos := DefaultFormatSettings;
  formatos.DecimalSeparator := ',';
  if valor >= 0 then
    AText := FormatFloat('0.00', valor, formatos) + ' C'
  else
    AText := FormatFloat('0.00', -valor, formatos) + ' D';
end;

// Mostra o nome do banco no lugar do id, tanto na exibição (DisplayText=True)
// quanto no texto do editor (False) — assim o combo da edição abre já no
// nome, que é o que está na lista. Id sem banco correspondente (registro
// antigo ou banco excluído) aparece como número, sem inventar nada.
procedure TFormMoney.BancoIdGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
begin
  if Sender.IsNull then
  begin
    AText := '';
    Exit;
  end;
  AText := BancoNomePorId(Sender.AsInteger);
  if AText = '' then
    AText := IntToStr(Sender.AsInteger);
end;

// Gravação da célula: o DBGrid faz Field.Text := texto do combo, e o
// SetEditText do campo vem parar aqui. Número passa direto (quem preferir
// digitar o id); nome vira o id correspondente. Texto vazio ou que não
// exista na lista não mexe no valor (o combo é de lista — só dá para
// escolher o que está lá —, e apagar tudo não deve virar lixo no id).
procedure TFormMoney.BancoIdSetText(Sender: TField; const AText: string);
var
  novoId: Integer;
begin
  if TryStrToInt(AText, novoId) then
  begin
    Sender.AsInteger := novoId;
    Exit;
  end;
  novoId := BancoIdPorNome(AText);
  if novoId >= 0 then
    Sender.AsInteger := novoId;
end;
// Liga os handlers de banco logo depois de cada Open: os campos são
// dinâmicos (FieldDefs vazio) e cada abertura os recria sem handlers — o
// mesmo problema, e a mesma solução, dos extratos.
procedure TFormMoney.SQLQueryContasAfterOpen(DataSet: TDataSet);
begin
  SQLQueryContas.FieldByName('bankid').OnGetText := @BancoIdGetText;
  SQLQueryContas.FieldByName('bankid').OnSetText := @BancoIdSetText;
end;

// Registro novo da tbSaldos nasce com a conta escolhida no cbAccount: a
// coluna account_id não aparece na grade (é ela que filtra), então quem
// preenche é aqui. Sem conta selecionada (combo vazia) o valor fica vazio e
// a gravação recusa - account_id é NOT NULL, e não há conta para sugerir.
procedure TFormMoney.SQLQuerySaldosNewRecord(DataSet: TDataSet);
var
  idConta: Integer;
begin
  if (cbAccount.ItemIndex >= 0) and (cbAccount.ItemIndex < cbAccount.Items.Count) then
  begin
    idConta := Integer(PtrInt(cbAccount.Items.Objects[cbAccount.ItemIndex]));
    DataSet.FieldByName('account_id').AsInteger := idConta;
  end;
end;

// Preenche o combo da coluna "Banco" com os nomes atuais de tbBancos. O
// LCL montou o editor e já copiou o PickList fixo de projeto para os itens;
// aqui isso é trocado pela lista viva da query (OnSelectEditor roda depois).
procedure TFormMoney.gridContasSelectEditor(Sender: TObject; Column: TColumn;
  var Editor: TWinControl);
var
  combo: TCustomComboBox;
  lista: TStringList;
  i: Integer;
begin
  if (Column = nil) or (Column.FieldName <> 'bankid') or (Editor = nil) or
    not (Editor is TCustomComboBox) then
    Exit;
  combo := TCustomComboBox(Editor);
  lista := TStringList.Create;
  try
    MontarListaBancos(lista);
    combo.Items.BeginUpdate;
    try
      combo.Items.Clear;
      for i := 0 to lista.Count - 1 do
        combo.Items.Add(lista.Names[i]);
    finally
      combo.Items.EndUpdate;
    end;
  finally
    lista.Free;
  end;
end;

// Copia a lista de bancos de tbBancos no formato "nome=id". Varre a própria
// query (que fica aberta desde a abertura do programa), guardando e
// devolvendo a posição com bookmark para a grade de bancos não mudar de
// linha — vazia se a query não estiver aberta (fora do programa, nunca).
// O id do banco é o código em TEXTO ('001', '077'...): no driver SQLite do
// FPC a coluna TEXT vira campo memo, que NÃO tem AsInteger (estouraria
// "Invalid type conversion to Integer in field id") — então o valor lido é
// sempre o texto puro.
procedure TFormMoney.MontarListaBancos(ALista: TStrings);
var
  bm: TBookmark;
begin
  ALista.BeginUpdate;
  try
    ALista.Clear;
    if not SQLQueryBanks.Active then
      Exit;
    bm := SQLQueryBanks.GetBookmark;
    SQLQueryBanks.DisableControls;
    try
      SQLQueryBanks.First;
      while not SQLQueryBanks.Eof do
      begin
        ALista.Values[SQLQueryBanks.FieldByName('name').AsString] :=
          SQLQueryBanks.FieldByName('id').AsString;
        SQLQueryBanks.Next;
      end;
    finally
      SQLQueryBanks.GotoBookmark(bm);
      SQLQueryBanks.EnableControls;
    end;
  finally
    ALista.EndUpdate;
  end;
end;

// O código do banco tem zero à esquerda ('077') e o bankid da conta guarda
// o número (77), então a comparação é numérica — código que não for número
// não casa com nada (bankid é INTEGER, não tem para onde ir).
function TFormMoney.BancoNomePorId(const AId: Integer): string;
var
  lista: TStringList;
  i, codigo: Integer;
begin
  Result := '';
  lista := TStringList.Create;
  try
    MontarListaBancos(lista);
    for i := 0 to lista.Count - 1 do
      if TryStrToInt(lista.ValueFromIndex[i], codigo) and (codigo = AId) then
      begin
        Result := lista.Names[i];
        Break;
      end;
  finally
    lista.Free;
  end;
end;

function TFormMoney.BancoIdPorNome(const ANome: string): Integer;
var
  lista: TStringList;
begin
  Result := -1;
  if ANome = '' then
    Exit;
  lista := TStringList.Create;
  try
    MontarListaBancos(lista);
    Result := StrToIntDef(lista.Values[ANome], -1);
  finally
    lista.Free;
  end;
end;

// Cor do valor conforme o sinal: crédito azul, débito vermelho. Célula
// selecionada fica com as cores padrão da seleção — vermelho/azul sobre o
// fundo de destaque não dão para ler.
procedure TFormMoney.gridTransPrepareCanvas(Sender: TObject; DataCol: Integer;
  Column: TColumn; AState: TGridDrawState);
begin
  if (Column = nil) or (Column.FieldName <> 'trnamt') then
    Exit;
  if gdSelected in AState then
    Exit;
  if (Column.Field = nil) or Column.Field.IsNull then
    Exit;
  if Column.Field.AsFloat >= 0 then
    gridTrans.Canvas.Font.Color := clBlue
  else
    gridTrans.Canvas.Font.Color := clRed;
end;

procedure TFormMoney.PrepararCamposExtratos;
begin
  SQLQueryExtratos.FieldByName('dtposted').OnGetText :=
    @ExtratoDtpostedGetText;
  SQLQueryExtratos.FieldByName('memo').OnGetText := @ExtratoTextoGetText;
  SQLQueryExtratos.FieldByName('chknum').OnGetText := @ExtratoTextoGetText;
  SQLQueryExtratos.FieldByName('trnamt').OnGetText := @ExtratoTrnAmtGetText;
end;

// enddate é TEXT (ftMemo no driver SQLite): sem dgDisplayMemoText a grade
// desenha pelo DisplayText - e, como a coluna tem máscara, o texto tem de vir
// DD/MM/AAAA também no editor (DisplayText=False), senão a edição começaria
// com o valor ISO cru, fora das posições da máscara. O que fica gravado não
// muda: é o OnSetText que converte de volta.
procedure TFormMoney.SaldosEnddateGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
var
  data: string;
begin
  data := Sender.AsString;
  if (Length(data) = 10) and (data[5] = '-') and (data[8] = '-') then
    data := Copy(data, 9, 2) + '/' + Copy(data, 6, 2) + '/' + Copy(data, 1, 4);
  AText := data;
end;

// Aceita o que a máscara produz (DD/MM/AAAA) e também o formato gravado
// (AAAA-MM-DD, para colar de fora). Dia/mês fora do lugar não é data: o valor
// fica como estava, sem inventar conversão.
procedure TFormMoney.SaldosEnddateSetText(Sender: TField; const AText: string);
var
  data: string;
  dataGravada: TDateTime;
begin
  data := Trim(AText);
  if data = '' then
  begin
    // enddate é NOT NULL: vazio vira texto vazio, que o SQLite aceita.
    Sender.AsString := '';
    Exit;
  end;
  if (Length(data) = 10) and (data[3] = '/') and (data[6] = '/') and
     TryEncodeDate(StrToIntDef(Copy(data, 7, 4), 0),
       StrToIntDef(Copy(data, 4, 2), 0), StrToIntDef(Copy(data, 1, 2), 0),
       dataGravada) then
    Sender.AsString := FormatDateTime('yyyy-mm-dd', dataGravada)
  else if (Length(data) = 10) and (data[5] = '-') and (data[8] = '-') and
     TryEncodeDate(StrToIntDef(Copy(data, 1, 4), 0),
       StrToIntDef(Copy(data, 6, 2), 0), StrToIntDef(Copy(data, 9, 2), 0),
       dataGravada) then
    Sender.AsString := FormatDateTime('yyyy-mm-dd', dataGravada);
end;

// Liga a máscara de data e os handlers de enddate logo depois de cada Open:
// os campos são dinâmicos (FieldDefs vazio) e cada abertura os recria - mesmo
// problema, mesma solução, dos extratos e do bankid das contas. A máscara
// "!99/00/0000;1;_" fixa as posições pela esquerda, aceita só dígito nas 8
// casas, insere as duas barras sozinha e, pelo "1" do segundo campo, devolve
// o texto COM as barras ao campo (é o que o OnSetText espera ler).
procedure TFormMoney.PrepararCamposSaldos;
begin
  SQLQuerySaldos.FieldByName('enddate').OnGetText := @SaldosEnddateGetText;
  SQLQuerySaldos.FieldByName('enddate').OnSetText := @SaldosEnddateSetText;
  SQLQuerySaldos.FieldByName('enddate').EditMask := '!99/00/0000;1;_';
end;

procedure TFormMoney.LimparFiltrosExtratos;
begin
  // Só esvazia: sem itens o ItemIndex cai sozinho para -1, e o OnChange que
  // isso dispara não reabre nada (AplicarFiltroExtratos exige database).
  cbAccount.Items.Clear;
  cbYear.Items.Clear;
  // Sem ano escolhido não há guia de mês para mostrar (a chamada explícita
  // cobre o caso do OnChange não disparar na limpeza).
  AtualizarAbasMes;
end;

procedure TFormMoney.miCloseClick(Sender: TObject);
begin
  // "Fechar Database" encerra a ligação com o arquivo de contas (o mesmo que
  // "Novo Database"/"Abrir Database" trocam de lugar). O banks.db não é
  // afetado: ele abre na inicialização e sustenta a lista de bancos.
  // Sempre dá para fechar duas vezes: fechar algo já fechado é só no-op.
  try
    SQLQueryContas.Close;
    SQLQueryExtratos.Close;
    SQLQuerySaldos.Close;
    SQLite3ConnContas.Close;
  except
    on E: Exception do
      MessageDlg('Não foi possível fechar o database.' + LineEnding + E.Message,
        mtError, [mbOK], 0);
  end;
  // Fechou: a página é ocultada e o "Fechar Database" volta a ser desativado
  // (se o fechar falhou, a query continua ligada e nada muda).
  AtualizarEstadoDatabase;
end;

procedure TFormMoney.miGerConClick(Sender: TObject);
begin
  NavigateToTab(tbContas);
end;

procedure TFormMoney.miGerSalClick(Sender: TObject);
begin
  // Mesmo caminho do "Gerenciar Contas": a página fica atrás do próprio
  // "Voltar" do painel inferior e o título vira "Gerenciamento de Saldos".
  NavigateToTab(tbSaldos);
end;

procedure TFormMoney.miImportClick(Sender: TObject);
var
  caminho, contaTexto, anoTexto, mensagem, erro: string;
  acctidArquivo, acctidConta: string;
  registros: TRegistrosOfx;
  consulta: TSQLQuery;
  ignorados, novos, contaId, indice: Integer;
begin
  // Cada linha de "extratos" aponta para uma conta (account_id NOT NULL),
  // entao sem conta escolhida nao ha para onde gravar. O menu "Transacoes"
  // ja' nasce desativado sem database aberto, entao nao falta checar isso.
  if (cbAccount.ItemIndex < 0) or (cbAccount.ItemIndex >= cbAccount.Items.Count)
  then
  begin
    MessageDlg('Cadastre uma conta antes de importar o extrato.' + LineEnding +
      'Cada linha de "extratos" pertence a uma conta da aba ' +
      '"Gerenciar Contas".', mtWarning, [mbOK], 0);
    Exit;
  end;
  contaId := Integer(PtrInt(cbAccount.Items.Objects[cbAccount.ItemIndex]));
  // A recarga das combos no fim comeca sempre no primeiro item: guarda as
  // duas escolhas para devolver o usuario para onde ele estava.
  contaTexto := cbAccount.Text;
  anoTexto := cbYear.Text;

  dlgImportar.InitialDir := ExtractFilePath(ParamStr(0));
  if not dlgImportar.Execute then
    Exit; // usuario cancelou

  caminho := dlgImportar.FileName;
  // Caminho relativo (sem pasta) fica na pasta inicial do dialogo, para o
  // arquivo nao cair no diretorio de trabalho corrente da aplicacao.
  if ExtractFilePath(caminho) = '' then
    caminho := IncludeTrailingPathDelimiter(dlgImportar.InitialDir) + caminho;

  // Le o arquivo ANTES de mexer no database: falha de leitura nao grava nada.
  // A conta de origem (<ACCTID>) vem da mesma leitura - e' ela que confere se
  // o extrato pertence a conta escolhida no filtro.
  try
    registros := LerOfx(caminho, ignorados);
    acctidArquivo := LerAcctIdDoArquivo(caminho);
  except
    on E: Exception do
    begin
      MessageDlg('Não foi possível ler o arquivo "' + ExtractFileName(caminho) +
        '".' + LineEnding + E.Message, mtError, [mbOK], 0);
      Exit;
    end;
  end;

  // Arquivo vazio, so' estrutura ou formato que nao e' OFX/OFC: nada a gravar.
  if Length(registros) = 0 then
  begin
    MessageDlg('Nenhum registro encontrado em "' + ExtractFileName(caminho) +
      '".' + LineEnding + 'O arquivo não tem linhas com data e valor.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  // O arquivo diz de qual conta veio (<ACCTID>): conferir com a conta
  // escolhida no filtro evita gravar linhas na conta errada - importar nao
  // tem outro destino possivel. Arquivo sem ACCTID nao tem o que conferir e
  // segue normalmente. A consulta e' propria pelo mesmo motivo de
  // AtualizarComboContas: nao mover o cursor da grade.
  acctidConta := '';
  if acctidArquivo <> '' then
  begin
    consulta := TSQLQuery.Create(nil);
    try
      consulta.Database := SQLite3ConnContas;
      consulta.Transaction := SQLTransactionContas;
      consulta.SQL.Text := 'SELECT acctid FROM contas WHERE id = ' +
        IntToStr(contaId) + ';';
      consulta.Open;
      if not consulta.EOF then
        acctidConta := Trim(consulta.FieldByName('acctid').AsString);
      consulta.Close;
    finally
      consulta.Free;
    end;
    if (acctidConta <> '') and (not SameText(acctidArquivo, acctidConta)) then
    begin
      MessageDlg('O arquivo "' + ExtractFileName(caminho) + '" é da conta "' +
        acctidArquivo + '",' + LineEnding + 'mas a conta selecionada no ' +
        'filtro é "' + acctidConta + '".' + LineEnding +
        'Escolha a conta certa no filtro e importe de novo.',
        mtWarning, [mbOK], 0);
      Exit;
    end;
  end;

  novos := 0;
  erro := '';
  // As queries fecham antes da gravacao: o commit da importacao nao pode
  // encontrar sentenca em andamento (o SQLite recusa o commit nesse caso).
  // O finally religa tudo e refaz filtros + grade, no mesmo caminho do
  // miNew/miOpen - e vale mesmo quando a importacao falha no meio.
  try
    try
      SQLQueryContas.Close;
      SQLQuerySaldos.Close;
      SQLQueryExtratos.Close;
      novos := ImportarExtratos(SQLite3ConnContas, contaId, registros);
    except
      on E: Exception do
        erro := 'Não foi possível importar "' + ExtractFileName(caminho) +
          '".' + LineEnding + E.Message;
    end;
  finally
    try
      if not SQLQueryContas.Active then
        SQLQueryContas.Open;
      // Recarrega as combos (a lista de anos vem de saldos, entao importar
      // nao mexe nela), devolve a conta e o ano escolhidos e reabre as queries
      // das duas grades ja' filtradas (extratos e saldos) - e' a grade de
      // extratos que mostra o que acabou de entrar.
      CarregarFiltrosExtratos;
      indice := cbAccount.Items.IndexOf(contaTexto);
      if indice >= 0 then
        cbAccount.ItemIndex := indice;
      indice := cbYear.Items.IndexOf(anoTexto);
      if indice >= 0 then
        cbYear.ItemIndex := indice;
      AplicarFiltroExtratos;
      AplicarFiltroSaldos;
      // Importar também cai na tela de extratos (mesmo caminho do
      // abrir/criar): é lá que as linhas novas aparecem. Sem navegar, quem
      // importou de uma tela de gestão ficava lá, sem ver nada do que entrou.
      IrParaTelaExtratos;
    except
      on E: Exception do
        if erro = '' then
          erro := 'O extrato foi gravado, mas a tela não pôde ser ' +
            'atualizada.' + LineEnding + E.Message;
    end;
  end;

  if erro <> '' then
  begin
    MessageDlg(erro, mtError, [mbOK], 0);
    Exit;
  end;

  // O que o usuario precisa saber: quanto entrou e quanto o arquivo tinha
  // incompleto. Repetido nao e' assunto aqui: o que vier do arquivo entra,
  // inclusive linha identica a outra do proprio extrato.
  mensagem := 'Importação concluída.' + LineEnding + LineEnding +
    IntToStr(novos) + ' ' + Plural(novos, 'registro adicionado em "',
    'registros adicionados em "') + contaTexto + '".';
  if ignorados > 0 then
    mensagem := mensagem + LineEnding + IntToStr(ignorados) + ' ' +
      Plural(ignorados, 'bloco ignorado (sem data ou sem valor).',
      'blocos ignorados (sem data ou sem valor).');
  MessageDlg(mensagem, mtInformation, [mbOK], 0);
end;

procedure TFormMoney.miListClick(Sender: TObject);
begin
  NavigateToTab(tbBancos);
end;

procedure TFormMoney.miNewClick(Sender: TObject);
var
  novoArquivo: string;
begin
  // O diálogo abre na pasta da aplicação, sugerindo "novo.db".
  dlgNewDatabase.InitialDir := ExtractFilePath(ParamStr(0));
  dlgNewDatabase.FileName := 'novo.db';
  if not dlgNewDatabase.Execute then
    Exit; // usuário cancelou

  novoArquivo := dlgNewDatabase.FileName;
  // O filtro já é *.db, mas ele pode digitar outro nome de extensão.
  if LowerCase(ExtractFileExt(novoArquivo)) <> '.db' then
    novoArquivo := ChangeFileExt(novoArquivo, '.db');
  // Caminho relativo (sem pasta) fica na pasta inicial do diálogo, para o
  // database não cair no diretório de trabalho corrente da aplicação.
  if ExtractFilePath(novoArquivo) = '' then
    novoArquivo := IncludeTrailingPathDelimiter(dlgNewDatabase.InitialDir) +
      novoArquivo;

  // banks.db é o database de bancos da aplicação (dado do usuário):
  // o "Novo Database" nunca pode sobrescrevê-lo.
  if SameText(ExtractFileName(novoArquivo), DatabaseFileName) then
  begin
    MessageDlg('"' + DatabaseFileName + '" é o database de bancos da aplicação.' +
      LineEnding + 'Escolha outro nome para o novo database.',
      mtError, [mbOK], 0);
    Exit;
  end;

  // Arquivo já existe? Só segue recriando do zero se ele confirmar.
  if FileExists(novoArquivo) then
    if MessageDlg('O arquivo "' + ExtractFileName(novoArquivo) +
      '" já existe e será recriado do zero.' + LineEnding +
      'Os dados atuais serão perdidos. Continuar?',
      mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      Exit;

  if not NewDatabase(novoArquivo) then
    Exit;

  // Mesma lógica da aba tbBancos: abrir as queries faz a grade e o navigator
  // das abas tbContas/tbSaldos passarem a enxergar as tabelas "contas" e
  // "saldos"; a de extratos entra por último, já filtrada (as três moram no
  // mesmo database, na mesma conexão).
  try
    SQLQueryContas.Close;
    SQLQueryExtratos.Close;
    SQLQuerySaldos.Close;
    SQLite3ConnContas.Close;
    SQLite3ConnContas.DatabaseName := novoArquivo;
    SQLQueryContas.Open;
    // O filtro precisa das combos preenchidas (conta de "contas", ano de
    // "extratos") antes das queries das grades abrirem com a seleção
    // aplicada - e' a mesma selecao que abre a de saldos, ja' filtrada.
    CarregarFiltrosExtratos;
    AplicarFiltroExtratos;
    AplicarFiltroSaldos;
  except
    on E: Exception do
      MessageDlg('O database "' + ExtractFileName(novoArquivo) +
        '" foi criado, mas não pôde ser aberto.' + LineEnding + E.Message,
        mtError, [mbOK], 0);
  end;
  // Criou e abriu: a página aparece e o "Fechar Database" é ativado (se não
  // abriu, a ligação anterior já tinha sido fechada e os dois são desligados).
  // A tela de entrada é a de transações: as telas de gestão ficam disponíveis
  // sem database (miList), então dá para criar um arquivo estando nelas -
  // sem o passo abaixo a revelação mostraria aquela tela, com o gridTrans
  // escondido (ele só existe na tela de extratos). Se o arquivo não abriu,
  // nada muda por aqui.
  if DatabaseAberto then
    IrParaTelaExtratos;
  AtualizarEstadoDatabase;
end;

procedure TFormMoney.miOpenClick(Sender: TObject);
var
  abrirArquivo: string;
begin
  // O diálogo abre na pasta da aplicação, com o filtro *.db.
  dlgOpenDatabase.InitialDir := ExtractFilePath(ParamStr(0));
  if not dlgOpenDatabase.Execute then
    Exit; // usuário cancelou

  abrirArquivo := dlgOpenDatabase.FileName;
  // Caminho relativo (sem pasta) fica na pasta inicial do diálogo, para o
  // database não cair no diretório de trabalho corrente da aplicação.
  if ExtractFilePath(abrirArquivo) = '' then
    abrirArquivo := IncludeTrailingPathDelimiter(dlgOpenDatabase.InitialDir) +
      abrirArquivo;

  if not FileExists(abrirArquivo) then
  begin
    MessageDlg('O arquivo "' + ExtractFileName(abrirArquivo) +
      '" não foi encontrado.', mtError, [mbOK], 0);
    Exit;
  end;

  // Só troca a conexão se o arquivo for mesmo um database de contas. A
  // confirmação abre conexão própria, então um arquivo ruim (ex.: o próprio
  // banks.db) não derruba a ligação que a aba tbContas já tinha.
  if not IsAccountsDatabase(abrirArquivo) then
  begin
    MessageDlg('"' + ExtractFileName(abrirArquivo) +
      '" não é um database de contas (falta a tabela "contas", "extratos" ' +
      'ou "saldos").' +
      LineEnding +
      'Crie um com o menu "Novo Database" ou escolha outro arquivo.',
      mtError, [mbOK], 0);
    Exit;
  end;

  // Mesma ligação do "Novo Database": fecha o que estiver aberto, aponta a
  // conexão para o arquivo escolhido e reabre as queries da tbContas, da
  // tbSaldos e a de extratos (uma ligação só para as três tabelas).
  try
    SQLQueryContas.Close;
    SQLQueryExtratos.Close;
    SQLQuerySaldos.Close;
    SQLite3ConnContas.Close;
    SQLite3ConnContas.DatabaseName := abrirArquivo;
    SQLQueryContas.Open;
    // O filtro precisa das combos preenchidas (conta de "contas", ano de
    // "extratos") antes das queries das grades abrirem com a seleção
    // aplicada - e' a mesma selecao que abre a de saldos, ja' filtrada.
    CarregarFiltrosExtratos;
    AplicarFiltroExtratos;
    AplicarFiltroSaldos;
  except
    on E: Exception do
      MessageDlg('Não foi possível abrir o database "' +
        ExtractFileName(abrirArquivo) + '".' + LineEnding + E.Message,
        mtError, [mbOK], 0);
  end;
  // Abriu: a página aparece e o "Fechar Database" é ativado (um arquivo
  // rejeitado antes daqui nem mexe na ligação: nada muda por aqui). A tela de
  // entrada é a de transações - como as telas de gestão ficam disponíveis sem
  // database, dá para abrir um arquivo estando nelas, e sem o passo abaixo a
  // revelação mostraria aquela tela com o gridTrans escondido (ele só existe
  // na tela de extratos). Se o arquivo não abriu, nada muda por aqui.
  if DatabaseAberto then
    IrParaTelaExtratos;
  AtualizarEstadoDatabase;
end;

// Volta para a tela de transações (a de entrada): devolve as guias que a
// navegação de gestão escondeu e recalcula as guias de mês pelo dado gravado,
// com a aba ativa no mês mais recente. Compartilhada pelo "Voltar" e pela
// abertura de database - abrir num arquivo novo tem de cair aqui, mesmo que o
// usuário estivesse numa tela de gestão (elas ficam disponíveis sem database,
// então dá para abrir um arquivo por elas - e, sem este passo, o gridTrans
// ficaria escondido, já que ele só existe na tela de extratos).
procedure TFormMoney.IrParaTelaExtratos;
var
  i: Integer;
begin
  // Devolve as guias que estavam visíveis (as de mês com saldo no ano
  // escolhido) e volta para a página de mês.
  if FInListView and (Length(FSavedTabVisible) = PageControl1.PageCount) then
  begin
    for i := 0 to PageControl1.PageCount - 1 do
      PageControl1.Pages[i].TabVisible := FSavedTabVisible[i];
    FInListView := False;
  end;
  PageControl1.ActivePage := tbJan;
  // As guias de mês são dinâmicas: o que a tela de gestão mexeu em "saldos"
  // pode ter mudado os meses do ano, então recarrega pelo dado da tabela e
  // põe a aba ativa no mês mais recente. Sem database a chamada não faz nada
  // (é o estado do "Voltar" sem database, que esconde a interface logo abaixo).
  AtualizarAbasMes;
end;

procedure TFormMoney.sbtnVoltarClick(Sender: TObject);
begin
  IrParaTelaExtratos;
  // Depois do OnChange acima (senão o rodapé voltaria a aparecer): sem
  // database em aberto não há o que mostrar na tela de extratos, então o
  // "Voltar" devolve o estado inicial (só o MainMenu). Com database (miNew/
  // miOpen) a interface revelada na navegação fica de pé — e o "Fechar
  // Database" volta a zerar a regra.
  if not DatabaseAberto then
    SetInterfaceVisible(False)
  else
    // Voltou de uma tela de gestão: a tbContas inclui/exclui "contas" e o
    // combo de conta é o retrato disso na tela de extratos (é o único lugar
    // onde ele aparece). Sem remontar aqui, a troca feita na grade só
    // apareceria no próximo miNew/miOpen/miImport.
    AtualizarComboContas;
end;

procedure TFormMoney.toggleShowControlsClick(Sender: TObject);
begin
  DBNavTrans.Visible := toggleShowControls.Checked;
end;

end.

