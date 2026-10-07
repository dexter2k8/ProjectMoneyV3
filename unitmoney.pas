unit unitMoney;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLite3Conn, SQLDB, Forms, Controls, Graphics, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls, Menus, DB, DBGrids, DBCtrls, Buttons;

type

  { TFormMoney }

  TFormMoney = class(TForm)
    cbYear: TComboBox;
    cbAccount: TComboBox;
    DBNavBancos: TDBNavigator;
    DBNavContas: TDBNavigator;
    DBNavTrans: TDBNavigator;
    DBNavSaldos: TDBNavigator;
    gridContas: TDBGrid;
    gridSaldos: TDBGrid;
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
    dlgNewDatabase: TSaveDialog;
    dlgOpenDatabase: TOpenDialog;
    SQLite3ConnContas: TSQLite3Connection;
    SQLTransactionContas: TSQLTransaction;
    SQLQueryContas: TSQLQuery;
    DataSourceContas: TDataSource;
    SQLQuerySaldos: TSQLQuery;
    DataSourceSaldos: TDataSource;
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
    procedure FormCreate(Sender: TObject);
    procedure miCloseClick(Sender: TObject);
    procedure miGerConClick(Sender: TObject);
    procedure miGerSalClick(Sender: TObject);
    procedure miListClick(Sender: TObject);
    procedure miNewClick(Sender: TObject);
    procedure miOpenClick(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure sbtnVoltarClick(Sender: TObject);
    procedure toggleShowControlsClick(Sender: TObject);
  private
    // Estado das guias antes de "Lista de Bancos" (para o "Voltar" devolver)
    FSavedTabVisible: array of Boolean;
    FInListView: Boolean;
    // Título da tela normal (mês) antes de entrar na lista de bancos
    FSavedTitle: String;
    // Base do miList/miGerCon: guarda guias + título e ativa a aba com as
    // guias ocultas (o "Voltar" devolve o estado salvo aqui).
    procedure NavigateToTab(ATab: TTabSheet);
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

  public

  end;

var
  FormMoney: TFormMoney;

implementation

uses
  unitDatabase;

{$R *.lfm}

{ TFormMoney }

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
  // ... mas o DBNavTrans mantém a regra própria (toggleShowControls) para,
  // quando o painel voltar, o navegador respeitar o toggle.
  DBNavTrans.Visible := exibindoExtrato and toggleShowControls.Checked;

  // Combos do cabeçalho: somem nas três telas de gestão (lista de bancos,
  // contas e saldos), junto com os textos "Anterior".
  cbYear.Visible := exibindoExtrato;
  cbAccount.Visible := exibindoExtrato;
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
  // Ativar antes de esconder: evita que o LCL troque de página sozinho.
  PageControl1.ActivePage := ATab;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := False;
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
  // Sem database aberto não há nada para fechar: o item fica desativado.
  miClose.Enabled := aberto;
  // Todo o menu "Transações" (importar, exportar, gerenciar) opera sobre o
  // database de contas: sem ligação, não há transações a fazer. Desativar o
  // menu de primeiro nível cobre os itens de uma vez - inclusive os futuros.
  mmTransactions.Enabled := aberto;
end;

procedure TFormMoney.miCloseClick(Sender: TObject);
begin
  // "Fechar Database" encerra a ligação com o arquivo de contas (o mesmo que
  // "Novo Database"/"Abrir Database" trocam de lugar). O banks.db não é
  // afetado: ele abre na inicialização e sustenta a lista de bancos.
  // Sempre dá para fechar duas vezes: fechar algo já fechado é só no-op.
  try
    SQLQueryContas.Close;
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
  // "saldos" (as duas moram no mesmo database, na mesma conexão).
  try
    SQLQueryContas.Close;
    SQLQuerySaldos.Close;
    SQLite3ConnContas.Close;
    SQLite3ConnContas.DatabaseName := novoArquivo;
    SQLQueryContas.Open;
    SQLQuerySaldos.Open;
  except
    on E: Exception do
      MessageDlg('O database "' + ExtractFileName(novoArquivo) +
        '" foi criado, mas não pôde ser aberto.' + LineEnding + E.Message,
        mtError, [mbOK], 0);
  end;
  // Criou e abriu: a página aparece e o "Fechar Database" é ativado (se não
  // abriu, a ligação anterior já tinha sido fechada e os dois são desligados).
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
      '" não é um database de contas (falta a tabela "contas" ou "saldos").' +
      LineEnding +
      'Crie um com o menu "Novo Database" ou escolha outro arquivo.',
      mtError, [mbOK], 0);
    Exit;
  end;

  // Mesma ligação do "Novo Database": fecha o que estiver aberto, aponta a
  // conexão para o arquivo escolhido e reabre as queries da tbContas e da
  // tbSaldos (uma ligação só para as duas tabelas).
  try
    SQLQueryContas.Close;
    SQLQuerySaldos.Close;
    SQLite3ConnContas.Close;
    SQLite3ConnContas.DatabaseName := abrirArquivo;
    SQLQueryContas.Open;
    SQLQuerySaldos.Open;
  except
    on E: Exception do
      MessageDlg('Não foi possível abrir o database "' +
        ExtractFileName(abrirArquivo) + '".' + LineEnding + E.Message,
        mtError, [mbOK], 0);
  end;
  // Abriu: a página aparece e o "Fechar Database" é ativado (um arquivo
  // rejeitado antes daqui nem mexe na ligação: nada muda por aqui).
  AtualizarEstadoDatabase;
end;

procedure TFormMoney.sbtnVoltarClick(Sender: TObject);
var
  i: Integer;
begin
  // Devolve as guias que estavam visíveis (JAN e Saldos no estado atual)
  // e volta para a página de mês.
  if FInListView and (Length(FSavedTabVisible) = PageControl1.PageCount) then
  begin
    for i := 0 to PageControl1.PageCount - 1 do
      PageControl1.Pages[i].TabVisible := FSavedTabVisible[i];
    FInListView := False;
  end;
  PageControl1.ActivePage := tbJan;
  // Depois do OnChange acima (senão o rodapé voltaria a aparecer): sem
  // database em aberto não há o que mostrar na tela de extratos, então o
  // "Voltar" devolve o estado inicial (só o MainMenu). Com database (miNew/
  // miOpen) a interface revelada na navegação fica de pé — e o "Fechar
  // Database" volta a zerar a regra.
  if not DatabaseAberto then
    SetInterfaceVisible(False);
end;

procedure TFormMoney.toggleShowControlsClick(Sender: TObject);
begin
  DBNavTrans.Visible := toggleShowControls.Checked;
end;

end.

