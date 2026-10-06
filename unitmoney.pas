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
    DBNavTrans: TDBNavigator;
    DBNavSaldos: TDBNavigator;
    gridSaldos: TDBGrid;
    gridTrans: TDBGrid;
    gridBancos: TDBGrid;
    lblTitle: TLabel;
    lblSaldo: TLabel;
    MainMenu: TMainMenu;
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
    pnSaldosControl: TPanel;
    pnHeader: TPanel;
    pnFooter: TPanel;
    Separator1: TMenuItem;
    Separator2: TMenuItem;
    Separator3: TMenuItem;
    Separator4: TMenuItem;
    sbtnVoltar: TSpeedButton;
    SQLite3ConnBancos: TSQLite3Connection;
    SQLTransactionBancos: TSQLTransaction;
    SQLQueryBanks: TSQLQuery;
    DataSourceBancos: TDataSource;
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
    procedure miListClick(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure sbtnVoltarClick(Sender: TObject);
    procedure toggleShowControlsClick(Sender: TObject);
  private
    // Estado das guias antes de "Lista de Bancos" (para o "Voltar" devolver)
    FSavedTabVisible: array of Boolean;
    FInListView: Boolean;
    // Título da tela normal (mês) antes de entrar na lista de bancos
    FSavedTitle: String;

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
  // O form já nasce com tbJan ativa e, durante o streaming, o OnChange não
  // dispara (csLoading): aplica a regra da lista de bancos aqui.
  FSavedTitle := lblTitle.Caption;
  PageControl1Change(nil);
end;

procedure TFormMoney.PageControl1Change(Sender: TObject);
var
  i: Integer;
  exibindoBancos: Boolean;
begin
  exibindoBancos := (PageControl1.ActivePage = tbBancos);

  // Rodapé: são controles da futura tabela de transações; só ficam fora da lista.
  for i := 0 to pnFooter.ControlCount - 1 do
    pnFooter.Controls[i].Visible := not exibindoBancos;
  // ... exceto o DBNavTrans, que tem regra própria (toggleShowControls).
  DBNavTrans.Visible := (not exibindoBancos) and toggleShowControls.Checked;

  // Cabeçalho e textos de saldo/anterior também somem na lista de bancos.
  cbYear.Visible := not exibindoBancos;
  cbAccount.Visible := not exibindoBancos;
  tsAnterior.Visible := not exibindoBancos;
  tslblAnterior.Visible := not exibindoBancos;

  // Título principal: troca para "Lista de Bancos" e volta ao anterior.
  if exibindoBancos then
    lblTitle.Caption := 'Lista de Bancos'
  else
    lblTitle.Caption := FSavedTitle;
end;

procedure TFormMoney.miListClick(Sender: TObject);
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
  PageControl1.ActivePage := tbBancos;
  for i := 0 to PageControl1.PageCount - 1 do
    PageControl1.Pages[i].TabVisible := False;
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
end;

procedure TFormMoney.toggleShowControlsClick(Sender: TObject);
begin
  DBNavTrans.Visible := toggleShowControls.Checked;
end;

end.

