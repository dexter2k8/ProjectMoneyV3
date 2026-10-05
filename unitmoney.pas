unit unitMoney;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLite3Conn, SQLDB, Forms, Controls, Graphics, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls, Menus, DB, DBGrids, DBCtrls;

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
    SQLite3ConnBancos: TSQLite3Connection;
    SQLTransactionBancos: TSQLTransaction;
    SQLQueryBanks: TSQLQuery;
    DataSourceBancos: TDataSource;
    toggleShowControls: TToggleBox;
    tsSaldo: TStaticText;
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
    tsAnterior: TStaticText;
    txtSaldo: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure toggleShowControlsClick(Sender: TObject);
  private

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
end;

procedure TFormMoney.toggleShowControlsClick(Sender: TObject);
begin
  DBNavTrans.Visible := toggleShowControls.Checked;
end;

end.

