unit unitMoney;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ComCtrls, ExtCtrls,
  StdCtrls, Menus, DBGrids, DBCtrls;

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
    procedure toggleShowControlsClick(Sender: TObject);
  private

  public

  end;

var
  FormMoney: TFormMoney;

implementation

{$R *.lfm}

{ TFormMoney }

procedure TFormMoney.toggleShowControlsClick(Sender: TObject);
begin
  DBNavTrans.Visible := toggleShowControls.Checked;
end;

end.

