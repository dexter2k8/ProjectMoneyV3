unit unitMoney;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ComCtrls, ExtCtrls,
  StdCtrls, Menus;

type

  { TFormMoney }

  TFormMoney = class(TForm)
    lblSaldo: TLabel;
    MainMenu: TMainMenu;
    MenuItemAbout: TMenuItem;
    MenuItemHelp: TMenuItem;
    MenuItemImpSal: TMenuItem;
    MenuItemImpTrans: TMenuItem;
    MenuItemExit: TMenuItem;
    MenuItemList: TMenuItem;
    MenuItemClose: TMenuItem;
    MenuItemOpen: TMenuItem;
    MenuItemExpSal: TMenuItem;
    MenuItemExpTrans: TMenuItem;
    MenuItemImport: TMenuItem;
    MenuItemTransactions: TMenuItem;
    MenuItemNew: TMenuItem;
    MenuItemArquivo: TMenuItem;
    PanelFooter: TPanel;
    Separator1: TMenuItem;
    Separator2: TMenuItem;
    Separator3: TMenuItem;
    txtSaldo: TLabel;
  private

  public

  end;

var
  FormMoney: TFormMoney;

implementation

{$R *.lfm}

end.

