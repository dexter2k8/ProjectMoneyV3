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
    procedure cbAccountChange(Sender: TObject);
    procedure cbYearChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure miCloseClick(Sender: TObject);
    procedure miGerConClick(Sender: TObject);
    procedure miGerSalClick(Sender: TObject);
    procedure miImportClick(Sender: TObject);
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
    // Combos de filtro da tela de extratos (conta e ano): montadas por
    // miNew/miOpen a partir do database aberto, refazem a query da grade
    // quando o usuário troca de seleção e são esvaziadas ao fechar.
    procedure CarregarFiltrosExtratos;
    procedure AplicarFiltroExtratos;
    procedure LimparFiltrosExtratos;
    // Remonta o combo de contas direto do database, mantendo a conta que o
    // usuário tinha escolhido (pelo id, para sobreviver à edição do texto).
    // É o caminho do "Voltar" das telas de gestão: a tbContas inclui/exclui
    // "contas" e o combo só existe na tela de extratos, então sem remontar
    // aqui a troca só aparece no próximo miNew/miOpen/miImport.
    procedure AtualizarComboContas;

  public

  end;

var
  FormMoney: TFormMoney;

implementation

uses
  unitDatabase, unitOfx;

{$R *.lfm}

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
  // Trocou a conta escolhida na tela de extratos: a query da grade reabre já
  // filtrada pela conta nova (sem database aberto não há o que fazer, e é
  // esse o caminho do OnChange que a limpeza das combos dispara).
  AplicarFiltroExtratos;
end;

procedure TFormMoney.cbYearChange(Sender: TObject);
begin
  // Mesmo caminho do cbAccountChange: mudou o ano, refaz o filtro da grade.
  AplicarFiltroExtratos;
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
  // Só entra ano de verdade (4 dígitos): dtposted é texto, qualquer outra
  // coisa nessa posição não é ano para ninguém.
  consulta := TSQLQuery.Create(nil);
  try
    consulta.Database := SQLite3ConnContas;
    consulta.Transaction := SQLTransactionContas;
    consulta.SQL.Text :=
      'SELECT DISTINCT substr(dtposted, 1, 4) FROM extratos' +
      ' ORDER BY substr(dtposted, 1, 4);';
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

  // Sem ano nenhum (extratos vazio, logo depois do "Novo Database") o combo
  // não pode ficar em branco: entra o ano atual e o filtro segue de pé.
  if cbYear.Items.Count = 0 then
    cbYear.Items.Add(FormatDateTime('yyyy', Date));
  cbYear.ItemIndex := 0;

  // Garante a seleção final mesmo se nenhuma das combos disparar OnChange
  // ao ganhar o primeiro ItemIndex.
  AplicarFiltroExtratos;
end;

procedure TFormMoney.AplicarFiltroExtratos;
var
  sql, condicao: string;
  idConta, ano: Integer;
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

  // Período = ano: não existe seletor de mês (as guias JAN..DEZ estão todas
  // ocultas). dtposted é texto ISO (YYYY-MM-DD...), então o ano são os 4
  // primeiros caracteres — que também batem no formato OFX (YYYYMMDD...).
  // O valor é validado como número antes de entrar na query e a comparação
  // é contra texto, para o SQLite não trocar de tipo no meio da expressão.
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

  sql := 'SELECT * FROM extratos';
  if condicao <> '' then
    sql := sql + ' WHERE ' + condicao;
  sql := sql + ' ORDER BY dtposted, id;';

  // Mudou o filtro, muda a query: só dá para trocar o SQL com a query
  // fechada, então fecha, reescreve e reabre na sequência.
  SQLQueryExtratos.Close;
  SQLQueryExtratos.SQL.Text := sql;
  SQLQueryExtratos.Open;
end;

procedure TFormMoney.LimparFiltrosExtratos;
begin
  // Só esvazia: sem itens o ItemIndex cai sozinho para -1, e o OnChange que
  // isso dispara não reabre nada (AplicarFiltroExtratos exige database).
  cbAccount.Items.Clear;
  cbYear.Items.Clear;
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
      if not SQLQuerySaldos.Active then
        SQLQuerySaldos.Open;
      // Recarrega as combos (o arquivo pode trazer anos que ainda nao
      // apareciam), devolve a conta e o ano escolhidos e reabre a query da
      // grade ja' filtrada - e' a grade que mostra o que acabou de entrar.
      CarregarFiltrosExtratos;
      indice := cbAccount.Items.IndexOf(contaTexto);
      if indice >= 0 then
        cbAccount.ItemIndex := indice;
      indice := cbYear.Items.IndexOf(anoTexto);
      if indice >= 0 then
        cbYear.ItemIndex := indice;
      AplicarFiltroExtratos;
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
    SQLQuerySaldos.Open;
    // O filtro precisa das combos preenchidas (conta de "contas", ano de
    // "extratos") antes da query da grade abrir com a seleção aplicada.
    CarregarFiltrosExtratos;
    AplicarFiltroExtratos;
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
    SQLQuerySaldos.Open;
    // O filtro precisa das combos preenchidas (conta de "contas", ano de
    // "extratos") antes da query da grade abrir com a seleção aplicada.
    CarregarFiltrosExtratos;
    AplicarFiltroExtratos;
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

