unit unitMoney;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLite3Conn, SQLDB, Forms, Controls, Graphics, Dialogs,
  ComCtrls, ExtCtrls, StdCtrls, Menus, DB, Grids, DBGrids, DBCtrls, Buttons;

type

  // gridSaldos é uma TDBGrid de base mesmo, e não um descendente próprio: o
  // Lazarus só resolve objeto de .lfm pela paleta, por designer base ou por
  // GetClass, e classe declarada num unit do projeto não é nenhuma das três -
  // o IDE para em "incapaz de localizar a classe do componente" (e o fallback
  // para o ancestral registrado, TDBGrid, é desligado no caminho de classe
  // aninhada). O que se queria no descendente era só digitar em enddate; ver
  // o CAST que faz o campo nascer ftString em AplicarFiltroSaldos.

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
    miImpTrans: TMenuItem;
    miExit: TMenuItem;
    miList: TMenuItem;
    miClose: TMenuItem;
    miOpen: TMenuItem;
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
    Separator4: TMenuItem;
    sbtnVoltar: TSpeedButton;
    btnMonthClear: TSpeedButton;
    SQLite3ConnBancos: TSQLite3Connection;
    SQLTransactionBancos: TSQLTransaction;
    SQLQueryBanks: TSQLQuery;
    DataSourceBancos: TDataSource;
    // Database novo/aberto (menu "Novo Database" / "Abrir Database") -> aba
    // tbContas: mesma cadeia conexão -> transaction -> query -> datasource da
    // aba tbBancos; os dois menus só trocam o arquivo dessa mesma cadeia.
    // Transacoes (miExpTrans): extrato inteiro da conta escolhida no
    // cbAccount gravado em CSV - o MESMO formato que o dlgImpCsv le' de volta.
    dlgExpCsv: TSaveDialog;
    // Extrato (miImport): arquivo OFC/OFX escolhido pelo usuario - o
    // conteudo vai para a tabela "extratos", na conta selecionada na tela.
    dlgImportar: TOpenDialog;
    // Transacoes (miImpTrans): CSV escolhido pelo usuario - as tabelas
    // "extratos" e "saldos" sao substituidas pelo conteudo dele.
    dlgImpCsv: TOpenDialog;
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
    // "Excluir Mês" (btnMonthClear): apaga o mês da guia ativa - as linhas de
    // "extratos" e o registro de "saldos" do mesmo fim de período, para TODAS
    // as contas (é o mês que vira guia, não a conta). Não há desfazer, então
    // o dialogo de confirmação vem antes de qualquer DELETE.
    procedure btnMonthClearClick(Sender: TObject);
    procedure cbAccountChange(Sender: TObject);
    procedure cbYearChange(Sender: TObject);
    // "Oculta" (a coluna "hidden" da tbContas): a grade mostra Sim/Nao no lugar
    // do 0/1 gravado - e' o unico lugar onde o usuario ve (e muda) se a conta
    // sai do filtro do cbAccount.
    procedure ContaHiddenGetText(Sender: TField; var AText: string;
      DisplayText: Boolean);
    procedure ContaHiddenSetText(Sender: TField; const AText: string);
    // Formatação da grade de extratos: os campos TEXT do SQLite viram ftMemo,
    // então a exibição passa pelo DisplayText, que é onde o OnGetText entra.
    // (Sem isso a data aparece como 20250305 e o valor sem C/D.)
    // dtposted tem máscara de data (mesmo caso do enddate da tbSaldos): o
    // texto COM as barras que a máscara entrega no editor (DD/MM/AAAA) tem
    // de voltar ISO (AAAA-MM-DD) para o arquivo - os filtros e as guias leem
    // a coluna por posição.
    procedure ExtratoDtpostedGetText(Sender: TField; var AText: string;
      DisplayText: Boolean);
    procedure ExtratoDtpostedSetText(Sender: TField; const AText: string);
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
    // "Ajuda > Sobre": só informativo - um dialogo com a mensagem do
    // programa e um UNICO botão (mbOK). Não mexe em database nem em estado
    // de tela, então funciona com a interface aberta ou não (o menu é o
    // controle que sempre fica acessível).
    procedure miAboutClick(Sender: TObject);
    procedure miCloseClick(Sender: TObject);
    // "Sair": fecha o database (o mesmo fechar do miClose) e encerra a
    // aplicação. É o fim do programa, não uma navegação - o estado da tela
    // depois dele não interessa mais.
    procedure miExitClick(Sender: TObject);
    procedure miGerConClick(Sender: TObject);
    procedure miGerSalClick(Sender: TObject);
    // "Exportar Transacoes (CSV)": le o extrato inteiro da conta escolhida no
    // cbAccount e grava um CSV no MESMO formato que o miImpTrans le' - da' para
    // importar o arquivo de volta e para levar o historico para outra maquina.
    procedure miExpTransClick(Sender: TObject);
    procedure miImportClick(Sender: TObject);
    // "Importar Transacoes (CSV)": le um CSV (o modelo e' backup/extrato.csv)
    // e SUBSTITUI o conteudo de "extratos" e "saldos" pelas linhas dele - a
    // mesma gravacao e os mesmos saldos de fim de mes do miImport.
    procedure miImpTransClick(Sender: TObject);
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
    // balance da tbSaldos: DUAS casas sempre, com vírgula e com o sinal na
    // frente - "20,00" para positivo e "-20,00" para negativo. É numerico,
    // sem o par C/D da grade de extratos: ali a letra assume o papel do
    // sinal da coluna de valor, aqui o próprio sinal é o crédito/débito.
    // DisplayText=False devolve o valor cru, que é o que o editor da grade
    // abre para editar (mesma regra do trnamt).
    procedure SaldosBalanceGetText(Sender: TField; var AText: string;
      DisplayText: Boolean);
    // O que o usuário digita: aceita os DOIS separadores (vírgula, o que a
    // grade exibe e o que o editor dela deixa digitar, e ponto, o formato
    // do arquivo e do SQL) e grava o MESMO número - 19,9 e 19.9 ficam 19,90
    // na grade, 20 fica 20,00. Texto que não é número não mexe no valor.
    procedure SaldosBalanceSetText(Sender: TField; const AText: string);
    // Liga os handlers de banco (OnGetText/OnSetText) a cada Open da query de
    // contas: os campos são dinâmicos e cada Open os recria sem handlers —
    // o mesmo problema, e a mesma solução, dos extratos (PrepararCamposExtratos).
    procedure SQLQueryContasAfterOpen(DataSet: TDataSet);
    // Registro novo da tbContas nasce VISIVEL: "hidden" e' NOT NULL, nao tem
    // campo na grade para digita-la e a query da grade e' um SELECT *, entao
    // sem este handler o INSERT que o SQLDB montaria levaria NULL e o SQLite
    // recusaria a linha com "NOT NULL constraint failed".
    procedure SQLQueryContasNewRecord(DataSet: TDataSet);
    // Registro novo da tbSaldos recebe a conta escolhida no cbAccount (a
    // coluna account_id está oculta na grade - é ela que filtra).
    procedure SQLQuerySaldosNewRecord(DataSet: TDataSet);
    // Registro novo da grade de extratos: as duas colunas NOT NULL que não
    // aparecem na grade (account_id vem do cbAccount, que filtra a tela -
    // mesmo caso da tbSaldos; trntype nasce 'OTHER', o mesmo padrão do
    // leitor de OFX/OFC quando o arquivo não traz o tipo). Sem isto o INSERT
    // que o sqldb monta nomeia todas as colunas do SELECT * e manda NULL nas
    // duas, que o SQLite recusa - não dava para criar transação nenhuma na
    // mão.
    procedure SQLQueryExtratosNewRecord(DataSet: TDataSet);
    procedure toggleShowControlsClick(Sender: TObject);
  private
    // Estado das guias antes de navegar para uma tela de gestão (para o
    // "Voltar" devolver - as guias de mês são dinâmicas e o "Voltar" as
    // recalcula pelo saldo do ano, em vez de só restaurar este estado).
    FSavedTabVisible: array of Boolean;
    FInListView: Boolean;
    // Guia de cada mês da tela de extratos (1 = tbJan ... 12 = tbDez): este
    // array é a tradutora mês<->guia, e não mais duas listas de doze braços
    // (o case de AbaDoMes e o if-chain de MesDaAba). Preenchido no FormCreate,
    // antes de qualquer troca de página.
    FMeses: array[1..12] of TTabSheet;
    // Título da tela normal (mês) antes de entrar na lista de bancos
    FSavedTitle: String;
    // Valor do tsAnterior (a abertura do mês ativo) guardado como NÚMERO. É
    // ele que semeia a coluna Saldo da grade: o SQLite soma trnamt linha a
    // linha a partir dele (função de janela SUM(...) OVER (...)), então o
    // número tem de existir ANTES de o SELECT ser montado - o Caption do
    // tsAnterior é texto formatado ("37,50 C"), não serve para a query.
    FSaldoAnterior: Double;
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
    // do cbAccount - mesma regra dos filtros); zero quando não existe. Guarda
    // o valor em FSaldoAnterior, que é a SEMENTE da coluna Saldo da grade -
    // daí rodar ANTES do SELECT ser montado em AplicarFiltroExtratos.
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
    // Regra de aparecimento do "Excluir Mês" (btnMonthClear): só na tela de
    // extratos (guia de mês visível - ver o TabVisible no corpo), com
    // database aberto (sem arquivo não há o que excluir) e com o "Show
    // Controls" ligado, que é o que revela os controles do rodapé.
    // PageControl1Change e o toggle usam esta MESMA função, para o botão não
    // sobreviver a uma troca de página ou de arquivo.
    function PodeExcluirMes: Boolean;
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
    // ---- Helpers compartilhados pelos fluxos (implementação no topo do .pas)
    // Conta/ano escolhidos nas combos: o combo carrega o id em Items.Objects,
    // e ItemIndex fora de faixa é o estado da combo vazia (-1).
    function IdContaSelecionada: Integer;
    // A mesma conta virando condição de WHERE (vazia quando não há conta).
    function CondicaoConta: string;
    function AnoSelecionado: Integer;
    // ' AND ' entre condições de um WHERE em montagem, só quando necessário.
    procedure AcrescentarCondicao(var ACondicao: string; const ANova: string);
    // Consulta própria na conexão das grades (o chamador é quem libera).
    function NovaConsulta: TSQLQuery;
    // Fecha as três queries da tela antes de uma gravação.
    procedure FecharQueriesContas;
    // Caminho relativo do diálogo de arquivo vira caminho na pasta inicial.
    function CaminhoAbsoluto(const AArquivo, APastaInicial: string): string;
    // Refaz combos, guias e grades depois de gravar, devolvendo conta/ano.
    procedure RecarregarTelaAposGravacao(const AContaTexto, AAnoTexto: string;
      AIrParaExtratos: Boolean);
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
  unitDatabase, unitOfx, unitMoneyFormatos;

{$R *.lfm}

{ TFormMoney }

// ---- Helpers compartilhados pelos fluxos da tela -------------------------------------------
// (conta e ano escolhidos nas combos, consulta própria na conexão das
// grades, fechar/religar as queries antes e depois de gravar). Estavam
// escritos à mão em cada lugar: o id da conta do cbAccount aparecia DEZ vezes
// como Integer(PtrInt(cbAccount.Items.Objects[ItemIndex])), com a checagem de
// faixa do lado, e o bloco "gravou, refaz a tela" era copiado três vezes.
// Agora cada regra existe uma vez só. As conversões de dinheiro, de data e o
// Plural foram para unitMoneyFormatos (funções puras, sem nada do form).

// Conta escolhida no cbAccount: o combo guarda o id de "contas" em
// Items.Objects (PtrInt, porque o id é numérico e cabe em endereço). -1
// quando não há conta selecionada - é o estado da combo vazia, sem contas
// cadastradas, e nela o filtro não restringe nada.
function TFormMoney.IdContaSelecionada: Integer;
begin
  Result := -1;
  if (cbAccount.ItemIndex >= 0) and (cbAccount.ItemIndex < cbAccount.Items.Count)
  then
    Result := Integer(PtrInt(cbAccount.Items.Objects[cbAccount.ItemIndex]));
end;

// A MESMA conta virando parte de um WHERE: 'account_id = N', ou a cadeia
// vazia quando não há conta selecionada - a mesma regra de combo vazia não
// restringe, compartilhada pelo filtro de extratos e pelo de saldos.
function TFormMoney.CondicaoConta: string;
var
  idConta: Integer;
begin
  Result := '';
  idConta := IdContaSelecionada;
  if idConta >= 0 then
    Result := 'account_id = ' + IntToStr(idConta);
end;

// Ano escolhido no cbYear: -1 quando não há ano nenhum (combo vazia - é o
// estado depois de um fechar, e o da lista vazia de "saldos"). O critério dos
// "4 dígitos" é o MESMO do preenchimento do combo: qualquer outra coisa nessa
// posição não é ano para ninguém.
function TFormMoney.AnoSelecionado: Integer;
begin
  Result := -1;
  if (cbYear.ItemIndex >= 0) and (cbYear.ItemIndex < cbYear.Items.Count) then
    Result := StrToIntDef(Trim(cbYear.Items[cbYear.ItemIndex]), -1);
end;

// Acrescenta uma condição a um WHERE em montagem, colocando o AND na frente
// só quando já havia algo - era o `if condicao <> '' then condicao :=
// condicao + ' AND '` que cada filtro novo tinha de lembrar de repetir.
procedure TFormMoney.AcrescentarCondicao(var ACondicao: string;
  const ANova: string);
begin
  if ANova = '' then
    Exit;
  if ACondicao <> '' then
    ACondicao := ACondicao + ' AND ';
  ACondicao := ACondicao + ANova;
end;

// Consulta própria na MESMA conexão das grades: Database e Transaction
// juntos, como o SQLDB exige (sem a transação a consulta não sabe onde
// entrar). Não mexe no cursor de nenhuma grade - é o motivo de existir, já
// que percorrer a query da grade derrubaria a linha que estiver em edição e
// moveria o registro selecionado. O chamador é quem libera.
function TFormMoney.NovaConsulta: TSQLQuery;
begin
  Result := TSQLQuery.Create(nil);
  Result.Database := SQLite3ConnContas;
  Result.Transaction := SQLTransactionContas;
end;

// As três queries da tela fecham antes de qualquer GRAVAÇÃO: o commit não
// pode encontrar sentença em andamento (o SQLite recusa o commit nesse caso).
// Quem grava é quem reabre, junto com o recarregar da tela abaixo.
procedure TFormMoney.FecharQueriesContas;
begin
  SQLQueryContas.Close;
  SQLQueryExtratos.Close;
  SQLQuerySaldos.Close;
end;

// Caminho relativo (sem pasta) fica na pasta inicial do diálogo, para o
// arquivo não cair no diretório de trabalho corrente da aplicação - a mesma
// regra dos cinco diálogos de arquivo (abrir, salvar, importar, exportar e
// novo database).
function TFormMoney.CaminhoAbsoluto(const AArquivo, APastaInicial: string
  ): string;
begin
  Result := AArquivo;
  if ExtractFilePath(Result) = '' then
    Result := IncludeTrailingPathDelimiter(APastaInicial) + Result;
end;

// Recarrega combos, guias e as duas grades depois de uma GRAVAÇÃO, e devolve
// o usuário para a conta e o ano que ele estava olhando (a recarga das combos
// começa sempre no primeiro item, então as escolhas são devolvidas por
// texto). AIrParaExtratos cobre o importar/excluir: essas telas também caem
// na tela de extratos, que é onde as linhas aparecem. Compartilhado pelo
// miImport, pelo miImpTrans e pelo "Excluir Mês" - e o miNew/miOpen usam a
// mesma sequência de filtros, sem a devolução de escolhas (ali não havia
// escolha a devolver).
procedure TFormMoney.RecarregarTelaAposGravacao(const AContaTexto,
  AAnoTexto: string; AIrParaExtratos: Boolean);
var
  indice: Integer;
begin
  if not SQLQueryContas.Active then
    SQLQueryContas.Open;
  // A lista de anos vem de "saldos", que a gravação acabou de mexer, e a de
  // contas de "contas"; as duas combos são remontadas vazias por Carregar-
  // FiltrosExtratos e recebem de volta a escolha logo abaixo.
  CarregarFiltrosExtratos;
  indice := cbAccount.Items.IndexOf(AContaTexto);
  if indice >= 0 then
    cbAccount.ItemIndex := indice;
  indice := cbYear.Items.IndexOf(AAnoTexto);
  if indice >= 0 then
    cbYear.ItemIndex := indice;
  // As duas grades reabrem já filtradas: a de extratos mostra o que acabou
  // de entrar, e a de saldos acompanha a mesma conta.
  AplicarFiltroExtratos;
  AplicarFiltroSaldos;
  if AIrParaExtratos then
    IrParaTelaExtratos;
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
  // A tradutora mês<->guia é montada aqui, ANTES de qualquer troca de página:
  // PageControl1Change, chamado logo abaixo, já consulta MesDaAba.
  FMeses[1] := tbJan;
  FMeses[2] := tbFev;
  FMeses[3] := tbMar;
  FMeses[4] := tbAbr;
  FMeses[5] := tbMai;
  FMeses[6] := tbJun;
  FMeses[7] := tbJul;
  FMeses[8] := tbAgo;
  FMeses[9] := tbSet;
  FMeses[10] := tbOut;
  FMeses[11] := tbNov;
  FMeses[12] := tbDez;
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
  // ... e o "Excluir Mês" é do MESMO grupo de controles do rodapé: nasce
  // escondido, some nas telas de gestão e sem database, e respeita o toggle
  // quando o painel voltar.
  btnMonthClear.Visible := PodeExcluirMes;

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
  // (IdContaSelecionada devolve -1 justamente nesse caso da combo vazia.)
  idEscolhido := IdContaSelecionada;

  cbAccount.Items.Clear;
  consulta := NovaConsulta;
  try
    // hidden = 1 e' a conta fora do filtro: ela continua GRAVADA (os extratos
    // e os saldos dela nao somem), mas nao aparece no combo - o cbAccount e' o
    // filtro da tela de extratos, nao a lista de cadastro (a grade de tbContas
    // continua mostrando todas). O esquema garante NOT NULL DEFAULT 0, entao
    // nao ha' linha sem marcador que escape do WHERE.
    consulta.SQL.Text := 'SELECT id, acctid, description FROM contas' +
      ' WHERE hidden = 0 ORDER BY id;';
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
  consulta := NovaConsulta;
  try
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

// Guia de mês da tela de extratos: 1 = tbJan ... 12 = tbDez, pelo array
// FMeses montado no FormCreate - a tradução é UMA só (a de antes tinha um
// case de doze braços aqui e um if-chain de doze do outro lado, com o mesmo
// conteúdo escrito duas vezes). Protege o chamador de um mês vindo fora do
// intervalo: devolve nil, como antes.
function TFormMoney.AbaDoMes(AMes: Integer): TTabSheet;
begin
  if (AMes >= 1) and (AMes <= 12) then
    Result := FMeses[AMes]
  else
    Result := nil;
end;

// Tradutor inverso de AbaDoMes: a aba ATIVA é quem diz qual mês a grade de
// extratos mostra (o pedido: ano do cbYear + mês da guia ativa). Devolve 0
// para o que não é guia de mês - as telas de gestão, onde a grade fica
// escondida e não há mês a filtrar - e também para guia nula, que é o estado
// da página sem guia ativa. Os comparadores são objetos (TTabSheet não é
// ordinal), então a procura é linha a linha no array.
function TFormMoney.MesDaAba(ATab: TTabSheet): Integer;
var
  i: Integer;
begin
  Result := 0;
  if ATab = nil then
    Exit;
  for i := 1 to 12 do
    if FMeses[i] = ATab then
    begin
      Result := i;
      Break;
    end;
end;

function TFormMoney.PodeExcluirMes: Boolean;
begin
  // MesDaAba cobre a troca de pagina: ele devolve 0 fora das guias de mes (as
  // tres telas de gestao), entao o botao some junto com o rodape da tela
  // errada - e devolve 0 tambem para ActivePage nulo, que e' o estado da
  // pagina sem guia ativa. (E' tambem o que segura a ordem dos testes: o
  // ActivePage so' e' olhado depois de ter passado por esses dois.)
  // O TabVisible cobre o caso inverso: database aberto e NENHUM mes com saldo
  // - ai' o AtualizarAbasMes nao acha guia e o LCL deixa a pagina sem guia
  // ativa, com o AtualizarAbasMes voltando para tbJan, que e' a pagina de
  // ENTRADA mas com a guia escondida. Sem guia na tela nao ha' mes na tela, e
  // apagar o mes 01 do ano seria apagar por suposicao. (E' tambem o que
  // segura a ordem dos testes: o ActivePage so' e' olhado depois de existir.)
  Result := DatabaseAberto and toggleShowControls.Checked and
    (MesDaAba(PageControl1.ActivePage) > 0) and
    PageControl1.ActivePage.TabVisible;
end;

// As guias JAN..DEZ são o cabeçalho de navegação da tela de extratos e só
// aparecem nos meses que têm saldo no ano do cbYear (mesma origem da lista
// de anos: saldos.enddate). A aba ativa vai para o mês mais recente - é o
// que define "hoje" quando o database abre e quando o ano muda.
procedure TFormMoney.AtualizarAbasMes;
var
  i, ano, mes, ultimoMes, mesLancado, mesAbertura, idConta: Integer;
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

  // Ano do combo (AnoSelecionado aplica o mesmo critério de "4 dígitos" do
  // preenchimento dele): sem ano escolhido (combo vazia) não há mês para
  // mostrar. Database já conferido no início do procedimento.
  ano := AnoSelecionado;

  mesLancado := 0;
  if ano >= 0 then
  begin
    // Consulta própria na mesma conexão (mesmo caminho do ano do cbYear):
    // meses do ano escolhido, direto de enddate. enddate é texto ISO, então
    // o mês são os caracteres 6..7 ("2026-03-31" -> "03").
    consulta := NovaConsulta;
    try
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
      idConta := IdContaSelecionada;
      condicao := '';
      if idConta >= 0 then
        condicao := ' AND account_id = ' + IntToStr(idConta);
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
  ano, mes: Integer;
begin
  // Sem database aberto não há o que filtrar: é o caminho do OnChange que a
  // limpeza das combos dispara ao fechar o arquivo.
  if not DatabaseAberto then
    Exit;

  condicao := '';
  // Conta: o combo carrega o "id" de "contas" em Items.Objects. Combo vazia
  // (nenhuma conta cadastrada) não restringe e valem todas as contas.
  AcrescentarCondicao(condicao, CondicaoConta);

  // Período = ano + mês: o ano vem do cbYear e o mês da aba ativa (guia
  // JAN..DEZ) - são as duas coisas que definem a tela de extratos. dtposted
  // é texto ISO (YYYY-MM-DD...), então o ano são os 4 primeiros caracteres -
  // que também batem no formato OFX (YYYYMMDD...). O valor é validado como
  // número antes de entrar na query e a comparação é contra texto, para o
  // SQLite não trocar de tipo no meio da expressão.
  ano := AnoSelecionado;
  if ano >= 0 then
    AcrescentarCondicao(condicao, 'substr(dtposted, 1, 4) = ' +
      QuotedStr(IntToStr(ano)));

  // Mês = guia ativa (MesDaAba devolve 0 fora das guias de mês: as telas de
  // gestão, onde a grade está escondida e não há mês a filtrar). O mês não
  // fica na mesma posição nos dois formatos de dtposted: em YYYY-MM-DD são
  // os caracteres 6..7 e em YYYYMMDD são os 5..6 - o '-' na posição 5 é o
  // que separa os casos (é o mesmo critério do DataExibicao, que formata a
  // coluna da grade). O mês é comparado como texto de 2 dígitos ('01'..'12'),
  // que é o que substr devolve.
  mes := MesDaAba(PageControl1.ActivePage);
  if mes > 0 then
    AcrescentarCondicao(condicao,
      '(CASE WHEN substr(dtposted, 5, 1) = ''-'' THEN substr(dtposted, 6, 2)' +
      ' ELSE substr(dtposted, 5, 2) END) = ' +
      QuotedStr(Format('%.2d', [mes])));

  // A abertura do mês (o valor do tsAnterior) é também a SEMENTE da coluna
  // Saldo da grade: o SQLite soma o trnamt linha a linha a partir dela
  // (função de janela SUM(...) OVER (...)). O número tem de existir ANTES de
  // o SELECT ser montado, então este cálculo veio para cima - antes vivia no
  // fim, junto com o txtSaldo, que continua ali.
  AtualizarSaldoAnterior;

  // O * não serve por causa do "trnamt": a coluna é NUMERIC e o driver
  // SQLite a tipa como TLargeintField, que LÊ por sqlite3_column_int64 - os
  // centavos de um -198,50 apareciam como "198,00 D" e a GRAVAÇÃO de uma
  // edição destruía a fração (o bind também era int64). O CAST faz o campo
  // nascer TFloatField (leitura por sqlite3_column_double, gravação por
  // sqlite3_bind_double) e o AS mantém o NOME da coluna - que é o que o sqldb
  // usa para gerar o INSERT/UPDATE ("trnamt"=:"trnamt"), então a edição
  // continua valendo para o arquivo. As demais colunas vêm nomeadas e sem
  // CAST, iguais ao *.
  // A coluna saldo é a MESMA ideia aplicada à soma corrente: o OVER usa o
  // MESMO ORDER BY do fim da query (dtposted, id - id é chave, então não há
  // empate), e a semente FSaldoAnterior entra como constante. Assim a 1a
  // linha já vem com abertura + trnamt e cada linha seguinte soma a anterior
  // - o saldo acumulado do mês, na ordem em que a grade mostra. O CAST final
  // é o mesmo motivo do trnamt: segurar o tipo como REAL.
  // E o parser do sqldb continua achando a query atualizável: em ppSelect só
  // o FROM muda o estado, e o SUM/OVER/ORDER de dentro dos parênteses são
  // pulados inteiros como um fraseado só (é por isso que a edição do trnamt
  // continua gravando).
  // O MESMO CAST com outro motivo no "dtposted": a coluna é TEXT e o driver
  // SQLite a tipa como ftMemo, que a TDBGrid trata como BLOB - e a grade só
  // entrega a tecla ao editor de bloco com dgDisplayMemoText (sem a opção,
  // EditorCanAcceptKey zera a tecla: o editor da coluna Data abria e nada
  // entrava nele). O CAST AS VARCHAR faz o campo nascer TStringField (o
  // mesmo truque do enddate dos saldos), a tecla volta a entrar, e o AS
  // mantém o NOME - o desenho continua passando pelo OnGetText, que formata
  // DD/MM/AAAA, e a edição continua valendo para o arquivo. O tamanho que o
  // driver dá ao VARCHAR sem precisão (255) folga para as datas.
  // memo e chknum ficam TEXT de propósito: sem a opção da grade também não
  // aceitariam tecla, mas CAST a ftString os limitaria a 255 caracteres
  // (o driver trunca na leitura) - com dgDisplayMemoText na grade eles
  // editam pelo texto cru e mantêm o tamanho livre do texto.
  sql := 'SELECT id, account_id, trntype, CAST(dtposted AS VARCHAR) AS dtposted,' +
    ' CAST(trnamt AS REAL) AS trnamt, memo, chknum,' +
    ' CAST(' + SqlNumero(FSaldoAnterior) +
    ' + SUM(CAST(trnamt AS REAL)) OVER (ORDER BY dtposted, id) AS REAL)' +
    ' AS saldo FROM extratos';
  if condicao <> '' then
    sql := sql + ' WHERE ' + condicao;
  sql := sql + ' ORDER BY dtposted, id;';

  // Mudou o filtro, muda a query: só dá para trocar o SQL com a query
  // fechada, então fecha, reescreve e reabre na sequência.
  SQLQueryExtratos.Close;
  SQLQueryExtratos.SQL.Text := sql;
  SQLQueryExtratos.Open;
  PrepararCamposExtratos;

  // O txtSaldo é o fechamento do mês (saldo lançado DENTRO do mês ativo) e
  // não entra em nada da query, então continua sendo calculado aqui: este é
  // o único ponto que roda em todas as mudanças de guia, de ano, de conta, na
  // abertura e na importação. (O tsAnterior também é daqui, mas ANTES do
  // SELECT - ver acima.)
  AtualizarSaldoMes;
end;

// O mês vem da guia ativa (é a guia que diz qual mês a grade filtra) e o
// ano do cbYear (é ele que define a lista de anos e as guias). Sem
// database/mês/ano válido não há o que consultar - é também o estado do
// fechar, onde a tela inteira já está escondida.
function TFormMoney.PeriodoExtratos(out ano, mes: Integer): Boolean;
begin
  mes := MesDaAba(PageControl1.ActivePage);
  ano := AnoSelecionado;
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
  if IdContaSelecionada >= 0 then
    condConta := ' AND ' + CondicaoConta;

  consulta := NovaConsulta;
  try
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
// assume o papel do sinal e o valor é o absoluto, com 2 casas, vírgula
// decimal e separador de milhar no ponto - zero conta como crédito (>= 0), a
// mesma regra do trnamt. A cor acompanha o sinal como na grade, MAS só o
// negativo é vermelho: o positivo fica na cor do próprio campo (o pedido foi
// "na cor atual", não o azul da grade). Os DOIS caminhos religam a cor de
// propósito - sem isso, depois de um saldo negativo o vermelho ficaria
// grudado no campo. Campo é TControl porque os dois exibidores são disso
// para baixo: o tsAnterior (TStaticText) e o txtSaldo (TLabel).
procedure TFormMoney.ExibirSaldo(Campo: TControl; valor: Double);
begin
  // A cor é daqui: só o negativo é vermelho e o positivo volta para a cor do
  // próprio campo (os DOIS caminhos religam de propósito - sem isso, depois
  // de um saldo negativo o vermelho ficaria grudado no campo). O TEXTO é o
  // mesmo par C/D da grade, em ValorCD.
  if valor >= 0 then
    Campo.Font.Color := clDefault
  else
    Campo.Font.Color := clRed;
  Campo.Caption := ValorCD(valor);
end;

// O tsAnterior mostra o balance do saldo imediatamente anterior ao mês ativo
// - o saldo com que o mês começa. "Imediatamente anterior" = o registro mais
// recente de saldos ANTERIOR ao primeiro dia do mês da guia ativa (no ano do
// cbYear, que é o que define o ano desse limite). Não havendo nenhum, o
// pedido é exibir zero - que aparece como crédito ("0,00 C").
// O valor é guardado em FSaldoAnterior antes de exibir porque é ele a
// semente da coluna Saldo da grade (a soma corrente do SQLite começa nele) -
// e essa semente precisa existir ANTES de o SELECT ser montado.
procedure TFormMoney.AtualizarSaldoAnterior;
var
  ano, mes: Integer;
begin
  if PeriodoExtratos(ano, mes) then
    // enddate é texto ISO (AAAA-MM-DD - formato que o próprio form grava),
    // então a comparação de texto ordena por data: o limite é o 1º dia do
    // mês ATIVO e o que vale é o registro mais recente antes dele.
    FSaldoAnterior := ConsultarSaldo('enddate < ' +
      QuotedStr(Format('%.4d-%.2d-01', [ano, mes])))
  else
    FSaldoAnterior := 0;
  ExibirSaldo(tsAnterior, FSaldoAnterior);
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
begin
  if not DatabaseAberto then
    Exit;

  condicao := CondicaoConta;
  if condicao <> '' then
    condicao := ' WHERE ' + condicao;

  // O CAST do balance é o mesmo motivo do CAST do trnamt na grade de
  // extratos: "NUMERIC" sem precisão o driver SQLite tipa como
  // TLargeintField, que LÊ por sqlite3_column_int64 - os centavos de um
  // 19,90 sumiam na exibição e a GRAVAÇÃO de uma edição destruía a fração
  // (o bind também era int64, e "19.9" virava "19"). O CAST faz o campo
  // nascer TFloatField (leitura por sqlite3_column_double, gravação por
  // sqlite3_bind_double) e o AS mantém o NOME da coluna - que é o que o
  // sqldb usa para gerar o INSERT/UPDATE ("balance"=:"balance"), então a
  // edição continua valendo para o arquivo.
  // O CAST do enddate é o mesmo truque com outro tipo: "TEXT" o driver tipa
  // como ftMemo, que é BLOB para a TDBGrid - e a grade só entrega a tecla ao
  // editor de blob com dgDisplayMemoText, opção que lê o texto cru do campo
  // e IGNORA o OnGetText que mostra DD/MM/AAAA. CAST AS VARCHAR faz o campo
  // nascer TStringField (não blob): a TDBGrid de base aceita a digitação, o
  // EditMask !99/00/0000 continua casando e o OnSetText grava AAAA-MM-DD -
  // e o AS mantém o nome, então a edição segue gravando no arquivo. As
  // demais colunas (id, account_id) vêm sem CAST.
  sql := 'SELECT id, account_id, CAST(balance AS REAL) AS balance, ' +
    'CAST(enddate AS VARCHAR) AS enddate FROM saldos' + condicao + ';';

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
// do arquivo passa direto). A coluna tem máscara de data (mesmo caso do
// enddate da tbSaldos), e a máscara exige o texto JÁ no formato DD/MM/AAAA
// também no editor (DisplayText=False) - com o valor cru a edição começaria
// fora das posições da máscara e a data viraria lixo. O que fica gravado não
// muda: é o OnSetText que converte de volta.
procedure TFormMoney.ExtratoDtpostedGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
begin
  // A tradução dos DOIS formatos de dtposted (ISO e o AAAAMMDD do OFX) mora
  // em DataExibicao, junto com a do enddate dos saldos.
  AText := DataExibicao(Sender.AsString);
end;

// Aceita o que a máscara produz (DD/MM/AAAA) e também o formato gravado
// (AAAA-MM-DD, para colar de fora). Dia/mês fora do lugar não é data: o
// valor fica como estava, sem inventar conversão.
procedure TFormMoney.ExtratoDtpostedSetText(Sender: TField; const AText: string);
var
  iso: string;
begin
  if Trim(AText) = '' then
  begin
    // dtposted é NOT NULL: vazio vira texto vazio, que o SQLite aceita.
    Sender.AsString := '';
    Exit;
  end;
  // DataISO aceita o que a máscara entrega (DD/MM/AAAA) e também o formato
  // gravado (AAAA-MM-DD), validando a data de verdade (31/02 não passa).
  // False = não é data: o valor fica como estava, sem inventar conversão.
  if DataISO(AText, iso) then
    Sender.AsString := iso;
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

// trnamt e saldo viram "20,00 C" (crédito) ou "20,00 D" (débito): a letra
// assume o papel do sinal e o valor é o absoluto, com 2 casas, vírgula
// decimal e separador de milhar no ponto ("1.234,56 C"). Serve as DUAS
// colunas porque lê tudo pelo Sender. O editor da grade lê o valor cru
// (DisplayText=False), então a máscara não muda o que se digita - no saldo
// não há editor, que é justamente o ReadOnly do PrepararCamposExtratos.
procedure TFormMoney.ExtratoTrnAmtGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
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
  // O par C/D (a letra no lugar do sinal, valor absoluto, duas casas e
  // vírgula) é o mesmo do tsAnterior e do txtSaldo: mora em ValorCD.
  AText := ValorCD(Sender.AsFloat);
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

// Mostra "Sim"/"Nao" no lugar do 0/1 gravado: o campo e' INTEGER, entao a
// grade desenharia "0"/"1" cru sem este handler - a mesma traducao que o
// BancoIdGetText faz do id do banco para o nome. DisplayText = False e' o
// texto que vai para o EDITOR: a mesma palavra, entao a celula abre ja
// legivel e o combo da PickList acha a entrada certa. Campo vazio fica
// vazio - nao ha' o que traduzir.
procedure TFormMoney.ContaHiddenGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
begin
  if Sender.IsNull then
  begin
    AText := '';
    Exit;
  end;
  if Sender.AsInteger = 0 then
    AText := 'Não'
  else
    AText := 'Sim';
end;

// Gravacao da celula: o DBGrid escreve via Field.Text, entao o "Sim"/"Nao"
// tem de voltar para 0/1. Aceita o que a grade mostra, a mesma palavra SEM
// acento (quem digitar) e o proprio 0/1 (colar de planilha). Qualquer outro
// texto NAO mexe no valor: apagar a celula nao pode virar "oculta" sem
// querer, e lixo nao pode corromper uma coluna NOT NULL.
procedure TFormMoney.ContaHiddenSetText(Sender: TField; const AText: string);
var
  texto: string;
begin
  texto := Trim(AText);
  if (texto = '1') or SameText(texto, 'Sim') then
    Sender.AsInteger := 1
  else if (texto = '0') or (texto = '') or SameText(texto, 'Não') or
    SameText(texto, 'Nao') then
    Sender.AsInteger := 0;
end;
// Liga os handlers de banco logo depois de cada Open: os campos são
// dinâmicos (FieldDefs vazio) e cada abertura os recria sem handlers — o
// mesmo problema, e a mesma solução, dos extratos.
procedure TFormMoney.SQLQueryContasAfterOpen(DataSet: TDataSet);
begin
  SQLQueryContas.FieldByName('bankid').OnGetText := @BancoIdGetText;
  SQLQueryContas.FieldByName('bankid').OnSetText := @BancoIdSetText;
  SQLQueryContas.FieldByName('hidden').OnGetText := @ContaHiddenGetText;
  SQLQueryContas.FieldByName('hidden').OnSetText := @ContaHiddenSetText;
end;

// Registro novo da tbContas nasce VISIVEL: "hidden" e' NOT NULL e no registro
// novo o campo ainda esta vazio (a coluna "Oculta" existe na grade, mas nada
// foi digitado nela). Sem isto o INSERT que o SQLDB montaria mandaria NULL e o
// SQLite recusaria a linha - o DEFAULT 0 do esquema so' vale quando a coluna
// e' OMITIDA do INSERT, e o SQLDB nomeia todas as colunas do SELECT *.
procedure TFormMoney.SQLQueryContasNewRecord(DataSet: TDataSet);
begin
  DataSet.FieldByName('hidden').AsInteger := 0;
end;

// Registro novo da tbSaldos nasce com a conta escolhida no cbAccount: a
// coluna account_id não aparece na grade (é ela que filtra), então quem
// preenche é aqui. Sem conta selecionada (combo vazia) o valor fica vazio e
// a gravação recusa - account_id é NOT NULL, e não há conta para sugerir.
procedure TFormMoney.SQLQuerySaldosNewRecord(DataSet: TDataSet);
var
  idConta: Integer;
begin
  idConta := IdContaSelecionada;
  if idConta >= 0 then
    DataSet.FieldByName('account_id').AsInteger := idConta;
end;

// Registro novo da grade de extratos: as duas colunas NOT NULL que a grade
// não mostra. account_id é a conta escolhida no cbAccount - é ela que filtra
// a tela, então é ela que o registro novo tem de pertencer (mesmo caso da
// tbSaldos, cujo handler faz o mesmo); sem conta selecionada o campo fica
// vazio e a gravação recusa. trntype não aparece em lugar nenhum da
// interface e nasce 'OTHER', o mesmo padrão dos leitores de OFX/OFC quando
// o arquivo não informa o tipo. As demais colunas NOT NULL (dtposted,
// trnamt) estão na grade e o usuário digita; as nulas ficam vazias.
procedure TFormMoney.SQLQueryExtratosNewRecord(DataSet: TDataSet);
var
  idConta: Integer;
begin
  idConta := IdContaSelecionada;
  if idConta >= 0 then
    DataSet.FieldByName('account_id').AsInteger := idConta;
  DataSet.FieldByName('trntype').AsString := 'OTHER';
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
// fundo de destaque não dão para ler. A coluna saldo entra na mesma regra do
// trnamt: ela também é dinheiro com sinal, e o saldo acumulado do mês costuma
// cruzar o zero (débito vermelho no fim do mês).
procedure TFormMoney.gridTransPrepareCanvas(Sender: TObject; DataCol: Integer;
  Column: TColumn; AState: TGridDrawState);
begin
  if (Column = nil) or ((Column.FieldName <> 'trnamt') and
    (Column.FieldName <> 'saldo')) then
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

// Liga os handlers de exibição da grade de extratos - e a máscara de data da
// coluna Data (mesmo caso do enddate da tbSaldos; religada a cada Open porque
// os campos são dinâmicos e cada abertura os recria). O saldo vem daqui:
// a coluna saldo nasce da expressão da query, com os ProviderFlags padrão do
// TField ([pfInUpdate, pfInWhere]) - e com eles o sqldb geraria um
// UPDATE ... "saldo"=:"saldo", que o SQLite recusaria ("no such column") a
// cada edição da grade. Zerar os ProviderFlags tira o campo do SET e do
// WHERE; o ReadOnly tira ele também da edição da célula (é calculado - não
// há o que digitar). E o texto é o MESMO par C/D do trnamt, que já formata
// qualquer campo numérico pelo Sender.
procedure TFormMoney.PrepararCamposExtratos;
begin
  SQLQueryExtratos.FieldByName('dtposted').OnGetText :=
    @ExtratoDtpostedGetText;
  SQLQueryExtratos.FieldByName('dtposted').OnSetText :=
    @ExtratoDtpostedSetText;
  SQLQueryExtratos.FieldByName('dtposted').EditMask := '!99/00/0000;1;_';
  SQLQueryExtratos.FieldByName('memo').OnGetText := @ExtratoTextoGetText;
  SQLQueryExtratos.FieldByName('chknum').OnGetText := @ExtratoTextoGetText;
  SQLQueryExtratos.FieldByName('trnamt').OnGetText := @ExtratoTrnAmtGetText;
  SQLQueryExtratos.FieldByName('saldo').OnGetText := @ExtratoTrnAmtGetText;
  SQLQueryExtratos.FieldByName('saldo').ReadOnly := True;
  SQLQueryExtratos.FieldByName('saldo').ProviderFlags := [];
end;

// enddate é TEXT (ftMemo no driver SQLite): sem dgDisplayMemoText a grade
// desenha pelo DisplayText - e, como a coluna tem máscara, o texto tem de vir
// DD/MM/AAAA também no editor (DisplayText=False), senão a edição começaria
// com o valor ISO cru, fora das posições da máscara. O que fica gravado não
// muda: é o OnSetText que converte de volta.
procedure TFormMoney.SaldosEnddateGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
begin
  // enddate é ISO (AAAA-MM-DD, o formato que o próprio form grava, e o único
  // que os filtros e as guias sabem ler): a tradução é a MESMA do dtposted
  // dos extratos, em DataExibicao, e texto que não é data volta intacto.
  AText := DataExibicao(Sender.AsString);
end;

// Aceita o que a máscara produz (DD/MM/AAAA) e também o formato gravado
// (AAAA-MM-DD, para colar de fora). Dia/mês fora do lugar não é data: o valor
// fica como estava, sem inventar conversão.
procedure TFormMoney.SaldosEnddateSetText(Sender: TField; const AText: string);
var
  iso: string;
begin
  if Trim(AText) = '' then
  begin
    // enddate é NOT NULL: vazio vira texto vazio, que o SQLite aceita.
    Sender.AsString := '';
    Exit;
  end;
  // DataISO aceita o que a máscara entrega (DD/MM/AAAA) e também o formato
  // gravado (AAAA-MM-DD, para colar de fora), validando a data de verdade
  // (31/02 não passa). False = não é data: o valor fica como estava, sem
  // inventar conversão.
  if DataISO(AText, iso) then
    Sender.AsString := iso;
end;

// balance é numérico: DUAS casas sempre, vírgula decimal, separador de milhar
// no ponto e o sinal na frente - "20,00" para positivo, "-20,00" para
// negativo e "-1.234,50" no milhar. A grade de saldos não tem o par C/D da
// grade de extratos (ali a letra assume o papel do sinal da coluna de
// valor); aqui o próprio sinal conta crédito/débito, que é também como o
// arquivo guarda. DisplayText=False devolve o valor cru, que é o que o
// editor da grade abre - o gravado não muda com a formatação.
procedure TFormMoney.SaldosBalanceGetText(Sender: TField; var AText: string;
  DisplayText: Boolean);
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
  // Duas casas, vírgula decimal e milhar no ponto: o MESMO texto de antes,
  // agora em FormatarDinheiro - que é de onde o ValorCD da grade de extratos
  // também tira o número.
  AText := FormatarDinheiro(Sender.AsFloat);
end;

// Gravação da célula: o DBGrid faz Field.Text := texto do editor e o
// SetEditText do campo vem parar aqui (mesmo caminho do enddate). Aceita os
// DOIS separadores - "19,9", o que a grade exibe e o que o editor da grade
// deixa digitar (o separador do Windows), e "19.9", o formato do arquivo e
// do SQL - trocando a vírgula por ponto e lendo com separador FIXO, que não
// depende do formato do sistema. "20" e "19,90" gravam o MESMO número que a
// grade mostra com duas casas. Texto vazio ou que não seja número não mexe
// no valor, como no banco de tbBancos (balance é NOT NULL, então vazio não
// pode virar NULL).
procedure TFormMoney.SaldosBalanceSetText(Sender: TField; const AText: string);
var
  formatos: TFormatSettings;
  texto: string;
  valor: Double;
begin
  texto := Trim(AText);
  if texto = '' then
    Exit;
  texto := StringReplace(texto, ',', '.', []);
  formatos := DefaultFormatSettings;
  formatos.DecimalSeparator := '.';
  formatos.ThousandSeparator := #0;
  if TryStrToFloat(texto, valor, formatos) then
    Sender.AsFloat := valor;
end;

// Liga máscara de data e os handlers de enddate E de balance logo depois de
// cada Open: os campos são dinâmicos (FieldDefs vazio) e cada abertura os
// recria - mesmo problema, mesma solução, dos extratos e do bankid das
// contas. A máscara "!99/00/0000;1;_" fixa as posições pela esquerda, aceita
// só dígito nas 8 casas, insere as duas barras sozinha e, pelo "1" do
// segundo campo, devolve o texto COM as barras ao campo (é o que o
// OnSetText espera ler).
procedure TFormMoney.PrepararCamposSaldos;
begin
  SQLQuerySaldos.FieldByName('balance').OnGetText := @SaldosBalanceGetText;
  SQLQuerySaldos.FieldByName('balance').OnSetText := @SaldosBalanceSetText;
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

procedure TFormMoney.miAboutClick(Sender: TObject);
begin
  // O dialogo e' o de INFORMACAO com so' o mbOK: sem escolha para o usuario,
  // entao nao ha' botao de cancelar nem desfazer - e nada muda de estado, o
  // "Sobre" e' uma placa de identificacao do programa.
  MessageDlg(
    'Este programa tem como objetivo gerenciar extratos bancários. ' +
    'Os dados podem ser inseridos manualmente ou obtidos de arquivo ' +
    'ofc/ofx. Também é possível importar ou exportar o database em ' +
    'formato excel csv. Dúvidas ou sugestões, envie email para ' +
    'neimestre@yahoo.com.br com o assunto projectMoney.' + LineEnding +
    'Copyright © 2019 - 2021 - 2026 - @neimestre',
    mtInformation, [mbOK], 0);
end;

procedure TFormMoney.miCloseClick(Sender: TObject);
begin
  // "Fechar Database" encerra a ligação com o arquivo de contas (o mesmo que
  // "Novo Database"/"Abrir Database" trocam de lugar). O banks.db não é
  // afetado: ele abre na inicialização e sustenta a lista de bancos.
  // Sempre dá para fechar duas vezes: fechar algo já fechado é só no-op.
  try
    FecharQueriesContas;
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

procedure TFormMoney.miExitClick(Sender: TObject);
begin
  // "Sair" é o "Fechar Database" seguido do fim da aplicação: o miClose faz
  // o fechar das três queries e da conexão - com o aviso de falha dele e com
  // o recálculo do estado da tela, que aqui não atrapalha (a janela some em
  // seguida). O Terminate é quem derruba o processo, e ele NÃO passa pelos
  // eventos de fechamento do form: sem o fechar antes, a aplicação sairia
  // com a conexão ainda ligada ao arquivo.
  miCloseClick(nil);
  Application.Terminate;
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
  ignorados, novos, jaExistiam, contaId: Integer;
begin
  // Cada linha de "extratos" aponta para uma conta (account_id NOT NULL),
  // entao sem conta escolhida nao ha para onde gravar. O menu "Transacoes"
  // ja' nasce desativado sem database aberto, entao nao falta checar isso.
  contaId := IdContaSelecionada;
  if contaId < 0 then
  begin
    MessageDlg('Cadastre uma conta antes de importar o extrato.' + LineEnding +
      'Cada linha de "extratos" pertence a uma conta da aba ' +
      '"Gerenciar Contas".', mtWarning, [mbOK], 0);
    Exit;
  end;
  // A recarga das combos no fim comeca sempre no primeiro item: guarda as
  // duas escolhas para devolver o usuario para onde ele estava.
  contaTexto := cbAccount.Text;
  anoTexto := cbYear.Text;

  dlgImportar.InitialDir := ExtractFilePath(ParamStr(0));
  if not dlgImportar.Execute then
    Exit; // usuario cancelou

  // Caminho relativo (sem pasta) fica na pasta inicial do dialogo - a mesma
  // regra dos outros quatro dialogos de arquivo, em CaminhoAbsoluto.
  caminho := CaminhoAbsoluto(dlgImportar.FileName, dlgImportar.InitialDir);

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
    consulta := NovaConsulta;
    try
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
      FecharQueriesContas;
      novos := ImportarExtratos(SQLite3ConnContas, contaId, registros,
        jaExistiam);
    except
      on E: Exception do
        erro := 'Não foi possível importar "' + ExtractFileName(caminho) +
          '".' + LineEnding + E.Message;
    end;
  finally
    try
      // Recarrega as combos (a lista de anos vem de saldos e a importacao
      // GRAVA um saldo por mes - o balanco acumulado ate' o fim de cada mes -,
      // entao o ano do fim do arquivo pode aparecer aqui), devolve a conta e
      // o ano que o usuario tinha escolhido, reabre as duas grades ja'
      // filtradas e cai na tela de extratos (sem navegar, quem importou de
      // uma tela de gestao ficaria la, sem ver nada do que entrou).
      RecarregarTelaAposGravacao(contaTexto, anoTexto, True);
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
  // incompleto. Linha igual a uma ja' gravada NAO entra de novo, entao
  // reimportar o MESMO arquivo nao faz segunda copia: o dialogo diz que os
  // dados ja' existem. Quando so' parte do arquivo e' nova, a segunda linha
  // do aviso mostra quantas linhas foram mantidas como estavam.
  if (jaExistiam > 0) and (novos = 0) then
    mensagem := 'Os dados deste extrato já existem em "' + contaTexto + '".' +
      LineEnding + LineEnding + IntToStr(jaExistiam) + ' ' +
      Plural(jaExistiam, 'registro já estava gravado.',
      'registros já estavam gravados.')
  else
  begin
    mensagem := 'Importação concluída.' + LineEnding + LineEnding +
      IntToStr(novos) + ' ' + Plural(novos, 'registro adicionado em "',
      'registros adicionados em "') + contaTexto + '".';
    if jaExistiam > 0 then
      mensagem := mensagem + LineEnding + IntToStr(jaExistiam) + ' ' +
        Plural(jaExistiam, 'registro já existia e foi mantido.',
        'registros já existiam e foram mantidos.');
  end;
  if ignorados > 0 then
    mensagem := mensagem + LineEnding + IntToStr(ignorados) + ' ' +
      Plural(ignorados, 'bloco ignorado (sem data ou sem valor).',
      'blocos ignorados (sem data ou sem valor).');
  MessageDlg(mensagem, mtInformation, [mbOK], 0);
end;

procedure TFormMoney.miImpTransClick(Sender: TObject);
var
  caminho, nomeArquivo, contaTexto, anoTexto, mensagem, erro: string;
  registros: TRegistrosOfx;
  ignorados, novos, contaId: Integer;
begin
  // Mesmo destino do miImport: a conta escolhida no filtro. O CSV nao traz
  // conta nenhuma (o modelo so' tem ID, TRNTYPE, DTPOSTED, TRNAMT, CHKNUM e
  // MEMO), entao sem conta escolhida nao ha' para onde gravar - e, como a
  // importacao SUBSTITUI as tabelas, errar de destino nao teria volta.
  contaId := IdContaSelecionada;
  if contaId < 0 then
  begin
    MessageDlg('Cadastre uma conta antes de importar as transações.' +
      LineEnding + 'O CSV vira o extrato inteiro da conta escolhida no ' +
      'filtro.', mtWarning, [mbOK], 0);
    Exit;
  end;

  // A recarga das combos no fim comeca sempre no primeiro item: guarda as
  // duas escolhas para devolver o usuario para onde ele estava.
  contaTexto := cbAccount.Text;
  anoTexto := cbYear.Text;

  dlgImpCsv.InitialDir := ExtractFilePath(ParamStr(0));
  if not dlgImpCsv.Execute then
    Exit; // usuario cancelou

  // Caminho relativo (sem pasta) fica na pasta inicial do dialogo - a mesma
  // regra dos outros dialogos de arquivo, em CaminhoAbsoluto.
  caminho := CaminhoAbsoluto(dlgImpCsv.FileName, dlgImpCsv.InitialDir);
  nomeArquivo := ExtractFileName(caminho);

  // Le o arquivo ANTES de mexer no database: aqui a gravacao e' uma
  // SUBSTITUICAO, entao falha de leitura nao pode ter apagado nada.
  ignorados := 0;
  try
    registros := LerCsv(caminho, ignorados);
  except
    on E: Exception do
    begin
      MessageDlg('Não foi possível ler o arquivo "' + nomeArquivo + '".' +
        LineEnding + E.Message, mtError, [mbOK], 0);
      Exit;
    end;
  end;

  // Sem nenhuma linha legivel nao ha' o que substituir - e dizer o formato
  // esperado e' o que resolve (o leitor ja' diz o que nao achou no cabecalho).
  if Length(registros) = 0 then
  begin
    MessageDlg('Nenhuma transação encontrada em "' + nomeArquivo + '".' +
      LineEnding + 'Esperava um CSV com separador ";" e uma primeira linha ' +
      'de cabeçalho com as colunas DTPOSTED e TRNAMT, e em cada linha data ' +
      'em AAAA-MM-DD e valor com ponto decimal (ex.: -80.90).',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  novos := 0;
  erro := '';
  // As queries fecham antes da gravacao (mesmo motivo do miImport): o commit
  // da substituicao nao pode encontrar sentenca em andamento. O finally
  // religa tudo e refaz filtros + grade, e vale mesmo quando falha no meio.
  try
    try
      FecharQueriesContas;
      novos := ImportarCsv(SQLite3ConnContas, contaId, registros);
    except
      on E: Exception do
        erro := 'Não foi possível importar "' + nomeArquivo + '".' +
          LineEnding + E.Message;
    end;
  finally
    try
      // Recarrega as combos (a lista de anos vem de saldos, que a
      // substituicao REFAZ - sobra so' o periodo do CSV, entao o ano pode
      // mudar de valor e ate' de quantidade), devolve a conta e o ano que o
      // usuario tinha escolhido, reabre as duas grades ja' filtradas e cai
      // na tela de extratos, que e' onde as linhas novas aparecem.
      RecarregarTelaAposGravacao(contaTexto, anoTexto, True);
    except
      on E: Exception do
        if erro = '' then
          erro := 'As transações foram gravadas, mas a tela não pôde ser ' +
            'atualizada.' + LineEnding + E.Message;
    end;
  end;

  if erro <> '' then
  begin
    MessageDlg(erro, mtError, [mbOK], 0);
    Exit;
  end;

  // O usuario tem de saber as DUAS coisas: o que entrou e o que foi
  // substituido - apagar "extratos" e "saldos" e' parte do que ele pediu.
  mensagem := 'Importação concluída.' + LineEnding + LineEnding +
    IntToStr(novos) + ' ' + Plural(novos, 'transação importada em "',
    'transações importadas em "') + contaTexto + '".' + LineEnding +
    'Os dados anteriores de "extratos" e "saldos" foram substituídos.';
  if ignorados > 0 then
    mensagem := mensagem + LineEnding + IntToStr(ignorados) + ' ' +
      Plural(ignorados, 'linha ignorada (sem data ou sem valor válido).',
      'linhas ignoradas (sem data ou sem valor válido).');
  MessageDlg(mensagem, mtInformation, [mbOK], 0);
end;

procedure TFormMoney.miExpTransClick(Sender: TObject);
var
  caminho, contaTexto, mensagem: string;
  contaId, exportados: Integer;
begin
  // Mesmo destino do miImpTrans: a conta escolhida no filtro. Exportar é o
  // passo inverso, então vale o mesmo aviso da falta de conta - sem conta
  // escolhida não há extrato nenhum para levar para o arquivo.
  contaId := IdContaSelecionada;
  if contaId < 0 then
  begin
    MessageDlg('Escolha uma conta antes de exportar as transações.' +
      LineEnding + 'O CSV leva o extrato inteiro da conta escolhida no filtro.',
      mtWarning, [mbOK], 0);
    Exit;
  end;
  contaTexto := cbAccount.Text;

  // O diálogo abre na pasta da aplicação, sugerindo "extratos.csv".
  dlgExpCsv.InitialDir := ExtractFilePath(ParamStr(0));
  dlgExpCsv.FileName := 'extratos.csv';
  if not dlgExpCsv.Execute then
    Exit; // usuário cancelou

  caminho := dlgExpCsv.FileName;
  // O filtro já é *.csv, mas ele pode digitar outro nome de extensão.
  if LowerCase(ExtractFileExt(caminho)) <> '.csv' then
    caminho := ChangeFileExt(caminho, '.csv');
  // Caminho relativo (sem pasta) fica na pasta inicial do diálogo - a mesma
  // regra dos outros diálogos de arquivo, em CaminhoAbsoluto.
  caminho := CaminhoAbsoluto(caminho, dlgExpCsv.InitialDir);

  // Exportar não mexe no que está gravado nem na tela: a leitura roda em outra
  // query, só de leitura, na MESMA conexão - então nada aqui fecha as queries
  // nem recarrega os filtros, como nos importar.
  try
    exportados := ExportarCsv(SQLite3ConnContas, contaId, caminho);
  except
    on E: Exception do
    begin
      MessageDlg('Não foi possível exportar "' + caminho + '".' + LineEnding +
        E.Message, mtError, [mbOK], 0);
      Exit;
    end;
  end;

  // O usuário precisa saber as DUAS coisas: quanto saiu e para onde - e
  // exportar zero linhas também é resposta, não é silêncio.
  mensagem := 'Exportação concluída.' + LineEnding + LineEnding +
    IntToStr(exportados) + ' ' + Plural(exportados, 'transação',
    'transações') + ' ' + Plural(exportados, 'exportada', 'exportadas') +
    ' de "' + contaTexto + '".' + LineEnding + 'Arquivo: ' + caminho;
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
  // Caminho relativo (sem pasta) fica na pasta inicial do diálogo - a mesma
  // regra dos outros diálogos de arquivo, em CaminhoAbsoluto.
  novoArquivo := CaminhoAbsoluto(novoArquivo, dlgNewDatabase.InitialDir);

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
    FecharQueriesContas;
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
  // Caminho relativo (sem pasta) fica na pasta inicial do diálogo - a mesma
  // regra dos outros diálogos de arquivo, em CaminhoAbsoluto.
  abrirArquivo := CaminhoAbsoluto(dlgOpenDatabase.FileName,
    dlgOpenDatabase.InitialDir);

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
    FecharQueriesContas;
    SQLite3ConnContas.Close;
    // Database criado antes do campo "hidden" nao tem a coluna, e sem ela a
    // query do cbAccount falharia logo abaixo. O ALTER roda numa CONEXAO
    // PROPRIA, exatamente enquanto a da interface esta fechada: duas
    // ligacoes ao mesmo arquivo brigariam pelo lock do SQLite, e um Commit
    // na conexao da interface fecharia as queries que vao abrir agora (e'
    // o mesmo motivo de ImportarCsv/ExcluirMes recarregarem apos o Commit).
    GarantirColunaHidden(abrirArquivo);
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
  btnMonthClear.Visible := PodeExcluirMes;
end;

procedure TFormMoney.btnMonthClearClick(Sender: TObject);
var
  ano, mes, apagados: Integer;
  erro, mesTexto, contaTexto, anoTexto: string;
begin
  // Botão escondido é o caminho das telas de gestão (e o da navegação sem
  // database); e mesmo visível a guia ativa pode ter virado outra entre o
  // clique e aqui. Só sai daqui com um mês em mãos.
  if not PodeExcluirMes then
    Exit;
  if not PeriodoExtratos(ano, mes) then
    Exit;

  // Mês e ano vêm do MESMO filtro que define a tela: a guia ativa dá o mês e
  // o cbYear o ano (PeriodoExtratos é o que a grade usa para dizer "este mês").
  mesTexto := Format('%.2d/%d', [mes, ano]);
  // Não há desfazer e não é recorte de conta: o aviso diz o mês INTEIRO, para
  // TODAS as contas - é o mês que vira guia, não a conta.
  if MessageDlg('Excluir todas as transações do mês ' + mesTexto +
    '? ' + 'O registro do mês também sai de "saldos", então a guia deixa de ' +
    'aparecer. Esta ação não pode ser desfeita.',
    mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
    Exit;

  // A recarga das combos no fim recomeça no primeiro item, então as duas
  // escolhas (conta e ano) são guardadas para devolver o usuário para onde
  // ele estava.
  contaTexto := cbAccount.Text;
  anoTexto := cbYear.Text;

  // As queries fecham antes da gravacao (mesmo motivo do miImport e do
  // miImpTrans): o commit do ExcluirMes nao pode encontrar sentenca em
  // andamento. O finally religa tudo e refaz combos + guias + grades, e vale
  // mesmo quando a exclusao falha no meio.
  erro := '';
  apagados := 0;
  try
    try
      FecharQueriesContas;
      apagados := ExcluirMes(SQLite3ConnContas, ano, mes);
    except
      on E: Exception do
        erro := 'Não foi possível excluir o mês ' + mesTexto + '.' +
          LineEnding + E.Message;
    end;
  finally
    try
      // A lista de anos vem de "saldos", e o mês excluído pode ter sido o
      // ÚNICO do ano - o ano então some do combo. O CarregarFiltrosExtratos
      // já refaz as guias (AtualizarAbasMes) pelo novo conteúdo da tabela, e
      // o ItemIndex do cbYear restaurado reexecuta esse recálculo, agora no
      // ano que o usuário estava olhando. Diferente dos importar, não há
      // navegação: excluir mês é feito na própria tela de extratos.
      RecarregarTelaAposGravacao(contaTexto, anoTexto, False);
    except
      on E: Exception do
        if erro = '' then
          erro := 'O mês ' + mesTexto + ' foi excluído, mas a tela não pôde ' +
            'ser atualizada.' + LineEnding + E.Message;
    end;
  end;

  if erro <> '' then
  begin
    MessageDlg(erro, mtError, [mbOK], 0);
    Exit;
  end;

  // O usuário tem de saber as DUAS coisas: o que sumiu da grade e que a guia
  // do mês também sumiu - foi o registro de "saldos" que a sustentava.
  MessageDlg(IntToStr(apagados) + ' ' +
    Plural(apagados, 'transação excluída do mês ',
    'transações excluídas do mês ') + mesTexto + '.' + LineEnding +
    'O registro do mês também saiu de "saldos", então a guia deixou de ' +
    'aparecer.', mtInformation, [mbOK], 0);
end;

end.

