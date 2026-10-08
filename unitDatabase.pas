unit unitDatabase;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLDb, SQLite3Conn, unitOfx;

const
  DatabaseFileName = 'banks.db';

function DatabasePath: string;
procedure EnsureDatabase;
// Cria do zero o arquivo AFileName com as tabelas contas/extratos/saldos.
// Devolve False (com diálogo de erro) quando não consegue; nesse caso remove
// o arquivo incompleto. Se já existir, é substituído: quem confirma é o chamador.
function NewDatabase(const AFileName: string): Boolean;
// Confere se AFileName é um database do ProjectMoney: abre como SQLite e tem
// as tabelas "contas" (aba tbContas), "extratos" (grade da tela de extratos)
// e "saldos" (aba tbSaldos) — as três que as queries da interface ligam ao
// arquivo.
// Não cria nem altera arquivo e não mostra diálogo — False também para arquivo
// inexistente/quebrado, e quem avisa o usuário é o chamador.
function IsAccountsDatabase(const AFileName: string): Boolean;
// Grava as transacoes lidas de um OFX/OFC na tabela "extratos" da conta
// AAccountId. Usa a MESMA conexao/transacao que a interface ja' tem aberta:
// uma conexao nova contra o mesmo arquivo brigaria com o lock do SQLite.
// Todo registro lido entra: nao ha' conferencia de repeticao, nem dentro do
// proprio arquivo (bancos lancam lancamentos identicos no mesmo dia) nem
// contra o que ja' esta' gravado (reimportar e' uma escolha do usuario).
// Devolve quantas linhas entraram.
// Erro de banco propaga: o dialogo e' do chamador (que tambem sabe o arquivo).
function ImportarExtratos(AConn: TSQLite3Connection; AAccountId: Integer;
  const ARegistros: TRegistrosOfx): Integer;

implementation

uses
  Dialogs;

const
  SqlCreateBanks =
    'CREATE TABLE "banks" (' + LineEnding +
    #9'"id"'#9'TEXT NOT NULL,' + LineEnding +
    #9'"name"'#9'TEXT NOT NULL,' + LineEnding +
    #9'"alias"'#9'TEXT NOT NULL,' + LineEnding +
    #9'PRIMARY KEY("id")' + LineEnding +
    ');';

  // Database novo (menu "Novo Database"): contas + extratos + saldos.
  // PRIMARY KEY("id" AUTOINCREMENT) foi validado contra o sqlite3.dll do
  // projeto antes de virar constante (aceito pelo parser do SQLite).
  SqlCreateContas =
    'CREATE TABLE "contas" (' + LineEnding +
    #9'"id"'#9'INTEGER NOT NULL,' + LineEnding +
    #9'"acctid"'#9'TEXT NOT NULL,' + LineEnding +
    #9'"accttype"'#9'TEXT,' + LineEnding +
    #9'"bankid"'#9'INTEGER NOT NULL,' + LineEnding +
    #9'"branchid"'#9'TEXT NOT NULL,' + LineEnding +
    #9'"description"'#9'TEXT,' + LineEnding +
    #9'PRIMARY KEY("id" AUTOINCREMENT)' + LineEnding +
    ');';

  SqlCreateExtratos =
    'CREATE TABLE "extratos" (' + LineEnding +
    #9'"id"'#9'INTEGER NOT NULL,' + LineEnding +
    #9'"account_id"'#9'INTEGER NOT NULL,' + LineEnding +
    #9'"trntype"'#9'TEXT NOT NULL,' + LineEnding +
    #9'"dtposted"'#9'TEXT NOT NULL,' + LineEnding +
    #9'"trnamt"'#9'NUMERIC NOT NULL,' + LineEnding +
    #9'"memo"'#9'TEXT,' + LineEnding +
    #9'"chknum"'#9'TEXT,' + LineEnding +
    #9'PRIMARY KEY("id" AUTOINCREMENT),' + LineEnding +
    #9'FOREIGN KEY("account_id") REFERENCES "contas"("id")' + LineEnding +
    ');';

  SqlCreateSaldos =
    'CREATE TABLE "saldos" (' + LineEnding +
    #9'"id"'#9'INTEGER NOT NULL,' + LineEnding +
    #9'"account_id"'#9'INTEGER NOT NULL,' + LineEnding +
    #9'"balance"'#9'NUMERIC NOT NULL,' + LineEnding +
    #9'"enddate"'#9'TEXT NOT NULL,' + LineEnding +
    #9'PRIMARY KEY("id" AUTOINCREMENT),' + LineEnding +
    #9'FOREIGN KEY("account_id") REFERENCES "contas"("id")' + LineEnding +
    ');';

function DatabasePath: string;
begin
  Result := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0))) +
    DatabaseFileName;
end;

procedure EnsureDatabase;
var
  Conn: TSQLite3Connection;
  Trans: TSQLTransaction;
begin
  // Já existe? Nada a fazer.
  if FileExists(DatabasePath) then
    Exit;

  Conn := TSQLite3Connection.Create(nil);
  Trans := TSQLTransaction.Create(nil);
  try
    try
      // Abrir a conexão cria o arquivo vazio
      Conn.DatabaseName := DatabasePath;
      Conn.Transaction := Trans;
      Conn.Open;
      Trans.Active := True;
      Conn.ExecuteDirect(SqlCreateBanks);
      Trans.Commit;
    except
      on E: Exception do
      begin
        // Remove o arquivo incompleto para tentar de novo na próxima execução
        if FileExists(DatabasePath) then
          DeleteFile(DatabasePath);
        MessageDlg('Não foi possível criar o arquivo "' + DatabaseFileName +
          '".' + LineEnding + E.Message, mtError, [mbOK], 0);
      end;
    end;
  finally
    Trans.Free;
    Conn.Free;
  end;
end;

function NewDatabase(const AFileName: string): Boolean;
var
  Conn: TSQLite3Connection;
  Trans: TSQLTransaction;
begin
  Result := False;

  // "Novo" = arquivo do zero: substitui o que estiver no caminho escolhido
  // (a confirmação de substituição é responsabilidade do chamador).
  if FileExists(AFileName) then
    if not DeleteFile(AFileName) then
    begin
      MessageDlg('Não foi possível substituir o arquivo "' +
        ExtractFileName(AFileName) + '".', mtError, [mbOK], 0);
      Exit;
    end;

  Conn := TSQLite3Connection.Create(nil);
  Trans := TSQLTransaction.Create(nil);
  try
    try
      // Abrir a conexão cria o arquivo vazio
      Conn.DatabaseName := AFileName;
      Conn.Transaction := Trans;
      Conn.Open;
      Trans.Active := True;
      // Uma sentença por chamada, mesmo formato do EnsureDatabase
      Conn.ExecuteDirect(SqlCreateContas);
      Conn.ExecuteDirect(SqlCreateExtratos);
      Conn.ExecuteDirect(SqlCreateSaldos);
      Trans.Commit;
      Result := True;
    except
      on E: Exception do
      begin
        // Remove o arquivo incompleto para não deixar database quebrado
        if FileExists(AFileName) then
          DeleteFile(AFileName);
        MessageDlg('Não foi possível criar o database "' +
          ExtractFileName(AFileName) + '".' + LineEnding + E.Message,
          mtError, [mbOK], 0);
      end;
    end;
  finally
    Trans.Free;
    Conn.Free;
  end;
end;

function IsAccountsDatabase(const AFileName: string): Boolean;
var
  Conn: TSQLite3Connection;
  Trans: TSQLTransaction;
  Query: TSQLQuery;
begin
  Result := False;

  // Não existe? Nem abre: o sqlite3_open criaria um arquivo vazio no lugar.
  if not FileExists(AFileName) then
    Exit;

  Conn := TSQLite3Connection.Create(nil);
  Trans := TSQLTransaction.Create(nil);
  Query := TSQLQuery.Create(nil);
  try
    try
      Conn.DatabaseName := AFileName;
      Conn.Transaction := Trans;
      Conn.Open;
      Trans.Active := True;
      Query.Database := Conn;
      Query.Transaction := Trans;
      // O sqlite_master guarda cada tabela: procurar pelos nomes ali aceita
      // qualquer formatação do CREATE TABLE. As TRÊS tabelas que as queries
      // da interface ligam têm de existir ("contas" na tbContas, "extratos"
      // na grade da tela de extratos e "saldos" na tbSaldos), senão o arquivo
      // não serve para a aplicação.
      Query.SQL.Text := 'SELECT COUNT(*) FROM sqlite_master' +
        ' WHERE type = ''table'' AND name IN (''contas'', ''extratos'',' +
        ' ''saldos'');';
      Query.Open;
      Result := Query.Fields[0].AsInteger = 3;
      Query.Close;
    except
      // Arquivo que não é SQLite (ou está quebrado) estoura aqui.
      Result := False;
    end;
  finally
    Query.Free;
    Trans.Free;
    Conn.Free;
  end;
end;

function ImportarExtratos(AConn: TSQLite3Connection; AAccountId: Integer;
  const ARegistros: TRegistrosOfx): Integer;
var
  trans: TSQLTransaction;
  i: Integer;
  reg: TRegistroOfx;
begin
  Result := 0;
  if Length(ARegistros) = 0 then
    Exit;

  trans := AConn.Transaction;
  if trans = nil then
    raise Exception.Create('A conexao nao tem transacao para gravar o extrato.');

  try
    for i := 0 to High(ARegistros) do
    begin
      reg := ARegistros[i];

      // O valor entra como texto e a coluna NUMERIC converte na gravacao.
      // Nao ha' StrToFloat aqui de proposito: no pt-BR ele le "12.34"
      // trocando ponto por virgula.
      AConn.ExecuteDirect(
        'INSERT INTO extratos (account_id, trntype, dtposted, trnamt,' +
        ' memo, chknum) VALUES (' + IntToStr(AAccountId) + ', ' +
        QuotedStr(reg.TrnType) + ', ' + QuotedStr(reg.DtPosted) + ', ' +
        QuotedStr(reg.TrnAmt) + ', ' + QuotedStr(reg.Memo) + ', ' +
        QuotedStr(reg.ChkNum) + ');', trans);
      Inc(Result);
    end;

    // Uma transacao so' para a importacao inteira: ou entra tudo, ou nada.
    // O Commit fecha o que a interface tem aberto nessa transacao (opcao
    // sqoKeepOpenOnCommit mantem as queries ligadas), por isso o chamador
    // recarrega os filtros e a grade depois daqui.
    if trans.Active then
      trans.Commit;
  except
    // Falha no meio: desfaz as linhas que ja tinham entrado para o
    // extrato nao ficar pela metade.
    if trans.Active then
      trans.Rollback;
    raise;
  end;
end;

end.
