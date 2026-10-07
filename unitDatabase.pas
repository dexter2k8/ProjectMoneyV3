unit unitDatabase;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

const
  DatabaseFileName = 'banks.db';

function DatabasePath: string;
procedure EnsureDatabase;
// Cria do zero o arquivo AFileName com as tabelas contas/extratos/saldos.
// Devolve False (com diálogo de erro) quando não consegue; nesse caso remove
// o arquivo incompleto. Se já existir, é substituído: quem confirma é o chamador.
function NewDatabase(const AFileName: string): Boolean;
// Confere se AFileName é um database do ProjectMoney: abre como SQLite e tem
// as tabelas "contas" (aba tbContas) e "saldos" (aba tbSaldos) — as duas que
// as queries da interface ligam ao arquivo.
// Não cria nem altera arquivo e não mostra diálogo — False também para arquivo
// inexistente/quebrado, e quem avisa o usuário é o chamador.
function IsAccountsDatabase(const AFileName: string): Boolean;

implementation

uses
  Dialogs, SQLDb, SQLite3Conn;

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
      // qualquer formatação do CREATE TABLE. As DUAS tabelas que as queries
      // da interface ligam têm de existir ("contas" na tbContas e "saldos"
      // na tbSaldos), senão o arquivo não serve para a aplicação.
      Query.SQL.Text := 'SELECT COUNT(*) FROM sqlite_master' +
        ' WHERE type = ''table'' AND name IN (''contas'', ''saldos'');';
      Query.Open;
      Result := Query.Fields[0].AsInteger = 2;
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

end.
