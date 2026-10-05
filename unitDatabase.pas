unit unitDatabase;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils;

const
  DatabaseFileName = 'banks.db';

function DatabasePath: string;
procedure EnsureDatabase;

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

end.
