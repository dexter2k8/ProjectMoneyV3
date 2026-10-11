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
// Junto com as linhas entra tambem o balanco FINAL do periodo coberto pelo
// arquivo, gravado em "saldos" (ver GravarSaldoFimPeriodo): e' a mesma
// transacao, entao ou as linhas e o balanco gravam ou nada grava.
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

// Converte um DTPOSTED em data ISO (AAAA-MM-DD), que e' o formato com que o
// proprio form grava em saldos.enddate - e portanto o formato em que texto
// ordena como se ordena por data (e' o min/max do periodo que usa isso).
// Aceita os DOIS jeitos de o texto chegar: AAAAMMDD (o que o leitor de OFX
// grava, ja' com hora e fuso cortados) e AAAA-MM-DD (texto que ja' veio ISO
// e passou direto pelo NormalizarData). Devolve '' quando nao da' para saber
// a data - data invalida nao define periodo nenhum.
function DataIsoDtposted(const ADtPosted: string): string;
var
  i: Integer;
begin
  Result := '';
  if Length(ADtPosted) < 8 then
    Exit;
  if ADtPosted[5] = '-' then
  begin
    // Ja' esta' em AAAA-MM-DD: corta nos 10 primeiros caracteres, porque uma
    // sobra de hora/fuso no fim quebraria a ordenacao por texto.
    if Length(ADtPosted) >= 10 then
      Result := Copy(ADtPosted, 1, 10);
    Exit;
  end;
  for i := 1 to 8 do
    if not (ADtPosted[i] in ['0'..'9']) then
      Exit;
  Result := Copy(ADtPosted, 1, 4) + '-' + Copy(ADtPosted, 5, 2) + '-' +
    Copy(ADtPosted, 7, 2);
end;

// Ultimo dia do mes de uma data ISO (AAAA-MM-DD): e' o enddate do saldo
// final, o fechamento do ultimo mes coberto pelo arquivo. Bissexto no
// calendario gregoriano (multiplo de 4, exceto seculo nao multiplo de 400).
// Devolve '' quando o mes nao e' de 1 a 12.
function UltimoDiaDoMes(const ADataIso: string): string;
var
  ano, mes, dia: Integer;
begin
  ano := StrToIntDef(Copy(ADataIso, 1, 4), 0);
  mes := StrToIntDef(Copy(ADataIso, 6, 2), 0);
  case mes of
    1, 3, 5, 7, 8, 10, 12: dia := 31;
    4, 6, 9, 11: dia := 30;
    2:
      if ((ano mod 4 = 0) and (ano mod 100 <> 0)) or (ano mod 400 = 0) then
        dia := 29
      else
        dia := 28;
  else
    dia := 0;
  end;
  if (ano <= 0) or (dia = 0) then
    Result := ''
  else
    Result := Format('%.4d-%.2d-%.2d', [ano, mes, dia]);
end;

// Grava em "saldos" o balanco FINAL do periodo coberto pelo arquivo que
// acabou de ser importado: a abertura (o saldo registrado ANTES do 1o dia do
// 1o mes do arquivo) mais a soma de TODAS as linhas de "extratos" da conta
// dentro do periodo - nao so' as que acabaram de entrar, porque o balanco de
// um fim de mes e' o acumulado ate' la'. E' a MESMA conta que a coluna Saldo
// da grade calcula, so' que do 1o ao ultimo mes do arquivo em vez de um mes
// so'.
//
// O enddate e' o ULTIMO DIA do ultimo mes do arquivo (nao o dia da ultima
// transacao): e' o fechamento do periodo, o formato que o cbYear e as guias
// de mes leem de "saldos". Ja' existindo saldo da MESMA conta com esse
// MESMO enddate, o balance e' ATUALIZADO - uma linha por fim de mes, entao
// reimportar refaz o periodo em vez de acumular linha nova.
//
// Sem data valida em nenhum registro nao ha' periodo para fechar e nada e'
// gravado aqui (as linhas continuam entrando em "extratos"). E' o caminho
// tambem do arquivo recusado: quem chama so' chama depois de gravar.
//
// Roda DENTRO da transacao da importacao, antes do commit: ou as linhas e o
// balanco gravam juntos, ou nada grava.
procedure GravarSaldoFimPeriodo(AConn: TSQLite3Connection;
  AAccountId: Integer; const ARegistros: TRegistrosOfx);
var
  trans: TSQLTransaction;
  i: Integer;
  data, primeiraData, ultimaData, inicio, fim, mesInicial, mesFinal: string;
  sqlAbertura, sqlSoma, sqlValor: string;
begin
  trans := AConn.Transaction;
  if trans = nil then
    raise Exception.Create('A conexao nao tem transacao para gravar o saldo.');

  // O periodo vai do primeiro ao ultimo dia IMPORTADO. So' datas validas
  // decidem o periodo: uma data quebrada no meio do arquivo nao pode puxar
  // o comeco ou o fim para um lugar que o resto dos registros nega.
  primeiraData := '';
  ultimaData := '';
  for i := 0 to High(ARegistros) do
  begin
    data := DataIsoDtposted(ARegistros[i].DtPosted);
    if data = '' then
      Continue;
    if (primeiraData = '') or (data < primeiraData) then
      primeiraData := data;
    if (ultimaData = '') or (data > ultimaData) then
      ultimaData := data;
  end;
  if primeiraData = '' then
    Exit;

  // O saldo comeca no 1o dia do 1o mes e vale ate' o ultimo dia do ultimo
  // mes: sao dois fechamentos de mes inteiros, por isso o periodo cobre
  // meses COMPLETOS mesmo que o arquivo traga so' algumas datas.
  inicio := Copy(primeiraData, 1, 8) + '01';
  fim := UltimoDiaDoMes(ultimaData);
  if fim = '' then
    Exit;
  mesInicial := Copy(primeiraData, 1, 4) + Copy(primeiraData, 6, 2);
  mesFinal := Copy(ultimaData, 1, 4) + Copy(ultimaData, 6, 2);

  // Abertura: o saldo da conta registrado antes do periodo, mais recente
  // primeiro e com o MAIOR id desempatando - a mesma regra do saldo anterior
  // da tela (ConsultarSaldo). Sem registro anterior, comeca do zero.
  sqlAbertura := 'COALESCE((SELECT CAST(balance AS REAL) FROM saldos' +
    ' WHERE account_id = ' + IntToStr(AAccountId) +
    ' AND enddate < ' + QuotedStr(inicio) +
    ' ORDER BY enddate DESC, id DESC LIMIT 1), 0)';

  // Soma do periodo por MES (AAAAMM), que e' o unico nivel em que o
  // periodo faz sentido (as bordas sao dia 1 e ultimo dia de mes). O
  // concatena trata os DOIS formatos de dtposted - AAAAMMDD (o do OFX) e
  // AAAA-MM-DD (o da grade) - pelo MESMO CASE do filtro da tela, entao a
  // soma bate com o que a coluna Saldo exibe. COALESCE cobre o periodo sem
  // nenhuma linha, cujo SUM e' NULL.
  sqlSoma := 'COALESCE((SELECT SUM(CAST(trnamt AS REAL)) FROM extratos' +
    ' WHERE account_id = ' + IntToStr(AAccountId) +
    ' AND (substr(dtposted, 1, 4) ||' +
    ' (CASE WHEN substr(dtposted, 5, 1) = ''-'' THEN substr(dtposted, 6, 2)' +
    ' ELSE substr(dtposted, 5, 2) END))' +
    ' BETWEEN ' + QuotedStr(mesInicial) + ' AND ' + QuotedStr(mesFinal) +
    '), 0)';

  // O CAST AS REAL e' o mesmo motivo do CAST da grade: NUMERIC sem ele o
  // driver le como inteiro e os centavos somem na gravacao.
  sqlValor := 'CAST((' + sqlAbertura + ') + (' + sqlSoma + ') AS REAL)';

  // Atualiza primeiro e deixa o INSERT cobrir o caso de nao haver linha.
  // A subconsulta da abertura le "saldos" junto com o UPDATE, mas o
  // registro que esta' sendo atualizado fica fora da leitura (o dele e' o
  // fim do periodo, sempre posterior ao < inicio da abertura), entao nao ha'
  // auto-referencia - e a soma e' de "extratos", tabela outra.
  AConn.ExecuteDirect('UPDATE saldos SET balance = ' + sqlValor +
    ' WHERE account_id = ' + IntToStr(AAccountId) +
    ' AND enddate = ' + QuotedStr(fim) + ';', trans);

  // O NOT EXISTS e' o que segura UMA linha por (conta, fim de mes) mesmo
  // com as duas sentencas na mesma transacao: se o UPDATE achou, nao insere.
  AConn.ExecuteDirect('INSERT INTO saldos (account_id, balance, enddate)' +
    ' SELECT ' + IntToStr(AAccountId) + ', ' + sqlValor + ', ' +
    QuotedStr(fim) + ' FROM (SELECT 1 AS um)' +
    ' WHERE NOT EXISTS (SELECT 1 FROM saldos WHERE account_id = ' +
    IntToStr(AAccountId) + ' AND enddate = ' + QuotedStr(fim) + ');', trans);
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

    // O balanco do fim do periodo vai na MESMA transacao das linhas: as
    // duas coisas (extratos e saldos) gravam juntas ou nada grava.
    GravarSaldoFimPeriodo(AConn, AAccountId, ARegistros);

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
