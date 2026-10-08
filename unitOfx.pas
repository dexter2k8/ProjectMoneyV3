unit unitOfx;

{$mode objfpc}{$H+}

{
  Leitor de extratos no formato OFC/OFX.

  O formato existe em duas versoes e as duas sao lidas do mesmo jeito:
    - OFX 1.x e OFC: SGML solto, o valor da tag termina na quebra de linha
        <DTPOSTED>20240115
    - OFX 2.x: XML de verdade, a propria tag fecha
        <DTPOSTED>20240115</DTPOSTED>
  Nos dois casos o valor vai do ">" ate o proximo "<" (ou ate o fim da linha),
  que e' exatamente o que ProximaTag devolve - por isso nao ha dois codigos.

  O texto e' lido como UTF-8 porque e' o que o LCL (UTF8ToUTF16) e o SQLite
  esperam: arquivo em UTF-8 passa direto e o resto (extrato antigo com
  CHARSET:1252, o padrao dos bancos em portugues) e' convertido de
  Windows-1252 para UTF-8 byte a byte - sem depender do codepage da maquina.
}

interface

uses
  Classes, SysUtils;

type
  // Uma transacao do arquivo, ja com os campos na forma da tabela "extratos".
  TRegistroOfx = record
    TrnType: string;   // <TRNTYPE> DEBIT/CREDIT/... ('OTHER' quando falta)
    DtPosted: string;  // <DTPOSTED> AAAAMMDD (hora e fuso cortados)
    TrnAmt: string;    // <TRNAMT> valor com sinal e ponto decimal
    Memo: string;      // <MEMO> descricao
    ChkNum: string;    // <CHECKNUM>/<CHKNUM> numero do cheque ('' sem cheque)
  end;
  TRegistrosOfx = array of TRegistroOfx;

// Devolve a proxima tag de ATexto a partir de APos (1 = comeco do texto).
// ANome vem em maiusculas e sem os "<>"; "/" no inicio e' tag de fechamento
// (</STMTTRN>), o que sempre vem sem conteudo. AValor e' o conteudo aparado.
// APos anda junto: chamar de novo com a mesma variavel segue da onde parou,
// e False significa que o texto acabou.
function ProximaTag(const ATexto: string; var APos: Integer;
  out ANome, AValor: string): Boolean;

// Conteudo da PRIMEIRA <TAG> de ATexto ('' quando nao ha). E' a versao de um
// campo so' de ProximaTag, para tirar um valor avulso do arquivo.
function ExtrairTag(const ATag, ATexto: string): string;

// Le o arquivo e devolve uma transacao por linha do extrato. AIgnorados recebe
// quantos blocos vieram sem data ou sem valor - esses nao viram linha nenhuma.
// Pode levantar excecao (arquivo nao existe, sem permissao, alem do tamanho):
// o dialogo e' do chamador.
function LerOfx(const AFileName: string; out AIgnorados: Integer): TRegistrosOfx;

implementation

uses
  StrUtils;

const
  // Windows-1252: as 32 posicoes que nao sao Latin-1 (0x80..0x9F). Todo o
  // resto de 0xA0 a 0xFF e' igual ao Latin-1 e vira UTF-8 pela conta comum.
  Tabela1252: array[$80..$9F] of Word = (
    $20AC, $0081, $201A, $0192, $201E, $2026, $2020, $2021,
    $02C6, $2030, $0160, $2039, $0152, $008D, $017D, $008F,
    $0090, $2018, $2019, $201C, $201D, $2022, $2013, $2014,
    $02DC, $2122, $0161, $203A, $0153, $009D, $017E, $0178);

// True quando os bytes formam UTF-8 bem formado (byte de inicio 11xxxxxx
// seguido de continuacoes 10xxxxxx). ASCII sempre passa; texto em 1252 cai ja
// no primeiro acento, que ali e' um byte solto sem continuacao.
function EhUtf8(const ADados: TBytes): Boolean;
var
  i, continuacao: Integer;
  b: Byte;
begin
  Result := True;
  i := 0;
  while i < Length(ADados) do
  begin
    b := ADados[i];
    if b < $80 then
      continuacao := 0
    else if b < $C2 then
    begin
      // Continuacao solta ($80..$BF) ou UTF-8 sobreposto ($C0/$C1)
      Result := False;
      Exit;
    end
    else if b < $E0 then
      continuacao := 1
    else if b < $F0 then
      continuacao := 2
    else if b < $F5 then
      continuacao := 3
    else
    begin
      Result := False;
      Exit;
    end;
    if continuacao > 0 then
    begin
      // Falta byte no fim do arquivo?
      if i + continuacao >= Length(ADados) then
      begin
        Result := False;
        Exit;
      end;
      Inc(i);
      while continuacao > 0 do
      begin
        if (ADados[i] < $80) or (ADados[i] > $BF) then
        begin
          Result := False;
          Exit;
        end;
        Dec(continuacao);
        Inc(i);
      end;
    end
    else
      Inc(i);
  end;
end;

// Le o arquivo e devolve o texto em UTF-8, que e' o formato comum entre o
// arquivo, a interface e o banco. BOM e' tirado; UTF-8 valido vai direto; o
// resto e' Windows-1252 (extrato em portugues emitido em ANSI) convertido
// aqui, porque guardar os bytes do 1252 como estao deixaria acento torto na
// grade e no SQLite.
function LerTextoUtf8(const AFileName: string): string;
var
  fs: TFileStream;
  dados: TBytes;
  inicio, i, usados, caracter: Integer;
begin
  Result := '';
  dados := nil;
  fs := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyWrite);
  try
    SetLength(dados, fs.Size);
    if fs.Size > 0 then
      fs.ReadBuffer(dados[0], fs.Size);
  finally
    fs.Free;
  end;

  // BOM UTF-8 (EF BB BF) nao e' texto: sobra o conteudo depois dele. Os bytes
  // do BOM sao UTF-8 valido, entao passam no teste abaixo sem truque.
  inicio := 0;
  if (Length(dados) >= 3) and (dados[0] = $EF) and (dados[1] = $BB) and
    (dados[2] = $BF) then
    inicio := 3;

  if EhUtf8(dados) then
  begin
    SetLength(Result, Length(dados) - inicio);
    if Length(dados) > inicio then
      Move(dados[inicio], Result[1], Length(dados) - inicio);
    Exit;
  end;

  // Nao e' UTF-8: 1252 -> UTF-8. Da' folga de 3 bytes por caractere (o pior
  // caso do 1252) e corta no fim para o tamanho real.
  SetLength(Result, (Length(dados) - inicio) * 3);
  usados := 0;
  for i := inicio to High(dados) do
  begin
    if (dados[i] >= $80) and (dados[i] <= $9F) then
      caracter := Tabela1252[dados[i]]
    else
      caracter := dados[i];
    if caracter < $80 then
    begin
      Inc(usados);
      Result[usados] := AnsiChar(caracter);
    end
    else if caracter < $800 then
    begin
      Inc(usados, 2);
      Result[usados - 1] := AnsiChar($C0 or (caracter shr 6));
      Result[usados] := AnsiChar($80 or (caracter and $3F));
    end
    else
    begin
      Inc(usados, 3);
      Result[usados - 2] := AnsiChar($E0 or (caracter shr 12));
      Result[usados - 1] := AnsiChar($80 or ((caracter shr 6) and $3F));
      Result[usados] := AnsiChar($80 or (caracter and $3F));
    end;
  end;
  SetLength(Result, usados);
end;

function ProximaTag(const ATexto: string; var APos: Integer;
  out ANome, AValor: string): Boolean;
var
  abre, fecha, fim, i: Integer;
begin
  ANome := '';
  AValor := '';
  Result := False;
  if APos < 1 then
    APos := 1;
  if APos > Length(ATexto) then
    Exit;

  // Abre com "<" e so' vale tag fechada com ">"
  abre := PosEx('<', ATexto, APos);
  if (abre = 0) or (abre >= Length(ATexto)) then
    Exit;
  fecha := PosEx('>', ATexto, abre + 1);
  if fecha = 0 then
    Exit;

  ANome := UpperCase(Trim(Copy(ATexto, abre + 1, fecha - abre - 1)));
  if ANome = '' then
    Exit;

  // O conteudo vai do ">" ate o proximo "<" - mas no SGML a propria quebra de
  // linha encerra o valor, entao corta tambem em #10/#13.
  fim := Length(ATexto) + 1;
  if fecha < Length(ATexto) then
  begin
    i := PosEx('<', ATexto, fecha + 1);
    if i > 0 then
      fim := i;
    for i := fecha + 1 to fim - 1 do
      if ATexto[i] in [#10, #13] then
      begin
        fim := i;
        Break;
      end;
  end;
  AValor := Trim(Copy(ATexto, fecha + 1, fim - fecha - 1));

  // </TAG> e' fechamento (nao tem conteudo); <TAG/> e' tag sem conteudo,
  // entao vira uma tag comum com valor vazio.
  if ANome[1] = '/' then
    AValor := ''
  else if ANome[Length(ANome)] = '/' then
  begin
    SetLength(ANome, Length(ANome) - 1);
    AValor := '';
  end;

  APos := fim;
  Result := True;
end;

function ExtrairTag(const ATag, ATexto: string): string;
var
  posicao: Integer;
  nome, valor, alvo: string;
begin
  Result := '';
  alvo := UpperCase(Trim(ATag));
  if (alvo = '') or (ATexto = '') then
    Exit;
  posicao := 1;
  while ProximaTag(ATexto, posicao, nome, valor) do
    if nome = alvo then
    begin
      Result := valor;
      Exit;
    end;
end;

// DTPOSTED vem como AAAAMMDD e as vezes com hora e fuso no fim
// (20240115120000.000[-5:EST]). Guarda so' a data, que e' o que o filtro de
// ano le por substr(dtposted, 1, 4); texto ja' em formato ISO (2024-01-15)
// passa direto, porque os 8 primeiros caracteres nao sao todos digitos.
function NormalizarData(const AData: string): string;
var
  i: Integer;
begin
  Result := AData;
  if Length(Result) < 8 then
    Exit;
  for i := 1 to 8 do
    if not (Result[i] in ['0'..'9']) then
      Exit;
  Result := Copy(Result, 1, 8);
end;

function LerOfx(const AFileName: string; out AIgnorados: Integer): TRegistrosOfx;
var
  texto, nome, valor: string;
  posicao, total, capacidade: Integer;
  registros: TRegistrosOfx;
  reg: TRegistroOfx;

  // Encerra o bloco de transacao em andamento. Ele vira linha de "extratos"
  // quando tem data E valor (as duas colunas NOT NULL); conta como ignorado
  // quando tem so' parte dos dois; e some em silencio quando nao e'
  // transacao nenhuma (cabecalho, saldo, lista de contas) - por isso so'
  // campos de transacao entram em "reg". Depois zera para o proximo bloco.
  procedure FecharBloco;
  begin
    if (reg.DtPosted <> '') and (reg.TrnAmt <> '') then
    begin
      // Coluna trntype e' NOT NULL mas o arquivo pode nao informar o tipo
      if reg.TrnType = '' then
        reg.TrnType := 'OTHER';
      // Cresce em dobro: SetLength(1,1,1...) de novo em novo faria o tempo
      // inteiro ficar no realocar do array em arquivos grandes
      if total >= capacidade then
      begin
        if capacidade = 0 then
          capacidade := 64
        else
          capacidade := capacidade * 2;
        SetLength(registros, capacidade);
      end;
      registros[total] := reg;
      Inc(total);
    end
    else if (reg.TrnType <> '') or (reg.DtPosted <> '') or
      (reg.TrnAmt <> '') then
      Inc(AIgnorados);
    reg.TrnType := '';
    reg.DtPosted := '';
    reg.TrnAmt := '';
    reg.Memo := '';
    reg.ChkNum := '';
  end;

begin
  SetLength(registros, 0);
  total := 0;
  capacidade := 0;
  AIgnorados := 0;

  texto := LerTextoUtf8(AFileName);
  posicao := 1;
  while ProximaTag(texto, posicao, nome, valor) do
  begin
    // <STMTTRN> e </STMTTRN> sao as bordas do registro no OFX. Fechar tambem
    // na ABERTURA descarta o que vier antes (cabecalho do arquivo), que e'
    // texto fora de bloco nenhum. Arquivo sem essas marcas (OFC) e' pego
    // pelo teste de repeticao la embaixo.
    if (nome = 'STMTTRN') or (nome = '/STMTTRN') then
    begin
      FecharBloco;
      Continue;
    end;

    // Qualquer outra tag de fechamento e so' estrutura
    if (nome <> '') and (nome[1] = '/') then
      Continue;

    if nome = 'TRNTYPE' then
    begin
      if reg.TrnType <> '' then
        FecharBloco; // repetiu o tipo: comecou outro registro
      reg.TrnType := valor;
    end
    else if nome = 'DTPOSTED' then
    begin
      if reg.DtPosted <> '' then
        FecharBloco;
      reg.DtPosted := NormalizarData(valor);
    end
    else if nome = 'TRNAMT' then
    begin
      if reg.TrnAmt <> '' then
        FecharBloco;
      reg.TrnAmt := valor;
    end
    else if nome = 'MEMO' then
      reg.Memo := valor
    else if (nome = 'CHECKNUM') or (nome = 'CHKNUM') then
      reg.ChkNum := valor;
  end;
  // O ultimo bloco do arquivo so' fecha aqui
  FecharBloco;

  SetLength(registros, total);
  Result := registros;
end;

end.
