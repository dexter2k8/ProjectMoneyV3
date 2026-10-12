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

  Alem do OFC/OFX, este arquivo le tambem CSV de transacoes (LerCsv) - a
  mesma ideia por outro caminho: as duas leituras devolvem TRegistrosOfx,
  entao a GRAVACAO continua sendo uma so' (ImportarExtratos, em unitDatabase).
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

// Conta de origem informada pelo arquivo: o conteudo do primeiro <ACCTID>
// (BANKACCTFROM de OFX/OFC, CCACCTFROM e INVACCTFROM usam a propria tag).
// '' quando o arquivo nao informa conta nenhuma - nesse caso nao ha o que
// conferir no chamador, que decide se importa mesmo assim.
function LerAcctIdDoArquivo(const AFileName: string): string;

// Le o arquivo e devolve uma transacao por linha do extrato. AIgnorados recebe
// quantos blocos vieram sem data ou sem valor - esses nao viram linha nenhuma.
// Pode levantar excecao (arquivo nao existe, sem permissao, alem do tamanho):
// o dialogo e' do chamador.
function LerOfx(const AFileName: string; out AIgnorados: Integer): TRegistrosOfx;

// Le um CSV de transacoes (o modelo e' backup/extrato.csv) e devolve uma
// transacao por linha, nos MESMOS campos de TRegistroOfx - e' por isso que a
// gravacao nao tem caminho proprio: ela reaproveita o do importador de
// OFX/OFC. O contrato do arquivo, tirado do modelo:
//   - separador ";" e a PRIMEIRA linha de cabecalho, com as colunas lidas
//     PELO NOME (a ordem delas nao importa): DTPOSTED e TRNAMT sao
//     obrigatorias e as demais sao opcionais - TRNTYPE vira 'OTHER' e
//     CHKNUM/MEMO ficam vazios;
//   - ID e' ignorado de proposito: "extratos.id" e INTEGER AUTOINCREMENT,
//     entao a chave do arquivo e' do exportador e nao da tabela;
//   - data em AAAA-MM-DD (ou AAAAMMDD) e valor com PONTO decimal, como no
//     modelo: virgula nao e' numero para a coluna NUMERIC e a linha seria
//     gravada como texto no lugar do numero - por isso ela vira ignorada.
// AIgnorados conta as linhas puladas por data ou valor invalidos. Pode
// levantar excecao (arquivo ilegivel ou cabecalho fora do contrato); o
// dialogo do erro e' do chamador.
function LerCsv(const AFileName: string; out AIgnorados: Integer): TRegistrosOfx;

// O campo de um CSV na forma em que SepararCamposCsv o le' de volta: as aspas
// envolvem quando o texto tem ';' (ou '"', ou uma ponta com espaco, que o Trim
// do leitor tiraria) e os '"' internos sao duplicados - e' o inverso exato do
// que o leitor faz ao cortar a linha. CR/LF viram espaco ANTES de tudo, ja' que
// o leitor divide as linhas antes de separar os campos e texto com quebra de
// linha nao teria como voltar inteiro.
// E' o meio de gravar um arquivo que o proprio LerCsv rele' sem perder nada.
function EscaparCampoCsv(const ATexto: string): string;

implementation

uses
  StrUtils;

type
  // Campos de uma linha do CSV, depois do corte. Tipo da IMPLEMENTACAO de
  // proposito: so' o leitor de CSV divide linha, entao ele nao e' contrato de
  // quem chama.
  TCamposCsv = array of string;

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

// Mesma leitura (e mesma decodificacao de encoding) da importacao, via o
// proprio LerTextoUtf8: o chamador confere entao exatamente o texto que as
// transacoes teriam - nada de abrir o arquivo com outra regua.
function LerAcctIdDoArquivo(const AFileName: string): string;
begin
  Result := Trim(ExtrairTag('ACCTID', LerTextoUtf8(AFileName)));
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

// ---------------------------------------------------------------- CSV

// Corta uma linha do CSV em campos. Aspas so' contam quando ABREM o campo (e'
// a regra do formato): dentro delas o separador vale como texto e o par ""
// vira uma aspas literal. Uma aspas avulsa no MEIO do texto - comum em memo de
// banco (Pix recebido: "Cp :...") - nao fecha nada, justamente porque o campo
// nao comecou com aspas; sem esse detalhe aquelas memorias engoliriam o resto
// da linha. Cada campo sai com as aspas de protecao ja' tiradas e com Trim.
function SepararCamposCsv(const ALinha: string): TCamposCsv;
var
  i, comeco, total: Integer;
  emAspas: Boolean;

  // O texto de comeco ate AFim (bordas inclusas), sem as aspas de protecao -
  // que so' existem quando o campo COMECA e TERMINA com aspas.
  function CampoTexto(AFim: Integer): string;
  begin
    Result := Trim(Copy(ALinha, comeco, AFim - comeco + 1));
    if (Length(Result) >= 2) and (Result[1] = '"') and
      (Result[Length(Result)] = '"') then
      Result := StringReplace(Copy(Result, 2, Length(Result) - 2), '""', '"',
        [rfReplaceAll]);
  end;

begin
  Result := nil;
  total := 0;
  comeco := 1;
  emAspas := False;
  i := 1;
  while i <= Length(ALinha) do
  begin
    if emAspas then
    begin
      if ALinha[i] = '"' then
      begin
        if (i < Length(ALinha)) and (ALinha[i + 1] = '"') then
          Inc(i) // "" e' aspas literal: pula a segunda
        else
          emAspas := False;
      end;
    end
    else if (ALinha[i] = '"') and (i = comeco) then
      emAspas := True // so' a aspas da virgula inicial abre o campo
    else if ALinha[i] = ';' then
    begin
      SetLength(Result, total + 1);
      Result[total] := CampoTexto(i - 1);
      Inc(total);
      comeco := i + 1;
    end;
    Inc(i);
  end;
  SetLength(Result, total + 1);
  Result[total] := CampoTexto(Length(ALinha));
end;

// Escapa UM campo para SepararCamposCsv - veja o contrato na interface.
function EscaparCampoCsv(const ATexto: string): string;
var
  texto: string;
  precisaAspas: Boolean;
begin
  // A quebra de linha nao tem como sobreviver ao corte do leitor: vira
  // espaco, que e' o que a propria coluna de um banco normalmente guarda.
  texto := StringReplace(ATexto, #13, ' ', [rfReplaceAll]);
  texto := StringReplace(texto, #10, ' ', [rfReplaceAll]);

  // Precisa de aspas o campo que o corte da linha mudaria: o ';' viraria
  // outro campo, o Trim tiraria o espaco da ponta e o '"' das bordas seria
  // confundido com o envoltorio (CampoTexto so' desenvolve o campo que
  // COMECA e TERMINA com aspas).
  precisaAspas := (Pos(';', texto) > 0) or (Pos('"', texto) > 0) or
    ((texto <> '') and ((texto[1] = ' ') or
    (texto[Length(texto)] = ' ')));
  if not precisaAspas then
  begin
    Result := texto;
    Exit;
  end;

  // Por dentro, cada aspas vira duas - o inverso do '""' que CampoTexto
  // devolve unico ao desenvolver o envoltorio.
  Result := '"' + StringReplace(texto, '"', '""', [rfReplaceAll]) + '"';
end;

// Uma data do CSV nos DOIS formatos que a tabela aceita - AAAA-MM-DD (o do
// modelo) e AAAAMMDD (o do OFX) - e que o CALENDARIO confirma: nao basta ter
// 8 ou 10 caracteres, "2025-13-40" nao fecha mes nenhum.
function EhDataCsv(const ATexto: string): Boolean;
var
  numeros, ano, mes, dia: Integer;
  data: TDateTime;
begin
  Result := False;
  if (Length(ATexto) = 10) and (ATexto[5] = '-') and (ATexto[8] = '-') then
    numeros := StrToIntDef(Copy(ATexto, 1, 4) + Copy(ATexto, 6, 2) +
      Copy(ATexto, 9, 2), 0)
  else if Length(ATexto) = 8 then
    numeros := StrToIntDef(ATexto, 0)
  else
    Exit;
  if numeros <= 0 then
    Exit;
  ano := numeros div 10000;
  mes := (numeros div 100) mod 100;
  dia := numeros mod 100;
  Result := TryEncodeDate(ano, mes, dia, data);
end;

// Um TRNAMT no formato do modelo: [sinal] digitos [. digitos] - o PONTO e' o
// separador decimal e nada mais vale. Virgula (o formato do sistema) e' aceita
// DE PROPOSITO: guardar "1.234,56" na coluna NUMERIC deixaria texto no lugar
// do numero, entao a linha nao passa daqui e vira ignorada.
function EhNumeroCsv(const ATexto: string): Boolean;
var
  i, ponto: Integer;
  digitos, digitosDepois: Boolean;
begin
  Result := False;
  if ATexto = '' then
    Exit;
  ponto := 0;
  digitos := False;
  digitosDepois := False;
  for i := 1 to Length(ATexto) do
  begin
    if ATexto[i] in ['0'..'9'] then
    begin
      digitos := True;
      if ponto > 0 then
        digitosDepois := True;
    end
    else if ATexto[i] = '.' then
    begin
      if ponto > 0 then
        Exit; // dois pontos nao fecha numero nenhum
      Inc(ponto);
    end
    else if (ATexto[i] <> '+') and (ATexto[i] <> '-') then
      Exit // virgula, espaco ou letra: fora do formato do modelo
    else if i > 1 then
      Exit; // sinal: so' na frente, e a frente ja' passou
  end;
  // "12." nao tem fracao e "." sozinho nao e' numero: os dois o
  // StrToFloatDef do importador leria como zero.
  Result := digitos and ((ponto = 0) or digitosDepois);
end;

function LerCsv(const AFileName: string; out AIgnorados: Integer): TRegistrosOfx;
var
  texto, nome: string;
  campos: TCamposCsv;
  registros: TRegistrosOfx;
  reg: TRegistroOfx;
  colTipo, colData, colValor, colDoc, colMemo: Integer;
  total, capacidade, p, comeco: Integer;
  cabecalho: Boolean;

  // O campo de um indice que a linha talvez nao tenha: arquivo com menos
  // colunas que o cabecalho nao pode derrubar a leitura, e o que faltar e'
  // o mesmo que o leitor OFX faz com a tag ausente - fica vazio.
  function CampoEm(const ATodos: TCamposCsv; AIndice: Integer): string;
  begin
    if (AIndice >= 0) and (AIndice <= High(ATodos)) then
      Result := ATodos[AIndice]
    else
      Result := '';
  end;

  // Uma linha do arquivo: a PRIMEIRA nao vazia e' o cabecalho, as demais
  // viram linha de "extratos" quando passam na regua de data e de valor.
  procedure ProcessarLinha(const ATexto: string);
  var
    k: Integer;
  begin
    campos := SepararCamposCsv(ATexto);
    if (Length(campos) = 1) and (campos[0] = '') then
      Exit; // linha em branco

    if not cabecalho then
    begin
      // E' do cabecalho que sai a ORDEM das colunas - le-las pelo nome e' o
      // que deixa o arquivo aceitar as colunas trocadas sem virar lixo.
      cabecalho := True;
      colTipo := -1;
      colData := -1;
      colValor := -1;
      colDoc := -1;
      colMemo := -1;
      for k := 0 to High(campos) do
      begin
        nome := UpperCase(Trim(campos[k]));
        if nome = 'TRNTYPE' then
          colTipo := k
        else if nome = 'DTPOSTED' then
          colData := k
        else if nome = 'TRNAMT' then
          colValor := k
        else if (nome = 'CHKNUM') or (nome = 'CHECKNUM') then
          colDoc := k
        else if nome = 'MEMO' then
          colMemo := k;
      end;
      if (colData < 0) or (colValor < 0) then
        raise Exception.Create('O arquivo não é um CSV de transações no ' +
          'formato esperado: separador ";" e uma primeira linha de cabeçalho ' +
          'com as colunas DTPOSTED e TRNAMT.');
      Exit;
    end;

    reg.TrnType := CampoEm(campos, colTipo);
    reg.DtPosted := CampoEm(campos, colData);
    reg.TrnAmt := CampoEm(campos, colValor);
    reg.Memo := CampoEm(campos, colMemo);
    reg.ChkNum := CampoEm(campos, colDoc);

    // Data e valor sao as duas colunas NOT NULL da tabela, e as duas passam
    // pela regua do calendario e do formato numerico antes de virar linha.
    if (reg.DtPosted = '') or (reg.TrnAmt = '') or
      (not EhDataCsv(reg.DtPosted)) or (not EhNumeroCsv(reg.TrnAmt)) then
    begin
      Inc(AIgnorados);
      Exit;
    end;

    // Coluna trntype e' NOT NULL mas o arquivo pode nao informar o tipo
    if reg.TrnType = '' then
      reg.TrnType := 'OTHER';

    // Cresce em dobro (o mesmo cuidado do LerOfx): SetLength linha a linha
    // ficaria o tempo inteiro no realocar do array.
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
  end;

begin
  SetLength(registros, 0);
  total := 0;
  capacidade := 0;
  AIgnorados := 0;
  cabecalho := False;
  colTipo := -1;
  colData := -1;
  colValor := -1;
  colDoc := -1;
  colMemo := -1;

  texto := LerTextoUtf8(AFileName);
  // As quebras: o modelo usa CRLF, mas CSV gravado em Unix ou editado a maos
  // pode trazer LF - e um CR solto tambem vira linha, como nos arquivos
  // antigos. Normalizar aqui e' o que deixa UM so' codigo de leitura.
  texto := StringReplace(texto, #13#10, #10, [rfReplaceAll]);
  texto := StringReplace(texto, #13, #10, [rfReplaceAll]);

  comeco := 1;
  for p := 1 to Length(texto) do
    if texto[p] = #10 then
    begin
      ProcessarLinha(Copy(texto, comeco, p - comeco));
      comeco := p + 1;
    end;
  // A ultima linha so' existe aqui quando o arquivo nao fecha com a quebra
  if comeco <= Length(texto) then
    ProcessarLinha(Copy(texto, comeco, Length(texto) - comeco + 1));

  SetLength(registros, total);
  Result := registros;
end;

end.
