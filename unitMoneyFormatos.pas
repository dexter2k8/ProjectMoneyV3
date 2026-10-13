unit unitMoneyFormatos;

{$mode objfpc}{$H+}

{ Conversões de apresentação compartilhadas pela tela de extratos.

  Por que um unit próprio: estas funções são PURAS - não mexem em query, em
  combo nem em formulário - e é exatamente o que a interface compartilha com
  as grades: como o dinheiro aparece, como a data aparece e volta a ser
  gravada, o plural das mensagens e o número virando literal de SQL. Tirá-las
  deixa o form com o que é dele, a INTERFACE (gráficos, filtros e queries).

  Os handlers OnGetText/OnSetText continuam no form de propósito: no modo
  objfpc do FPC um procedimento unit-level não pode ser atribuído a um evento
  - evento é "of object", que só método de classe satisfaz (compilado e
  conferido antes de decidir isto). }

interface

// "1 registro" / "2 registros": evita "registro(s)" na mensagem de resultado.
function Plural(AN: Integer; const ASingular, APlural: string): string;

// Número virando literal de SQL: o separador decimal é SEMPRE ponto, porque o
// SQL não entende vírgula - e o DefaultFormatSettings do programa é o do
// sistema, que aqui é português (vírgula). FloatToStr com formato próprio
// cobre isso; o valor é Double puro, então não há o que injetar.
function SqlNumero(valor: Double): string;

// Dinheiro na forma brasileira: DUAS casas sempre, vírgula decimal e
// separador de milhar no ponto - "1.234,56". É o texto que as grades e os
// textos de saldo mostram; o que fica GRAVADO não muda com isto (quem lê e
// grava o número é o campo, que continua lendo o valor cru).
function FormatarDinheiro(valor: Double): string;

// O par C/D da coluna de valor da grade de extratos (e do tsAnterior e do
// txtSaldo): a letra assume o papel do sinal e o valor é o absoluto - "20,00
// C" para crédito, "20,00 D" para débito, zero contando como crédito (>= 0,
// a mesma regra do campo). Serve as DUAS colunas numéricas da grade porque
// lê tudo pelo Sender, e também os dois exibidores de saldo, que cuidam da
// cor por fora.
function ValorCD(valor: Double): string;

// Data gravada virando texto de tela: aceita os DOIS formatos do campo
// dtposted (ISO "AAAA-MM-DD" e o "AAAAMMDD" do OFX, que o importador grava
// direto) e devolve "DD/MM/AAAA". Qualquer outra coisa volta INTACTA - não há
// o que traduzir e não se inventa data.
function DataExibicao(const aData: string): string;

// O inverso de DataExibicao, com validação: aceita o que a máscara da grade
// entrega ("DD/MM/AAAA") e também o formato gravado ("AAAA-MM-DD", para colar
// de fora) e devolve a data em ISO, que é o formato que o SQLite recebe.
// False quando não é data válida: dia/mês fora do lugar não é conversão, e o
// chamador mantém o valor como estava.
function DataISO(const aData: string; out aISO: string): Boolean;

implementation

uses
  SysUtils;

function Plural(AN: Integer; const ASingular, APlural: string): string;
begin
  if AN = 1 then
    Result := ASingular
  else
    Result := APlural;
end;

function SqlNumero(valor: Double): string;
var
  formatos: TFormatSettings;
begin
  formatos := DefaultFormatSettings;
  formatos.DecimalSeparator := '.';
  formatos.ThousandSeparator := #0;
  Result := FloatToStr(valor, formatos);
end;

// Os formatos de dinheiro deste unit (este e o do ValorCD) montam o
// TFormatSettings aqui com separador FIXO, que não depende do formato do
// sistema - a mesma ideia do SqlNumero, com o separador trocado: o SQL pede
// ponto, a grade pede vírgula.
function FormatarDinheiro(valor: Double): string;
var
  formatos: TFormatSettings;
begin
  formatos := DefaultFormatSettings;
  formatos.DecimalSeparator := ',';
  formatos.ThousandSeparator := '.';
  Result := FormatFloat('#,##0.00', valor, formatos);
end;

function ValorCD(valor: Double): string;
begin
  if valor >= 0 then
    Result := FormatarDinheiro(valor) + ' C'
  else
    Result := FormatarDinheiro(-valor) + ' D';
end;

function DataExibicao(const aData: string): string;
begin
  Result := aData;
  // AAAAMMDD (o formato do OFX, que o importador grava como está): só
  // dígitos, com dia em 7..8, mês em 5..6 e ano em 1..4.
  if (Length(Result) = 8) and (StrToIntDef(Result, -1) >= 0) then
    Result := Copy(Result, 7, 2) + '/' + Copy(Result, 5, 2) + '/' +
      Copy(Result, 1, 4)
  // AAAA-MM-DD (ISO, o formato que o próprio form grava, inclusive nos
  // saldos): o '-' na posição 5 é o que separa os dois casos.
  else if (Length(Result) = 10) and (Result[5] = '-') and (Result[8] = '-') then
    Result := Copy(Result, 9, 2) + '/' + Copy(Result, 6, 2) + '/' +
      Copy(Result, 1, 4);
end;

function DataISO(const aData: string; out aISO: string): Boolean;
var
  texto: string;
  data: TDateTime;
begin
  Result := False;
  aISO := '';
  texto := Trim(aData);
  // TryEncodeDate valida de verdade (31/02 não é data), então o que não passa
  // por ele não vira conversão nenhuma - o valor gravado fica como estava.
  if (Length(texto) = 10) and (texto[3] = '/') and (texto[6] = '/') and
     TryEncodeDate(StrToIntDef(Copy(texto, 7, 4), 0),
       StrToIntDef(Copy(texto, 4, 2), 0), StrToIntDef(Copy(texto, 1, 2), 0),
       data) then
    aISO := FormatDateTime('yyyy-mm-dd', data)
  else if (Length(texto) = 10) and (texto[5] = '-') and (texto[8] = '-') and
     TryEncodeDate(StrToIntDef(Copy(texto, 1, 4), 0),
       StrToIntDef(Copy(texto, 6, 2), 0), StrToIntDef(Copy(texto, 9, 2), 0),
       data) then
    aISO := FormatDateTime('yyyy-mm-dd', data)
  else
    Exit;
  Result := True;
end;

end.
