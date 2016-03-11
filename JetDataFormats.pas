unit JetDataFormats;

interface
uses SysUtils;

function JetFormatSettings: TFormatSettings;

function EncodeBin(data: array of byte): WideString;
function EncodeOleBin(data: OleVariant): WideString;

function EncodeComment(str: WideString): WideString;
function DecodeComment(str: WideString): WideString;

function EncodeStr(val: WideString): WideString;

function JetEncodeTypedValue(Value: OleVariant; DataType: integer): Widestring;
function JetEncodeValue(Value: OleVariant): WideString;

implementation
uses Variants, OleDB, JetCommon;

var
  FJetFormatSettings: TFormatSettings;
  FJetFormatSettingsInitialized: boolean = false;

function JetFormatSettings: TFormatSettings;
begin
  if not FJetFormatSettingsInitialized then begin
    FJetFormatSettings := TFormatSettings.Create();
    with FJetFormatSettings do begin
      DecimalSeparator := '.';
      DateSeparator := '-';
      TimeSeparator := ':';
      ShortDateFormat := 'mm-dd-yyyy hh:nn:ss';
      LongDateFormat := 'mm-dd-yyyy hh:nn:ss';
    end;
    FJetFormatSettingsInitialized := true;
  end;
  Result := FJetFormatSettings;
end;

//Encodes binary data for inserting to SQL text.
//Presently unused (DEFAULT values are already escaped in the DB)
function EncodeBin(data: array of byte): WideString;
const HexChars:WideString='0123456789ABCDEF';
var i: integer;
begin
  if Length(data)<=0 then begin
    Result := 'NULL';
    exit;
  end;

  Result := '0x';
  SetLength(Result, 2+Length(data)*2);
  for i := 0 to Length(data) - 1 do begin
    Result[2+i*2+0] := HexChars[1+(data[i] shr 4)];
    Result[2+i*2+1] := HexChars[1+(data[i] and $0F)];
  end;
end;

function EncodeOleBin(data: OleVariant): WideString;
var bin_data: array of byte;
begin
  if VarIsNil(data) then
    Result := 'NULL'
  else begin
    bin_data := data;
    Result := EncodeBin(bin_data);
  end;
end;

//Encodes /**/ comment for inserting into SQL text. Escapes closure symbols.
//Replacements:
//  \ == \\
//  / == \/
//It's handy that you don't need a special parsing when looking for comment's end:
//dangerous combinations just get broken during the encoding.
function EncodeComment(str: WideString): WideString;
var pc: PWideChar;
  i: integer;
  c: WideChar;
begin
 //We'll never need more than twice the size
  SetLength(Result, 2*Length(str));
  pc := @Result[1];

  for i := 1 to Length(str) do begin
    c := str[i];
    if (c='\') or (c='/') then begin
      pc^ := '\';
      Inc(pc);
      pc^ := c;
    end else
    begin
      pc^ := c;
    end;
    Inc(pc);
  end;

 //Actual length
  SetLength(Result, (integer(pc)-integer(@Result[1])) div 2);
end;

//Decodes comment.
function DecodeComment(str: WideString): WideString;
var pc: PWideChar;
  i: integer;
  c: WideChar;
  SpecSymbol: boolean;
begin
 //We'll never need more than the source size
  SetLength(Result, Length(str));
  pc := @Result[1];

  SpecSymbol := false;
  for i := 1 to Length(str) do begin
    c := str[i];
    if (not SpecSymbol) and (c='\') then begin
      SpecSymbol := true;
      continue;
    end;
    SpecSymbol := false;
    pc^ := c;
    Inc(pc);
  end;

 //Actual length
  SetLength(Result, (integer(pc)-integer(@Result[1])) div 2);
end;

//Encodes a string for inserting into SQL text, replaces specsymbols.
//Presently unused (DEFAULT values are already escaped in the DB)
function EncodeStr(val: WideString): WideString;
var pc: PWideChar;
  i: integer;
  c: WideChar;
begin
 //We'll never need more than twice the size
  SetLength(Result, 2*Length(val));
  pc := @Result[1];

  for i := 1 to Length(val) do begin
    c := val[i];
    if c=#00 then begin  //nul
      pc^ := '\';
      Inc(pc);
      pc^ := '0';
    end else
    if c=#08 then begin  //backspace
      pc^ := '\';
      Inc(pc);
      pc^ := 'b';
    end else
    if c=#09 then begin  //tab
      pc^ := '\';
      Inc(pc);
      pc^ := 't';
    end else
    if (c='''') or (c='"') or (c='\') then begin
      pc^ := '\';
      Inc(pc);
      pc^ := c;
    end else
    begin
      pc^ := c;
    end;
    Inc(pc);
  end;

 //Actual length
  SetLength(Result, (integer(pc)-integer(@Result[1])) div 2);
end;

//Because Delphi does not allow int64(Value).
function uint_cast(Value: OleVariant): int64;
begin
  Result := value;
end;

//Formats a field value according to it's type
//Presently unused (DEFAULT values are already escaped in the DB)
function JetEncodeTypedValue(Value: OleVariant; DataType: integer): Widestring;
begin
  if VarIsNil(Value) then
    Result := 'NULL'
  else
  case DataType of
    DBTYPE_I1, DBTYPE_I2, DBTYPE_I4,
    DBTYPE_UI1, DBTYPE_UI2, DBTYPE_UI4:
      Result := IntToStr(integer(Value));
    DBTYPE_I8, DBTYPE_UI8:
      Result := IntToStr(uint_cast(Value));
    DBTYPE_R4, DBTYPE_R8:
      Result := FloatToStr(double(Value), JetFormatSettings);
    DBTYPE_NUMERIC, DBTYPE_DECIMAL, DBTYPE_CY:
      Result := FloatToStr(currency(Value), JetFormatSettings);
    DBTYPE_GUID:
     //Or else it's not GUID
      Result := GuidToString(StringToGuid(Value));
    DBTYPE_DATE:
      Result := '#'+DatetimeToStr(Value, JetFormatSettings)+'#';
    DBTYPE_BOOL:
      Result := BoolToStr(Value, {UseBoolStrs=}true);
    DBTYPE_BYTES:
      Result := EncodeOleBin(Value);
    DBTYPE_WSTR:
      Result := ''''+EncodeStr(Value)+'''';
  else
    Result := ''''+EncodeStr(Value)+''''; //best guess
  end;
end;

//Encodes a value according to it's variant type
//Presently unused (DEFAULT values are already escaped in the DB)
function JetEncodeValue(Value: OleVariant): WideString;
begin
  if VarIsNil(Value) then
    Result := 'NULL'
  else
  case VarType(Value) of
    varSmallInt, varInteger, varShortInt,
    varByte, varWord, varLongWord:
      Result := IntToStr(integer(Value));
    varInt64:
      Result := IntToStr(uint_cast(Value));
    varSingle, varDouble:
      Result := FloatToStr(double(Value), JetFormatSettings);
    varCurrency:
      Result := FloatToStr(currency(Value), JetFormatSettings);
    varDate:
      Result := '#'+DatetimeToStr(Value, JetFormatSettings)+'#';
    varOleStr, varString:
      Result := ''''+EncodeStr(Value)+'''';
    varBoolean:
      Result := BoolToStr(Value, {UseBoolStrs=}true);
    varArray:
      Result := EncodeOleBin(Value);
  else
    Result := ''''+EncodeStr(Value)+''''; //best guess
  end;
end;

end.
