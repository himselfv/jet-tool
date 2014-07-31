unit StreamUtils;
{$WEAKPACKAGEUNIT ON}
{
Stream helpers. Cached readers/writers.
Some comments are in Russian, deal with it *puts on glasses*.
(c) himselfv, me@boku.ru.
}

interface
uses SysUtils, Classes{$IFDEF DCC}, UniStrUtils, Windows{$ENDIF};

type
  EStreamException = class(Exception);

const
  sCannotReadData = 'Cannot read data from the stream';
  sCannotWriteData = 'Cannot write data to the stream';

type
 {
   Extends TStream with helper methods to read various data format.
   They're not very efficient by themselves so better use this on fast streams
   (TMemoryStream, memory-mapped file, cached stream)
 }
  TStreamHelper = class helper for TStream
  public
    function ReadInt: integer;
    function ReadInt64: int64;
    function ReadAnsiChar(out c: AnsiChar): boolean;
    function ReadWideChar(out c: WideChar): boolean;
    function ReadAnsiLine: AnsiString;
    function ReadUniLine: UnicodeString; //assumes Windows UTF-16
    procedure WriteInt(const Value: integer);
    procedure WriteIn64(const Value: int64);
    function WriteAnsiChar(const c: AnsiChar): boolean;
    function WriteWideChar(const c: WideChar): boolean;
    procedure WriteAnsiString(const s: AnsiString);
    procedure WriteUniString(const s: UnicodeString);
  end;

{
 A class to speed up reading from low latency streams (usually file streams).

 Every time you request something from TFileStream, it reads the data from a
 file. Even with drive cache enabled, it's a kernel-mode operation which is slow.

 This class tries to overcome the problem by reading data in large chunks.
 Whenever you request your byte or two, it reads the whole chunk and stores
 it in local memory. Next time you request another two bytes you'll get
 them right away, because they're already here.

 Please remember that you CAN'T use StreamReader on a Stream "for a while".
 The moment you read your first byte through StreamReader the underlying
 stream is NOT YOURS ANYMORE. You shouldn't make any reads to the Stream
 except than through StreamReader.
}

const
  DEFAULT_CHUNK_SIZE = 4096;

type
  TStreamReader = class(TStream)
  protected
    FStream: TStream;
    FOwnStream: boolean;
    FBytesRead: int64;
  public
    property Stream: TStream read FStream;
    property OwnStream: boolean read FOwnStream;
    property BytesRead: int64 read FBytesRead;

  protected
    flag_reallocbuf: boolean;
    FNextChunkSize: integer; //size of a chunk we'll try to read next time
    FChunkSize: integer; //size of current chunk
    buf: pbyte;   //cached data
    ptr: pbyte;   //current location in buf
    adv: integer; //number of bytes read from buf
    rem: integer; //number of bytes remaining in buf.
     //adv+rem not always equals to FChunkSize since we could have read only
     //partial chunk or FChunkSize could have been changed after that.
    function NewChunk: integer;
    function UpdateChunk: integer;
    procedure SetChunkSize(AValue: integer);
    procedure ResetBuf;
    function ReallocBuf(ASize: integer): boolean;
    procedure FreeBuf;
  public
    property ChunkSize: integer read FChunkSize write SetChunkSize;
    property ChunkBytesRemaining: integer read rem;

  public
    constructor Create(AStream: TStream; AOwnStream: boolean = false);
    destructor Destroy; override;
    procedure JoinStream(AStream: TStream; AOwnsStream: boolean = false);
    procedure ReleaseStream;

  protected
    function GetSize: Int64; override;
    function GetInternalPosition: integer;
    function LocalSeek(Offset: Int64; Origin: TSeekOrigin): boolean;
  public
    function Seek(const Offset: Int64; Origin: TSeekOrigin): Int64; override;
    function Read(var Buffer; Count: Longint): Longint; override;
    function ReadBuf(ABuf: pbyte; ALen: integer): integer; inline;
    function Write(const Buffer; Count: Longint): Longint; override;
    function Peek(var Buffer; ASize: integer): integer;
    function PeekByte(out b: byte): boolean;
  end;

  TStreamWriter = class(TStream)
  protected
    FStream: TStream;
    FOwnStream: boolean;
    FBytesWritten: int64;
  public
    property Stream: TStream read FStream;
    property OwnStream: boolean read FOwnStream;
    property BytesWritten: int64 read FBytesWritten;

  protected
    FChunkSize: integer;
    buf: pbyte;
    ptr: pbyte;
    used: integer;
    procedure ResetBuf;
    procedure FreeBuf;
    procedure SetChunkSize(AValue: integer);
  public
    property ChunkSize: integer read FChunkSize write SetChunkSize;

  public
    constructor Create(AStream: TStream; AOwnStream: boolean = false);
    destructor Destroy; override;
    procedure JoinStream(AStream: TStream; AOwnsStream: boolean = false);
    procedure ReleaseStream;

  protected
    function GetSize: Int64; override;
    procedure SetSize(NewSize: Longint); overload; override;
    procedure SetSize(const NewSize: Int64); overload; override;
    function GetInternalPosition: integer;

  public
    function Seek(const Offset: Int64; Origin: TSeekOrigin): Int64; override;
    function Read(var Buffer; Count: Longint): Longint; override;
    function Write(const Buffer; Count: Longint): Longint; override;
    function WriteBuf(ABuf: PByte; ALen: integer): integer; inline;
    procedure Flush;

  end;

 { TStringStream
  Stores or retrieves raw data from a string.
  If you pass a string, it will be read/updated, else internal buffer is used. }

 { Don't instantiate }
  TCustomStringStream = class(TStream)
  protected
    FPtr: PByte;
    FPos: integer; //in bytes
    function RemainingSize: integer;
  public
    function Read(var Buffer; Count: Longint): Longint; override;
    function Write(const Buffer; Count: Longint): Longint; override;
    function Seek(const Offset: Int64; Origin: TSeekOrigin): Int64; override;
  end;

 { Single-byte strings }
  TAnsiStringStream = class(TCustomStringStream)
  protected
    FOwnBuffer: AnsiString;
    FString: PAnsiString;
    function GetString: AnsiString;
    function GetSize: Int64; override;
    procedure SetSize(NewSize: Longint); override;
  public
    constructor Create(AString: PAnsiString = nil);
    property Data: AnsiString read GetString;
  end;

 { Double-byte strings }
  TUnicodeStringStream = class(TCustomStringStream)
  protected
    FOwnBuffer: UnicodeString;
    FString: PUnicodeString;
    function GetString: UnicodeString;
    function GetSize: Int64; override;
    procedure SetSize(NewSize: Longint); override;
  public
    constructor Create(AString: PUnicodeString = nil);
    property Data: UnicodeString read GetString;
  end;

  TStringStream = TUnicodeStringStream;

{
  Cached read/write char-by-char.
}

const
  BOM_UTF16LE = WideChar($FEFF);
  BOM_UTF16BE = WideChar($FFFE);
  BOM_UTF8 = $BFBBEF; //full version it's EF BB BF

type
  TCharSet = (
    csAnsi,
    csUtf16Le,
    csUtf16Be,
    csUtf8      //requires at least 4 bytes of buffer!
  );

type
  TCharReader = class(TStreamReader)
  protected
    Charset: TCharSet;
    function DetectCharset: TCharset;

  public
   //Create with autodetected encoding
    constructor Create(AStream: TStream; AOwnStream: boolean = false); overload;
   //Create with specified encoding (if you don't want auto, set to your default one)
    constructor Create(AStream: TStream; ACharSet: TCharSet;
      AOwnStream: boolean = false); overload;
    function ReadChar(out c: WideChar): boolean;
    function PeekChar(out c: WideChar): boolean;
    function ReadLine(out s: UnicodeString): boolean;
  end;

  TCharWriter = class(TStreamWriter)
  protected
    Charset: TCharset;

  public
   //Create with default encoding (UTF16LE)
    constructor Create(AStream: TStream; AOwnStream: boolean = false); overload;
   //Create with specified encoding
    constructor Create(AStream: TStream; ACharSet: TCharSet; AOwnStream: boolean = false); overload;
    destructor Destroy; override;
    procedure WriteBom;
    procedure WriteChar(const c: WideChar);
    procedure WriteChars(c: PWideChar; len: integer);
    procedure WriteString(const c: UnicodeString);
    procedure WriteLine(const s: UnicodeString);
  end;


type
  TStreamReader2 = class(TStreamReader)
  public
    destructor Destroy; override;

  protected
   //These are used for ReadStr. Instead of reallocating temporary pchar
   //every time, we just keep this buffer up while alive.
    tmpstr_ansi_s: PAnsiChar;
    tmpstr_ansi_sz: cardinal;
    tmpstr_wide_s: PWideChar;
    tmpstr_wide_sz: cardinal;
    procedure FreeTmpStr;
  public
    function TryReadLineA(var s: PAnsiChar; var sz: cardinal): boolean;
    function TryReadLineW(var s: PWideChar; var sz: cardinal): boolean;
    procedure ReadLineA(var s: PAnsiChar; var sz: cardinal);
    procedure ReadLineW(var s: PWideChar; var sz: cardinal);
    function TryReadStrA(out s: AnsiString): boolean;
    function TryReadStrW(out s: WideString): boolean;
    function ReadStrA: AnsiString;
    function ReadStrW: WideString;

  public
    function ReadAnsiChar(out c: AnsiChar): boolean;
    function ReadWideChar(out c: WideChar): boolean;
    function ReadUtf8Char(out c: WideChar): boolean;
  end;


implementation

{ UTF8 Utils }
{ Parts based on JWBConvert by Fillip Karbt taken from Wakan source code. }

const
  UTF8_VALUE1 = $00;        // Value for set bits for single byte UTF-8 Code.
  UTF8_MASK1 = $80;        // Mask (i.e. bits not set by the standard) 0xxxxxxx
  UTF8_WRITE1 = $ff80;      // Mask of bits we cannot allow if we are going to write one byte code
  UTF8_VALUE2 = $c0;        // Two byte codes
  UTF8_MASK2 = $e0;        // 110xxxxx 10yyyyyy
  UTF8_WRITE2 = $f800;      // Mask of bits we cannot allow if we are going to write two byte code
  UTF8_VALUE3 = $e0;        // Three byte codes
  UTF8_MASK3 = $f0;        // 1110xxxx 10yyyyyy 10zzzzzz
  UTF8_VALUE4 = $f0;        // Four byte values
  UTF8_MASK4 = $f8;        // 11110xxx ----    (These values are not supported by JWPce).
  UTF8_VALUEC = $80;        // Continueation byte (10xxxxxx).
  UTF8_MASKC = $c0;

{ In both of the following functions UTF8 character sequences are stored as integer.
 Remember, lowest byte first.
 You may want to cast this for easy access: }
type
  TUTF8Bytes = array[0..3] of byte;

{ Converts UCS2-LE character into UTF8 sequence of up to 4 bytes.
 Returns the number of bytes used }
function UCS2toUTF8(const ch: WideChar; out uch: integer): integer; inline;
begin
  if (word(ch) and UTF8_WRITE1)=0 then
  begin
    Result:=1;
    uch:=byte(ch);
  end else
  if (word(ch) and UTF8_WRITE2)=0 then
  begin
    Result := 2;
    uch := (UTF8_VALUE2 or byte(word(ch) shr 6))
         + (UTF8_VALUEC or byte(word(ch) and $3f)) shl 8;
  end else
  begin
    Result := 3;
    uch := (UTF8_VALUE3 or byte(word(ch) shr 12))
         + (UTF8_VALUEC or byte((word(ch) shr 6) and $3f)) shl 8
         + (UTF8_VALUEC or byte(word(ch) and $3f)) shl 16;
  end;
end;

{ Converts a sequence of up to 4 bytes in UTF8, all of which must be present,
 to one UCS2-LE character, or puts DefaultChar there if the conversion is impossible.
 Returns the number of bytes this UTF8 symbol occupied. }
function UTF8toUCS2(const uch: integer; const DefaultChar: WideChar; out ch: WideChar): integer; inline;
var b: TUTF8Bytes absolute uch;
begin
  if (b[0] and UTF8_MASK1)=UTF8_VALUE1 then
  begin
    ch := WideChar(b[0]);
    Result := 1;
  end else
  if (b[0] and UTF8_MASK2)=UTF8_VALUE2 then
  begin
    ch := WideChar(((b[0] and $1f) shl 6) or (b[1] and $3f));
    Result := 2;
  end else
  if (b[0] and UTF8_MASK3)=UTF8_VALUE3 then
  begin
    ch := WideChar(((b[0] and $0f) shl 12) or ((b[1] and $3f) shl 6) or (b[2] and $3f));
    Result := 3;
  end else
  if (b[0] and UTF8_MASK4)=UTF8_VALUE4 then
  begin
    ch := DefaultChar;
    Result := 4;
  end else
  begin
   //Invalid character
    ch := DefaultChar;
    Result := 1;//because we don't know what else to do
  end;
end;


{ TStreamHelper }

function TStreamHelper.ReadInt: integer;
begin
  if Read(Result, SizeOf(Result))<>SizeOf(Result) then
    raise EStreamException.Create(sCannotReadData);
end;

function TStreamHelper.ReadInt64: int64;
begin
  if Read(Result, SizeOf(Result))<>SizeOf(Result) then
    raise EStreamException.Create(sCannotReadData);
end;

procedure TStreamHelper.WriteInt(const Value: integer);
begin
  if Write(Value, SizeOf(Value))<>SizeOf(Value) then
    raise EStreamException.Create(sCannotWriteData);
end;

procedure TStreamHelper.WriteIn64(const Value: int64);
begin
  if Write(Value, SizeOf(Value))<>SizeOf(Value) then
    raise EStreamException.Create(sCannotWriteData);
end;

function TStreamHelper.ReadAnsiChar(out c: AnsiChar): boolean;
begin
  Result := Self.Read(c, SizeOf(c)) = SizeOf(c);
end;

function TStreamHelper.ReadWideChar(out c: WideChar): boolean;
begin
  Result := Self.Read(c, SizeOf(c)) = SizeOf(c);
end;

function TStreamHelper.ReadAnsiLine: AnsiString;
var c: AnsiChar;
begin
  Result := '';
  while ReadAnsiChar(c) and (c <> #13) and (c <> #10) do
    Result := Result + c;
end;

function TStreamHelper.ReadUniLine: UnicodeString;
var c: WideChar;
begin
  Result := '';
  while ReadWideChar(c) and (c <> #13) and (c <> #10) do
    Result := Result + c;
end;

function TStreamHelper.WriteAnsiChar(const c: AnsiChar): boolean;
begin
  Result := Self.Write(c, SizeOf(c))=SizeOf(c);
end;

function TStreamHelper.WriteWideChar(const c: WideChar): boolean;
begin
  Result := Self.Write(c, SizeOf(c))=SizeOf(c);
end;

{
 Throws exception if it could not write the whole string.
 There's no function which returns false or the number of characters written,
 since we can't tell how many characters there were - we write bytes.
 And returning the number of bytes isn't much helpful to the client (what if
 we wrote half of a 4-byte char?) and can lead to misunderstandings and usage
 errors.
 If you want to write byte-by-byte, just use standard Write(var Buffer) instead.
}
procedure TStreamHelper.WriteAnsiString(const s: AnsiString);
begin
  if Write(s[1], Length(s)*SizeOf(AnsiChar)) <> Length(s)*SizeOf(AnsiChar) then
    raise Exception.Create('Cannot write a line to '+self.ClassName);
end;

procedure TStreamHelper.WriteUniString(const s: UnicodeString);
begin
  if Write(s[1], Length(s)*SizeOf(WideChar)) <> Length(s)*SizeOf(WideChar) then
    raise Exception.Create('Cannot write a line to '+self.ClassName);
end;


{ TStreamReader }

constructor TStreamReader.Create(AStream: TStream; AOwnStream: boolean = false);
begin
  inherited Create;
  FStream := AStream;
  FOwnStream := AOwnStream;

  buf := nil;
  ptr := nil;
  adv := 0;
  rem := 0;
  FChunkSize := 0;
  SetChunkSize(DEFAULT_CHUNK_SIZE);
end;

destructor TStreamReader.Destroy;
begin
  FreeBuf;
  if OwnStream then
    FreeAndNil(FStream)
  else
    FStream := nil;
  inherited;
end;

procedure TStreamReader.JoinStream(AStream: TStream; AOwnsStream: boolean = false);
begin
  ReleaseStream;

  FStream := AStream;
  FOwnStream := AOwnsStream;
  ResetBuf;
end;

//Releases the stream and synchronizes it's position with the expected one
procedure TStreamReader.ReleaseStream;
begin
  if FStream=nil then exit;

 //Scroll back (Seek support required!)
  FStream.Seek(int64(-adv-rem), soCurrent);

 //Release the stream
  FStream := nil;
  FOwnStream := false;
  ResetBuf;
end;

procedure TStreamReader.SetChunkSize(AValue: integer);
begin
 //If we can't realloc right now, set delayed reallocation
  if ReallocBuf(AValue) then
    FChunkSize := AValue
  else begin
    flag_reallocbuf := true;
    FNextChunkSize := AValue;
  end;
end;

{ Downloads a complete new chunk of data. Returns the number of bytes read. }
function TStreamReader.NewChunk: integer;
begin
 //Delayed reallocation
  if flag_reallocbuf and ReallocBuf(FNextChunkSize) then begin
    FChunkSize := FNextChunkSize;
    flag_reallocbuf := false;
  end;

  rem := FStream.Read(buf^, FChunkSize);
  adv := 0;
  ptr := buf;
  Result := rem;
end;

{ Moves remaining data to the beginning of the cache and downloads more.
 Returns the number of bytes read. }
function TStreamReader.UpdateChunk: integer;
var DataPtr: Pbyte;
begin
 //Full cache download
  if rem <= 0 then begin
    Result := NewChunk;
    exit;
  end;

 //Partial download
  Move(ptr^, buf^, rem);
  ptr := buf;
  adv := 0;

  if flag_reallocbuf and ReallocBuf(FNextChunkSize) then begin
    FChunkSize := FNextChunkSize;
    flag_reallocbuf := false;
  end;

  DataPtr := PByte(cardinal(buf) + cardinal(adv+rem));
  Result := Stream.Read(DataPtr^, FChunkSize-adv-rem);
  rem := rem + Result;
end;

//Clears the contents of the cache
procedure TStreamReader.ResetBuf;
begin
  adv := 0;
  rem := 0;
  ptr := buf;
end;

function TStreamReader.ReallocBuf(ASize: integer): boolean;
begin
 //We can't decrease buffer size cause there's still data inside.
  if adv + rem > ASize then begin
    Result := false;
    exit;
  end;

  ReallocMem(buf, ASize);
  ptr := pointer(IntPtr(buf) + adv);
  Result := true;
end;

procedure TStreamReader.FreeBuf;
begin
  if Assigned(buf) then
    FreeMem(buf);
end;

function TStreamReader.GetSize: Int64;
begin
  Result := FStream.Size;
end;

function TStreamReader.GetInternalPosition: integer;
begin
  Result := FStream.Position - rem;
end;

//If possible, try to Seek inside the buffer
function TStreamReader.LocalSeek(Offset: Int64; Origin: TSeekOrigin): boolean;
var pos,sz: Int64;
begin
  if Origin=soEnd then begin
    sz := FStream.Size;
    //Convert to from beginning
    Offset := sz - Offset;
    Origin := soBeginning;
  end;

  if Origin=soBeginning then begin
    pos := FStream.Position;
    if (Offset>=pos) or (Offset<pos-adv-rem) then begin
      Result := false; //not in this chunk
      exit;
    end;
   //Convert to relative
    Offset := rem-(pos-Offset);
    Origin := soCurrent;
  end;

  if Origin=soCurrent then begin
    if (Offset<-adv) or (Offset>=rem) then begin
      Result := false;
      exit;
    end;

    adv := adv+Offset;
    rem := rem-Offset;
    Inc(ptr, Offset);
    Result := true;
  end else
    Result := false;
end;

function TStreamReader.Seek(const Offset: Int64; Origin: TSeekOrigin): Int64;
begin
 //TStream calls Seek(0,soCurrent) to determine Position, so be fast in this case
  if (Origin=soCurrent) and (Offset=0) then
    Result := GetInternalPosition()
  else
  if LocalSeek(Offset, Origin) then
    Result := GetInternalPosition()
  else begin
    Result := FStream.Seek(Offset, Origin);
    ResetBuf;
  end;
end;

function TStreamReader.Read(var Buffer; Count: Longint): Longint;
var pbuf: PByte;
begin
  if Count=0 then begin
    Result := 0;
    exit;
  end;

 //The most common case
  if Count <= rem then begin
    Move(ptr^, Buffer, Count);
    Inc(ptr, Count);
    Inc(adv, Count);
    Dec(rem, Count);

    Inc(FBytesRead, Count);
    Result := Count;
    exit;
  end;

 //Read first part
  pbuf := @Buffer;
  if rem > 0 then begin
    Move(ptr^, pbuf^, rem);
    Inc(ptr, rem);
    Inc(adv, rem);

   //Update variables
    Inc(pbuf, rem);
    Dec(Count, rem);

    Result := rem;
    rem := 0;
  end else
    Result := 0;

 //Download the remaining part

 //If it's smaller than a chunk, read the whole chunk
  if Count < FChunkSize then begin
    NewChunk;

    if rem < Count then //rem was already updated in NewChunk
      Count := rem;

    Move(ptr^, pbuf^, Count);
    Inc(ptr, Count);
    Inc(adv, Count);
    Dec(rem, Count);

    Inc(Result, Count);
  end else
   //Else just read it from stream
    Result := Result + FStream.Read(pbuf^, Count);

  Inc(FBytesRead, Result);
end;

function TStreamReader.ReadBuf(ABuf: pbyte; ALen: integer): integer;
begin
  Result := Read(ABuf^, ALen);
end;

function TStreamReader.Write(const Buffer; Count: Longint): Longint;
begin
  raise Exception.Create('StreamReader cannot write.');
end;

//Won't Peek for more than CacheSize
function TStreamReader.Peek(var Buffer; ASize: integer): integer;
begin
 //If the data is in cache, it's simple
  if ASize <= rem then begin
    Move(ptr^, Buffer, ASize);
    Result := ASize;
    exit;
  end;

 //Else return the best possible amount. Make the cache completely fresh
  if rem <= FChunkSize then
    UpdateChunk;

 //If the complete data fit, return it
  if ASize <= rem then begin
    Move(ptr^, Buffer, ASize);
    Result := ASize;
    exit;
  end;

 //Didn't fit => return all that's available
  Move(ptr^, Buffer, rem);
  Result := rem;
end;

//Уж один-то байт мы всегда можем подсмотреть, если в файле осталось.
function TStreamReader.PeekByte(out b: byte): boolean;
begin
 //Если буфер непуст, берём первый байт
  if rem >= 0 then begin
    b := ptr^;
    Result := true;
    exit;
  end;

 //Иначе буфер пуст. Качаем следующий кусочек.
  NewChunk;

  if rem >= 0 then begin
    b := ptr^;
    Result := true;
  end else
    Result := false;
end;


{ TStreamWriter }

constructor TStreamWriter.Create(AStream: TStream; AOwnStream: boolean = false);
begin
  inherited Create;
  FStream := AStream;
  FOwnStream := AOwnStream;

  FChunkSize := 0;
  buf := nil;
  ptr := nil;
  used := 0;

  SetChunkSize(DEFAULT_CHUNK_SIZE);
end;

destructor TStreamWriter.Destroy;
begin
  Flush();
  FreeBuf;

  if OwnStream then
    FreeAndNil(FStream)
  else
    FStream := nil;
  inherited;
end;

procedure TStreamWriter.JoinStream(AStream: TStream; AOwnsStream: boolean = false);
begin
 //Освобождаем старый поток
  ReleaseStream;

 //Цепляемся к потоку
  FStream := AStream;
  FOwnStream := AOwnsStream;

 //Сбрасываем буфер на всякий случай
  ResetBuf;
end;

//Освобождает поток, синхронизируя его положение с ожидаемым
procedure TStreamWriter.ReleaseStream;
begin
  if FStream=nil then exit;

 //Сбрасываем
  Flush;

 //Отпускаем поток
  FStream := nil;
  FOwnStream := false;
end;

//Сбрасывает на диск содержимое буфера и обнуляет буфер.
procedure TStreamWriter.Flush;
begin
  if used <= 0 then exit;
  FStream.Write(buf^, used);
  ptr := buf;
  used := 0;
end;

procedure TStreamWriter.ResetBuf;
begin
  ptr := buf;
  used := 0;
end;

procedure TStreamWriter.FreeBuf;
begin
  if Assigned(buf) then
    FreeMem(buf);
  buf := nil;
  ptr := nil;
  used := 0;
end;

procedure TStreamWriter.SetChunkSize(AValue: integer);
begin
 //Если в новый буфер текущие данные не влезают, сбрасываем их в поток
  if AValue < used then
    Flush;

  FChunkSize := AValue;
  ReallocMem(buf, FChunkSize);

 //Обновляем указатель на текущий байт
  ptr := pointer(IntPtr(buf) + used);
end;

//Use this instead of underlying Stream's Position
function TStreamWriter.GetInternalPosition: integer;
begin
  Result := FStream.Position + used;
end;

function TStreamWriter.GetSize: Int64;
begin
  Result := FStream.Size + used;
end;

procedure TStreamWriter.SetSize(NewSize: Longint);
begin
  SetSize(int64(NewSize));
end;

procedure TStreamWriter.SetSize(const NewSize: Int64);
begin
  Flush();
  FStream.Size := NewSize;
end;

function TStreamWriter.Seek(const Offset: Int64; Origin: TSeekOrigin): Int64;
begin
  if (Origin=soCurrent) and (Offset=0) then
    Result := GetInternalPosition() //TStream uses this to determine Position
  else begin
    Flush;
    Result := FStream.Seek(Offset, Origin);
  end;
end;

function TStreamWriter.Read(var Buffer; Count: Longint): Longint;
begin
  raise Exception.Create('StreamWriter cannot read.');
end;

function TStreamWriter.Write(const Buffer; Count: Longint): Longint;
var rem: integer;
  pbuf: PByte;
begin
  if Count<=0 then begin
    Result := 0;
    exit;
  end;

 //Если влезает в кэш, кладём туда
  rem := FChunkSize - used;
  if Count <= rem then begin
    Move(Buffer, ptr^, Count);
    Inc(used, Count);
    Inc(ptr, Count);
    Result := Count;
    Inc(FBytesWritten, Count);
    exit;
  end;

 //Иначе нас просят записать нечто большее. Вначале добиваем текущий буфер
  pbuf := @Buffer;
  if used > 0 then begin
    if rem > 0 then begin
      Move(pbuf^, ptr^, rem);
      Inc(ptr, rem);
      Inc(used, rem);

     //Update variables
      Inc(pbuf, rem);
      Dec(Count, rem);

      Result := rem;
    end else
      Result := 0;

    Flush;
  end else
    Result := 0;

 //Если остаток меньше буфера, сохраняем его в буфер
  if Count < FChunkSize then begin
    Move(pbuf^, ptr^, Count);
    Inc(ptr, Count);
    Inc(used, Count);
    Inc(Result, Count);
  end else
  //Иначе пишем его напрямую
   Result := Result + FStream.Write(pbuf^, Count);

  Inc(FBytesWritten, Result);
end;

function TStreamWriter.WriteBuf(ABuf: PByte; ALen: integer): integer;
begin
  Result := Write(ABuf^, ALen);
end;


{
TCustomStringStream
}

function TCustomStringStream.RemainingSize: integer;
begin
  Result := Self.Size-FPos;
end;

function TCustomStringStream.Read(var Buffer; Count: Longint): Longint;
begin
  if Count>RemainingSize then
    Count := RemainingSize;
  Move(PByte(IntPtr(FPtr)+FPos)^,Buffer,Count);
  Inc(FPos,Count);
  Result := Count;
end;

function TCustomStringStream.Write(const Buffer; Count: Longint): Longint;
begin
  if RemainingSize<Count then
    SetSize(FPos+Count);
  Move(Buffer,PByte(IntPtr(FPtr)+FPos)^,Count);
  Inc(FPos,Count);
  Result := Count;
end;

function TCustomStringStream.Seek(const Offset: Int64; Origin: TSeekOrigin): Int64;
begin
  if Origin=soCurrent then begin
    if Offset=0 then begin
      Result := FPos;
      exit;
    end;
    Result := FPos+Offset;
  end else
  if Origin=soEnd then begin
    Result := GetSize-Offset;
  end else
    Result := Offset;

  if Result<0 then
    Result := 0
  else
  if Result>GetSize then
    Result := GetSize;
  FPos := Result;
end;

{
TAnsiStringStream
}

constructor TAnsiStringStream.Create(AString: PAnsiString = nil);
begin
  inherited Create();
  if AString=nil then
    FString := @FOwnBuffer
  else
    FString := AString;
  FPos := 0;
  FPtr := pointer(FString^);
end;

function TAnsiStringStream.GetString: AnsiString;
begin
  Result := FString^;
end;

function TAnsiStringStream.GetSize: Int64;
begin
  Result := Length(FString^)*SizeOf(AnsiChar);
end;

procedure TAnsiStringStream.SetSize(NewSize: Longint);
begin
  SetLength(FString^, NewSize);
  FPtr := pointer(FString^);
end;

{
TUnicodeStringStream
}

constructor TUnicodeStringStream.Create(AString: PUnicodeString = nil);
begin
  inherited Create();
  if AString=nil then
    FString := @FOwnBuffer
  else
    FString := AString;
  FPos := 0;
  FPtr := pointer(FString^);
end;

function TUnicodeStringStream.GetString: UnicodeString;
begin
  Result := FString^;
end;

function TUnicodeStringStream.GetSize: Int64;
begin
  Result := Length(FString^)*SizeOf(WideChar);
end;

procedure TUnicodeStringStream.SetSize(NewSize: Longint);
begin
  if NewSize mod 2 = 0 then
    SetLength(FString^, NewSize div SizeOf(WideChar))
  else
    SetLength(FString^, NewSize div SizeOf(WideChar) + 1); //no choice but to allocate one more symbol
  FPtr := pointer(FString^);
end;


{ TCharReader }

function SwapChar(c: WideChar): WideChar; inline;
begin
  Result := WideChar(byte(c) shl 8 + word(c) shr 8);
end;

procedure SwapCharsIn(s: PWideChar; len: integer); inline;
var i: integer;
begin
  for i := 0 to len - 1 do begin
    s^ := SwapChar(s^);
    Inc(s);
  end;
end;

function SwapChars(s: PWideChar; len: integer): WideString;
var i: integer;
begin
  SetLength(Result, len);
  for i := 0 to len - 1 do begin
    Result[i+1] := SwapChar(s^);
    Inc(s);
  end;
end;

constructor TCharReader.Create(AStream: TStream; AOwnStream: boolean = false);
begin
  inherited Create(AStream, AOwnStream);
  Charset := DetectCharset;
end;

constructor TCharReader.Create(AStream: TStream; ACharSet: TCharSet;
  AOwnStream: boolean = false);
begin
  inherited Create(AStream, AOwnStream);
  Charset := ACharset;
end;

//Детектируем бом и сразу вычитываем.
function TCharReader.DetectCharset: TCharset;
var Bom: WideChar;
  Bom3: integer;
begin
 //No data => ANSI
  if Peek(Bom, 2) < 2 then
    Result := csAnsi
  else

  if Bom = BOM_UTF16LE then begin
    Result := csUtf16Le;
    Read(Bom, 2);
  end else

  if Bom = BOM_UTF16BE then begin
    Result := csUtf16Be;
    Read(Bom, 2);
  end else

  if word(Bom) = word(BOM_UTF8) then begin
   //full bom check -- 3 bytes
    Result := csUtf8;
    Bom3 := 0;
    Peek(Bom3, 3);
    if Bom3=BOM_UTF8 then
      Read(Bom3, 3);
  end else

 //No BOM => ANSI
  Result := csAnsi;
end;

function TCharReader.ReadChar(out c: WideChar): boolean;
var _c: AnsiChar;
  u, sz: integer;
begin
  case Charset of
    csAnsi: begin
      Result := (Read(_c, 1) = 1);
     {$IFDEF DCC}
      c := ToWideChar(_c, CP_ACP);
     {$ELSE}
      c := _c; //TODO: Use proper conversion.
     {$ENDIF}
    end;

    csUtf16Be: begin
      Result := (Read(c, 2) = 2);
      c := SwapChar(c);
    end;

    csUtf8: begin
      Result := (Peek(u, 4) = 4);
      sz := UTF8toUCS2(u, WideChar($FFFF), c);
      Read(u, sz);
    end

  else //Utf16Le
    Result := (Read(c, 2) = 2);
  end;
end;

function TCharReader.PeekChar(out c: WideChar): boolean;
var _c: AnsiChar;
  u: integer;
begin
  case Charset of
    csAnsi: begin
      Result := (Peek(_c, 1) = 1);
     {$IFDEF DCC}
      c := ToWideChar(_c, CP_ACP);
     {$ELSE}
      c := _c; //TODO: Use proper conversion
     {$ENDIF}
    end;

    csUtf16Be: begin
      Result := (Peek(c, 2) = 2);
      c := SwapChar(c);
    end;

    csUtf8: begin
      Result := (Peek(u, 4) = 4);
      UTF8toUCS2(u, WideChar($FFFF), c);
    end

  else //Utf16Le
    Result := (Peek(c, 2) = 2);
  end;
end;

(*
  Читает строку из потока, возвращает true, если удалось.
*)
function TCharReader.ReadLine(out s: UnicodeString): boolean;
var c: WideChar;
begin
  s := '';
  Result := ReadChar(c);
  if not Result then exit; {больше ничего нет}

 {иначе строчка точно есть}
  while Result do begin
    if c=#13 then begin
      if PeekChar(c) and (c=#10) then
        ReadChar(c);
      break;
    end else
    if c=#10 then begin
      if PeekChar(c) and (c=#13) then
        ReadChar(c);
      break;
    end else
      s := s + c;
    Result := ReadChar(c);
  end;
  Result := true; {уж что-то мы прочитали}
end;


{ TCharWriter }

constructor TCharWriter.Create(AStream: TStream; AOwnStream: boolean = false);
begin
  inherited Create(AStream, AOwnStream);
  Charset := csUtf16Le;
end;

constructor TCharWriter.Create(AStream: TStream; ACharSet: TCharSet;
  AOwnStream: boolean = false);
begin
  inherited Create(AStream, AOwnStream);
  CharSet := ACharSet;
end;

destructor TCharWriter.Destroy;
begin
  Flush();
  inherited;
end;

procedure TCharWriter.WriteChar(const c: WideChar);
var _c: AnsiChar;
  _c_be: WideChar;
  u, sz: integer;
begin
  case Charset of
    csAnsi: begin
     {$IFDEF DCC}
      _c := ToChar(c, CP_ACP);
     {$ELSE}
      _c := c; //TODO: Use proper conversion.
     {$ENDIF}
      Write(_c, 1);
    end;

    csUtf16Le:
      Write(c, 2);

    csUtf16Be: begin
     //Swap bytes
      _c_be := SwapChar(c);
      Write(_c_be, 2);
    end;

    csUtf8: begin
      sz := UCS2toUTF8(c, u);
      Write(u, sz);
    end
  end;
end;

procedure TCharWriter.WriteChars(c: PWideChar; len: integer);
var _c: AnsiString;
  _c_be: WideString;
begin
  case Charset of
    csAnsi: begin
     {$IFDEF DCC}
      _c := BufToString(c, len, CP_ACP);
     {$ELSE}
      _c := c; //TODO: Use proper conversion.
     {$ENDIF}
      Write(_c[1], Length(_c)*SizeOf(AnsiChar));
    end;

    csUtf16Le:
      Write(c^, len*SizeOf(WideChar));

    csUtf16Be: begin
      _c_be := SwapChars(c, len);
      Write(_c_be[1], len*SizeOf(WideChar));
    end;

    csUtf8:
      while len>0 do begin
        WriteChar(c^);
        Inc(c);
        Dec(len);
      end;
  end;
end;

procedure TCharWriter.WriteString(const c: UnicodeString);
begin
  WriteChars(@c[1], Length(c));
end;

procedure TCharWriter.WriteLine(const s: UnicodeString);
const CRLF: UnicodeString = #$000D#$000A;
begin
  WriteString(s);
  WriteString(CRLF);
end;

procedure TCharWriter.WriteBom;
var Bom: WideChar;
  Bom3: integer;
begin
  case Charset of
   //Nothing to write
    csAnsi: begin end;

    csUtf16Le:
    begin
      Bom := BOM_UTF16LE;
      Write(Bom, SizeOf(Bom));
    end;

    csUtf16Be:
    begin
      Bom := BOM_UTF16BE;
      Write(Bom, SizeOf(Bom));
    end;

    csUtf8:
    begin
      Bom3 := BOM_UTF8;
      Write(Bom3, 3);
    end;
  end;
end;


////////////////////////////////////////////////////////////////////////////////
///  Misc useful read functions

destructor TStreamReader2.Destroy;
begin
  FreeTmpStr;
  inherited;
end;

//Reads everything up to #13#10 into the buffer + sets terminating NULL.
//If the buffer length in chars (received in "sz") is enough, it's left as it is.
//Else the buffer is expanded, and the "sz" is updated accordingly.
function TStreamReader2.TryReadLineA(var s: PAnsiChar; var sz: cardinal): boolean;
var c: AnsiChar;
  l_used: cardinal;
begin
  l_used := 0;

  while (Read(c, SizeOf(c))=SizeOf(c))
    and (c <> #13) do
  begin
   //Reallocate memory, if needed
    if l_used >= sz then begin
      sz := 2*sz + 10;
      ReallocMem(s, sz*SizeOf(AnsiChar));
    end;

    s[l_used] := c;
    Inc(l_used);
  end;

  if not (c=#13) then
   //If we've read nothing this time
    if l_used = 0 then begin
      Result := false;
      exit;
    end else begin
      s[l_used] := #00;
      Result := true;
      exit;
    end;

  if not (Read(c, SizeOf(c)) = SizeOf(c))
  or not (c = #10) then
    raise Exception.Create('Illegal linebreak detected at symbol ' + IntToStr(Position));

 //Reallocate memory, if needed
  if l_used >= sz then begin
    sz := 2*sz + 1;
    ReallocMem(s, sz);
  end;

  s[l_used] := #00;
  Result := true;
end;

function TStreamReader2.TryReadLineW(var s: PWideChar; var sz: cardinal): boolean;
var c: WideChar;
  l_used: cardinal;
begin
  l_used := 0;

  while (Read(c, SizeOf(c))=SizeOf(c))
    and (c <> #10) do
  begin
   //Reallocate memory, if needed
    if l_used >= sz then begin
      sz := 2*sz + 10;
      ReallocMem(s, sz*SizeOf(WideChar));
    end;

    s[l_used] := c;
    Inc(l_used);
  end;

  if not (c=#10) then
   //If we've read nothing this time
    if l_used = 0 then begin
      Result := false;
      exit;
    end;

 //Reallocate memory, if needed
  if l_used >= sz then begin
    sz := 2*sz + 1;
    ReallocMem(s, sz);
  end;

  s[l_used] := #00;
  Result := true;
end;

procedure TStreamReader2.ReadLineA(var s: PAnsiChar; var sz: cardinal);
begin
  if not TryReadLineA(s, sz) then
    raise Exception.Create('Cannot read line of text.');
end;

procedure TStreamReader2.ReadLineW(var s: PWideChar; var sz: cardinal);
begin
  if not TryReadLineW(s, sz) then
    raise Exception.Create('Cannot read line of text.');
end;

function TStreamReader2.ReadStrA: AnsiString;
begin
  ReadLineA(tmpstr_ansi_s, tmpstr_ansi_sz);
  Result := tmpstr_ansi_s;
end;

function TStreamReader2.ReadStrW: WideString;
begin
  ReadLineW(tmpstr_wide_s, tmpstr_wide_sz);
  Result := tmpstr_wide_s;
end;

function TStreamReader2.TryReadStrA(out s: AnsiString): boolean;
begin
  Result := TryReadLineA(tmpstr_ansi_s, tmpstr_ansi_sz);
  if Result then
    s := tmpstr_ansi_s;
end;

function TStreamReader2.TryReadStrW(out s: WideString): boolean;
begin
  Result := TryReadLineW(tmpstr_wide_s, tmpstr_wide_sz);
  if Result then
    s := tmpstr_wide_s;
end;

procedure TStreamReader2.FreeTmpStr;
begin
  FreeMem(tmpstr_ansi_s);
  FreeMem(tmpstr_wide_s);
  tmpstr_ansi_s := nil;
  tmpstr_wide_s := nil;
  tmpstr_ansi_sz := 0;
  tmpstr_wide_sz := 0;
end;

////////////////////////////////////////////////////////////////////////////////

//Пытается прочесть один байт, интерпретирует его как AnsiChar.
function TStreamReader2.ReadAnsiChar(out c: AnsiChar): boolean;
begin
  Result := (Read(c, SizeOf(c))=SizeOf(AnsiChar));
end;

//Пытается прочесть два байта, интерпретирует их как WideChar.
function TStreamReader2.ReadWideChar(out c: WideChar): boolean;
begin
  Result := (Read(c, SizeOf(c))=SizeOf(WideChar));
end;

const
  REPL_CHAR = WideChar($FFFD);

//Пытается прочесть несколько байт, интерпретирует их как Utf8-char, преобразует в WideChar.
//Возвращает false, если поток кончился, а прочесть не удалось.
//Возвращает REPL_CHAR, если символ не влез в UCS2 или закодирован неверно.
function TStreamReader2.ReadUtf8Char(out c: WideChar): boolean;
var c1, c2, c3: byte;
begin
  Result := (Read(c1, 1)=1);
  if not Result then exit;

 //Один байт: 0xxxxxxx
  if (c1 and $80) = 0 then begin
    c := WideChar(c1);
    Result := true;
    exit;
  end;

 //Два байта: 110xxxxxx 10yyyyyy
  Result := (Read(c2, 1)=1);
  if not Result then exit;

 //У ведомых байт должно быть 10xxxxxx
  if (c2 and $C0 <> $80) then begin
    c := REPL_CHAR;
    exit;
  end;

  c1 := c1 and $3F; //сбрасываем в нули два левых бита
  if (c1 and $20) = 0 then begin //не стоит третий бит
    c := WideChar((c1 shl 6) or (c2 and $3F));
    Result := true;
    exit;
  end;

 //Три байта: 1110xxxx 10yyyyyy 10zzzzzz
  Result := (Read(c3, 1)=1);
  if not Result then exit;

 //У ведомых байт должно быть 10xxxxxx
  if (c3 and $C0 <> $80) then begin
    c := REPL_CHAR;
    exit;
  end;

  c1 := c1 and $1F; //сбрасываем в ноль третий бит
  if (c1 and $10) = 0 then begin //не стоит четвёртый бит
    c := WideChar((c1 shl 12) or ((c2 and $3F) shl 6) or (c3 and $3F));
    Result := true;
    exit;
  end;

 //Четыре байта: у нас не поддерживается. Но мы прочтём четвёртый.
  Result := (Read(c1, 1)=1); //уже неважно, куда
  if not Result then exit;
  c := REPL_CHAR;

 //Больше четырёх байт мы в страшном сне представить не можем.
end;

end.
