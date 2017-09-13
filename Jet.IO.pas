unit Jet.IO;
{
Console IO, logging and other system facilities.
}

interface
uses Windows, UniStrUtils, Jet.CommandLine;

type
  TLoggingMode = (lmDefault, lmSilent, lmNormal, lmVerbose);

//Application settings. Mostly set from command-line or auto-configured.
var
  LoggingMode: TLoggingMode = lmDefault;
 //I/O Redirects
  stdi, stdo, stde: UniString;


//Writes a string to error output.
//All errors, usage info, hints go here. Redirect it somewhere if you don't need it.
procedure err(msg: UniString);

//Writes a string to error output if not in silent mode
procedure warn(msg: UniString);

//Writes a string to error output if verbose log is enabled.
procedure log(msg: UniString);
procedure verbose(msg: UniString);


function IsConsoleHandle(stdHandle: cardinal): boolean;


type
  TIoSettingsParser = class(TCommandLineParser)
  public
    procedure Reset; override;
    function HandleOption(ctx: PParsingContext; const s: UniString): boolean; override;
    procedure Finalize; override;
  end;

var
  IoSettings: TIoSettingsParser;

implementation
uses SysUtils;

procedure err(msg: UniString);
begin
  writeln(ErrOutput, msg);
end;

procedure warn(msg: UniString);
begin
  if LoggingMode<>lmSilent then
    err(msg);
end;

procedure log(msg: UniString);
begin
  if LoggingMode = lmVerbose then
    err(msg);
end;

procedure verbose(msg: UniString);
begin
  if LoggingMode = lmVerbose then
    err(msg);
end;


//Returns true when a specified STD_HANDLE actually points to a console object.
//It's a LUCKY / UNKNOWN type situation: if it does, we're lucky and assume
//keyboard input. If it doesn't, we don't know: it might be a file or a pipe
//to another console.
function IsConsoleHandle(stdHandle: cardinal): boolean;
begin
 //Failing GetFileType/GetStdHandle is fine, we'll return false.
  Result := (GetFileType(GetStdHandle(stdHandle))=FILE_TYPE_CHAR);
end;


procedure TIoSettingsParser.Reset;
begin
end;

function TIoSettingsParser.HandleOption(ctx: PParsingContext; const s: UniString): boolean;
begin
  Result := true;
  if WideSameText(s, '-stdi') then
    Define(stdi, 'Filename', ctx.NextParam('-stdi', 'filename'))
  else
  if WideSameText(s, '-stdo') then
    Define(stdo, 'Filename', ctx.NextParam('-stdo', 'filename'))
  else
  if WideSameText(s, '-stde') then begin
    Define(stde, 'Filename', ctx.NextParam('-stde', 'filename'));
  end else
    Result := false;
end;

//Leaks file handles! By design. We don't care, they'll be released on exit anyway.
procedure RedirectIo(nStdHandle: dword; filename: UniString);
var hFile: THandle;
begin
  if nStdHandle=STD_INPUT_HANDLE then
    hFile := CreateFileW(PWideChar(Filename), GENERIC_READ, FILE_SHARE_READ,
      nil, OPEN_EXISTING, 0, 0)
  else
    hFile := CreateFileW(PWideChar(Filename), GENERIC_WRITE, 0,
      nil, OPEN_ALWAYS, 0, 0);
  if hFile=INVALID_HANDLE_VALUE then
    RaiseLastOsError();
  SetStdHandle(nStdHandle, hFile);
end;

procedure TIoSettingsParser.Finalize;
begin
 //Auto-redirect I/O
  if stdi<>'' then
    RedirectIo(STD_INPUT_HANDLE, stdi);
  if stdo<>'' then
    RedirectIo(STD_OUTPUT_HANDLE, stdo);
  if stde<>'' then
    RedirectIo(STD_OUTPUT_HANDLE, stde);
end;


initialization
  IoSettings := TIoSettingsParser.Create;

finalization
{$IFDEF DEBUG}
  FreeAndNil(IoSettings);
{$ENDIF}

end.
