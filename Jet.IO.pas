unit Jet.IO;
{
Console IO, logging and other system facilities.
}

interface
uses Windows, UniStrUtils, Jet.CommandLine;

type
 //"Default" states are needed because some defaults are unknown until later.
 //They will be resolved before returning from ParseCommandLine.
  TLoggingMode = (lmDefault, lmSilent, lmNormal, lmVerbose);
  TErrorHandlingMode = (emDefault, emIgnore, emStop);
  TTriBool = (tbDefault, tbTrue, tbFalse);

//Application settings. Mostly set from command-line or auto-configured.
var
  LoggingMode: TLoggingMode = lmDefault;
  Errors: TErrorHandlingMode = emDefault;
  CrlfBreak: TTriBool = tbDefault;


//Writes a string to error output.
//All errors, usage info, hints go here. Redirect it somewhere if you don't need it.
procedure err(msg: UniString);

//Writes a string to error output if not in silent mode
procedure warn(msg: UniString);

//Writes a string to error output if verbose log is enabled.
procedure log(msg: UniString);
procedure verbose(msg: UniString);


type
  TIoSettingsParser = class(TCommandLineParser)
  protected
   //I/O Redirects
    stdi, stdo, stde: UniString;
  public
    procedure ShowHelp; override;
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


procedure TIoSettingsParser.ShowHelp;
begin
  err('What to do with errors when executing:');
  err('  --silent :: do not print anything at all');
  err('  --verbose :: echo commands which are being executed');
  err('  --ignore-errors :: continue on errors');
  err('  --stop-on-errors :: exit with error code');
  err('  --crlf-break :: CR/LF ends command');
  err('  --no-crlf-break');
  err('With private extensions enabled, **WEAK** commands do not produce errors in any way (no messages, no aborts).');
  err('');
  err('IO redirection helpers:');
  err('  -stdi [filename] :: sets standard input');
  err('  -stdo [filename] :: sets standard output');
  err('  -stde [filename] :: sets standard error console');
  err('These are only applied after the command-line parsing is over');
  err('');
end;

function TIoSettingsParser.HandleOption(ctx: PParsingContext; const s: UniString): boolean;
begin
  Result := true;

  //Logging
  if WideSameText(s, '--silent') then begin
    LoggingMode := lmSilent
  end else
  if WideSameText(s, '--verbose') then begin
    LoggingMode := lmVerbose
  end else

  //Errors
  if WideSameText(s, '--ignore-errors') then begin
    Errors := emIgnore;
  end else
  if WideSameText(s, '--stop-on-errors') then begin
    Errors := emStop;
  end else

  //CRLF break
  if WideSameText(s, '--crlf-break') then begin
    CrlfBreak := tbTrue;
  end else
  if WideSameText(s, '--no-crlf-break') then begin
    CrlfBreak := tbFalse;
  end else

  //IO redirection
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

//Returns true when a specified STD_HANDLE actually points to a console object.
//It's a LUCKY / UNKNOWN type situation: if it does, we're lucky and assume
//keyboard input. If it doesn't, we don't know: it might be a file or a pipe
//to another console.
function IsConsoleHandle(stdHandle: cardinal): boolean;
begin
 //Failing GetFileType/GetStdHandle is fine, we'll return false.
  Result := (GetFileType(GetStdHandle(stdHandle))=FILE_TYPE_CHAR);
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
var KeyboardInput: boolean;
begin
 //Resolve default values depending on a type of input stream.
 //If we fail to guess the type, default to File (this can always be overriden manually!)
  KeyboardInput := IsConsoleHandle(STD_INPUT_HANDLE) and (IoSettings.stdi='');
  if Errors=emDefault then
    if KeyboardInput then
      Errors := emIgnore
    else
      Errors := emStop;
  if LoggingMode=lmDefault then
    if KeyboardInput then
      LoggingMode := lmVerbose
    else
      LoggingMode := lmNormal;
  if CrlfBreak=tbDefault then
    if KeyboardInput then
      CrlfBreak := tbTrue
    else
      CrlfBreak := tbFalse;

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
