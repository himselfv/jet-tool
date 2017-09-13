unit Jet.CommandLine;

interface
uses SysUtils, UniStrUtils;

type
  EUsage = class(Exception);

procedure BadUsage(msg: UniString='');
procedure Redefined(term: string; old, new: UniString);
procedure Define(var Term: UniString; TermName: string; Value: UniString);


type
 {
 Command line parsing context. For now its a record because it's easier to track its lifetime.
 We may later make this a virtual class and overrides will pull params from different sources:
   string, TStringList, ParamStr()

 As far as this module is concerned:
   argument == any continuous element in the command line
   option == argument starting with - or --
   param == any element in the command line which serves as a parameter for the preceding argument

 }
  TParsingContext = record
    i: integer;
    procedure Reset;
    function TryNextArg(out value: UniString): boolean;
    function NextParam(key, param: UniString): UniString;
    function TryNextParam(key, param: UniString; out value: UniString): boolean;
  end;
  PParsingContext = ^TParsingContext;

  TCommandLineParser = class
  public
    procedure Reset; virtual;
    function HandleOption(ctx: PParsingContext; const s: UniString): boolean; virtual;
    procedure Finalize; virtual;
  end;

implementation

procedure BadUsage(msg: UniString='');
begin
  raise EUsage.Create(msg);
end;

procedure Redefined(term: string; old, new: UniString);
begin
  raise EUsage.Create(term+' already specified: '+old+'. Cannot process command "'+new+'".');
end;

procedure Define(var Term: UniString; TermName: string; Value: UniString);
begin
  if Term <> '' then
    Redefined(TermName, Term, Value);
  Term := Value;
end;


procedure TParsingContext.Reset;
begin
  i := 1;
end;

//Tries to consume one more argument from the command line or returns false
function TParsingContext.TryNextArg(out value: UniString): boolean;
begin
  Result := (i < ParamCount);
  if Result then begin
    Inc(i);
    value := ParamStr(i+1);
  end;
end;

//Tries to consume one more _parameter_ for the current _argument_ from the command line.
//Parameters can't start with - or --, this is considered the beginning of the next option.
//If you need this format, use TryNextArg.
function TParsingContext.TryNextParam(key, param: UniString; out value: UniString): boolean;
begin
  if i >= ParamCount then begin
    Result := false;
    exit;
  end;
  value := ParamStr(i+1);
  Result := not value.StartsWith('-');
  if Result then
    Inc(i)
  else
    value := '';
end;

function TParsingContext.NextParam(key, param: UniString): UniString;
begin
  Inc(i);
  if i > ParamCount then
    BadUsage(Key+' requires specifying the '+param);
  Result := ParamStr(i);
end;



procedure TCommandLineParser.Reset;
begin
  //Override to reset any configurable values to their default state
end;

function TCommandLineParser.HandleOption(ctx: PParsingContext; const s: string): boolean;
begin
  //Override to handle some options and return true
  Result := false;
end;

procedure TCommandLineParser.Finalize;
begin
  //Override to perform any post-processing and consistency checks after parsing
  //the available parameters.
end;

end.
