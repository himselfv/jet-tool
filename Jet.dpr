program Jet;

{$APPTYPE CONSOLE}

uses
  SysUtils, Windows, ActiveX, Variants, AdoDb, OleDb, AdoInt, ComObj, UniStrUtils,
  DAO_TLB, ADOX_TLB, StreamUtils,
  StringUtils in 'StringUtils.pas',
  DaoDumper in 'DaoDumper.pas',
  JetCommon in 'JetCommon.pas',
  AdoxDumper in 'AdoxDumper.pas';

(*
  Supported propietary extension comments:
    /**COMMENT* [comment text] */  - table, field or view comment
    /**WEAK**/ - ignore errors for this command (handy for DROP TABLE/VIEW)

  By default when reading from keyboard these are set:
    --ignore-errors
    --crlf-break
    --verbose

  These are dumped by default:
    --tables
    --views
    --relations
    --comments
    --private-extensions

  Returns error codes:
    1 :: Generic error
    2 :: Usage error
    3 :: OLE Error

  What this thing dumps:
    Tables
      Comments
      Fields
        Standard field markers [not null, default]
        Auto-increment
        Comments
    Indexes
      Multi- and single-type indexes
      Standard markers [unique, disallow null, ignore null, primary key]
    Foreign keys
    Views

  What can be added in the future:
    Check constraints
    Rest of the "table constraints"
    Stored procedures

  A note about constraints. There are:
    Check constraints :: added through SQL, unrelated to anything else
    Referential constraints :: automatically generated for foreign keys
    Unique index constraints :: automatically added for every UNIQUE index
    Manual table constraints :: any other constraints added manually through
      ADD CONSTRAINT which are not CHECK CONSTRAINT.
  All of these compose "Table constraints".

  Documentation:
    http://msdn.microsoft.com/en-us/library/ms714540(VS.85).aspx  :: Access data types
    http://msdn.microsoft.com/en-us/library/bb267262(office.12).aspx :: Access 2007 DDL

  Schema documentation:
    http://msdn.microsoft.com/en-us/library/ms675274(VS.85).aspx  :: Schema identifiers
    http://msdn.microsoft.com/en-us/library/ee265709(BTS.10).aspx  :: Some of the schemas, documented

  Adding and modifying check constraints (not documented in MSDN):
    http://www.w3schools.com/Sql/sql_check.asp
*)

{$UNDEF DEBUG}

type
  PWideCharArray = ^TWideCharArray;
  TWideCharArray = array[0..16383] of WideChar;
  OemString = type AnsiString(CP_OEMCP);

//Writes a string to error output.
//All errors, usage info, hints go here. Redirect it somewhere if you don't need it.
procedure err(msg: UniString);
begin
  writeln(ErrOutput, msg);
end;

type
  EUsage = class(Exception);

procedure PrintShortUsage;
begin
  err('Do "jet help" for extended info.');
end;

procedure PrintUsage;
begin
  err('Usage:');
  err('  jet <command> [params]');
  err('');
  err('Commands:');
  err('  jet touch :: connect to database and quit');
  err('  jet dump :: dump sql schema');
  err('  jet exec :: parse sql from input');
  err('  jet schema :: output internal jet schema reports');
  err('  jet daoschema :: output DAO schema report');
  err('  jet adoxschema :: output ADOX schema report');
  err('');
  err('Connection params:');
  err('  -c [connection-string] :: uses an ADO connection string. -dp is ignored');
  err('  -dsn [data-source-name] :: uses an ODBC data source name');
  err('  -f [file.mdb] :: opens a jet database file');
  err('  -u [user]');
  err('  -p [password]');
  err('  -dp [database password]'); {Works fine with database creation too}
  err('  -new :: works only with exec and filename');
  err('  -force :: overwrite existing database (requires -new)');
  err('You cannot use -c with --comments when executing (dumping is fine).');
  err('You can only use -new with -f.');
 (* -dsn will probably not work with --comments too, as long as it really is MS Access DSN. They deny DAO DSN connections. *)
  err('');
  err('Useful IO tricks:');
  err('  -stdi [filename] :: sets standard input');
  err('  -stdo [filename] :: sets standard output');
  err('  -stde [filename] :: sets standard error console');
  err('These are only applied after the command-line parsing is over');
  err('');
  err('What to include and whatnot for dumping:');
  err('  --no-tables, --tables');
  err('  --no-views, --views');
  err('  --no-procedures, --procedures');
  err('  --no-comments, --comments :: how comments are dumped depends on if private extensions are enabled');
  err('  --no-drop, --drop :: "DROP" tables etc before creating');
  err('');
  err('With --tables and --views you can specify individual names:');
  err('  --tables [tablename],[tablename]');
  err('  --views [viewname],[viewname]');
  err('Specify --case-sensitive-ids or --case-insensitive-ids if needed (default: sensitive).');
  err('');
  err('Works both for dumping and executing:');
  err('  --no-private-extensions, --private-extensions :: disables dumping/parsing private extensions (check help)');
  err('');
  err('What to do with errors when executing:');
  err('  --silent :: do not print anything (at all)');
  err('  --verbose :: echo commands which are being executed');
  err('  --ignore-errors :: continue on error');
  err('  --stop-on-errors :: exit with error code');
  err('  --crlf-break :: CR/LF ends command');
  err('  --no-crlf-break');
  err('With private extensions enabled, **WEAK** commands do not produce errors in any way (messages, stop).')
end;

procedure BadUsage(msg: UniString='');
begin
  raise EUsage.Create(msg);
end;

procedure Redefined(term: string; old, new: UniString);
begin
  raise EUsage.Create(term+' already defined: '+old+'. Cannot redefine to "'+new+'".');
end;

function IsConsoleHandle(stdHandle: cardinal): boolean; forward;

type
 //"Default" states are needed because some defaults are unknown until later.
 //They will be resolved before returning from ParseCommandLine.
  TLoggingMode = (lmDefault, lmSilent, lmNormal, lmVerbose);
  TErrorHandlingMode = (emDefault, emIgnore, emStop);
  TTriBool = (tbDefault, tbTrue, tbFalse);

var
  Command: UniString;
 //Connection
  ConnectionString: UniString;
  DataSourceName: UniString;
  Filename: UniString;
  User, Password: UniString;
  DatabasePassword: UniString;
  NewDb: boolean;
  ForceNewDb: boolean;
 //Database options
  CaseInsensitiveIDs: boolean;
 //Dump contents
  NeedDumpTables: boolean = true;
  DumpTableList: TUniStringArray; //empty = all
  NeedDumpViews: boolean = true;
  DumpViewList: TUniStringArray;
  NeedDumpProcedures: boolean = false;
  NeedDumpRelations: boolean = true;
  NeedDumpCheckConstraints: boolean = false;
 //Dump options
  HandleComments: boolean = true;
  PrivateExtensions: boolean = true;
  DropObjects: boolean = true;
 //Log
  LoggingMode: TLoggingMode = lmDefault;
  Errors: TErrorHandlingMode = emDefault;
 //Redirects
  stdi, stdo, stde: UniString;
  CrlfBreak: TTriBool = tbDefault;

var //Dynamic properties
  CanUseDao: boolean; //sometimes there are other ways

procedure ParseCommandLine;
var i: integer;
  s, list: UniString;
  KeyboardInput: boolean;

  procedure Define(var Term: UniString; TermName: string; Value: UniString);
  begin
    if Term <> '' then
      Redefined(TermName, Term, Value);
    Term := Value;
  end;

  function NextParam(key, param: UniString): UniString;
  begin
    Inc(i);
    if i > ParamCount then
      BadUsage(Key+' requires specifying the '+param);
    Result := ParamStr(i);
  end;

  //Same as NextParam, but returns false if there's no param or next param is a flag
  //Can return true and value=='' if there's an explicitly defined empty param (that's a feature)
  function TryNextParam(key, param: UniString; out value: UniString): boolean;
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

begin
  i := 1;
  while i <= ParamCount do begin
    s := ParamStr(i);
    if Length(s) <= 0 then begin
      Inc(i);
      continue;
    end;

    if s[1] <> '-' then begin
      Define(Command, 'Command', s);
    end else

    if WideSameText(s, '-c') then begin
      s := NextParam('-c', 'connection string');
      Define(ConnectionString, 'Connection string', s);
    end else
    if WideSameText(s, '-dsn') then begin
      s := NextParam('-dsn', 'data source name');
      Define(DataSourceName, 'Data source name', s);
    end else
    if WideSameText(s, '-f') then begin
      s := NextParam('-f', 'filename');
      Define(Filename, 'Filename', s);
    end else
    if WideSameText(s, '-u') then begin
      s := NextParam('-u', 'username');
      Define(User, 'Username', s);
    end else
    if WideSameText(s, '-p') then begin
      s := NextParam('-p', 'password');
      Define(Password, 'Password', s);
    end else
    if WideSameText(s, '-dp') then begin
      s := NextParam('-dp', 'database password');
      Define(DatabasePassword, 'Database password', s);
    end else

    if WideSameText(s, '-new') then begin
      NewDb := true;
    end else
    if WideSameText(s, '-force') then begin
      ForceNewDb := true;
    end else

   //Database options
    if WideSameText(s, '--case-sensitive-ids') then begin
      CaseInsensitiveIDs := false;
    end else
    if WideSameText(s, '--case-insensitive-ids') then begin
      CaseInsensitiveIDs := true;
    end else

   //What to dump
    if WideSameText(s, '--tables') then begin
      NeedDumpTables := true;
      if TryNextParam('--tables', 'table list', list) then begin
        DumpTableList := Split(list, ',');
        if Length(DumpTableList) <= 0 then
         //Empty DumpTableList internally means dump all, so just disable dump if asked for none
          NeedDumpTables := false;
      end;
    end else
    if WideSameText(s, '--no-tables') then begin
      NeedDumpTables := false;
    end else
    if WideSameText(s, '--views') then begin
      NeedDumpViews := true;
      if TryNextParam('--views', 'view list', list) then begin
        DumpViewList := Split(list, ',');
        if Length(DumpViewList) <= 0 then
          NeedDumpViews := false;
      end;
    end else
    if WideSameText(s, '--no-views') then begin
      NeedDumpViews := false;
    end else
    if WideSameText(s, '--procedures') then begin
      NeedDumpProcedures := true;
    end else
    if WideSameText(s, '--no-procedures') then begin
      NeedDumpProcedures := false;
    end else
    if WideSameText(s, '--relations') then begin
      NeedDumpRelations := true;
    end else
    if WideSameText(s, '--no-relations') then begin
      NeedDumpRelations := false;
    end else
    if WideSameText(s, '--check-constraints') then begin
      NeedDumpCheckConstraints := true;
    end else
    if WideSameText(s, '--no-check-constraints') then begin
      NeedDumpCheckConstraints := false;
    end else

   //Dump options
    if WideSameText(s, '--comments') then begin
      HandleComments := true; {when dumping this means DUMP comments, else PARSE comments - requires PrivateExt}
    end else
    if WideSameText(s, '--no-comments') then begin
      HandleComments := false;
    end else
    if WideSameText(s, '--private-extensions') then begin
      PrivateExtensions := true;
    end else
    if WideSameText(s, '--no-private-extensions') then begin
      PrivateExtensions := false;
    end else
    if WideSameText(s, '--drop') then begin
      DropObjects := true;
    end else
    if WideSameText(s, '--no-drop') then begin
      DropObjects := false;
    end else

    if WideSameText(s, '--silent') then begin
      LoggingMode := lmSilent
    end else
    if WideSameText(s, '--verbose') then begin
      LoggingMode := lmVerbose
    end else

    if WideSameText(s, '--ignore-errors') then begin
      Errors := emIgnore;
    end else
    if WideSameText(s, '--stop-on-errors') then begin
      Errors := emStop;
    end else

    if WideSameText(s, '--crlf-break') then begin
      CrlfBreak := tbTrue;
    end else
    if WideSameText(s, '--no-crlf-break') then begin
      CrlfBreak := tbFalse;
    end else

    if WideSameText(s, '-stdi') then begin
      s := NextParam('-stdi', 'filename');
      Define(stdi, 'Filename', s);
    end else
    if WideSameText(s, '-stdo') then begin
      s := NextParam('-stdo', 'filename');
      Define(stdo, 'Filename', s);
    end else
    if WideSameText(s, '-stde') then begin
      s := NextParam('-stde', 'filename');
      Define(stde, 'Filename', s);
    end else

     //Default case
      BadUsage('Unsupported option: '+s);

    Inc(i);
    continue;
  end;

 //If asked for help, there's nothing to check
  if WideSameText(Command, 'help') then exit;

 //Check params, only one is allowed: ConnectionString, ConnectionName, Filename
  i := 0;
  if ConnectionString<>'' then Inc(i);
  if DataSourceName<>'' then Inc(i);
  if Filename<>'' then Inc(i);
  if i > 1 then BadUsage('Only one source (ConnectionString/DataSourceName/Filename) can be specified.');
  if i < 1 then BadUsage('A source (ConnectionString/DataSourceName/Filename) needs to be specified.');

 //With ConnectionString, additional params are disallowed: everything is in.
  if (ConnectionString<>'') and (DatabasePassword<>'') then
    BadUsage('ConnectionString conflicts with DatabasePassword: this should be included inside of the connection string.');

 //If requested to create a db, allow only filename connection.
  if NewDb and (Filename='') then
    BadUsage('Database creation is supported only when connecting by Filename.');
  if NewDb and (WideSameText(Command, 'dump')
    or WideSameText(Command, 'schema') or WideSameText(Command, 'daoschema')
    or WideSameText(Command, 'adoxschema'))
  and (LoggingMode=lmVerbose) then begin
    err('NOTE: You asked to create a database and then dump its contents.');
    err('What the hell are you trying to do?');
  end;
  if ForceNewDb and not NewDb then
    BadUsage('-force requires -new');

  if NewDb and not ForceNewDb and FileExists(Filename) then
    raise Exception.Create('File '+Filename+' already exists. Cannot create a new one.');

 //Whether we can use DAO. If not, prefer other options.
  CanUseDao := (Filename<>'');

 //To parse comments we need DAO, i.e. filename connection
  if WideSameText(Command, 'exec') and HandleComments then begin
    if not PrivateExtensions then
      BadUsage('You need --private-extensions to handle --comments when doing "exec".');
    if (Filename='') and (DataSourceName='') then
      BadUsage('You cannot use ConnectionString source with --comments when doing "exec"');
  end;

 //Convert all sources to ConnectionStrings
  if Filename<>'' then
    ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source="'+Filename+'";'
  else
  if DataSourceName <> '' then
    ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source="'+DataSourceName+'";';

  if DatabasePassword<>'' then
    ConnectionString := ConnectionString + 'Jet OLEDB:Database Password="'+DatabasePassword+'";';

 //If asked for case-insensitive IDs, lowercase all relevant cached info
  if CaseInsensitiveIDs then begin
    DumpTableList := ToLowercase(DumpTableList);
    DumpViewList := ToLowercase(DumpViewList);
  end;

 //Resolve default values depending on a type of input stream.
 //If we fail to guess the type, default to File (this can always be overriden manually!)
  KeyboardInput := IsConsoleHandle(STD_INPUT_HANDLE) and (stdi='');
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
end;


const
  Jet10 = 1;
  Jet11 = 2;
  Jet20 = 3;
  Jet3x = 4;
  Jet4x = 5;

var
  AdoxCatalog: Catalog;

//Creates a new database and resets a database-creation-required flag.
procedure CreateNewDatabase;
begin
  if ForceNewDb and FileExists(Filename) then
    DeleteFileW(PWideChar(Filename));
  AdoxCatalog := CoCatalog.Create;
  AdoxCatalog.Create(ConnectionString
    + 'Jet OLEDB:Engine Type='+IntToStr(Jet4x)+';');
  if LoggingMode=lmVerbose then
    writeln('Database created.');
  NewDb := false; //works only once
end;

function GetAdoConnection: _Connection; forward;

//Returns an ADOX Catalog. Caching is implemented.
function GetAdoxCatalog: Catalog;
begin
  if AdoxCatalog=nil then begin
    AdoxCatalog := CoCatalog.Create;
    AdoxCatalog._Set_ActiveConnection(GetAdoConnection);
  end;
  Result := AdoxCatalog;
end;

var
  AdoConnection: _Connection;

//Returns an ADO connection. Caching is implemented
function GetAdoConnection: _Connection;
begin
  if NewDb then CreateNewDatabase;
  if AdoConnection=nil then begin
    Result := CoConnection.Create;
    Result.Open(ConnectionString, User, Password, 0);
    AdoConnection := Result;
  end else
    Result := AdoConnection;
end;

//Returns a NEW ADO connection.
function EstablishAdoConnection: _Connection;
begin
  if NewDb then CreateNewDatabase;
  Result := CoConnection.Create;
  Result.Open(ConnectionString, User, Password, 0);
end;

var
  DaoEngine: DbEngine;
  DaoConnection: Database;

procedure NeededDaoError;
begin
  raise Exception.Create('The operation you''re performing apparently requires '
    +'DAO. DAO and DAO-dependent functions can only be accessed through Filename '
    +'source. That you see this error must mean that somehow this condition was not '
    +'properly verified during command-line parsing. '#13
    +'Please file a bug to the developers. For the time being, try to guess which '
    +'setting required DAO (usually something obscure like importing comments) '
    +'and either disable that or switch to connecting through Filename.');
end;

//Returns a DAO connection. Caching is implemented.
function GetDaoConnection: Database;
var Params: OleVariant;
begin
  if Assigned(DaoConnection) then begin
    Result := DaoConnection;
    exit;
  end;

  if NewDb then CreateNewDatabase;

  DaoEngine := CoDbEngine.Create;
  if Filename<>'' then begin
    if DatabasePassword<>'' then
      Params := UniString('MS Access;pwd=')+DatabasePassword
    else
      Params := '';
    Result := DaoEngine.OpenDatabase(Filename, False, False, Params);
  end else
  if DataSourceName<>'' then begin
   //Although this will probably not work
    Params := 'ODBC;DSN='+DataSourceName+';UID='+User+';PWD='+Password+';';
    Result := DaoEngine.OpenDatabase('', False, False, Params);
  end else
    NeededDaoError();
end;

//Establishes a NEW DAO connection. Also refreshes the engine cache.
//Usually called every time DAO connection is needed by a DAO-dependent proc.
function EstablishDaoConnection: Database;
var DbEngine: _DbEngine;
  Params: OleVariant;
begin
  DbEngine := CoDbEngine.Create;
 //Do not disable, or Dao refreshing will break too
  DbEngine.Idle(dbRefreshCache);
  if Filename<>'' then begin
    if DatabasePassword<>'' then
      Params := UniString('MS Access;pwd=')+DatabasePassword
    else
      Params := '';
    Result := DbEngine.OpenDatabase(Filename, False, False, Params);
  end else
  if DataSourceName<>'' then begin
   //Although this will probably not work
    Params := 'ODBC;DSN='+DataSourceName+';UID='+User+';PWD='+Password+';';
    Result := DbEngine.OpenDatabase('', False, False, Params);
  end else
    NeededDaoError();
end;

//This is needed before you CoUninitialize Ole. More notes where this is called.
procedure ClearOleObjects;
begin
  AdoConnection := nil;
  DaoConnection := nil;
  AdoxCatalog := nil;
end;

////////////////////////////////////////////////////////////////////////////////
/// Touch --- establishes connection and quits

procedure Touch();
var conn: _Connection;
begin
  conn := GetAdoConnection;
end;

////////////////////////////////////////////////////////////////////////////////
//// Schema --- Dumps many schema tables returned by Access

procedure DumpRecordset(rs: _Recordset);
var i: integer;
begin
  while not rs.EOF do begin
    for i := 0 to rs.Fields.Count - 1 do
      writeln(rs.Fields[i].Name+'='+str(rs.Fields[i].Value));
    writeln('');
    rs.MoveNext();
  end;
end;

procedure DumpSchema(conn: _Connection; Schema: integer);
var rs: _Recordset;
begin
  rs := conn.OpenSchema(Schema, EmptyParam, EmptyParam);
  DumpRecordset(rs);
end;

procedure Dump(conn: _Connection; SectionName: string; Schema: integer);
begin
  Section(SectionName);
  DumpSchema(conn, Schema);
end;

procedure PrintSchema();
var conn: _Connection;
begin
  conn := GetAdoConnection;
//  Dump(conn, 'Catalogs', adSchemaCatalogs); ---not supported by Access
  Dump(conn, 'Tables', adSchemaTables);
  Dump(conn, 'Columns', adSchemaColumns);

//  Dump(conn, 'Asserts', adSchemaAsserts); ---not supported by Access
  Dump(conn, 'Check constraints', adSchemaCheckConstraints);
  Dump(conn, 'Referential constraints', adSchemaReferentialConstraints);
  Dump(conn, 'Table constraints', adSchemaTableConstraints);

  Dump(conn, 'Column usage', adSchemaConstraintColumnUsage);
  Dump(conn, 'Key column usage', adSchemaKeyColumnUsage);
//  Dump(conn, 'Table usage', adSchemaConstraintTableUsage); --- not supported by Access

  Dump(conn, 'Indexes', adSchemaIndexes);
  Dump(conn, 'Primary keys', adSchemaPrimaryKeys);
  Dump(conn, 'Foreign keys', adSchemaForeignKeys);

//  Dump(conn, 'Properties', adSchemaProperties); ---not supported by Access

//  Dump(conn, 'Procedure columns', adSchemaProcedureColumns); -- not supported by Access
//  Dump(conn, 'Procedure parameters', adSchemaProcedureParameters); -- not supported by Access
  Dump(conn, 'Procedures', adSchemaProcedures);

//  Dump(conn, 'Provider types', adSchemaProviderTypes); -- nothing of interest
end;

////////////////////////////////////////////////////////////////////////////////
//// DaoSchema -- dumps DAO structure

procedure PrintDaoSchema();
begin
  DaoDumper.PrintDaoSchema(GetDaoConnection);
end;

////////////////////////////////////////////////////////////////////////////////
//// AdoxSchema -- dumps ADOX structure

procedure PrintAdoxSchema();
begin
  AdoxDumper.PrintAdoxSchema(GetAdoxCatalog);
  GetAdoxcatalog.Tables[0].Properties
end;

////////////////////////////////////////////////////////////////////////////////
/// DumpSql --- Dumps database contents

(*
 Writes a warning to a generated file. If we're not in silent mode, outputs it as an error too.
 This should be used in cases where the warning is really important (something cannot be done,
 some information ommited). If you just want to give a hint, use "err(msg)";
*)
procedure Warning(msg: UniString);
begin
  writeln('/* !!! Warning: '+msg+' */');
  if LoggingMode<>lmSilent then
    err('Warning: '+msg);
end;

{$REGION 'Encoding'}
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
    Result := NULL
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
{$ENDREGION}

{$REGION 'Columns'}
type
  TColumnDesc = record
    Name: WideString;
    Description: WideString;
    Flags: integer;
    OrdinalPosition: integer;
    DataType: integer;
    IsNullable: OleVariant;
    HasDefault: OleVariant;
    Default: OleVariant;
    NumericScale: OleVariant;
    NumericPrecision: OleVariant;
    CharacterMaximumLength: OleVariant;
    AutoIncrement: boolean;
  end;
  PColumnDesc = ^TColumnDesc;

  TColumns = record
    data: array of TColumnDesc;
    procedure Clear;
    procedure Add(Column: TColumnDesc);
    procedure Sort;
  end;

procedure TColumns.Clear;
begin
  SetLength(data, 0);
end;

procedure TColumns.Add(Column: TColumnDesc);
begin
  SetLength(data, Length(data)+1);
  data[Length(data)-1] := Column;
end;

procedure TColumns.Sort;
var i, j, k: integer;
  tmp: TColumnDesc;
begin
  for i := 1 to Length(data) - 1 do begin
    j := i-1;
    while (j >= 0) and (data[i].OrdinalPosition < data[j].OrdinalPosition) do
      Dec(j);
    Inc(j);
    if j<>i then begin
      tmp := data[i];
      for k := i downto j+1 do
        data[k] := data[k-1];
      data[j] := tmp;
    end;
  end;
end;
{$ENDREGION}

function GetTableText(conn: _Connection; Table: UniString): UniString;
var rs: _Recordset;
  Columns: TColumns;
  Column: TColumnDesc;
  s, dts: string;
  tmp: OleVariant;
  pre, scal: OleVariant;
  AdoxTable: ADOX_TLB.Table;
  DaoTable: DAO_TLB.TableDef;
begin
 //Doing this through DAO is slightly faster (OH GOD HOW MUCH DOES ADOX SUCK),
 //but DAO can only be used with -f, so effectively we just strip out all
 //the other options.
  if CanUseDao then //filename connection
    DaoTable := GetDaoConnection.TableDefs[Table]
  else
    AdoxTable := GetAdoxCatalog.Tables[Table];

  rs := conn.OpenSchema(adSchemaColumns,
    VarArrayOf([Unassigned, Unassigned, Table, Unassigned]), EmptyParam);
 //Reading data
  Columns.Clear;
  while not rs.EOF do begin
    Column.Name := str(rs.Fields['COLUMN_NAME'].Value);
    Column.Description := str(rs.Fields['DESCRIPTION'].Value);
    Column.Flags := int(rs.Fields['COLUMN_FLAGS'].Value);
    Column.OrdinalPosition := int(rs.Fields['ORDINAL_POSITION'].Value);

    tmp := rs.Fields['DATA_TYPE'].Value;
    if VarIsNil(tmp) then begin
      Warning('Empty data type for column '+Column.Name);
      rs.MoveNext();
      continue;
    end;
    Column.DataType := integer(tmp);

    Column.IsNullable := rs.Fields['IS_NULLABLE'].Value;
    Column.HasDefault := rs.Fields['COLUMN_HASDEFAULT'].Value;
    Column.Default := rs.Fields['COLUMN_DEFAULT'].Value;
    Column.NumericPrecision := rs.Fields['NUMERIC_PRECISION'].Value;
    Column.NumericScale := rs.Fields['NUMERIC_SCALE'].Value;
    Column.CharacterMaximumLength := rs.Fields['CHARACTER_MAXIMUM_LENGTH'].Value;

    if CanUseDao then
      Column.AutoIncrement := Includes(cardinal(DaoTable.Fields[Column.Name].Attributes), dbAutoIncrField)
    else
      Column.AutoIncrement := AdoxTable.Columns[Column.Name].Properties['AutoIncrement'].Value;

    Columns.Add(Column);
    rs.MoveNext;
  end;
  Columns.Sort;


 //Building string
  Result := '';
  for Column in Columns.data do begin
   //Data type
    if Column.AutoIncrement then
      dts := 'COUNTER' //special access data type, also known as AUTOINCREMENT
    else
    case Column.DataType of
      DBTYPE_I1: dts := 'TINYINT';
      DBTYPE_I2: dts := 'SMALLINT';
      DBTYPE_I4: dts := 'INTEGER';
      DBTYPE_I8: dts := 'BIGINT';
      DBTYPE_UI1: dts := 'BYTE';
      DBTYPE_UI2: dts := 'SMALLINT UNSIGNED';
      DBTYPE_UI4: dts := 'INTEGER UNSIGNED';
      DBTYPE_UI8: dts := 'BIGINT UNSIGNED';
      DBTYPE_CY: dts := 'MONEY';
      DBTYPE_R4: dts := 'REAL';
      DBTYPE_R8: dts := 'FLOAT';
      DBTYPE_GUID: dts := 'UNIQUEIDENTIFIER';
      DBTYPE_DATE: dts := 'DATETIME';
      DBTYPE_NUMERIC,
      DBTYPE_DECIMAL: begin
        pre := Column.NumericPrecision;
        scal := Column.NumericScale;
        if not VarIsNil(pre) and not VarIsNil(Scal) then begin
          dts := 'DECIMAL('+string(pre)+', '+string(scal)+')';
        end else
        if not VarIsNil(pre) then begin
          dts := 'DECIMAL('+string(pre)+')';
        end else
        if not VarIsNil(scal) then begin
          dts := 'DECIMAL(18, '+string(scal)+')'; //default pre
        end else
          dts := 'DECIMAL';
      end;
      DBTYPE_BOOL: dts := 'BIT';

      DBTYPE_BYTES:
        if Includes(Column.Flags, DBCOLUMNFLAGS_ISLONG) then
          dts := 'LONGBINARY'
        else
          dts := 'BINARY';

      DBTYPE_WSTR:
       //If you specify TEXT, Access makes LONGTEXT (Memo) field, if TEXT(len) then the usual limited text.
       //But we'll go the safe route.
        if Includes(Column.Flags, DBCOLUMNFLAGS_ISLONG) then
          dts := 'LONGTEXT' //"Memo field"
        else begin
          if VarIsNil(Column.CharacterMaximumLength) then begin
            Warning('Null CHARACTER_MAXIMUM_LENGTH although DBCOLUMNFLAGS_ISLONG '
              +'is not set on TEXT field');
            tmp := 0;
          end;
          dts := 'TEXT('+string(Column.CharacterMaximumLength)+')';
        end
    else
      Warning('Unsupported data type '+IntToStr(Column.DataType)+' for column '+Column.Name);
      rs.MoveNext();
      continue;
    end;

   //Main
    s := '['+Column.Name + '] ' + dts;
    if not bool(Column.IsNullable) then
      s := s + ' NOT NULL';
    if bool(Column.HasDefault) then
      if VarIsNil(Column.Default) then
        s := s + ' DEFAULT NULL'
      else
       //Default values do not need to be encoded, they're stored in an encoded way.
       //String values are properly escaped and quoted, Function() ones aren't, it's all fine.
        s := s + ' DEFAULT ' + str(Column.Default);

   //Access does not support comments, therefore we just output them in our propietary format.
   //If you use this importer, it'll understand them.
    if HandleComments and (Column.Description <> '') then
      if PrivateExtensions then
        s := s + ' /**COMMENT* '+EncodeComment(Column.Description)+' */'
      else
        s := s + ' /* '+EncodeComment(Column.Description)+' */';

    if Result <> '' then
      Result := Result + ','#13#10+s
    else
      Result := s;
  end;
end;

{$REGION 'Indexes'}
type
  TIndexColumnDesc = record
    Name: UniString;
    Collation: integer;
    OrdinalPosition: integer;
  end;
  PIndexColumnDesc = ^TIndexColumnDesc;

  TIndexDesc = record
    Name: UniString;
    PrimaryKey: boolean;
    Unique: boolean;
    Nulls: integer; //DBPROP_IN_*
    Columns: array of TIndexColumnDesc;
    _Initialized: boolean; //indicates that the object properties has been set
    _Declared: boolean; //set after index has been declared inline [field definition]
    function ContainsColumn(ColumnName: UniString): boolean;
    function AddColumn(ColumnName: UniString): PIndexColumnDesc;
    procedure SortColumns;
  end;
  PIndexDesc = ^TIndexDesc;

  TIndexes = record
    data: array of TIndexDesc;
    procedure Clear;
    function Find(IndexName: UniString): PIndexDesc;
    function Get(IndexName: UniString): PIndexDesc;
  end;

function TIndexDesc.ContainsColumn(ColumnName: UniString): boolean;
var i: integer;
begin
  Result := false;
  for i := 0 to Length(Columns) - 1 do
    if WideSameText(ColumnName, Columns[i].Name) then begin
      Result := true;
      break;
    end;
end;

function TIndexDesc.AddColumn(ColumnName: UniString): PIndexColumnDesc;
begin
  if ContainsColumn(ColumnName) then begin
    Result := nil;
    exit;
  end;
  SetLength(Columns, Length(Columns)+1);
  Result := @Columns[Length(Columns)-1];
  Result.Name := ColumnName;
  Result.Collation := 0; 
end;

procedure TIndexDesc.SortColumns;
var i, j, k: integer;
  tmp: TIndexColumnDesc;
begin
  for i := 1 to Length(Columns) - 1 do begin
    j := i-1;
    while (j >= 0) and (Columns[i].OrdinalPosition < Columns[j].OrdinalPosition) do
      Dec(j);
    Inc(j);
    if j<>i then begin
      tmp := Columns[i];
      for k := i downto j+1 do
        Columns[k] := Columns[k-1];
      Columns[j] := tmp;
    end;
  end;
end;

procedure TIndexes.Clear;
begin
  SetLength(data, 0);
end;

function TIndexes.Find(IndexName: UniString): PIndexDesc;
var i: integer;
begin
  Result := nil;
  for i := 0 to Length(data) - 1 do
    if WideSameText(IndexName, data[i].Name) then begin
      Result := @data[i];
      break;
    end;
end;

function TIndexes.Get(IndexName: UniString): PIndexDesc;
begin
  Result := Find(IndexName);
  if Result<>nil then exit;

  SetLength(data, Length(data)+1);
  Result := @data[Length(data)-1];
  Result^.Name := IndexName;
  Result._Initialized := false;
  Result._Declared := false;
end;
{$ENDREGION}

{$REGION 'Constraints'}
type
  TConstraintDesc = record
    Name: UniString;
  end;
  PConstraintDesc = ^TConstraintDesc;

  TConstraints = record
    data: array of TConstraintDesc;
    procedure Clear;
    function Find(ConstraintName: UniString): PConstraintDesc;
    procedure Add(ConstraintName: UniString);
  end;

procedure TConstraints.Clear;
begin
  SetLength(Data, 0);
end;

function TConstraints.Find(ConstraintName: UniString): PConstraintDesc;
var i: integer;
begin
  Result := nil;
  for i := 0 to Length(data) - 1 do
    if WideSameText(ConstraintName, data[i].Name) then begin
      Result := @data[i];
      break;
    end;
end;

procedure TConstraints.Add(ConstraintName: UniString);
var Constraint: PConstraintDesc;
begin
  Constraint := Find(ConstraintName);
  if Constraint <> nil then exit;
  SetLength(data, Length(data)+1);
  data[Length(data)-1].Name := ConstraintName;
end;
{$ENDREGION}

//Returns CONSTRAINT list for a given table.
function GetTableIndexes(conn: _Connection; Table: UniString): TIndexes;
var rs: _Recordset;
  Index: PIndexDesc;
  IndexName: UniString;
  Column: PIndexColumnDesc;
  Constraints: TConstraints;
begin
 //First we read table constraints to filter out those indexes which are constraint-related
  rs := conn.OpenSchema(adSchemaTableConstraints,
    VarArrayOf([Unassigned, Unassigned, Unassigned, Unassigned,
      Unassigned, Table, 'FOREIGN KEY']), EmptyParam);
  Constraints.Clear;
  while not rs.EOF do begin
    Constraints.Add(str(rs.Fields['CONSTRAINT_NAME'].Value));
    rs.MoveNext();
  end;

 //Indexes
  rs := conn.OpenSchema(adSchemaIndexes,
    VarArrayOf([Unassigned, Unassigned, Unassigned, Unassigned, Table]), EmptyParam);

 //Reading data
  Result.Clear;
  while not rs.EOF do begin
    IndexName := str(rs.Fields['INDEX_NAME'].Value);

   //Ignore constraint-related indexes
    if Constraints.Find(IndexName)<>nil then begin
      rs.MoveNext;
      continue;
    end;

    Index := Result.Get(IndexName);
    if not Index._Initialized then begin
     //Supposedly these values should be the same for all the records belonging to the same index
      Index.PrimaryKey := rs.Fields['PRIMARY_KEY'].Value;
      Index.Unique := rs.Fields['UNIQUE'].Value;
      Index.Nulls := int(rs.Fields['NULLS'].Value);
      Index._Initialized := true;
    end;

    Column := Index.AddColumn(str(rs.Fields['COLUMN_NAME'].Value));
    if Column<>nil then begin //wasn't defined yet -- the usual case
      Column.Collation := int(rs.Fields['COLLATION'].Value);
      Column.OrdinalPosition := int(rs.Fields['ORDINAL_POSITION'].Value);
    end;
    
    rs.MoveNext;
  end;
end;

//Dumps index creation commands for a given table
procedure DumpIndexes(conn: _Connection; TableName: string);
var Indexes: TIndexes;
  Index: PIndexDesc;
  i, j: integer;
  s, fl, tmp: UniString;
  Multiline: boolean;
begin
  Indexes := GetTableIndexes(conn, TableName);
  for i := 0 to Length(Indexes.data) - 1 do begin
    Index := @Indexes.data[i];
    Index.SortColumns;

    if Index.Unique then
      s := 'CREATE UNIQUE INDEX ['
    else
      s := 'CREATE INDEX [';
    s := s + Index.Name + '] ON [' + TableName + ']';

    fl := '';
    Multiline := false;
    for j := 0 to Length(Index.Columns) - 1 do begin
      tmp := Index.Columns[j].Name;
      case Index.Columns[j].Collation of
        DB_COLLATION_ASC: tmp := tmp + ' ASC';
        DB_COLLATION_DESC: tmp := tmp + ' DESC';
      end;
      if fl<>'' then begin
        fl := fl + ','#13#10 + tmp;
        Multiline := true;
      end else
        fl := tmp;
    end;

    if Multiline then
      s := s + ' ('#13#10 + fl + #13#10 + ')'
    else
      s := s + ' (' + fl + ')';

    if Index.PrimaryKey or (Index.Nulls=DBPROPVAL_IN_DISALLOWNULL)
    or (Index.Nulls=DBPROPVAL_IN_IGNORENULL) then begin
      s := s + ' WITH';
      if Index.PrimaryKey then
        s := s + ' PRIMARY';
      if Index.Nulls=DBPROPVAL_IN_DISALLOWNULL then
        s := s + ' DISALLOW NULL'
      else
      if Index.Nulls=DBPROPVAL_IN_IGNORENULL then
        s := s + ' IGNORE NULL';
    end;

    s := s + ';';
    writeln(s);
  end;
end;

{$REGION 'ForeignKeys'}
type
  TForeignColumn = record
    PrimaryKey: UniString;
    ForeignKey: UniString;
    Ordinal: integer;
  end;
  PForeignColumn = ^TForeignColumn;

  TForeignKey = record
    Name: UniString;
    PrimaryTable: UniString;
    Columns: array of TForeignColumn;
    PkName: OleVariant;
    OnUpdate: UniString;
    OnDelete: UniString;
    _Initialized: boolean;
    procedure SortColumns;
    function AddColumn: PForeignColumn;
  end;
  PForeignKey = ^TForeignKey;

  TForeignKeys = record
    data: array of TForeignKey;
    procedure Clear;
    function Find(ForeignKeyName: UniString): PForeignKey;
    function Get(ForeignKeyName: UniString): PForeignKey;
  end;

procedure TForeignKey.SortColumns;
var i, j, k: integer;
  tmp: TForeignColumn;
begin
  for i := 1 to Length(Columns) - 1 do begin
    j := i-1;
    while (j >= 0) and (Columns[i].Ordinal < Columns[j].Ordinal) do
      Dec(j);
    Inc(j);
    if j<>i then begin
      tmp := Columns[i];
      for k := i downto j+1 do
        Columns[k] := Columns[k-1];
      Columns[j] := tmp;
    end;
  end;
end;

function TForeignKey.AddColumn: PForeignColumn;
begin
  SetLength(Columns, Length(Columns)+1);
  Result := @Columns[Length(Columns)-1];
end;

procedure TForeignKeys.Clear;
begin
  SetLength(data, 0);
end;

function TForeignKeys.Find(ForeignKeyName: UniString): PForeignKey;
var i: integer;
begin
  Result := nil;
  for i := 0 to Length(data) - 1 do
    if WideSameText(ForeignKeyName, data[i].Name) then begin
      Result := @data[i];
      break;
    end;
end;

function TForeignKeys.Get(ForeignKeyName: UniString): PForeignKey;
begin
  Result := Find(ForeignKeyName);
  if Result <> nil then exit;
  SetLength(data, Length(data)+1);
  Result := @data[Length(data)-1];
  Result.Name := ForeignKeyName;
  Result._Initialized := false;
end;
{$ENDREGION}

//Dumps foreign key creation commands for a given table
procedure DumpForeignKeys(conn: _Connection; TableName: UniString);
var rs: _Recordset;
  ForKeys: TForeignKeys;
  ForKey: PForeignKey;
  s, s_for, s_ref: UniString;
  i, j: integer;
begin
  rs := conn.OpenSchema(adSchemaForeignKeys,
    VarArrayOf([Unassigned, Unassigned, Unassigned,
      Unassigned, Unassigned, TableName]), EmptyParam);
  ForKeys.Clear;
  while not rs.EOF do begin
    ForKey := ForKeys.Get(str(rs.Fields['FK_NAME'].Value));
    if not ForKey._Initialized then begin
      ForKey.PkName := str(rs.Fields['PK_NAME'].Value);
      ForKey.PrimaryTable := str(rs.Fields['PK_TABLE_NAME'].Value);
      ForKey.OnUpdate := str(rs.Fields['UPDATE_RULE'].Value);
      ForKey.OnDelete := str(rs.Fields['DELETE_RULE'].Value);

     //These are used internally when no action is defined.
     //Maybe they would have worked in CONSTRAINT too, but let's follow the standard.
      if WideSameText(ForKey.OnUpdate, 'NO ACTION') then
        ForKey.OnUpdate := '';
      if WideSameText(ForKey.OnDelete, 'NO ACTION') then
        ForKey.OnDelete := '';

      ForKey._Initialized := true;
    end;

    with ForKey.AddColumn^ do begin
      PrimaryKey := str(rs.Fields['PK_COLUMN_NAME'].Value);
      ForeignKey := str(rs.Fields['FK_COLUMN_NAME'].Value);
      Ordinal := int(rs.Fields['ORDINAL'].Value);
    end;

    rs.MoveNext;
  end;

  if Length(ForKeys.data)>0 then
    writeln('/* Foreign keys for '+TableName+' */');

  for i := 0 to Length(ForKeys.data) - 1 do begin
    ForKey := @ForKeys.Data[i];
    ForKey.SortColumns;

    s := 'ALTER TABLE ['+TableName+'] ADD CONSTRAINT ['+ForKey.Name+'] FOREIGN KEY (';

    if Length(ForKey.Columns)<=0 then begin
      Warning('Foreign key '+ForKey.Name+' has a definition but no columns. '
        +'This is pretty damn strange, mail a detailed bug report please.');
    end;

    s_for := ForKey.Columns[0].ForeignKey;
    s_ref := ForKey.Columns[0].PrimaryKey;
    for j := 1 to Length(ForKey.Columns) - 1 do begin
      s_for := s_for + ', ' + ForKey.Columns[j].ForeignKey;
      s_ref := s_ref + ', ' + ForKey.Columns[j].PrimaryKey;
    end;

    s := s + s_for + ') REFERENCES ['+ForKey.PrimaryTable+'] (' + s_ref + ')';
    if ForKey.OnUpdate<>'' then
      s := s  + ' ON UPDATE '+ForKey.OnUpdate;
    if ForKey.OnDelete<>'' then
      s := s  + ' ON DELETE '+ForKey.OnDelete;

    s := s + ';';

   //If PK_NAME=Null, FK is relation-only. These cannot be created by DDL
   //at this point, and are pointless anyway (no check).
   //Output them as a comment.
    if VarIsNil(ForKey.PkName) or (ForKey.PkName='') then begin
      s := '/* '+s+' */'#13#10
        +'/* Relation-only Foreign Key commented out: cannot be defined with SQL. */';
    end;

    writeln(s);
  end;

  if Length(ForKeys.data)>0 then
    writeln('');
end;

(*
//Dumps check constraint creation commands for a given table
procedure DumpCheckConstraint(conn: _Connection; TableName: UniString);
var rs: _Recordset;
  ForKeys: TForeignKeys;
  ForKey: PForeignKey;
  s, s_for, s_ref: UniString;
  i, j: integer;
begin
  rs := conn.OpenSchema(adSchemaCheckConstraints,
    VarArrayOf([Unassigned, Unassigned, Unassigned,
      Unassigned, Unassigned, TableName]), EmptyParam);
  ForKeys.Clear;
  while not rs.EOF do begin
    ForKey := ForKeys.Get(str(rs.Fields['FK_NAME'].Value));
    if not ForKey._Initialized then begin
      ForKey.PkName := str(rs.Fields['PK_NAME'].Value);
      ForKey.PrimaryTable := str(rs.Fields['PK_TABLE_NAME'].Value);
      ForKey.OnUpdate := str(rs.Fields['UPDATE_RULE'].Value);
      ForKey.OnDelete := str(rs.Fields['DELETE_RULE'].Value);

     //These are used internally when no action is defined.
     //Maybe they would have worked in CONSTRAINT too, but let's follow the standard.
      if WideSameText(ForKey.OnUpdate, 'NO ACTION') then
        ForKey.OnUpdate := '';
      if WideSameText(ForKey.OnDelete, 'NO ACTION') then
        ForKey.OnDelete := '';

      ForKey._Initialized := true;
    end;

    with ForKey.AddColumn^ do begin
      PrimaryKey := str(rs.Fields['PK_COLUMN_NAME'].Value);
      ForeignKey := str(rs.Fields['FK_COLUMN_NAME'].Value);
      Ordinal := int(rs.Fields['ORDINAL'].Value);
    end;

    rs.MoveNext;
  end;

  if Length(ForKeys.data)>0 then
    writeln('/* Foreign keys for '+TableName+' */');

  for i := 0 to Length(ForKeys.data) - 1 do begin
    ForKey := @ForKeys.Data[i];
    ForKey.SortColumns;

    s := 'ALTER TABLE ['+TableName+'] ADD CONSTRAINT ['+ForKey.Name+'] FOREIGN KEY (';

    if Length(ForKey.Columns)<=0 then begin
      Warning('Foreign key '+ForKey.Name+' has a definition but no columns. '
        +'This is pretty damn strange, mail a detailed bug report please.');
    end;

    s_for := ForKey.Columns[0].ForeignKey;
    s_ref := ForKey.Columns[0].PrimaryKey;
    for j := 1 to Length(ForKey.Columns) - 1 do begin
      s_for := s_for + ', ' + ForKey.Columns[j].ForeignKey;
      s_ref := s_ref + ', ' + ForKey.Columns[j].PrimaryKey;
    end;

    s := s + s_for + ') REFERENCES ['+ForKey.PrimaryTable+'] (' + s_ref + ')';
    if ForKey.OnUpdate<>'' then
      s := s  + ' ON UPDATE '+ForKey.OnUpdate;
    if ForKey.OnDelete<>'' then
      s := s  + ' ON DELETE '+ForKey.OnDelete;

    s := s + ';';

   //If PK_NAME=Null, FK is relation-only. These cannot be created by DDL
   //at this point, and are pointless anyway (no check).
   //Output them as a comment.
    if VarIsNil(ForKey.PkName) or (ForKey.PkName='') then begin
      s := '/* '+s+' */'#13#10
        +'/* Relation-only Foreign Key commented out: cannot be defined with SQL. */';
    end;

    writeln(s);
  end;

  if Length(ForKeys.data)>0 then
    writeln('');
end;
*)

//Lowercases a database ID only if IDs are configured as case-insensitive
function LowercaseID(const AId: UniString): UniString;
begin
  if CaseInsensitiveIDs then
    Result := ToLowercase(AId)
  else
    Result := AId;
end;

//Dumps table creation commands, then foreign keys and check constraints (if needed)
procedure DumpTables(conn: _Connection);
var rs: _Recordset;
  TableName: UniString;
  Description: UniString;
begin
  rs := conn.OpenSchema(adSchemaTables,
    VarArrayOf([Unassigned, Unassigned, Unassigned, 'TABLE']), EmptyParam);
  while not rs.EOF do begin
    TableName := str(rs.Fields['TABLE_NAME'].Value);
    Description := str(rs.Fields['DESCRIPTION'].Value);

    if (Length(DumpTableList) > 0) and not Contains(DumpTableList, LowercaseID(TableName)) then begin
      rs.MoveNext;
      continue;
    end;

    if DropObjects then
      if PrivateExtensions then
        writeln('DROP TABLE ['+TableName+'] /**WEAK**/;')
      else
        writeln('DROP TABLE ['+TableName+'];');
    writeln('CREATE TABLE ['+TableName+'] (');
    writeln(GetTableText(conn, TableName));
    if HandleComments and (Description<>'') then
      if PrivateExtensions then
        writeln(') /**COMMENT* '+EncodeComment(Description)+'*/;')
      else
        writeln(') /* '+EncodeComment(Description)+' */')
    else
      writeln(');');
    DumpIndexes(conn, TableName);
    writeln('');
    rs.MoveNext();
  end;

 //One more time, with foreign keys
  if NeedDumpRelations and not rs.BOF then begin
    rs.MoveFirst();
    while not rs.EOF do begin
      TableName := str(rs.Fields['TABLE_NAME'].Value);
      DumpForeignKeys(conn, TableName);
      rs.MoveNext();
    end;
  end;

 //TODO: One more time, with check constraints
 (*
  if NeedDumpCheckConstraints and not rs.BOF then begin
    rs.MoveFirst();
    while not rs.EOF do begin
      TableName := str(rs.Fields['TABLE_NAME'].Value);
      DumpCheckConstraints(conn, TableName);
      rs.MoveNext();
    end;
  end;
 *)
end;

procedure DumpViews(conn: _Connection);
var rs: _Recordset;
  TableName: UniString;
  Description: UniString;
  Definition: UniString;
begin
  rs := conn.OpenSchema(adSchemaViews, EmptyParam, EmptyParam);
  while not rs.EOF do begin
    TableName := str(rs.Fields['TABLE_NAME'].Value);
    Description := str(rs.Fields['DESCRIPTION'].Value);

    if (Length(DumpViewList) > 0) and not Contains(DumpViewList, LowercaseID(TableName)) then begin
      rs.MoveNext;
      continue;
    end;

  (*
    Раньше использовалось CREATE VIEW, но хотя суть та же,
    Access не поддерживает в нём сложные запросы
    (даже если в базе они объявлены, как VIEW).
  *)
    if DropObjects then
      if PrivateExtensions then
        writeln('DROP PROCEDURE ['+TableName+'] /**WEAK**/;')
      else
        writeln('DROP PROCEDURE ['+TableName+'];');
    writeln('CREATE PROCEDURE ['+TableName+'] AS');

   //Access seems to keep it's own ';' at the end of DEFINITION
    Definition := Trim(str(rs.Fields['VIEW_DEFINITION'].Value));
    if (Length(Definition)>0) and (Definition[Length(Definition)]=';') then
      SetLength(Definition, Length(Definition)-1);

    if HandleComments and (Description <> '') then begin
      writeln(Definition);
      if PrivateExtensions then
        writeln('/**COMMENT* '+EncodeComment(Description)+' */;')
      else
        writeln('/* '+EncodeComment(Description)+' */;')
    end else
      writeln(Definition+';');
    writeln('');

    rs.MoveNext();
  end;
end;


procedure DumpProcedures(conn: _Connection);
var rs: _Recordset;
  ProcedureName: UniString;
  Description: UniString;
  Definition: UniString;
begin
  rs := conn.OpenSchema(adSchemaProcedures, EmptyParam, EmptyParam);
  while not rs.EOF do begin
    ProcedureName := str(rs.Fields['PROCEDURE_NAME'].Value);
    Description := str(rs.Fields['DESCRIPTION'].Value);

    if DropObjects then
      if PrivateExtensions then
        writeln('DROP PROCEDURE ['+ProcedureName+'] /**WEAK**/;')
      else
        writeln('DROP PROCEDURE ['+ProcedureName+'];');
    writeln('CREATE PROCEDURE ['+ProcedureName+'] AS');

   //Access seems to keep it's own ';' at the end of DEFINITION
    Definition := Trim(str(rs.Fields['PROCEDURE_DEFINITION'].Value));
    if (Length(Definition)>0) and (Definition[Length(Definition)]=';') then
      SetLength(Definition, Length(Definition)-1);

    if HandleComments and (Description <> '') then begin
      writeln(Definition);
      if PrivateExtensions then
        writeln('/**COMMENT* '+EncodeComment(Description)+' */;')
      else
        writeln('/* '+EncodeComment(Description)+' */;')
    end else
      writeln(Definition+';');
    writeln('');

    rs.MoveNext();
  end;
end;

procedure DumpSql();
var conn: _Connection;
begin
  conn := GetAdoConnection;
  writeln('/* Access SQL export data follows. Auto-generated. */'#13#10);

 //Tables
  if NeedDumpTables then begin
    writeln('/* Tables */');
    DumpTables(conn);
  end;

 //Views
  if NeedDumpViews then begin
    writeln('/* Views */');
    DumpViews(conn);
  end;

 //Procedures
  if NeedDumpProcedures then begin
    writeln('/* Procedures */');
    DumpProcedures(conn); //not implemented yet
  end;

  writeln('/* Access SQL export data end. */');
end;

////////////////////////////////////////////////////////////////////////////////
///  ExecSql --- Executes SQL from Input

//Outputs a warning which should only be visible if we're not in silent mode.
procedure Complain(msg: UniString);
begin
  if LoggingMode<>lmSilent then
    err(msg);
end;

//Reads next command from input stream, joining consequtive lines when needed.
//Saves the rest of the line. Minds CrlfBreak setting.

type
  TCommentState = (
    csNone,
    csBrace,  // {
    csSlash,  // /*
    csLine,   // ends with newline (e.g. -- comment)
    csQuote,      // '
    csBacktick,   // `
    csDoubleQuote // "
  );

var read_buf: UniString;
function ReadNextCmd(out cmd: UniString): boolean;
var pc: PWideChar;
  Comment: TCommentState;
  ts: UniString;
  ind: integer;

  procedure appendStr(pc: PWideChar; shift:integer);
  begin
    ts := Trim(Copy(read_buf, 1, charPos(@read_buf[1], pc)+shift));
    if ts <> '' then
      if cmd <> '' then
        cmd := cmd + #13#10 + ts
      else
        cmd := ts;
  end;

  procedure appendWholeStr();
  begin
    ts := Trim(read_buf);
    if ts <> '' then
      if cmd <> '' then
        cmd := cmd + #13#10 + ts
      else
        cmd := ts;
  end;

  procedure restartStr(var pc: PWideChar; shift: integer);
  begin
    ind := charPos(@read_buf[1], pc)+shift;
    read_buf := Copy(read_buf, ind, Length(read_buf) - ind + 1);
    pc := PWideChar(integer(@read_buf[1])-SizeOf(WideChar));
  end;

begin
  cmd := '';
  Result := false;

  Comment := csNone;
  while not Result and ((read_buf<>'') or not Eof) do begin
   //Ends with newline
    if Comment=csLine then
      Comment := csNone;

   //Read next part if finished with previous
    if read_buf='' then
      readln(read_buf);

   //Parse
    if read_buf<>'' then begin //could've read another empty string
      pc := @read_buf[1];
      while pc^ <> #00 do begin
       //Comment closers
        if (pc^='}') and (Comment=csBrace) then
          Comment := csNone
        else
        if (pc^='/') and prevCharIs(@read_buf[1], pc, '*') and (Comment=csSlash) then
          Comment := csNone
        else
        if (pc^='''') and (Comment=csQuote) then
          Comment := csNone
        else
        if (pc^='`') and (Comment=csBacktick) then
          Comment := csNone
        else
        if (pc^='"') and (Comment=csDoubleQuote) then
          Comment := csNone
        else
       //Comment openers
        if (pc^='{') and (Comment=csNone) then
          Comment := csBrace
        else
        if (pc^='*') and prevCharIs(@read_buf[1], pc, '/') and (Comment=csNone) then
          Comment := csSlash
        else
        if (pc^='-') and prevCharIs(@read_buf[1], pc, '-') and (Comment=csNone) then
          Comment := csLine
        else
        if (pc^='''') and (Comment=csNone) then
          Comment := csQuote
        else
        if (pc^='`') and (Comment=csNone) then
          Comment := csBacktick
        else
        if (pc^='"') and (Comment=csNone) then
          Comment := csDoubleQuote
        else
       //Command is over, return (save the rest of the line for later)
        if (pc^=';') and (Comment=csNone) then begin
          appendStr(pc, -1);
          restartStr(pc, +2);
          Result := true;
          break;
        end;
        Inc(pc);
      end;
    end; //if read_buf <> ''

   //No ';' in this line => append the bufer and zero it
    if not Result then begin
      appendWholeStr;
      read_buf:='';

     //If we're in CRLF Break mode, exit (command is over), else continue
      if CrlfBreak=tbTrue then break;
    end;
  end;

  if not Result then //read to EOF without finding ';'
    Result := cmd <> '';
end;

procedure daoSetOrAdd(Dao: Database; Props: DAO_TLB.Properties; Name, Value: UniString);
var Prop: DAO_TLB.Property_;
begin
  try
    Props.Item[Name].Value := Value;
  except
    on E:EOleException do begin
      if (E.ErrorCode=HRESULT($800A0CC6)) {Property not found} then begin
        Prop := dao.CreateProperty(Name, dbText, Value, False);
        Props.Append(Prop);
      end;
    end;
  end;
end;

procedure jetSetTableComment(TableName: UniString; Comment: UniString);
var dao: Database;
  td: TableDef;
begin
  if LoggingMode=lmVerbose then
    writeln('Table '+TableName+' comment '+Comment);
  dao := EstablishDaoConnection;
  dao.TableDefs.Refresh;
  td := dao.TableDefs[TableName];
  daoSetOrAdd(dao, td.Properties, 'Description', Comment);
end;

procedure jetSetFieldComment(TableName, FieldName: UniString; Comment: UniString);
 var dao: Database;
   td: TableDef;
begin
  if LoggingMode=lmVerbose then
    writeln('Table '+TableName+' field '+FieldName+' comment '+Comment);
  dao := EstablishDaoConnection;
  dao.TableDefs.Refresh;
  td := dao.TableDefs[TableName];
  daoSetOrAdd(dao, td.Fields[FieldName].Properties, 'Description', Comment);
end;

procedure jetSetProcedureComment(ProcedureName: UniString; Comment: UniString);
var dao: Database;
  td: QueryDef;
begin
  if LoggingMode=lmVerbose then
    writeln('Procedure '+ProcedureName+' comment '+Comment);
  dao := EstablishDaoConnection;
  dao.QueryDefs.Refresh;
  td := dao.QueryDefs[ProcedureName];
  daoSetOrAdd(dao, td.Properties, 'Description', Comment);
end;

procedure PrintRecordset(rs: _Recordset);
var i: integer;
begin
  writeln('');
  while not rs.EOF do begin
    for i := 0 to rs.Fields.Count - 1 do
      writeln(str(rs.Fields[i].Name)+': '+str(rs.Fields[i].Value));
    writeln('');
    rs.MoveNext;
  end;
end;

procedure ExecCmd(conn: _Connection; cmd: UniString);
var RecordsAffected: OleVariant;
  Weak: boolean;
  Data: TUniStringArray;
  Fields: TUniStringArray;
  strComment: UniString;
  TableName, FieldName: UniString;
  i: integer;
  tr_cmd: UniString;
  rs: _Recordset;
begin
  if LoggingMode=lmVerbose then begin
    writeln('');
    writeln('Executing: ');
    writeln(cmd);
  end;

  Weak := false;
  if PrivateExtensions then
    Weak := MetaPresent(cmd, 'WEAK');

  try
    tr_cmd := Trim(RemoveComments(cmd));
    if tr_cmd='' then exit; //nothing to do

    rs := conn.Execute(tr_cmd, RecordsAffected, 0);
    if (rs<>nil) and (rs.State=adStateOpen) and (LoggingMode=lmVerbose) then
      PrintRecordset(rs);
  except
    on E: EOleException do
    if not Weak then begin
      if LoggingMode<>lmSilent then begin
        err('');
        err('Error while executing: ');
        err(tr_cmd);
      end;

      if (Errors=emIgnore) and (LoggingMode<>lmSilent) then
        err(E.Classname + ': ' + E.Message + '(0x' + IntToHex(E.ErrorCode, 8) + ')');
      if Errors<>emIgnore then
        raise; //Re-raise it to be caught in Main()
      exit; //else just exit
    end;
    on E: Exception do begin //Unrecognized exception, preferences are ignored - STOP.
      err('');
      err('Error while executing: ');
      err(tr_cmd);
      err(E.Classname + ': ' + E.Message);
      raise;
    end;
  end;

 //Do private extension processing  
  if PrivateExtensions and HandleComments then begin
    Data := Match(cmd, ['CREATE', 'TABLE', '(', ')']);
    if (Length(Data)=5) and (Trim(Data[1])='' {between CREATE and TABLE}) then begin
      TableName := CutIdBrackets(RemoveComments(Data[2]));
      if GetMeta(Data[4], 'COMMENT', strComment) then
        jetSetTableComment(TableName, DecodeComment(strComment));

     //Field comments
      Fields := Split(Data[3], ',');
      for i := 0 to Length(Fields)-1 do
        if GetMeta(Fields[i], 'COMMENT', strComment) then begin
          FieldName := CutIdBrackets(FieldNameFromDefinition(RemoveComments(Fields[i])));
          if FieldName='' then
            Complain('Cannot decode field name for field definition "'+Fields[i]+'". '
              +'Comment will not be added to the database.')
          else
            jetSetFieldComment(TableName, FieldName, DecodeComment(strComment));
        end;
    end;

   (* VIEW - это во всех отношениях процедуры *)
    Data := Match(cmd, ['CREATE', 'VIEW', 'AS']);
    if (Length(Data)=4) and (Trim(Data[1])='' {between CREATE and VIEW}) then begin
      TableName := CutIdBrackets(RemoveComments(Data[2]));
      if GetMeta(Data[3], 'COMMENT', strComment) then
        jetSetProcedureComment(TableName, DecodeComment(strComment));
    end;

    Data := Match(cmd, ['CREATE', 'PROCEDURE', 'AS']);
    if (Length(Data)=4) and (Trim(Data[1])='' {between CREATE and PROCEDURE}) then begin
      TableName := CutIdBrackets(RemoveComments(Data[2]));
      if GetMeta(Data[3], 'COMMENT', strComment) then
        jetSetProcedureComment(TableName, DecodeComment(strComment));
    end;
  end; //of PrivateExtensions

end;

procedure ExecSql();
var conn: _Connection;
  cmd: UniString;
begin
  conn := GetAdoConnection;
  while ReadNextCmd(cmd) do
    ExecCmd(conn, cmd);
end;

////////////////////////////////////////////////////////////////////////////////

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

//Returns true when a specified STD_HANDLE actually points to a console object.
//It's a LUCKY / UNKNOWN type situation: if it does, we're lucky and assume
//keyboard input. If it doesn't, we don't know: it might be a file or a pipe
//to another console.
function IsConsoleHandle(stdHandle: cardinal): boolean;
begin
 //Failing GetFileType/GetStdHandle is fine, we'll return false.
  Result := (GetFileType(GetStdHandle(stdHandle))=FILE_TYPE_CHAR);
end;



begin
  try
    CoInitializeEx(nil, COINIT_MULTITHREADED);
    ParseCommandLine;

   //Redirects
    if stdi<>'' then
      RedirectIo(STD_INPUT_HANDLE, stdi);
    if stdo<>'' then
      RedirectIo(STD_OUTPUT_HANDLE, stdo);
    if stde<>'' then
      RedirectIo(STD_OUTPUT_HANDLE, stde);

    if WideSameText(Command, 'touch') then
      Touch()
    else
    if WideSameText(Command, 'dump') then
      DumpSql()
    else
    if WideSameText(Command, 'exec') then
      ExecSql()
    else
    if WideSameText(Command, 'help') then
      PrintUsage()
    else
    if WideSameText(Command, 'schema') then
      PrintSchema()
    else
    if WideSameText(Command, 'daoschema') then
      PrintDaoSchema()
    else
    if WideSameText(Command, 'adoxschema') then
      PrintAdoxSchema()
    else
    if Command='' then
      BadUsage('No command specified')
    else
      BadUsage('Unsupported command: '+Command);
    ClearOleObjects(); //or else it'll stay till finalization when Ole is long gone.
   //Also it's paramount that we nil it inside of the function: Delphi will not
   //actually derefcount it until we exit the scope of where we nil it.
    CoUninitialize();
  except
    on E:EUsage do begin
      if E.Message <> '' then
        err(E.Message);
      PrintShortUsage;
      ExitCode := 2;
    end;
    on E:EOleException do begin
      err(E.Classname + ': ' + E.Message + '(0x' + IntToHex(E.ErrorCode, 8) + ')');
      ExitCode := 3;
    end;
    on E:Exception do begin
      err(E.Classname+': '+E.Message);
      ExitCode := 1;
    end;
  end;
end.
