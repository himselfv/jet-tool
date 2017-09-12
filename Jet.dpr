program Jet;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  Windows,
  ActiveX,
  Variants,
  AdoDb,
  OleDb,
  AdoInt,
  ComObj,
  UniStrUtils,
  DAO_TLB,
  ADOX_TLB,
  StreamUtils,
  StringUtils in 'StringUtils.pas',
  DaoDumper in 'DaoDumper.pas',
  JetCommon in 'JetCommon.pas',
  AdoxDumper in 'AdoxDumper.pas',
  JetDataFormats in 'JetDataFormats.pas';

//{$DEFINE DUMP_CHECK_CONSTRAINTS}
//Feature under development.

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
    Procedures
    Data from tables and queries

  What can be added in the future:
    Check constraints
    Rest of the "table constraints"

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

//Writes a string to error output if verbose log is enabled.
procedure log(msg: UniString); forward;

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
  err('  jet dump :: dump sql schema / data');
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
  err('Database format is auto-guessed from file name, you can override:');
  err('  --as-accdb, --accdb');
  err('  --as-mdb, --mdb');
  err('By default this tool will try to open any, and create most modern type (accdb).');
  err('');
  err('Useful IO tricks:');
  err('  -stdi [filename] :: sets standard input');
  err('  -stdo [filename] :: sets standard output');
  err('  -stde [filename] :: sets standard error console');
  err('These are only applied after the command-line parsing is over');
  err('');
  err('What to include and whatnot for dumping:');
  err('  --tables, --no-tables');
  err('  --views, --no-views');
  err('  --procedures, --no-procedures');
  err('  --relations, --no-relations');
 {$IFDEF DUMP_CHECK_CONSTRAINTS}
  err('  --check-constraints, --no-check-constraints');
 {$ENDIF}
  err('  --data, --no-data');
  err('  --query "SQL QUERY" TableName :: data from this SQL query (each subsequent usage adds a query to the list)');
  err('If none of these are explicitly given, the default set MAY be used. If any is given, only that is exported.');
  err('Shortcuts: --all, --none, --default');
  err('');
  err('  --comments, --no-comments :: how comments are dumped depends on if private extensions are enabled');
  err('  --drop, --no-drop :: DROP tables etc before creating');
  err('  --create, --no-create :: CREATE tables etc (on by default)');
  err('  --enable-if-exists, --disable-if-exists :: enables IF EXISTS option for DROP commands (not supported by Jet)');
  err('');
  err('With --tables, --views and --data you can specify individual names:');
  err('  --tables [tablename],[tablename]');
  err('  --views [viewname],[viewname]');
  err('  --data [tablename],[tablename] :: defaults to same as --tables');
  err('Specify --case-sensitive-ids or --case-insensitive-ids if needed (default: sensitive).');
  err('');
  err('Works both for dumping and executing:');
  err('  --no-private-extensions, --private-extensions :: disables dumping/parsing private extensions (see help)');
  err('');
  err('What to do with errors when executing:');
  err('  --silent :: do not print anything (at all)');
  err('  --verbose :: echo commands which are being executed');
  err('  --ignore-errors :: continue on error');
  err('  --stop-on-errors :: exit with error code');
  err('  --crlf-break :: CR/LF ends command');
  err('  --no-crlf-break');
  err('With private extensions enabled, **WEAK** commands do not produce errors in any way (messages, stop).');
  err('');
  err('Database access providers:');
  err('Jet/ACE OLEDB and DAO providers have several versions which are available on different platforms. '
    +'Newer providers can handle newer file formats.');
  err('You can specify exact one or the best available will be chosen:');
  err('  --oledb-eng [provider name]');
  err('  --dao-eng [provider name]');
end;

procedure BadUsage(msg: UniString='');
begin
  raise EUsage.Create(msg);
end;

procedure Redefined(term: string; old, new: UniString);
begin
  raise EUsage.Create(term+' already specified: '+old+'. Cannot process command "'+new+'".');
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

 //Providers
  Providers: record
    OleDbEng: UniString;        //set by user or auto-detected
    DaoEng: UniString;          //set by user or auto-detected
  end;

 //Connection
  ConnectionString: UniString;
  DataSourceName: UniString;
  Filename: UniString;
  User, Password: UniString;
  DatabasePassword: UniString;
  NewDb: boolean;
  ForceNewDb: boolean;
  DatabaseFormat: (dbfDefault, dbfMdb, dbfAccdb) = dbfDefault;

 //Database options
  CaseInsensitiveIDs: boolean;
  Supports_IfExists: boolean; //database supports IF EXISTS syntax
 //Dump contents
  DumpDefaultSourceSet: boolean = true; //cleared if the user has explicitly said what to dump
  NeedDumpTables: boolean = false;
  DumpTableList: TUniStringArray; //empty = all
  NeedDumpViews: boolean = false;
  DumpViewList: TUniStringArray;
  NeedDumpProcedures: boolean = false;
  NeedDumpRelations: boolean = false;
 {$IFDEF DUMP_CHECK_CONSTRAINTS}
  NeedDumpCheckConstraints: boolean = false;
 {$ENDIF}
  NeedDumpData: boolean = false;
  DumpDataList: TUniStringArray; //empty = use DumpTableList
  DumpQueryList: array of record
    Query: string;
    TableName: string;
  end;
 //Dump options
  HandleComments: boolean = true;
  PrivateExtensions: boolean = true;
  DropObjects: boolean = true;    //add DROP commands
  CreateObjects: boolean = true;  //add CREATE commands
 //Log
  LoggingMode: TLoggingMode = lmDefault;
  Errors: TErrorHandlingMode = emDefault;
 //Redirects
  stdi, stdo, stde: UniString;
  CrlfBreak: TTriBool = tbDefault;

var //Dynamic properties
  CanUseDao: boolean; //sometimes there are other ways

procedure ConfigureDumpAllSources;
begin
  NeedDumpTables := true;
  SetLength(DumpTableList, 0);
  NeedDumpViews := true;
  SetLength(DumpViewList, 0);
  NeedDumpProcedures := true;
  NeedDumpRelations := true;
 {$IFDEF DUMP_CHECK_CONSTRAINTS}
  NeedDumpCheckConstraints := true;
 {$ENDIF}
  NeedDumpData := true;
  SetLength(DumpDataList, 0);
end;

procedure ConfigureDumpNoSources;
begin
  NeedDumpTables := false;
  SetLength(DumpTableList, 0);
  NeedDumpViews := false;
  SetLength(DumpViewList, 0);
  NeedDumpProcedures := false;
  NeedDumpRelations := false;
 {$IFDEF DUMP_CHECK_CONSTRAINTS}
  NeedDumpCheckConstraints := false;
 {$ENDIF}
  NeedDumpData := false;
  SetLength(DumpDataList, 0);
end;

procedure ConfigureDumpDefaultSources;
begin
  NeedDumpTables := true;
  SetLength(DumpTableList, 0);
  NeedDumpViews := true;
  SetLength(DumpViewList, 0);
  NeedDumpProcedures := true;
  NeedDumpRelations := true;
 {$IFDEF DUMP_CHECK_CONSTRAINTS}
  NeedDumpCheckConstraints := true;
 {$ENDIF}
  NeedDumpData := false;
  SetLength(DumpDataList, 0);
end;

procedure AutodetectOleDbProvider; forward;

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

   //Database provider options
    if WideSameText(s, '--oledb-eng') then begin
      Define(Providers.OleDbEng, 'OLEDB Engine', NextParam('--oledb-eng', 'OLEDB Engine'));
    end else
    if WideSameText(s, '--dao-eng') then begin
      Define(Providers.DaoEng, 'DAO Engine', NextParam('--dao-eng', 'DAO Engine'));
    end else

   //Database format
    if WideSameText(s, '--as-accdb')
    or WideSameText(s, '--accdb') then begin
      DatabaseFormat := dbfAccdb
    end else
    if WideSameText(s, '--as-mdb')
    or WideSameText(s, '--mdb') then begin
      DatabaseFormat := dbfMdb
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
      DumpDefaultSourceSet := false; //override default
      NeedDumpTables := true;
      if TryNextParam('--tables', 'table list', list) then begin
        DumpTableList := Split(list, ',');
        if Length(DumpTableList) <= 0 then
         //Empty DumpTableList internally means dump all, so just disable dump if asked for none
          NeedDumpTables := false;
      end;
    end else
    if WideSameText(s, '--no-tables') then begin
      DumpDefaultSourceSet := false;
      NeedDumpTables := false;
    end else
    if WideSameText(s, '--views') then begin
      DumpDefaultSourceSet := false; //override default
      NeedDumpViews := true;
      if TryNextParam('--views', 'view list', list) then begin
        DumpViewList := Split(list, ',');
        if Length(DumpViewList) <= 0 then
          NeedDumpViews := false;
      end;
    end else
    if WideSameText(s, '--no-views') then begin
      DumpDefaultSourceSet := false;
      NeedDumpViews := false;
    end else
    if WideSameText(s, '--procedures') then begin
      DumpDefaultSourceSet := false; //override default
      NeedDumpProcedures := true;
    end else
    if WideSameText(s, '--no-procedures') then begin
      DumpDefaultSourceSet := false;
      NeedDumpProcedures := false;
    end else
    if WideSameText(s, '--relations') then begin
      DumpDefaultSourceSet := false; //override default
      NeedDumpRelations := true;
    end else
    if WideSameText(s, '--no-relations') then begin
      DumpDefaultSourceSet := false;
      NeedDumpRelations := false;
    end else
   {$IFDEF DUMP_CHECK_CONSTRAINTS}
    if WideSameText(s, '--check-constraints') then begin
      DumpDefaultSourceSet := false; //override default
      NeedDumpCheckConstraints := true;
    end else
    if WideSameText(s, '--no-check-constraints') then begin
      DumpDefaultSourceSet := false;
      NeedDumpCheckConstraints := false;
    end else
   {$ENDIF}
    if WideSameText(s, '--data') then begin
      DumpDefaultSourceSet := false;
      NeedDumpData := true;
      if TryNextParam('--data', 'table list', list) then begin
        DumpDataList := Split(list, ',');
        if Length(DumpDataList) <= 0 then
         //Empty DumpDataList internally means dump all, so just disable dump if asked for none
          NeedDumpData := false;
      end;
    end else
    if WideSameText(s, '--no-data') then begin
      DumpDefaultSourceSet := false;
      NeedDumpData := false;
    end else
    if WideSameText(s, '--query') then begin
      DumpDefaultSourceSet := false;
      SetLength(DumpQueryList, Length(DumpQueryList)+1);
      DumpQueryList[Length(DumpQueryList)-1].Query := NextParam('--query', 'SQL query text');
      DumpQueryList[Length(DumpQueryList)-1].TableName := NextParam('--query', 'Table name');
    end else

   //Shortcuts
   //Use to explicitly set the playing field before modifying
    if WideSameText(s, '--all') then begin
      DumpDefaultSourceSet := false;
      ConfigureDumpAllSources();
    end else
    if WideSameText(s, '--none') then begin
      DumpDefaultSourceSet := false;
      ConfigureDumpNoSources();
    end else
    if WideSameText(s, '--default') then begin
      DumpDefaultSourceSet := false;
      ConfigureDumpDefaultSources(); //configure right now, so it can be overriden
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
    if WideSameText(s, '--create') then begin
      CreateObjects := true;
    end else
    if WideSameText(s, '--no-create') then begin
      CreateObjects := false;
    end else
    if WideSameText(s, '--enable-if-exists') then begin
      Supports_IfExists := true;
    end else
    if WideSameText(s, '--disable-if-exists') then begin
      Supports_IfExists := false;
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

 //Auto-enable accdb by file name (if not explicitly disable)
  if (DatabaseFormat = dbfDefault) and (ExtractFileExt(Filename) = '.accdb') then
    DatabaseFormat := dbfAccdb;
 //We can't auto-enable Accdb in other cases (connection string / data source name),
 //so force it with --accdb if it matters.

 //To build a ConnectionString we need to select an OLEDB provider.
 //We can't delay it until first use because ConnectionString is needed both by ADO and ADOX.
  if Providers.OleDbEng = '' then
    AutoDetectOleDbProvider();
 //On the other hand, DAO detection can be safely delayed until first use.

 //Convert all sources to ConnectionStrings
  if Filename<>'' then
    ConnectionString := 'Provider='+Providers.OleDbEng+';Data Source="'+Filename+'";'
  else
  if DataSourceName <> '' then
    ConnectionString := 'Provider='+Providers.OleDbEng+';Data Source="'+DataSourceName+'";';

  if DatabasePassword<>'' then
   //Thankfully the parameter name has not been changed in ACE 12.0
    ConnectionString := ConnectionString + 'Jet OLEDB:Database Password="'+DatabasePassword+'";';

 //If no modifications have been made, dump default set of stuff
  if DumpDefaultSourceSet then
    ConfigureDumpDefaultSources;

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



function CLSIDFromProgID(const ProgID: UniString; out clsid: TGUID): boolean;
var hr: HRESULT;
begin
  hr := ActiveX.CLSIDFromProgID(PChar(ProgID), clsid);
  Result := SUCCEEDED(hr);
  if not Result then
    log('Trying class '+ProgID+'... not found.');
end;

//Automatically detects which supported OLEDB Jet providers are available and chooses one.
//Called if the user has not specified a provider explicitly.
procedure AutodetectOleDbProvider;
var clsid: TGUID;
begin
  //For now prefer the most modern version.
  //There are rumors that ACE 12 works differently from Jet 4.0 on old DBs,
  //then we'll switch to "Jet for mdb, Ace for accdb" scheme.

  if CLSIDFromProgID('Microsoft.ACE.OLEDB.12.0', clsid) then begin
    Providers.OleDbEng := 'Microsoft.ACE.OLEDB.12.0';
    exit;
  end;

  if CLSIDFromProgID('Microsoft.Jet.OLEDB.4.0', clsid) then begin
    Providers.OleDbEng := 'Microsoft.Jet.OLEDB.4.0';
    if DatabaseFormat = dbfAccdb then
      err('ERROR: ACCDB format requires Microsoft.ACE.OLEDB.12.0 provider which has not been found. The operations will likely fail.');
    exit;
  end;

  err('ERROR: Jet/ACE OLEDB provider not found. The operations will likely fail.');
  //Still set the most compatible provider just in case
  Providers.OleDbEng := 'Microsoft.Jet.OLEDB.4.0';
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

//Returns a NEW ADO connection.
function EstablishAdoConnection: _Connection;
begin
  if NewDb then CreateNewDatabase;
  Result := CoConnection.Create;
  Result.Open(ConnectionString, User, Password, 0);
end;

//Returns an ADO connection. Caching is implemented
function GetAdoConnection: _Connection;
begin
  if NewDb then CreateNewDatabase;
  if AdoConnection=nil then begin
    Result := EstablishAdoConnection;
    AdoConnection := Result;
  end else
    Result := AdoConnection;
end;


var
  DaoConnection: Database = nil;
  Dao: record
    SupportState: (
      ssUntested,
      ssDetected,           //ProviderCLSID is valid
      ssUnavailable         //No DAO provider or DAO disabled by connection type
      );
    ProviderCLSID: TGUID;
  end = (SupportState: ssUntested);

//These are not used and are provided only for information. We query the registry by ProgIDs.
const
  CLASS_DAO36_DBEngine_x86 = '{00000100-0000-0010-8000-00AA006D2EA4}'; //no x64 version exists
  CLASS_DAO120_DBEngine = '{CD7791B9-43FD-42C5-AE42-8DD2811F0419}'; //both x64 and x86

//Automatically detects which supported DAO providers are available and chooses one,
//or finds the provider the user has specified.
procedure AutodetectDao;
begin
 //If explicit DAO provider is set, simply convert it to CLSID.
  if Providers.DaoEng <> '' then begin
    if not CLSIDFromProgID(Providers.DaoEng, Dao.ProviderCLSID) then
     //Since its explicitly configured, we should raise
      raise Exception.Create('Cannot find DAO provider with ProgID='+Providers.DaoEng);
    Dao.SupportState := ssDetected;
    exit;
  end;

  //If we don't have a preconfigured DAO Engine, figure out which we can use
  //For now prefer the most modern version.

  if CLSIDFromProgID('DAO.DBEngine.120', Dao.ProviderCLSID) then begin
    Providers.DaoEng := 'DAO.DBEngine.120';
    Dao.SupportState := ssDetected;
    exit;
  end;

  if CLSIDFromProgID('DAO.DBEngine.36', Dao.ProviderCLSID) then begin
    Providers.DaoEng := 'DAO.DBEngine.36';
    Dao.SupportState := ssDetected;
    if DatabaseFormat = dbfAccdb then
      err('WARNING: ACCDB format requires DAO.DBEngine.120 provider which is not found. DAO operations will probably fail.');
    exit;
  end;

  err('WARNING: No compatible DAO provider found. DAO operations will be unavailable.');
 {$IFDEF CPUX64}
  err('  Note that this X64 build of jet-tool cannot use 32-bit only DAO.DBEngine.36 even if it''s installed. '
    +'You need DAO.DBEngine.120 which is included in "Microsoft Office 12.0 Access Database Engine Objects Library" and later.');
 {$ENDIF}
  Dao.SupportState := ssUnavailable;
end;

resourcestring
  sDaoUnavailable =
    'The operation you''re trying to perform requires DAO. No supported DAO providers have been detected.'
  + 'This tool requires either DAO.DBEngine.120 or DAO.DBEngine.36.'#13
  + 'Install the required DAO providers, override DAO provider selection or disable the features that require DAO.';
  sDaoConnectionTypeWrong =
    'The operation you''re trying to perform requires DAO. DAO and DAO-dependent functions can only be '
  + 'accessed through Filename source.'#13
  + 'Somehow this condition was not properly verified during command-line parsing. '
  + 'Please file a bug to the developers.'#13
  + 'For the time being, disable the setting that required DAO (usually something obscure like importing '
  + 'comments) or switch to connecting through Filename.';


//Establishes a NEW DAO connection. Also refreshes the engine cache.
//Usually called every time DAO connection is needed by a DAO-dependent proc.
function EstablishDaoConnection: Database;
var DaoEngine: _DbEngine;
  Params: OleVariant;
begin
  if NewDb then CreateNewDatabase;

  //Figure out supported DAO engine and its CLSID
  if Dao.SupportState = ssUntested then
    AutodetectDao();
  if Dao.SupportState = ssUnavailable then
    //Since we're here, someone tried to use Dao functions anyway. Raise!
    raise Exception.Create(sDaoUnavailable);

 //Doing the same as this would, only with variable CLSID:
 //  DaoEngine := CoDbEngine.Create;
  DaoEngine := CreateComObject(Dao.ProviderCLSID) as DBEngine;

 //Do not disable, or Dao refreshing will break too
  DaoEngine.Idle(dbRefreshCache);

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
    raise Exception.Create(sDaoConnectionTypeWrong);
end;

//Returns a DAO connection. Caching is implemented.
function GetDaoConnection: Database;
begin
  if Assigned(DaoConnection) then begin
    Result := DaoConnection;
    exit;
  end;
  Result := EstablishDaoConnection;
  DaoConnection := Result;
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

//Lowercases a database ID only if IDs are configured as case-insensitive
function LowercaseID(const AId: UniString): UniString;
begin
  if CaseInsensitiveIDs then
    Result := ToLowercase(AId)
  else
    Result := AId;
end;

function GetDropCmdSql(const AType, ATableName: string): string;
begin
  Result := 'DROP '+AType;
  if Supports_IfExists then
    Result := Result + ' IF EXISTS';
  Result := Result + ' ['+ATableName+']';
  if PrivateExtensions then
    Result := Result + ' /**WEAK**/';
  Result := Result + ';';
end;

//Dumps table creation commands
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
      writeln(GetDropCmdSql('TABLE', TableName));

    if CreateObjects then begin
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
    end;

    rs.MoveNext();
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

function GetDropConstraintCmdSql(const ATableName, AConstraint: string): string;
begin
  Result := 'ALTER TABLE ['+ATableName+'] DROP CONSTRAINT';
  if Supports_IfExists then
    Result := Result + ' IF EXISTS';
  Result := Result + ' ['+AConstraint+']';
  if PrivateExtensions then
    Result := Result + ' /**WEAK**/';
  Result := Result + ';';
end;

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

    if DropObjects and not NeedDumpTables then //the foreign keys are dropped automatically with tables
      writeln(GetDropConstraintCmdSql(TableName, ForKey.Name));

    if CreateObjects then begin
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
    end; //of CreateObjects

  end;

  if Length(ForKeys.data)>0 then
    writeln('');
end;

//Dumps foreign keys for all tables
procedure DumpRelations(conn: _Connection);
var rs: _Recordset;
  TableName: UniString;
begin
  rs := conn.OpenSchema(adSchemaTables,
    VarArrayOf([Unassigned, Unassigned, Unassigned, 'TABLE']), EmptyParam);
  while not rs.EOF do begin
    TableName := str(rs.Fields['TABLE_NAME'].Value);
    DumpForeignKeys(conn, TableName);
    rs.MoveNext();
  end;
end;


{$IFDEF DUMP_CHECK_CONSTRAINTS}
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

    if DropObjects and not NeedDumpTables then //the foreign keys are dropped automatically with tables
      writeln(GetDropConstraintCmdSql(TableName, ForKey.Name));

    if CreateObjects then begin
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
    end; //of CreateObjects

  end;

  if Length(ForKeys.data)>0 then
    writeln('');
end;

//Dumps check constraints for all tables
procedure DumpCheckConstraints(conn: _Connection);
var rs: _Recordset;
  TableName: UniString;
begin
  rs := conn.OpenSchema(adSchemaTables,
    VarArrayOf([Unassigned, Unassigned, Unassigned, 'TABLE']), EmptyParam);
  while not rs.EOF do begin
      TableName := str(rs.Fields['TABLE_NAME'].Value);
      DumpCheckConstraints(conn, TableName);
      rs.MoveNext();
  end;
end;
{$ENDIF}


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

   //We used CREATE VIEW before which is compatible, but does not support complicated queries in Access,
   //even if the objects of type VIEW can host those in principle
    if DropObjects then
      writeln(GetDropCmdSql('PROCEDURE', TableName));

    if CreateObjects then begin
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
    end;

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
      writeln(GetDropCmdSql('PROCEDURE', ProcedureName));

    if CreateObjects then begin
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
    end;

    rs.MoveNext();
  end;
end;


//Dumps INSERT commands for the records in a recordset.
//  TableName: table to insert records into
procedure DumpInsertCommands(const TableName: string; const rs: _Recordset);
var pref, vals: string;
  i: integer;
begin
 //Generate common insertion prefix. Unfortunately, Jet does not allow multi-record insertion, so
 //we'll have to repeat it every time.
  if rs.Fields.Count <= 0 then
    pref := ''
  else begin
    pref := '[' + rs.Fields[0].Name + ']';
    for i := 1 to rs.Fields.Count-1 do
      pref := pref + ',['+rs.Fields[i].Name+']';
  end;
  pref := 'INSERT INTO ['+TableName+'] ('+pref+') VALUES ';

  while not rs.EOF do begin
    if rs.Fields.Count <= 0 then
      vals := ''
    else begin
      vals := JetEncodeValue(rs.Fields[0].Value);
      for i := 1 to rs.Fields.Count-1 do
        vals := vals + ',' + JetEncodeValue(rs.Fields[i].Value);
    end;
    writeln(pref+'('+vals+');');
    rs.MoveNext;
  end;
end;

//Dumps table contents as INSERT commands for all selected tables
procedure DumpData(conn: _Connection);
var rs: _Recordset;
  TableName: UniString;
  RecordsAffected: OleVariant;
begin
  if DumpDataList = nil then
    DumpDataList := DumpTableList; //maybe also nil

  rs := conn.OpenSchema(adSchemaTables,
    VarArrayOf([Unassigned, Unassigned, Unassigned, 'TABLE']), EmptyParam);
  while not rs.EOF do begin
    TableName := str(rs.Fields['TABLE_NAME'].Value);
    if (Length(DumpDataList) > 0) and not Contains(DumpDataList, LowercaseID(TableName)) then begin
      rs.MoveNext;
      continue;
    end;

    writeln('/* Data for table '+TableName+' */');
    DumpInsertCommands(TableName, conn.Execute('SELECT * FROM ['+TableName+']', RecordsAffected, 0));

    writeln('');
    rs.MoveNext();
  end;
end;


procedure DumpQueries(conn: _Connection);
var RecordsAffected: OleVariant;
  i: integer;
begin
  for i := 0 to Length(DumpQueryList)-1 do begin
    writeln('/* "'+EncodeComment(DumpQueryList[i].Query)+'" */');
    DumpInsertCommands(DumpQueryList[i].TableName, conn.Execute(DumpQueryList[i].Query, RecordsAffected, 0));
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

 //Relations (Foreign keys)
  if NeedDumpRelations then begin
    writeln('/* Relations */');
    DumpRelations(conn);
  end;

 {$IFDEF DUMP_CHECK_CONSTRAINTS}
 //Check constraints
  if NeedDumpCheckConstraints then begin
    writeln('/* Check constraints */')
    DumpCheckConstraints(conn);
  end;
 {$ENDIF}

 //Views
  if NeedDumpViews then begin
    writeln('/* Views */');
    DumpViews(conn);
  end;

 //Procedures
  if NeedDumpProcedures then begin
    writeln('/* Procedures */');
    DumpProcedures(conn);
  end;

 //Data
  if NeedDumpData then begin
    writeln('/* Table data */');
    DumpData(conn);
  end;

 //Additional queries
  if Length(DumpQueryList) > 0 then begin
    writeln('/* Additional queries */');
    DumpQueries(conn);
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

procedure log(msg: UniString);
begin
  if LoggingMode = lmVerbose then
    err(msg);
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
