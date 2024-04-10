unit Jet.Connection;
{
Database connection settings and connection procedures.
}

interface
uses StringUtils, AdoDb, OleDb, ADOX_TLB, DAO_TLB, Jet.CommandLine;

type
 //Multiple database formats are supported by JET/ACE providers.
  TDatabaseFormat = (
    dbfDefault,         //auto-select from file name, default to MDB4.0
    dbfMdb10,           //various older versions
    dbfMdb11,
    dbfMdb20,
    dbfMdb3x,
    dbfMdb4x,           //latest MDB
    dbfAccdb            //ACCDB
  );

//Application settings. Mostly configured via command line.
var
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
  DatabaseFormat: TDatabaseFormat = dbfDefault;

var //Dynamic properties
  CanUseDao: boolean; //sometimes there are other ways

type
  TConnectionSettingsParser = class(TCommandLineParser)
  public
    procedure PrintUsage; override;
    function HandleOption(ctx: PParsingContext; const s: UniString): boolean; override;
    procedure Finalize; override;
  end;

var
  ConnectionSettings: TConnectionSettingsParser = nil;


function GetAdoConnection: _Connection;
function EstablishAdoConnection: _Connection;

function GetAdoxCatalog: Catalog;

function GetDaoConnection: Database;
function EstablishDaoConnection: Database;

procedure ClearOleObjects;


implementation
uses SysUtils, Windows, ActiveX, ComObj, AdoInt, Jet.IO;

procedure AutodetectOleDbProvider; forward;

procedure TConnectionSettingsParser.PrintUsage;
begin
  err('Connection params:');
  err('  -f [file.mdb] :: open a jet database file (preferred)');
  err('  -dsn [data-source-name] :: use an ODBC data source name');
  err('  -c [connection-string] :: use an ADO connection string (least preferred, overrides many settings)');
  err('  -u [user]');
  err('  -p [password]');
  err('  -dp [database password]'); {Works fine with database creation too}
  err('  -new :: create a new database (works only by file name)');
  err('  -force :: overwrite existing database (requires -new)');
  err('You cannot use -c with --comments when executing (dumping is fine).');
 (* -dsn will probably not work with --comments too, as long as it really is MS Access DSN. They deny DAO DSN connections. *)
  err('');
  err('Database format:');
  err('  --mdb :: use Jet 4.0 .mdb format (default)');
  err('  --accdb :: use .accdb format');
  err('  --db-format [jet10 / jet11 / jet20 / jet3x / jet4x (mdb) / ace12 (accdb)]');
  err('By default the tool guesses by the file name (assumes jet4x MDBs unless the extension is accdb).');
  err('');
  err('Jet/ACE OLEDB and DAO have several versions which are available on different platforms.');
  err('You can override the default selection (best compatible available):');
  err('  --oledb-eng [ProgID] :: e.g. Microsoft.Jet.OLEDB.4.0');
  err('  --dao-eng [ProgID] :: e.g. DAO.Engine.36');
  err('');
end;

function TConnectionSettingsParser.HandleOption(ctx: PParsingContext; const s: UniString): boolean;
var s1: UniString;
begin
  Result := true;
  if WideSameText(s, '-c') then begin
    Define(ConnectionString, 'Connection string', ctx.NextParam(s, 'connection string'));
  end else
  if WideSameText(s, '-dsn') then begin
    Define(DataSourceName, 'Data source name', ctx.NextParam(s, 'data source name'));
  end else
  if WideSameText(s, '-f') then begin
    Define(Filename, 'Filename', ctx.NextParam(s, 'filename'));
  end else
  if WideSameText(s, '-u') then begin
    Define(User, 'Username', ctx.NextParam(s, 'username'));
  end else
  if WideSameText(s, '-p') then begin
    Define(Password, 'Password', ctx.NextParam(s, 'password'));
  end else
  if WideSameText(s, '-dp') then begin
    Define(DatabasePassword, 'Database password', ctx.NextParam(s, 'database password'));
  end else

  if WideSameText(s, '-new') then begin
    NewDb := true;
  end else
  if WideSameText(s, '-force') then begin
    ForceNewDb := true;
  end else

 //Database provider options
  if WideSameText(s, '--oledb-eng') then begin
    Define(Providers.OleDbEng, 'OLEDB Engine', ctx.NextParam(s, 'OLEDB Engine'));
  end else
  if WideSameText(s, '--dao-eng') then begin
    Define(Providers.DaoEng, 'DAO Engine', ctx.NextParam(s, 'DAO Engine'));
  end else

 //Database format
  if WideSameText(s, '--db-format')
  or WideSameText(s, '--database-format') then begin
    s1 := ctx.NextParam(s, 'format name');
    if WideSameText(s1, 'jet10') then
      DatabaseFormat := dbfMdb10
    else
    if WideSameText(s1, 'jet11') then
      DatabaseFormat := dbfMdb11
    else
    if WideSameText(s1, 'jet20') then
      DatabaseFormat := dbfMdb20
    else
    if WideSameText(s1, 'jet3x') then
      DatabaseFormat := dbfMdb3x
    else
    if WideSameText(s1, 'jet4x') then
      DatabaseFormat := dbfMdb4x
    else
    if WideSameText(s1, 'ace12') then
      DatabaseFormat := dbfAccdb
    else
      BadUsage('Unsupported database format: '+s1);
  end else
 //Some shortcuts
  if WideSameText(s, '--as-mdb')
  or WideSameText(s, '--mdb') then
    DatabaseFormat := dbfMdb4x
  else
  if WideSameText(s, '--as-accdb')
  or WideSameText(s, '--accdb') then
    DatabaseFormat := dbfAccdb
  else
    Result := false;
end;

procedure TConnectionSettingsParser.Finalize;
var i: integer;
begin
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
  if ForceNewDb and not NewDb then
    BadUsage('-force requires -new');

  if NewDb and not ForceNewDb and FileExists(Filename) then
    raise Exception.Create('File '+Filename+' already exists. Use -force with -new to overwrite.');

 //Whether we can use DAO. If not, prefer other options.
  CanUseDao := (Filename<>'');

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
end;



function CLSIDFromProgID(const ProgID: UniString; out clsid: TGUID): boolean;
var hr: HRESULT;
begin
  hr := ActiveX.CLSIDFromProgID(PChar(ProgID), clsid);
  Result := SUCCEEDED(hr);
  if not Result then
    verbose('Trying class '+ProgID+'... not found.');
end;

//Automatically detects which supported OLEDB Jet providers are available and chooses one.
//Called if the user has not specified a provider explicitly.
procedure AutodetectOleDbProvider;
var clsid: TGUID;
const
  sOleDbProviderJet4: string = 'Microsoft.Jet.OLEDB.4.0';
  sOleDbProviderAce12: string = 'Microsoft.ACE.OLEDB.12.0';
begin
(*
Different providers support different sets of database formats:
  Jet 4.0 supports Jet11-Jet4x (MDB), but not Ace12 (ACCDB)
  ACE 12  supports Ace12 (ACCDB)

ACE also supports Jet11-Jet4x, but it's complicated:
* some features reportedly work differently, notably some field types and user/password support.
* ACE14+ deprecates Jet11-Jet20

Jet 4.0 is almost universally available anyway, so we'll prefer it for older DB types,
and prefer ACE 12 for accdb.

Note that DAO preference needs not to follow ADO preference strictly.
*)

 //For Accdb, try ACE12 first
  if DatabaseFormat = dbfAccdb then
    if CLSIDFromProgID(sOleDbProviderAce12, clsid) then begin
      Providers.OleDbEng := sOleDbProviderAce12;
      exit;
    end;

 //Try Jet 4.0
  if CLSIDFromProgID(sOleDbProviderJet4, clsid) then begin
    Providers.OleDbEng := sOleDbProviderJet4;
    if DatabaseFormat = dbfAccdb then
      //We have found something, but it's not ACE12
      err('ERROR: ACCDB format requires Microsoft.ACE.OLEDB.12.0 provider which has not been found. The operations will likely fail.');
    exit;
  end;

 //For MDBs try ACE12 as a last resort
  if DatabaseFormat <> dbfAccdb then
    if CLSIDFromProgID(sOleDbProviderAce12, clsid) then begin
      Providers.OleDbEng := sOleDbProviderAce12;
      log('NOTE: Fallback to ACE12 for older database access may introduce some inconsistencies.');
      exit;
    end;

  err('ERROR: Jet/ACE OLEDB provider not found. The operations will likely fail.');
  //Still set the most compatible provider just in case
  Providers.OleDbEng := sOleDbProviderJet4;
end;


const
  JetEngineType_Jet10 = 1;
  JetEngineType_Jet11 = 2;
  JetEngineType_Jet20 = 3;
  JetEngineType_Jet3x = 4;
  JetEngineType_Jet4x = 5;
  JetEngineType_Ace12 = 6;  //confirmed
  //Some other known types:
  // DBASE3 = 10;
  // Xslx = 30 / 37 in some examples.

var
  AdoxCatalog: Catalog;

//Creates a new database and resets a database-creation-required flag.
procedure CreateNewDatabase;
var engType: integer;
begin
  if ForceNewDb and FileExists(Filename) then
    DeleteFileW(PWideChar(Filename));

  case DatabaseFormat of
    dbfMdb10: engType := JetEngineType_Jet10;
    dbfMdb11: engType := JetEngineType_Jet11;
    dbfMdb20: engType := JetEngineType_Jet20;
    dbfMdb3x: engType := JetEngineType_Jet3x;
    dbfMdb4x: engType := JetEngineType_Jet4x;
    dbfAccdb: engType := JetEngineType_Ace12;
  else
   //By default, create Jet4x MDB
    engType := JetEngineType_Jet4x;
  end;

  AdoxCatalog := CoCatalog.Create;
  AdoxCatalog.Create(ConnectionString
    + 'Jet OLEDB:Engine Type='+IntToStr(engType)+';');
  verbose('Database created.');
  NewDb := false; //works only once
end;

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
const
  sDaoEngine36 = 'DAO.DBEngine.36';
  sDaoEngine120 = 'DAO.DBEngine.120';
begin
 //If explicit DAO provider is set, simply convert it to CLSID.
  if Providers.DaoEng <> '' then begin
    if not CLSIDFromProgID(Providers.DaoEng, Dao.ProviderCLSID) then
     //Since its explicitly configured, we should raise
      raise Exception.Create('Cannot find DAO provider with ProgID='+Providers.DaoEng);
    Dao.SupportState := ssDetected;
    exit;
  end;

(*
  As with ADO, prefer older DAO for older database types, and newer DAO for ACCDB.
  DAO preference needs not to be strictly in sync with Jet preference:
     OLEDB engine: Jet4.0,  DAO engine: DAO120    <-- this is okay (if both can handle the file)
*)

  //For ACCDB try DAO120 first
  if DatabaseFormat = dbfAccdb then
    if CLSIDFromProgID(sDaoEngine120, Dao.ProviderCLSID) then begin
      Providers.DaoEng := sDaoEngine120;
      Dao.SupportState := ssDetected;
      exit;
    end;

  //DAO36
  if CLSIDFromProgID(sDaoEngine36, Dao.ProviderCLSID) then begin
    Providers.DaoEng := sDaoEngine36;
    Dao.SupportState := ssDetected;
    if DatabaseFormat = dbfAccdb then
      err('WARNING: ACCDB format requires DAO.DBEngine.120 provider which is not found. DAO operations will probably fail.');
    exit;
  end;

  if DatabaseFormat <> dbfAccdb then
    if CLSIDFromProgID(sDaoEngine120, Dao.ProviderCLSID) then begin
      Providers.DaoEng := sDaoEngine120;
      Dao.SupportState := ssDetected;
      log('NOTE: Fallback to DAO120 for older database access may introduce some inconsistencies.');
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


initialization
  ConnectionSettings := TConnectionSettingsParser.Create;

finalization
{$IFDEF DEBUG}
  FreeAndNil(ConnectionSettings);
{$ENDIF}

end.
