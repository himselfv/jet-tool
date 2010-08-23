program Jet;

{$APPTYPE CONSOLE}

uses
  SysUtils, Windows, ActiveX, Variants, AdoDb, OleDb, AdoInt, ComObj, WideStrUtils,
  DAO_TLB,
  StringUtils in 'StringUtils.pas';

(*
  Supported propietary extension comments:
    /**COMMENT* [comment text] */  - table, field or view comment
    /**WEAK**/ - ignore errors for this command (handy for DROP TABLE/VIEW)

  By default when reading from keyboard:
    --ignore-errors
    --crlf-break
    --verbose

  Return codes:
    -1 :: Generic error
    -2 :: Usage error
    -3 :: OLE Error
*)

{$UNDEF DEBUG}

type
  EUsage = class(Exception);

procedure PrintShortUsage;
begin
  writeln('Do "jet help" for extended info.');
end;

procedure PrintUsage;
begin
  writeln('Usage:');
  writeln('  jet <command> [params]');
  writeln('');
  writeln('Commands:');
  writeln('  jet touch :: connect to database and quit');
  writeln('  jet dump :: dump sql schema');
  writeln('  jet exec :: parse sql from input');
  writeln('  jet schema :: output internal jet schema reports');
  writeln('');
  writeln('Connection params:');
  writeln('  -c [connection-string] :: uses an ADO connection string. -dp is ignored');
  writeln('  -dsn [data-source-name] :: uses an ODBC data source name');
  writeln('  -f [file.mdb] :: opens a jet database file');
  writeln('  -u [user]');
  writeln('  -p [password]');
  writeln('  -dp [database password]'); {Works fine with database creation too}
  writeln('  -new :: works only with exec and filename');
  writeln('  -force :: overwrite existing database (requires -new)');
  writeln('You cannot use -c with --comments when executing (dumping is fine).');
  writeln('You can only use -new with -f.');
 (* -dsn will probably not work with --comments too, as long as it really is MS Access DSN. They deny DAO DSN connections. *)
  writeln('');
  writeln('Useful IO tricks:');
  writeln('  -stdi [filename] :: sets standard input');
  writeln('  -stdo [filename] :: sets standard output');
  writeln('  -stde [filename] :: sets standard error console');
  writeln('This is only applied after the command-line parsing is over');
  writeln('');
  writeln('What to include and whatnot for dumping:');
  writeln('  --no-tables, --tables');
  writeln('  --no-views, --views');
  writeln('  --no-procedures, --procedures');
  writeln('  --no-comments, --comments :: how comments are dumped depends on if private extensions are enabled');
  writeln('  --no-drop, --drop :: "DROP" tables etc before creating');
  writeln('');
  writeln('Works both for dumping and executing:');
  writeln('  --no-private-extensions, --private-extensions :: disables dumping/parsing private extensions (check help)');
  writeln('');
  writeln('What to do with errors when executing:');
  writeln('  --silent :: do not print anything (at all)');
  writeln('  --verbose :: echo commands which are being executed');
  writeln('  --ignore-errors :: continue on error');
  writeln('  --stop-on-errors :: exit with error code');
  writeln('  --crlf-break :: CR/LF ends command');
  writeln('  --no-crlf-break');
  writeln('With private extensions enabled, **WEAK** commands do not produce errors in any way (messages, stop).')
end;

procedure BadUsage(msg: string='');
begin
  raise EUsage.Create(msg);
end;

procedure Redefined(term, old, new: string);
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
  Command: WideString;
 //Connection
  ConnectionString: WideString;
  DataSourceName: WideString;
  Filename: WideString;
  User, Password: WideString;
  DatabasePassword: WideString;
  NewDb: boolean;
  ForceNewDb: boolean;
 //Dump contents
  NeedDumpTables: boolean = true;
  NeedDumpViews: boolean = true;
  NeedDumpProcedures: boolean = false; //TODO: enable by default when ready
  HandleComments: boolean = true;
  PrivateExtensions: boolean = true;
  DropObjects: boolean = true;
 //Log
  LoggingMode: TLoggingMode = lmDefault;
  Errors: TErrorHandlingMode = emDefault;
 //Redirects
  stdi, stdo, stde: WideString;
  CrlfBreak: TTriBool = tbDefault;


procedure ParseCommandLine;
var i: integer;
  s: WideString;
  KeyboardInput: boolean;

  procedure Define(var Term: WideString; TermName: string; Value: WideString);
  begin
    if Term <> '' then
      Redefined(TermName, Term, Value);
    Term := Value;
  end;

  function NextParam(key, param: WideString): WideString;
  begin
    Inc(i);
    if i > ParamCount then
      BadUsage(Key+' requires specifying the '+param);
    Result := ParamStr(i);
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

    if WideSameText(s, '--tables') then begin
      NeedDumpTables := true;
    end else
    if WideSameText(s, '--no-tables') then begin
      NeedDumpTables := false;
    end else
    if WideSameText(s, '--views') then begin
      NeedDumpViews := true;
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

 //Если спросили help, ничего не проверяем
  if WideSameText(Command, 'help') then exit;

 //Проверяем параметры. Допустимо только одно: ConnectionString, ConnectionName, Filename
  i := 0;
  if ConnectionString<>'' then Inc(i);
  if DataSourceName<>'' then Inc(i);
  if Filename<>'' then Inc(i);
  if i > 1 then BadUsage('Only one source (ConnectionString/DataSourceName/Filename) can be specified.');
  if i < 1 then BadUsage('A source (ConnectionString/DataSourceName/Filename) needs to be specified.');

 //Если подключение - ConnectionString, дополнительные параметры не допускаются: всё включено
  if (ConnectionString<>'') and (DatabasePassword<>'') then
    BadUsage('ConnectionString conflicts with DatabasePassword: this should be included inside of the connection string.');

 //Если требуется создать базу, допустимо только подключение по имени файла
  if NewDb and (Filename='') then
    BadUsage('Database creation is supported only when connecting by Filename.');
  if NewDb and (WideSameText(Command, 'dump') or WideSameText(Command, 'schema'))
  and (LoggingMode=lmVerbose) then begin
    writeln('NOTE: You asked to create database and then dump its contents.');
    writeln('What the hell are you trying to do?');
  end;
  if ForceNewDb and not NewDb then
    BadUsage('-force requires -new');

  if NewDb and not ForceNewDb and FileExists(Filename) then
    raise Exception.Create('File '+Filename+' already exists. Cannot create a new one.');

 //Чтобы парсить комменты, требуется DAO, т.е. подключение по имени файла
  if WideSameText(Command, 'exec') and HandleComments then begin
    if not PrivateExtensions then
      BadUsage('You need --private-extensions to handle --comments when doing "exec".');
    if (Filename='') and (DataSourceName='') then
      BadUsage('You cannot use ConnectionString source with --comments when doing "exec"');
  end;

 //Преобразуем все виды источников к ConnectionString
  if Filename<>'' then
    ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source="'+Filename+'";'
  else
  if DataSourceName <> '' then
    ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source="'+DataSourceName+'";';

  if DatabasePassword<>'' then
    ConnectionString := ConnectionString + 'Jet OLEDB:Database Password="'+DatabasePassword+'";';

 //Ресольвим дефолтные значения в зависимости от типа входного потока.
 //Если тип не удаётся угадать, по умолчанию подразумеваем файл (всегда можно указать опции!)
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

function VarIsNil(val: OleVariant): boolean;
begin
  Result := VarIsClear(val) or VarIsNull(val);
end;

function str(val: OleVariant): string;
begin
  if VarIsNil(val) then
    Result := ''
  else Result := string(val);
end;

function int(val: OleVariant): integer;
begin
  if VarIsNil(val) then
    Result := 0
  else Result := integer(val);
end;

function bool(val: OleVariant): boolean;
begin
  if VarIsNil(val) then
    Result := false
  else Result := boolean(val);
end;

function includes(main, flag: cardinal): boolean;
begin
  Result := (main and flag) = flag;
end;

const
  Jet10 = 1;
  Jet11 = 2;
  Jet20 = 3;
  Jet3x = 4;
  Jet4x = 5;

procedure CreateNewDatabase;
var Catalog: OleVariant;
begin
  if ForceNewDb and FileExists(Filename) then
    DeleteFileW(PWideChar(Filename));
  Catalog := CreateOleObject('ADOX.Catalog');
  Catalog.Create(ConnectionString
    + 'Jet OLEDB:Engine Type='+IntToStr(Jet4x)+';');
  if LoggingMode=lmVerbose then
    writeln('Database created.');
  NewDb := false; //works only once
end;

function EstablishConnection: _Connection;
begin
  if NewDb then CreateNewDatabase;
  Result := CoConnection.Create;
  Result.Open(ConnectionString, User, Password, 0)
end;

var
  DaoConnection: Database;

function EstablishDaoConnection: Database;
var DbEngine: _DbEngine;
  Params: OleVariant;
begin
//If you enable this, Dao refreshing will break for GOD KNOWS WHAT REASONS
//  if Assigned(DaoConnection) then begin
//    Result := DaoConnection;
//    exit;
//  end;

  DbEngine := CoDbEngine.Create;
 //Do not disable, or Dao refreshing will break too
  DbEngine.Idle(dbRefreshCache);
  if Filename<>'' then begin
    if DatabasePassword<>'' then
      Params := WideString('MS Access;pwd=')+DatabasePassword
    else
      Params := '';
    Result := DbEngine.OpenDatabase(Filename, False, False, Params);
  end else
  if DataSourceName<>'' then begin
   //Although this will probably not work
    Params := 'ODBC;DSN='+DataSourceName+';UID='+User+';PWD='+Password+';';
    Result := DbEngine.OpenDatabase('', False, False, Params);
  end else
    raise Exception.Create('The operation you''re performing apparently requires '
      +'DAO. DAO and DAO-dependent functions can only be accessed through Filename '
      +'source. That you see this error must mean somehow this condition was not '
      +'properly checked during command-line parsing. '#13
      +'Please file a bug to developers. As for yourself, try to guess which '
      +'setting required DAO (usually something obscure like importing comments) '
      +'and either disable that or switch to connecting through Filename.');

  DaoConnection := Result;
end;

procedure ClearDao;
begin
  DaoConnection := nil;
end;

////////////////////////////////////////////////////////////////////////////////
/// Touch --- establishes connection and quits

procedure Touch();
var conn: _Connection;
begin
  conn := EstablishConnection;
end;

////////////////////////////////////////////////////////////////////////////////
//// Schema --- Dumps many schema tables returned by Access

procedure DumpFields(conn: _Connection; Schema: integer);
var rs: _Recordset;
  i: integer;
begin
  rs := conn.OpenSchema(Schema, EmptyParam, EmptyParam);
  while not rs.EOF do begin
    for i := 0 to rs.Fields.Count - 1 do
      writeln(str(rs.Fields[i].Name)+'='+str(rs.Fields[i].Value));
    writeln('');
    rs.MoveNext();
  end;
end;

procedure Section(SectionName: string);
begin
  writeln('');
  writeln('');
  writeln(SectionName);
  writeln('=========================================');
end;

procedure Dump(conn: _Connection; SectionName: string; Schema: integer);
begin
  Section(SectionName);
  DumpFields(conn, Schema);
end;

procedure PrintSchema();
var conn: _Connection;
begin
  conn := EstablishConnection;
  Dump(conn, 'Tables', adSchemaTables);
  Dump(conn, 'Columns', adSchemaColumns);

  Dump(conn, 'Check constaints', adSchemaCheckConstraints);
  Dump(conn, 'Referential constaints', adSchemaReferentialConstraints);
  Dump(conn, 'Table constaints', adSchemaTableConstraints);

  Dump(conn, 'Column usage', adSchemaConstraintColumnUsage);
  Dump(conn, 'Key column usage', adSchemaKeyColumnUsage);
//  Dump(conn, 'Table usage', adSchemaConstraintTableUsage); --- not supported by Access

  Dump(conn, 'Indexes', adSchemaIndexes);
  Dump(conn, 'Foreign keys', adSchemaForeignKeys);

//  Dump(conn, 'Asserts', adSchemaAsserts); ---not supported by Access
end;

////////////////////////////////////////////////////////////////////////////////
/// DumpSql --- Dumps database contents

procedure Warning(msg: WideString);
begin
  writeln('/* !!! Warning: '+msg+' */');
end;

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

function GetTableText(conn: _Connection; Table: WideString): WideString;
var rs: _Recordset;
  Columns: TColumns;
  Column: TColumnDesc;
  s, dts: string;
  tmp: OleVariant;
  pre, scal: OleVariant;
begin
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
    Columns.Add(Column);
    rs.MoveNext;
  end;
  Columns.Sort;


 //Building string
  Result := '';
  for Column in Columns.data do begin

   //Data type
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

      DBTYPE_BYTES: dts := 'BINARY';
      DBTYPE_WSTR:
       //If you specify TEXT, Access makes LONGTEXT (Memo) field, if TEXT(len) then the usual limited text.
        if Includes(Column.Flags, DBCOLUMNFLAGS_ISLONG) then
          dts := 'TEXT' //"Memo field"
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
      s := s + ' DEFAULT '+str(Column.Default);

   //Access does not support comments, therefore we just output them in our propietary format.
   //If you use this importer, it'll understand them.
    if HandleComments and (Column.Description <> '') then
      if PrivateExtensions then
        s := s + ' /**COMMENT* '+Column.Description+' */'
      else
        s := s + ' /* '+Column.Description+' */';

    if Result <> '' then
      Result := Result + ','#13#10+s
    else
      Result := s;
  end;
end;

procedure DumpTables(conn: _Connection);
var rs: _Recordset;
  TableName: WideString;
  Description: WideString;
begin
  rs := conn.OpenSchema(adSchemaTables,
    VarArrayOf([Unassigned, Unassigned, Unassigned, 'TABLE']), EmptyParam);
  while not rs.EOF do begin
    TableName := str(rs.Fields['TABLE_NAME'].Value);
    Description := str(rs.Fields['DESCRIPTION'].Value);

    if DropObjects then
      if PrivateExtensions then
        writeln('DROP TABLE ['+TableName+'] /**WEAK**/;')
      else
        writeln('DROP TABLE ['+TableName+'];');    
    writeln('CREATE TABLE ['+TableName+'] (');
    writeln(GetTableText(conn, TableName));
    if HandleComments and (Description<>'') then
      if PrivateExtensions then
        writeln(') /**COMMENT* '+Description+'*/;')
      else
        writeln(') /* '+Description+' */')
    else
      writeln(');');
    writeln('');
    rs.MoveNext();
  end;
end;

procedure DumpViews(conn: _Connection);
var rs: _Recordset;
  TableName: WideString;
  Description: WideString;
  Definition: WideString;
begin
  rs := conn.OpenSchema(adSchemaViews, EmptyParam, EmptyParam);
  while not rs.EOF do begin
    TableName := str(rs.Fields['TABLE_NAME'].Value);
    Description := str(rs.Fields['DESCRIPTION'].Value);

    if DropObjects then
      if PrivateExtensions then
        writeln('DROP VIEW ['+TableName+'] /**WEAK**/;')
      else
        writeln('DROP VIEW ['+TableName+'];');
    writeln('CREATE VIEW ['+TableName+'] AS');

   //Access seems to keep it's own ';' at the end of DEFINITION
    Definition := Trim(str(rs.Fields['VIEW_DEFINITION'].Value));
    if (Length(Definition)>0) and (Definition[Length(Definition)]=';') then
      SetLength(Definition, Length(Definition)-1);

    if HandleComments and (Description <> '') then begin
      writeln(Definition);
      if PrivateExtensions then
        writeln('/**COMMENT* '+Description+' */;')
      else
        writeln('/* '+Description+' */;')
    end else
      writeln(Definition+';');
    writeln('');

    rs.MoveNext();
  end;
end;

procedure DumpSql();
var conn: _Connection;
begin
  conn := EstablishConnection;
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
   // DumpProcedures(conn); //not writen yet
  end;

  writeln('/* Access SQL export data end. */');
end;

////////////////////////////////////////////////////////////////////////////////
///  ExecSql --- Executes SQL from Input

//Читает следующую команду из входного потока, соединяя при необходимости подряд идущие строки.
//Сохраняет остаток строки. Учитывает настройку CrlfBreak.
type
  TCommentState = (csNone, csBrace, csSlash, csLine);

var read_buf: WideString;
function ReadNextCmd(out cmd: WideString): boolean;
var pc: PWideChar;
  Comment: TCommentState;
  ts: WideString;
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
    pc := @read_buf[1];
    while pc^ <> #00 do begin
     //Comment closers
      if (pc^='}') and (Comment=csBrace) then
        Comment := csNone
      else
      if (pc^='/') and prevCharIs(@read_buf[1], pc, '*') and (Comment=csSlash) then
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
     //Command is over, return (save the rest of the line for later)
      if pc^=';' then begin
        appendStr(pc, -1);
        restartStr(pc, +2);
        Result := true;
        break;
      end;
      Inc(pc);
    end;

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

procedure Complain(msg: WideString);
begin
  if LoggingMode<>lmSilent then
    writeln(msg);
end;

procedure daoSetOrAdd(Dao: Database; Props: Properties; Name, Value: WideString);
var Prop: Property_;
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

procedure jetSetTableComment(TableName: WideString; Comment: WideString);
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

procedure jetSetFieldComment(TableName, FieldName: WideString; Comment: WideString);
 var dao: Database;
   td: TableDef;
begin
  if LoggingMode=lmVerbose then
    writeln('Table '+TableName+' field '+FieldName+' comment '+Comment);
  dao := EstablishDaoConnection;
  dao.TableDefs.Refresh;
  td := dao.TableDefs[TableName];
  daoSetOrAdd(dao, td.Fields[FieldName].Properties, 'Description', 'Comment');
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

procedure ExecCmd(conn: _Connection; cmd: WideString);
var RecordsAffected: OleVariant;
  Weak: boolean;
  Data: TWideStringArray;
  Fields: TWideStringArray;
  strComment: WideString;
  TableName, FieldName: WideString;
  i: integer;
  tr_cmd: WideString;
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
        writeln('');
        writeln('Error while executing: ');
        writeln(tr_cmd);
      end;

      if (Errors=emIgnore) and (LoggingMode<>lmSilent) then
        writeln(E.Classname + ': ' + E.Message + '(0x' + IntToHex(E.ErrorCode, 8) + ')');
      if Errors<>emIgnore then
        raise; //Re-raise it to be caught in Main()
      exit; //else just exit
    end;
    on E: Exception do begin //Unrecognized exception, preferences are ignored - STOP.
      writeln('');
      writeln('Error while executing: ');
      writeln(tr_cmd);
      writeln(E.Classname + ': ' + E.Message);
      raise;
    end;
  end;

 //Do private extension processing  
  if PrivateExtensions and HandleComments then begin
    Data := Match(cmd, ['CREATE', 'TABLE', '(', ')']);
    if (Length(Data)=5) and (Trim(Data[1])='' {between CREATE and TABLE}) then begin
      TableName := CutIdBrackets(RemoveComments(Data[2]));
      if GetMeta(Data[4], 'COMMENT', strComment) then
        jetSetTableComment(TableName, strComment);

     //Field comments
      Fields := Split(Data[3], ',');
      for i := 0 to Length(Fields)-1 do
        if GetMeta(Fields[i], 'COMMENT', strComment) then begin
          FieldName := CutIdBrackets(FieldNameFromDefinition(RemoveComments(Fields[i])));
          if FieldName='' then
            Complain('Cannot decode field name for field definition "'+Fields[i]+'". '
              +'Comment will not be added to the database.')
          else
            jetSetFieldComment(TableName, FieldName, strComment);
        end;
    end;

    Data := Match(cmd, ['CREATE', 'VIEW', 'AS']);
    if (Length(Data)=5) and (Trim(Data[1])='' {between CREATE and VIEW}) then begin
      TableName := CutIdBrackets(RemoveComments(Data[2]));
      if GetMeta(Data[3], 'COMMENT', strComment) then
        jetSetTableComment(TableName, strComment);
    end;
  end; //of PrivateExtensions
 
end;

procedure ExecSql();
var conn: _Connection;
  cmd: WideString;
begin
  conn := EstablishConnection;
  while ReadNextCmd(cmd) do
    ExecCmd(conn, cmd);
end;

////////////////////////////////////////////////////////////////////////////////

//Leaks file handles! By design. We don't care, they'll be released on exit anyway.
procedure RedirectIo(nStdHandle: dword; filename: WideString);
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
    if Command='' then
      BadUsage('No command specified')
    else
      BadUsage('Unsupported command: '+Command);
    ClearDao(); //or else it'll stay till finalization when Ole is long gone.
   //Also it's paramount that we nil it inside of the function: Delphi will not
   //actually derefcount it until we exit the scope of where we nil it.
    CoUninitialize();
  except
    on E:EUsage do begin
      if E.Message <> '' then
        writeln(E.Message);
      PrintShortUsage;
      ExitCode := -2;
    end;
    on E:EOleException do begin
      writeln(E.Classname + ': ' + E.Message + '(0x' + IntToHex(E.ErrorCode, 8) + ')');
      ExitCode := -3;
    end;
    on E:Exception do begin
      Writeln(E.Classname, ': ', E.Message);
      ExitCode := -1;
    end;
  end;
end.
