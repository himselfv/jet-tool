unit DaoDumper;

(*
  Dedicated to dumping DAO structures.

  Every object which has "Properties" can be queried through these.
  They contain all the individual fields and more. 
*)

interface
uses SysUtils, ComObj, DAO_TLB, JetCommon;

procedure PrintDaoSchema(dao: Database);

implementation

//Prints the contents of DAO recordset (all records and fields)
procedure DumpDaoRecordset(rs: Recordset);
var i: integer;
begin
  while not rs.EOF do begin
    for i := 0 to rs.Fields.Count - 1 do
      writeln(rs.Fields[i].Name+'='+str(rs.Fields[i].Value));
    writeln('');
    rs.MoveNext();
  end;
end;


procedure PrintDaoTable(dao: Database; TableName: string);
begin
  SubSection('Table: '+TableName);
  DumpDaoRecordset(dao.ListFields(TableName));
end;

procedure PrintDaoTables(dao: Database);
var rs: Recordset;
begin
  rs := dao.ListTables;
  DumpDaoRecordset(rs);

  rs.MoveFirst;
  while not rs.EOF do begin
    try
      PrintDaoTable(dao, rs.Fields['Name'].Value);
    except
     //Sometimes we don't have sufficient rights
      on E: EOleException do
        writeln(E.Classname + ': '+ E.Message);
    end;
    rs.MoveNext;
  end;
end;


type
  TPropertyNames = array of WideString;

//Handy function for inplace array initialization
function PropNames(Names: array of WideString): TPropertyNames;
var i: integer;
begin
  SetLength(Result, Length(Names));
  for i := 0 to Length(Names) - 1 do
    Result[i] := Names[i];
end;

function IsBanned(Banned: TPropertyNames; PropName: WideString): boolean;
var i: integer;
begin
  Result := false;
  for i := 0 to Length(Banned) - 1 do
    if WideSameText(Banned[i], PropName) then begin
      Result := true;
      exit;
    end;
end;

//Many DAO objects have Properties
//  Prop1=Value1
//  Prop2=Value2 (inherited)
procedure DumpDaoProperties(Props: Properties; Banned: TPropertyNames);
var i: integer;
  prop: Property_;
  proptype: integer;
  propname: WideString;
  propval: WideString;
  propflags: WideString;

  procedure AddPropFlag(flag: WideString);
  begin
    if propflags='' then
      propflags := flag
    else
      propflags := ', ' + flag;
  end;

begin
  for i := 0 to Props.Count-1 do begin
    prop := Props[i];
    propname := prop.Name;
    proptype := prop.type_;

   //Some properties are unsupported in some objects, or take too long to query
    if IsBanned(Banned, PropName) then begin
     // writeln(PropName+'=[skipping]');
      continue;
    end;

    propflags := '';
   // AddPropFlag('type='+IntToStr(proptype)); //we mostly don't care about types
    if prop.Inherited_ then
      AddPropFlag('inherited');
    if propflags <> '' then
      propflags := ' (' + propflags + ')';

    try
      if proptype=0 then
        propval := 'Unsupported type'
      else
        propval := str(prop.Value);
    except
      on E: EOleException do begin
        propval := E.Classname + ': ' + E.Message;
      end;
    end;

    writeln(
      propname + '=',
      propval,
      propflags
    );
  end;
end;


type
  TFieldKind = (fkRecordset, fkDefinition, fkParameter, fkRelation);

procedure PrintDaoField(f: _Field; FieldKind: TFieldKind);
var Banned: TPropertyNames;
begin
  case FieldKind of
    fkRecordset: Banned := PropNames(['ForeignName', 'Relation']);
    fkDefinition: Banned := PropNames(['Value', 'ValidateOnSet', 'ForeignName', 'FieldSize', 'OriginalValue', 'VisibleValue']);
    fkParameter: Banned := PropNames(['Value', 'ValidateOnSet', 'ForeignName', 'FieldSize', 'OriginalValue', 'VisibleValue']);
    fkRelation: Banned := PropNames(['Value', 'ValidateOnSet', 'FieldSize', 'OriginalValue', 'VisibleValue'])
  else
    Banned := nil;
  end;

  DumpDaoProperties(f.Properties, Banned);
end;

//Many DAO objects have Fields collections
procedure PrintDaoFields(f: Fields; FieldKind: TFieldKind);
var i: integer;
begin
  for i := 0 to f.Count - 1 do begin
    PrintDaoField(f[i], FieldKind);
    writeln('');
  end;
end;


procedure PrintDaoIndex(f: Index);
begin
  DumpDaoProperties(f.Properties, PropNames([]));
end;

procedure PrintDaoIndexes(f: Indexes);
var i: integer;
begin
  for i := 0 to f.Count - 1 do begin
    PrintDaoIndex(f[i]);
    writeln('');
  end;
end;


procedure PrintDaoParameter(f: Parameter);
begin
  writeln(
    f.Name + ': ',
    f.type_,
    ' = ' + str(f.Value) + ' (dir:',
    f.Direction,
    ')'
  );

  writeln('Properties:');
  DumpDaoProperties(f.Properties, PropNames([]));
end;

procedure PrintDaoParameters(f: Parameters);
var i: integer;
begin
  for i := 0 to f.Count - 1 do begin
    PrintDaoParameter(f[i]);
    writeln('');
  end;
end;


////////////////////////////////////////////////////////////////////////////////

procedure PrintDaoTableDef(def: TableDef);
begin
  DumpDaoProperties(def.Properties,
    PropNames(['ConflictTable', 'ReplicaFilter', 'Connect']));
  writeln('');

  Subsection('Fields: ');
  try
    PrintDaoFields(def.Fields, fkDefinition);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;

  Subsection('Indexes:');
  try
    PrintDaoIndexes(def.Indexes);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;
end;

procedure PrintDaoQueryDef(def: QueryDef);
begin
  DumpDaoProperties(def.Properties,
    PropNames(['StillExecuting', 'CacheSize', 'Prepare']));
  writeln('');

  Subsection('Fields:');
  try
    PrintDaoFields(def.Fields, fkParameter);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;

  Subsection('ListParameters:');
  try
    DumpDaoRecordset(def.ListParameters);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;

  Subsection('Parameters:');
  try
    PrintDaoParameters(def.Parameters);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;
end;

procedure PrintDaoRelation(def: Relation);
begin
  DumpDaoProperties(def.Properties, PropNames([]));
  writeln('');

  Subsection('Fields:');
  try
    PrintDaoFields(def.Fields, fkRelation);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure PrintDaoSchema(dao: Database);
var i: integer;
begin
  Section('Table Properties');
  DumpDaoProperties(dao.Properties,
    PropNames(['ReplicaID', 'DesignMasterID']));

  Section('Tables (basic definitions)');
  PrintDaoTables(dao);

  for i := 0 to dao.TableDefs.Count-1 do try
    Section('TableDef: '+dao.TableDefs[i].Name);
    PrintDaoTableDef(dao.TableDefs[i]);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;

  for i := 0 to dao.QueryDefs.Count-1 do try
    Section('QueryDef: '+dao.QueryDefs[i].Name);
    PrintDaoQueryDef(dao.QueryDefs[i]);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;

  for i := 0 to dao.Relations.Count-1 do try
    Section('Relation: '+dao.Relations[i].Name);
    PrintDaoRelation(dao.Relations[i]);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;
end;


end.
