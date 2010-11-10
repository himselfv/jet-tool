unit AdoxDumper;

(*
  Dedicated to dumping ADOX structures.

  Unlike DAO, ADOX does not store all individual fields in Properties.
  We need to output both Properties and those individual fields which are not covered. 
*)

interface
uses SysUtils, ComObj, ADOX_TLB, JetCommon;

procedure PrintAdoxSchema(adox: Catalog);

implementation

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

//Many ADOX objects have Properties
//  Prop1=Value1
//  Prop2=Value2 (inherited)
procedure DumpAdoxProperties(Props: Properties; Banned: TPropertyNames);
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
   // AddPropFlag('attr='+IntToStr(prop.Attributes));
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


procedure PrintAdoxColumn(f: Column);
begin
  writeln('Name: ', f.Name);
  writeln('Attributes: ', f.Attributes);
  writeln('DefinedSize: ', f.DefinedSize);
  writeln('NumericScale: ', f.NumericScale);
  writeln('Precision: ', f.Precision);
  writeln('RelatedColumn: ', f.RelatedColumn);
  writeln('SortOrder: ', f.SortOrder);
  writeln('Type: ', f.type_);
  DumpAdoxProperties(f.Properties, PropNames([]));
end;

procedure PrintAdoxIndex(f: Index);
var i: integer;
begin
  writeln('Name: ', f.Name);
  writeln('IndexNulls: ', f.IndexNulls);
  DumpAdoxProperties(f.Properties, PropNames([]));
  writeln('');

  writeln('Columns');
  for i := 0 to f.Columns.Count - 1 do try
    PrintAdoxColumn(f.Columns[i]);
    writeln('');
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;
end;

procedure PrintAdoxKey(f: Key);
var i: integer;
begin
 (* [no Properties] *)
  writeln('Name: ', f.Name);
  writeln('DeleteRule: ', f.DeleteRule);
  writeln('Type: ', f.Type_);
  writeln('RelatedTable: ', f.RelatedTable);
  writeln('UpdateRule: ', f.UpdateRule);
  writeln('');

  writeln('Columns');
  for i := 0 to f.Columns.Count - 1 do try
    PrintAdoxColumn(f.Columns[i]);
    writeln('');
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;
end;


procedure PrintAdoxTable(f: Table);
var i: integer;
begin
  writeln('Name: ', f.Name);
  writeln('Type: ', f.type_);
  DumpAdoxProperties(f.Properties, PropNames([]));
  writeln('');

  Subsection('Columns');
  for i := 0 to f.Columns.Count - 1 do try
    PrintAdoxColumn(f.Columns[i]);
    writeln('');
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;

  Subsection('Indexes');
  for i := 0 to f.Indexes.Count - 1 do try
    PrintAdoxIndex(f.Indexes[i]);
    writeln('');
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;

  Subsection('Keys');
  for i := 0 to f.Keys.Count - 1 do try
    PrintAdoxKey(f.Keys[i]);
    writeln('');
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;
end;

procedure PrintAdoxProcedure(f: Procedure_);
begin
 (* [no Properties] *)
  writeln('Name: ', f.Name);
  writeln('DateCreated: ', str(f.DateCreated));
  writeln('DateModified: ', str(f.DateModified));
end;

procedure PrintAdoxView(f: View);
begin
 (* [no Properties] *)
  writeln('Name: ', f.Name);
  writeln('DateCreated: ', str(f.DateCreated));
  writeln('DateModified: ', str(f.DateModified));
end;

procedure PrintAdoxSchema(adox: Catalog);
var i: integer;
begin
  for i := 0 to adox.Tables.Count-1 do try
    Section('Table: '+adox.Tables[i].Name);
    PrintAdoxTable(adox.Tables[i]);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;

  for i := 0 to adox.Procedures.Count-1 do try
    Section('Procedure: '+adox.Procedures[i].Name);
    PrintAdoxProcedure(adox.Procedures[i]);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;

  for i := 0 to adox.Views.Count-1 do try
    Section('View: '+adox.Views[i].Name);
    PrintAdoxView(adox.Views[i]);
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;

  (*
    Not dumping:
      Groups
      Users
  *)
end;

end.
