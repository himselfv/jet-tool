unit AdoxDumper;

(*
  Dedicated to dumping ADOX structures.

  Unlike DAO, ADOX does not store all individual fields in Properties.
  We need to output both Properties and those individual fields which are not covered. 
*)

interface
uses SysUtils, ComObj, ADOX_TLB, JetCommon;

procedure PrintAdoxSchema(adox: Catalog);

function AdoxDataTypeToStr(const t: DataTypeEnum): string;
function AdoxKeyTypeToStr(const t: KeyTypeEnum): string;
function AdoxRuleTypeToStr(const t: RuleEnum): string;

implementation

function AdoxDataTypeToStr(const t: DataTypeEnum): string;
begin
  case t of
    adEmpty: Result := 'Empty';
    adTinyInt: Result := 'TinyInt';
    adSmallInt: Result := 'SmallInt';
    adInteger: Result := 'Integer';
    adBigInt: Result := 'BigInt';
    adUnsignedTinyInt: Result := 'UnsignedTinyInt';
    adUnsignedSmallInt: Result := 'UnsignedSmallInt';
    adUnsignedInt: Result := 'UnsignedInt';
    adUnsignedBigInt: Result := 'UnsignedBigInt';
    adSingle: Result := 'Single';
    adDouble: Result := 'Double';
    adCurrency: Result := 'Currency';
    adDecimal: Result := 'Decimal';
    adNumeric: Result := 'Numeric';
    adBoolean: Result := 'Boolean';
    adError: Result := 'Error';
    adUserDefined: Result := 'UserDefined';
    adVariant: Result := 'Variant';
    adIDispatch: Result := 'IDispatch';
    adIUnknown: Result := 'IUnknown';
    adGUID: Result := 'GUID';
    adDate: Result := 'Date';
    adDBDate: Result := 'DBDate';
    adDBTime: Result := 'DBTime';
    adDBTimeStamp: Result := 'DBTimeStamp';
    adBSTR: Result := 'BSTR';
    adChar: Result := 'Char';
    adVarChar: Result := 'VarChar';
    adLongVarChar: Result := 'LongVarChar';
    adWChar: Result := 'WChar';
    adVarWChar: Result := 'VarWChar';
    adLongVarWChar: Result := 'LongVarWChar';
    adBinary: Result := 'Binary';
    adVarBinary: Result := 'VarBinary';
    adLongVarBinary: Result := 'LongVarBinary';
    adChapter: Result := 'Chapter';
    adFileTime: Result := 'FileTime';
    adPropVariant: Result := 'PropVariant';
    adVarNumeric: Result := 'VarNumeric';
  else Result := '';
  end;

  if Result <> '' then
    Result := Result + ' (' + IntToStr(t) + ')'
  else
    Result := IntToStr(t);
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


{
Not all Column object properties are available in all contexts.
Index and Key columns only support:
 - Name
 - SortOrder (if Index)
 - RelatedColumn (if Key)

Table columns do not support:
 - SortOrder
 - RelatedColumn

Tested in Jet 4.0 and ACE12 from Office 2010.
}
type
  TColumnMode = (cmTable, cmIndex, cmKey);

procedure PrintAdoxColumn(f: Column; mode: TColumnMode);
begin
  writeln('Name: ', f.Name);
  if not (mode in [cmIndex, cmKey]) then begin
    writeln('Type: ', AdoxDataTypeToStr(f.type_));
    writeln('Attributes: ', f.Attributes);
    writeln('DefinedSize: ', f.DefinedSize);
    writeln('NumericScale: ', f.NumericScale);
    writeln('Precision: ', f.Precision);
  end;

  if not (mode in [cmTable, cmKey]) then
    writeln('SortOrder: ', f.SortOrder);
  if not (mode in [cmTable, cmIndex]) then
    writeln('RelatedColumn: ', f.RelatedColumn);

  if not (mode in [cmIndex, cmKey]) then
    DumpAdoxProperties(f.Properties, PropNames([]));
end;

procedure PrintAdoxIndex(f: Index);
var i: integer;
begin
  writeln('Name: ', f.Name);
  writeln('IndexNulls: ', f.IndexNulls);
  DumpAdoxProperties(f.Properties, PropNames([]));
  writeln('');

  for i := 0 to f.Columns.Count - 1 do try
    writeln('Index column ['+IntToStr(i)+']:');
    PrintAdoxColumn(f.Columns[i], cmIndex);
    writeln('');
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message);
  end;
end;


function AdoxKeyTypeToStr(const t: KeyTypeEnum): string;
begin
  case t of
    adKeyPrimary: Result := 'Primary key';
    adKeyForeign: Result := 'Foreign key';
    adKeyUnique: Result := 'Unique key';
  else Result := 'Unknown ('+IntToStr(t)+')';
  end;
end;

function AdoxRuleTypeToStr(const t: RuleEnum): string;
begin
  case t of
    adRINone: Result := 'Do nothing';
    adRICascade: Result := 'Cascade';
    adRISetNull: Result := 'Set Null';
    adRISetDefault: Result := 'Set Default';
  else Result := 'Unknown action ('+IntToStr(t)+')';
  end;
end;

procedure PrintAdoxKey(f: Key);
var i: integer;
begin
 (* [no Properties] *)
  writeln('Name: ', f.Name);
  writeln('Type: ', AdoxKeyTypeToStr(f.Type_));
  writeln('UpdateRule: ', AdoxRuleTypeToStr(f.UpdateRule));
  writeln('DeleteRule: ', AdoxRuleTypeToStr(f.DeleteRule));
  writeln('RelatedTable: ', f.RelatedTable);
  writeln('');

  for i := 0 to f.Columns.Count - 1 do try
    writeln('Key column ['+IntToStr(i)+']:');
    PrintAdoxColumn(f.Columns[i], cmKey);
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

  for i := 0 to f.Columns.Count - 1 do try
    Subsection('Column['+IntToStr(i)+']');
    PrintAdoxColumn(f.Columns[i], cmTable);
    writeln('');
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;

  for i := 0 to f.Indexes.Count - 1 do try
    Subsection('Index['+IntToStr(i)+']');
    PrintAdoxIndex(f.Indexes[i]);
    writeln('');
  except
    on E: EOleException do
      writeln(E.Classname + ': ' + E.Message+''#13#10);
  end;

  for i := 0 to f.Keys.Count - 1 do try
    Subsection('Key['+IntToStr(i)+']');
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
