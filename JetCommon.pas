unit JetCommon;

(*
  Contains common procedures for printing/converting stuff
*)

interface

function VarIsNil(val: OleVariant): boolean;
function str(val: OleVariant): WideString;
function arr_str(val: OleVariant): WideString;
function int(val: OleVariant): integer;
function uint(val: OleVariant): int64;
function bool(val: OleVariant): boolean;
function includes(main, flag: cardinal): boolean;

(*
  Previous section is closed automatically when a new one is declared.
  If you need to close manually, call EndSection.
*)
procedure Section(SectionName: string);
procedure Subsection(SectionName: string);
procedure EndSection;
procedure EndSubsection;

implementation
uses Variants;

function VarIsNil(val: OleVariant): boolean;
begin
  Result := VarIsClear(val) or VarIsNull(val);
end;

function str(val: OleVariant): WideString;
begin
  if VarIsNil(val) then
    Result := ''
  else begin
    Result := string(val);
  end;
end;

//То же, что str(), но форматирует и массивы 
function arr_str(val: OleVariant): WideString;
var i: integer;
begin
  if not VarIsArray(val) then begin
    Result := str(val);
    exit;
  end;

  if Length(val) > 0 then
    Result := arr_str(val[0])
  else
    Result := '';
  for i := 1 to Length(val) - 1 do
    Result := Result + ', ' + arr_str(val[i]);
  Result := '(' + Result + ')';
end;

function int(val: OleVariant): integer;
begin
  if VarIsNil(val) then
    Result := 0
  else Result := integer(val);
end;

function uint(val: OleVariant): int64;
begin
  if VarIsNil(val) then
    Result := 0
  else Result := val;
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


var
  InSection: boolean = false;
  InSubSection: boolean = false;

procedure Section(SectionName: string);
begin
  if InSection then
    EndSection;
  writeln(SectionName);
  writeln('=========================================');
  InSection := true;
end;

procedure SubSection(SectionName: string);
begin
  if InSubsection then
    EndSubsection;
  writeln('== ' + SectionName);
  InSubsection := true;
end;

procedure EndSection;
begin
  if InSection then begin
    writeln('');
    writeln('');
  end else
  if InSubsection then
    EndSubsection;
  InSubsection := false;
  InSection := false;
end;

procedure EndSubsection;
begin
  if InSubsection then
    writeln('');
  InSubsection := false;
end;

end.
