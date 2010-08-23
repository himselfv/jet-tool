unit StringUtils;

interface
uses SysUtils, WideStrUtils;

(*
  Contains various string handling routines for parsing Jet SQL data.
*)

type
  TMarkerFlag = (
    mfKeepOp,   //do not consider op/ed to be part of the contents.
    mfKeepEd,   //for example, keep them when deleting comments
    mfNesting   //allow nesting
  );
  TMarkerFlags = set of TMarkerFlag;
  TMarker = record
    s: Widestring;
    e: Widestring;
    f: TMarkerFlags;
  end;
  PMarker = ^TMarker;

  TMarkers = array of TMarker;

function Marker(s, e: WideString; f: TMarkerFlags=[]): TMarker;
procedure Append(var mk: TMarkers; r: TMarkers);

var
 //These will be populated on initialization.
 //If they were consts, they would be type incompatible, plus it would have been
 //a pain to pass both pointer and length everywhere

 //For RemoveComments()
  CommentMarkers: TMarkers;

 //Sequences in which final ';' is ignored
 //Sequences, in which CREATE TABLE field set ',' separator is ignored
 //Match()
  NoEndCommandMarkers: TMarkers;


function charPos(start, ptr: PWideChar): integer;
function prevChar(ptr: PWideChar): PWideChar;
function prevCharIs(start, ptr: PWideChar; c: WideChar): boolean;
function nextChar(ptr: PWideChar): PWideChar;
function SubStr(pc, pe: PWideChar): WideString;
procedure Adjust(ups, upc, ps: PWideChar; out pc: PWideChar);

function GetMeta(s: Widestring; meta: WideString; out content: WideString): boolean;
function MetaPresent(s, meta: WideString): boolean;

function WStrPosIn(ps, pe, pattern: PWideChar): PWideChar;
function WStrMatch(main, sub: PWideChar): boolean;

function RemoveParts(s: WideString; op, ed: WideString; Flags: TMarkerFlags): string;
function RemoveCommentsI(s: WideString; m: TMarkers): WideString;
function RemoveComments(s: WideString): WideString;

function CommentTypeOpener(cType: integer): WideString;
function CommentTypeCloser(cType: integer): WideString;
function WStrPosAnyCommentI(ps: PWideChar; m: TMarkers; out pc: PWideChar): integer;
function WStrPosAnyComment(ps: PWideChar; out pc: PWideChar): integer;
function WStrPosCommentCloserI(pc: PWideChar; cType: integer; m: TMarkers): PWideChar;
function WStrPosCommentCloser(pc: PWideChar; cType: integer): PWideChar;
function WStrPosOrCommentI(ps: PWideChar; pattern: PWideChar; m: TMarkers; out pc: PWideChar): integer;
function WStrPosOrComment(ps: PWideChar; pattern: PWideChar; out pc: PWideChar): integer;
function WStrPosIgnoreCommentsI(ps: PWideChar; pattern: PWideChar; m: TMarkers): PWideChar;
function WStrPosIgnoreComments(ps: PWideChar; pattern: PWideChar): PWideChar;


type
  TWideStringArray = array of WideString;

function MatchI(s: WideString; parts: array of WideString; m: TMarkers): TWideStringArray;
function Match(s: WideString; parts: array of WideString): TWideStringArray;
function CutIdBrackets(s: WideString): WideString;
function SplitI(s: WideString; sep: WideChar; m: TMarkers): TWideStringArray;
function Split(s: WideString; sep: WideChar): TWideStringArray;

function FieldNameFromDefinition(s: WideString): WideString; 

implementation

//Returns character index of a character Ptr in string Start
function charPos(start, ptr: PWideChar): integer;
begin
  Result := 1+(integer(ptr) - integer(start)) div SizeOf(WideChar);
end;

//Returns a pointer to the previous character (obviously, no check for the start of the string)
function prevChar(ptr: PWideChar): PWideChar;
begin
  Result := PWideChar(integer(ptr)-SizeOf(WideChar));
end;

//Checks if the previous character exists (using Start) and if it's the required one.
function prevCharIs(start, ptr: PWideChar; c: WideChar): boolean;
begin
  if integer(ptr) <= integer(start) then
    Result := false
  else
    Result := PWideChar(integer(ptr)-SizeOf(WideChar))^=c;
end;

function nextChar(ptr: PWideChar): PWideChar;
begin
  Result := PWideChar(integer(ptr)+SizeOf(WideChar));
end;

//Returns a string consisting of [pc, pe). Not including pe.
function SubStr(pc, pe: PWideChar): WideString;
begin
  SetLength(Result, (integer(pe)-integer(pc)) div SizeOf(WideChar));
  Move(pc^, Result[1], (integer(pe)-integer(pc))); //size in bytes
end;

//Receives two starting pointers for two strings and one position pointer.
//Adjusts the second position pointer so that it points in the second string
//to the character with the same index first position pointer points at in the first string.
//In other words:
//  ups: 'Sample text';
//  upc:    'ple text';
//  ps:  'SAMPLE TEXT';
//Output:
//  pc:     'PLE TEXT';
procedure Adjust(ups, upc, ps: PWideChar; out pc: PWideChar);
begin
  pc := PWideChar(integer(ps) + integer(upc) - integer(ups));
end;

//Scans string S for meta Meta (/**Meta* content */) and returns it's content.
//Returns false if no meta was found.
//Metas are a private way of storing information in comments.
function GetMeta(s: Widestring; meta: WideString; out content: WideString): boolean;
var pc, pe: PWideChar;
begin
  pc := WStrPos(@s[1], PWideChar('/**'+meta+'*'));
  if pc=nil then begin
    Result := false;
    exit;
  end;

  Inc(pc, 4+Length(meta));
  pe := WStrPos(pc, PWideChar(WideString('*/')));
  if pe=nil then begin //unterminated comment - very strange
    Result := false;
    exit;
  end;

  content := Trim(SubStr(pc, pe));
  Result := true;
end;

//Checks if the specified meta is present in the string
function MetaPresent(s, meta: WideString): boolean;
begin
  Result := (WStrPos(@s[1], PWideChar('/**'+meta+'*')) <> nil);
end;

//Looks for a pattern inside of a substring [ps, pe) not including pe.
//Returns a pointer to Pattern or nil.
function WStrPosIn(ps, pe, pattern: PWideChar): PWideChar;
var
  Str, SubStr: PWideChar;
  Ch: WideChar;
begin
  Result := nil;
  if (ps = nil) or (ps^ = #0) or (pattern = nil) or (pattern^ = #0)
    or (integer(ps) >= integer(pe)) then Exit;
  Result := ps;
  Ch := pattern^;
  repeat
    if Result^ = Ch then
    begin
      Str := Result;
      SubStr := pattern;
      repeat
        Inc(Str);
        Inc(SubStr);
        if (SubStr^ = #0) or (integer(SubStr)>=integer(pe)) then exit;
        if (Str^ = #0) or (integer(Str)>=integer(pe)) then
        begin
          Result := nil;
          exit;
        end;
        if Str^ <> SubStr^ then break;
      until (FALSE);
    end;
    Inc(Result);
  until (Result^ = #0) or (integer(Result)>=integer(pe));
  Result := nil;
end;

//Returns True if Main starts with Sub.
function WStrMatch(main, sub: PWideChar): boolean;
begin
  while (main^<>#00) and (sub^<>#00) and (main^=sub^) do begin
    Inc(main);
    Inc(sub);
  end;
  Result := (sub^=#00);
end;

//Receives a pointer to a start of the block. Looks for an ending marker, minding nesting.
function WStrPosEnd(pc: PWideChar; op, ed: WideString; Nesting: boolean): PWideChar;
var lvl: integer;
begin
  lvl := 0;
  while pc^<>#00 do begin
    if Nesting and WStrMatch(pc, PWideChar(op)) then
      Inc(lvl);
    if WStrMatch(pc, PWideChar(ed)) then begin
      Dec(lvl);
      if lvl<0 then begin
        Result := pc;
        exit;
      end;
    end;

    Inc(pc);
  end;
  Result := nil;
end;


//Removes all the parts of the text between Op-Ed blocks, including Op blocks and, if instructed, Ed blocks.
function RemoveParts(s: WideString; op, ed: WideString; Flags: TMarkerFlags): string;
var ps, pc, pe: PWideChar;

  procedure appendResult(pc, pe: PWideChar);
  begin
   //Leave one or zero spaces at the start
    if (Length(Result) > 0) and (Result[Length(Result)]=' ') then
      while pc^=' ' do Inc(pc)
    else
      if pc^=' ' then begin //or we'd do Dec without an Inc
        while pc^=' ' do Inc(pc);
        Dec(pc);
      end else //no spaces at all
        if Result <> '' then
          Result := Result + ' ';

   //One space at most at the end
    Dec(pe); //now points to the last symbol
    if pe^=' ' then begin
      while (pe^=' ') and (integer(pc) < integer(pe)) do
        Dec(pe);
      Inc(pe); //one space
    end;
    Inc(pe); //points to the symbol after the last one

   //Copy
    Result := Result + SubStr(pc, pe);
  end;

begin
  Result := '';

  ps := @s[1];
  pc := WStrPos(ps, PWideChar(op));
  while pc <> nil do begin
    pe := WStrPosEnd(pc, op, ed, mfNesting in Flags);
    if pe=nil then begin
      appendResult(ps, WStrEnd(ps));
      exit;
    end;

   //Else we save the text till Pc and skip [pc, pe]
    if mfKeepOp in flags then
      Inc(pc, Length(op));
    appendResult(ps, pc);

   //Next part
    ps := pe;
    if not (mfKeepEd in flags) then
      Inc(ps, Length(ed));
    pc := WStrPos(ps, PWideChar(op));
  end;

 //If there's still text till the end, add it
  if ps^<>#00 then
    appendResult(ps, WStrEnd(ps));
end;

//Removes all supported comments, linefeeds.
function RemoveCommentsI(s: WideString; m: TMarkers): WideString;
var i: integer;
begin
  for i := 0 to Length(m)-1 do
    s := RemoveParts(s, m[i].s, m[i].e, m[i].f);
  Result := RemoveParts(s, #13, #10, []);
end;

function RemoveComments(s: WideString): WideString;
begin
  Result := RemoveCommentsI(s, CommentMarkers);
end;


function CommentTypeOpener(cType: integer): WideString;
begin
  if (cType>=0) and (cType<Length(CommentMarkers)-1) then
    Result := CommentMarkers[cType].s
  else
    Result := '';
end;

function CommentTypeCloser(cType: integer): WideString;
begin
  if (cType>=0) and (cType<=Length(CommentMarkers)-1) then
    Result := CommentMarkers[cType].e
  else
    Result := '';
end;

//Finds the first comment of any supported type or nil. Returns comment type (int) or -1.
//Use WStrPosCommentCloser() to find comment closer for a given type.
//Use CommentTypeOpener(), CommentTypeCloser() if you're curious about the details.
function WStrPosAnyCommentI(ps: PWideChar; m: TMarkers; out pc: PWideChar): integer;
var pt: PWideChar;
  i: integer;
begin
  pc := nil;
  Result := -1;

  for i := 0 to Length(m) - 1 do begin
    if pc=nil then
      pt := WStrPos(ps, PWideChar(m[i].s))
    else
      pt := WStrPosIn(ps, pc, PWideChar(m[i].s));
    if pt<>nil then begin
      pc := pt;
      Result := i;
    end;
  end;
end;

function WStrPosAnyComment(ps: PWideChar; out pc: PWideChar): integer;
begin
  Result := WStrPosAnyCommentI(ps, CommentMarkers, pc);
end;

//Finds the comment closer by comment type and returns the pointer to the first
//symbol after the comment is over. Or nil.
//Minds comment nesting, if it's enabled for this comment type.
function WStrPosCommentCloserI(pc: PWideChar; cType: integer; m: TMarkers): PWideChar;
begin
  if (cType >= 0) and (cType <= Length(m)-1) then begin
    Result := WStrPosEnd(pc, m[cType].s, m[cType].e, mfNesting in m[cType].f);
    if (Result<>nil) and not (mfKeepEd in m[cType].f) then Inc(Result, Length(m[cType].e));
  end else
    Result := nil;
end;

function WStrPosCommentCloser(pc: PWideChar; cType: integer): PWideChar;
begin
  Result := WStrPosCommentCloserI(pc, cType, CommentMarkers);
end;

function WStrMatchAnyMarkerOp(ps: PWideChar; m: TMarkers): integer;
var i: integer;
begin
  Result := 0;
  for i := 0 to Length(m) - 1 do
    if WStrMatch(ps, PWideChar(m[i].s)) then begin
      Result := i;
      break;
    end;
end;

//Finds a first instance of Pattern in Ps OR a first instance of comment of any type,
//whichever comes first.
//Returns the comment type OR -1. If the comment-type >=0, then pc points
//to a comment opener, else it points to the Pattern instance or nil.
function WStrPosOrCommentI(ps: PWideChar; pattern: PWideChar; m: TMarkers; out pc: PWideChar): integer;
begin
  pc := ps;
  while pc^<>#00 do begin
    if WStrMatch(pc, pattern) then begin
      Result := -1;
      exit;
    end;

    Result := WStrMatchAnyMarkerOp(pc, m);
    if Result<>0 then exit;

    Inc(pc);
  end;
  Result := -1;
  pc := nil;
end;

function WStrPosOrComment(ps: PWideChar; pattern: PWideChar; out pc: PWideChar): integer;
begin
  Result := WStrPosOrCommentI(ps, pattern, CommentMarkers, pc);
end;

//Finds a first instance of Pattern in Ps IGNORING any comments on the way.
function WStrPosIgnoreCommentsI(ps: PWideChar; pattern: PWideChar; m: TMarkers): PWideChar;
var ct: integer;
begin
  repeat
    ct := WStrPosOrCommentI(ps, pattern, m, Result);
    if ct<0 then break;

    Inc(Result, Length(m[ct].s)); //or else WStrPosCommentCloser will think it's another one
    ps := WStrPosCommentCloserI(Result, ct, m);
    if ps=nil then begin //malformed comment, ignore
      Result := nil;
      exit;
    end;
  until false;
end;

//Finds a first instance of Pattern in Ps, skipping comments. Returns nil
//if no instance can be found
function WStrPosIgnoreComments(ps: PWideChar; pattern: PWideChar): PWideChar;
begin
  Result := WStrPosIgnoreCommentsI(ps, pattern, CommentMarkers);
end;



//Splits string by parts, for example:
//  Input: asd bsd (klmn pqrs) rte
//  Parts: bsd,(,)
//  Result: asd,,klmn pqrs,rte
//Handles spaces and application-supported comments fine (comments are ignored
//in matching but included in results)
//  m: Ignore markers

function MatchI(s: WideString; parts: array of WideString; m: TMarkers): TWideStringArray;
var us: WideString;
  ps, pc: PWideChar;
  ups, upc: PWideChar;
  i: integer;
begin
  us := WideUpperCase(s);
  for i := 0 to Length(parts) - 1 do
    parts[i] := WideUpperCase(parts[i]);
  SetLength(Result, Length(parts)+1);

 //Empty string
  if Length(s)<=0 then begin
    if Length(parts)>0 then
      SetLength(Result, 0) //error
    else begin
      SetLength(Result, 1); //this is actually okay
      Result[0]:='';
    end;
    exit;
  end;

 //Parse
  ps := @s[1];
  ups := @us[1];
  for i := 0 to Length(parts) - 1 do begin
   //Find next part, skipping comments as needed
    upc := WStrPosIgnoreCommentsI(ups, PWideChar(parts[i]), m);
    if upc=nil then begin //no match
      SetLength(Result, 0);
      exit;
    end;

    Adjust(ups, upc, ps, pc);
    Result[i] := Trim(SubStr(ps, pc));

    ups := upc;
    ps := pc;
    Inc(ups, Length(parts[i]));
    Inc(ps, Length(parts[i]));
  end;

  Result[Length(Result)-1] := Trim(SubStr(ps, WStrEnd(ps)));
end;

function Match(s: WideString; parts: array of WideString): TWideStringArray;
begin
  Result := MatchI(s, parts, NoEndCommandMarkers);
end;

//Deletes Jet identification brackets [] if they're present. (Also trims the string first)
function CutIdBrackets(s: WideString): WideString;
var ps, pe: PWideChar;
begin
  if Length(s)=0 then begin
    Result := '';
    exit;
  end;

  ps := @s[1];
  while ps^=' ' do Inc(ps);

  pe := @s[Length(s)];
  while (pe^=' ') and (integer(ps) < integer(pe)) do Dec(pe);

  if (ps^='[') and (pe^=']') then begin
    Inc(ps);
    Dec(pe);
  end;

  if (ps<>@s[1]) or (pe<>@s[Length(s)]) then begin
    Inc(pe);
    Result := SubStr(ps, pe);
  end else
    Result := s; //nothign to cut
end;

//Splits string by separator, trims parts. Comments are handled fine.
function SplitI(s: WideString; sep: WideChar; m: TMarkers): TWideStringArray;
var ps, pc: PWideChar;
begin
  if Length(s)<=0 then begin
    SetLength(Result, 1);
    Result[0] := '';
    exit;
  end;

  ps := @s[1];
  SetLength(Result, 0);
  repeat
    pc := WStrPosIgnoreCommentsI(ps, PWideChar(Widestring(sep)), m);
    SetLength(Result, Length(Result)+1);
    if pc=nil then begin
      Result[Length(Result)-1] := SubStr(ps, WStrEnd(ps));
      exit;
    end;
    Result[Length(Result)-1] := SubStr(ps, pc);
    ps := pc;
    Inc(ps);
  until false;
end;

function Split(s: WideString; sep: WideChar): TWideStringArray;
begin
  Result := SplitI(s, sep, NoEndCommandMarkers);
end;


//Extracts a field name (first word, possibly in [] brackets) from a CREATE TABLE
//field definition.
//Allows spaces in field name, treats comments inside of brackets as normal symbols,
//treats comments outside as separators.
function FieldNameFromDefinition(s: WideString): WideString;
var pc: PWideChar;
  InBrackets: boolean;
begin
  if Length(s)<=0 then begin
    Result := '';
    exit;
  end;

  pc := @s[1];
  InBrackets := false;
  while pc^<>#00 do begin
    if pc^='[' then
      InBrackets := true
    else
    if pc^=']' then
      InBrackets := false
    else
    if InBrackets then begin
     //do nothing
    end else
    if (pc^=' ') or (pc^='{')
    or ((pc^='-')and(nextChar(pc)^='-'))
    or ((pc^='/')and(nextChar(pc)^='*')) then begin
      Result := Trim(SubStr(@s[1], pc));
      exit;
    end;

    Inc(pc);
  end;

  Result := Trim(s);
end;

function Marker(s, e: WideString; f: TMarkerFlags=[]): TMarker;
begin
  Result.s := s;
  Result.e := e;
  Result.f := f;
end;

procedure Append(var mk: TMarkers; r: TMarkers);
var i: integer;
begin
  SetLength(mk, Length(mk)+Length(r));
  for i := 0 to Length(r) - 1 do
    mk[Length(mk)+i] := r[i];
end;

initialization
  SetLength(CommentMarkers, 3);
  CommentMarkers[0] := Marker('{', '}');
  CommentMarkers[1] := Marker('/*', '*/');
  CommentMarkers[2] := Marker('--', #13#10, [mfKeepEd]);

  NoEndCommandMarkers := Copy(CommentMarkers);
  SetLength(NoEndCommandMarkers, 6);
  NoEndCommandMarkers[3] := Marker('[', ']');
  NoEndCommandMarkers[4] := Marker('(', ')', [mfNesting]);
  NoEndCommandMarkers[5] := Marker('''', '''');
end.
