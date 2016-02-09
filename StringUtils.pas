unit StringUtils;

interface
uses SysUtils, WideStrUtils;

(*
  Contains various string handling routines for parsing Jet SQL data.
*)

type
 {$IFDEF UNICODE}
  UniString = string;
 {$ELSE}
  UniString = WideString;
 {$ENDIF}

  TMarkerFlag = (
    mfKeepOp,   //do not consider op/ed to be part of the contents
    mfKeepEd,   //for example, keep them when deleting comments
    mfNesting,  //allow nesting
    mfEscaping  //ignore escaped EDs. Requires no nesting. 
  );
  TMarkerFlags = set of TMarkerFlag;

 //Note on nesting:
 //Blocks can either allow nesting, in which case subblocks of any type can exist,
 //or be "comment-like", so that all the contents is treated as a data.
 //In nesting blocks, escaping is not allowed.

  TMarker = record
    s: UniString;
    e: UniString;
    f: TMarkerFlags;
  end;
  PMarker = ^TMarker;

  TMarkers = array of TMarker;

function Marker(s, e: UniString; f: TMarkerFlags=[]): TMarker;
procedure Append(var mk: TMarkers; r: TMarkers);

var
 (*
   These will be populated on initialization.
   If they were consts, they would be type incompatible, plus it would have been
   a pain to pass both pointer and length everywhere.
 *)

 //For RemoveComments()
  CommentMarkers: TMarkers;

 //Everything in these is considered string literal (comment openers ignored)
  StringLiteralMarkers: TMarkers;

 //Sequences in which final ';' is ignored
 //Sequences, in which CREATE TABLE separator ',' is ignored
 //Match()
  NoEndCommandMarkers: TMarkers;

 //Pseudo marker (#13, #10) - used in CRLF deletion
  EndLineMarker: TMarker;


function charPos(start, ptr: PWideChar): integer;
function prevChar(ptr: PWideChar): PWideChar;
function prevCharIs(start, ptr: PWideChar; c: WideChar): boolean;
function nextChar(ptr: PWideChar): PWideChar;
function SubStr(pc, pe: PWideChar): UniString;
procedure Adjust(ups, upc, ps: PWideChar; out pc: PWideChar);

function GetMeta(s: UniString; meta: UniString; out content: UniString): boolean;
function MetaPresent(s, meta: UniString): boolean;

function WStrPosIn(ps, pe, pattern: PWideChar): PWideChar;
function WStrMatch(main, sub: PWideChar): boolean;
function WStrPosEnd(pc: PWideChar; m: TMarkers; mk: TMarker): PWideChar;

function RemoveParts(s: UniString; mk: TMarker; IgnoreBlocks: TMarkers): string;
function RemoveCommentsI(s: UniString; m: TMarkers): UniString;
function RemoveComments(s: UniString): UniString;

function CommentTypeOpener(cType: integer): UniString;
function CommentTypeCloser(cType: integer): UniString;
function WStrPosAnyCommentI(ps: PWideChar; m: TMarkers; out pc: PWideChar): integer;
function WStrPosAnyComment(ps: PWideChar; out pc: PWideChar): integer;
function WStrPosCommentCloserI(pc: PWideChar; cType: integer; m: TMarkers): PWideChar;
function WStrPosCommentCloser(pc: PWideChar; cType: integer): PWideChar;
function WStrPosOrCommentI(ps: PWideChar; pattern: PWideChar; m: TMarkers; out pc: PWideChar): integer;
function WStrPosOrComment(ps: PWideChar; pattern: PWideChar; out pc: PWideChar): integer;
function WStrPosIgnoreCommentsI(ps: PWideChar; pattern: PWideChar; m: TMarkers): PWideChar;
function WStrPosIgnoreComments(ps: PWideChar; pattern: PWideChar): PWideChar;


type
  TUniStringArray = array of UniString;

function MatchI(s: UniString; parts: array of UniString; m: TMarkers): TUniStringArray;
function Match(s: UniString; parts: array of UniString): TUniStringArray;
function CutIdBrackets(s: UniString): UniString;
function SplitI(s: UniString; sep: WideChar; m: TMarkers): TUniStringArray;
function Split(s: UniString; sep: WideChar): TUniStringArray;

function FieldNameFromDefinition(s: UniString): UniString;

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
function SubStr(pc, pe: PWideChar): UniString;
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
function GetMeta(s: UniString; meta: UniString; out content: UniString): boolean;
var pc, pe: PWideChar;
begin
  pc := WStrPos(@s[1], PWideChar('/**'+meta+'*'));
  if pc=nil then begin
    Result := false;
    exit;
  end;

  Inc(pc, 4+Length(meta));
  pe := WStrPos(pc, PWideChar(UniString('*/')));
  if pe=nil then begin //unterminated comment - very strange
    Result := false;
    exit;
  end;

  content := Trim(SubStr(pc, pe));
  Result := true;
end;

//Checks if the specified meta is present in the string
function MetaPresent(s, meta: UniString): boolean;
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

//Receives a pointer to a start of the block of type Ind.
//Looks for an ending marker, minding nesting. Returns a pointer to the start of ED.
function WStrPosEnd(pc: PWideChar; m: TMarkers; mk: TMarker): PWideChar;
var i: integer;
  SpecSymbol: boolean;
begin
 //Skip opener
  Inc(pc, Length(mk.s));

 //To find the block ED, we need to scan till the end markers:
 // I. For simple comments: minding specsymbols.
 // II. For nested comments:
 //   - scan till the ED marker or the OP marker for any comment
 //    - if it's the OP, recursively call the search to get it's ED, restart from there
  SpecSymbol := false;
  while pc^ <> #00 do begin

    if mfNesting in mk.f then
     //Look for any OP
      for i := 0 to Length(m) - 1 do
        if WStrMatch(pc, PWideChar(mk.s)) then begin
         //Find it's ED
          pc := WStrPosEnd(pc, m, m[i]);
          if pc=nil then begin
            Result := nil;
            exit;
          end;
         //Skip ED (should be available! it was matched)
          Inc(pc, Length(m[i].e));
          break;
        end;

    if mfEscaping in mk.f then
      if SpecSymbol then begin
        SpecSymbol := false;
        Inc(pc);
        continue;
      end else
      if pc^='\' then begin
        SpecSymbol := true;
        Inc(pc);
        continue;
      end;

    if WStrMatch(pc, PWideChar(mk.e)) then begin
      Result := pc;
      exit;
    end;

    Inc(pc);
  end;
  Result := nil;
end;


//Removes all the parts of the text between Op-Ed markers, except those in IgnoreBlocks.
//Usually you want to pass all the non-nesting block types in IgnoreBlocks.
//Flags govern if the OP/ED markers themselves will be stripped or preserved.
function RemoveParts(s: UniString; mk: TMarker; IgnoreBlocks: TMarkers): string;
var ps, pc, pe: PWideChar;
  ctype: integer;

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
    if (pe^=' ') and (integer(pc) < integer(pe)) then begin
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
  ctype := WStrPosOrCommentI(ps, PWideChar(mk.s), IgnoreBlocks, pc);
  while pc <> nil do begin
    if ctype>=0 then begin
     //It's an IgnoreBlock    
      pe := WStrPosEnd(pc, IgnoreBlocks, IgnoreBlocks[ctype]);
      if pe=nil then begin
        appendResult(ps, WStrEnd(ps));
        exit;
      end;

     //Next part      
      Inc(pe, Length(IgnoreBlocks[ctype].e)); //Disregard Flags.mfKeepEd because we copy the contents anyway
      appendResult(ps, pe);
      ps := pe;
    end else begin
     //It's a DeleteBlock
      pe := WStrPosEnd(pc, IgnoreBlocks, mk);
      if pe=nil then begin
        appendResult(ps, WStrEnd(ps));
        exit;
      end;

     //Save the text till Pc and skip [pc, pe]
      if mfKeepOp in mk.f then
        Inc(pc, Length(mk.s));
      appendResult(ps, pc);

     //Next part
      ps := pe;
      if not (mfKeepEd in mk.f) then
        Inc(ps, Length(mk.e));
    end;
    ctype := WStrPosOrCommentI(ps, PWideChar(mk.s), IgnoreBlocks, pc);
  end;

 //If there's still text till the end, add it
  if ps^<>#00 then
    appendResult(ps, WStrEnd(ps));
end;

//Removes all supported comments, linefeeds.
function RemoveCommentsI(s: UniString; m: TMarkers): UniString;
var i: integer;
 im: TMarkers;
begin
  im := Copy(m); //Blocks to ignore comments in.
  Append(im, StringLiteralMarkers);

  for i := 0 to Length(m)-1 do
    s := RemoveParts(s, m[i], im);
  Result := RemoveParts(s, EndLineMarker, im);
end;

function RemoveComments(s: UniString): UniString;
begin
  Result := RemoveCommentsI(s, CommentMarkers);
end;


function CommentTypeOpener(cType: integer): UniString;
begin
  if (cType>=0) and (cType<Length(CommentMarkers)-1) then
    Result := CommentMarkers[cType].s
  else
    Result := '';
end;

function CommentTypeCloser(cType: integer): UniString;
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
    Result := WStrPosEnd(pc, m, m[cType]);
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

function MatchI(s: UniString; parts: array of UniString; m: TMarkers): TUniStringArray;
var us: UniString;
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

function Match(s: UniString; parts: array of UniString): TUniStringArray;
begin
  Result := MatchI(s, parts, NoEndCommandMarkers);
end;

//Deletes Jet identification brackets [] if they're present. (Also trims the string first)
function CutIdBrackets(s: UniString): UniString;
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
function SplitI(s: UniString; sep: WideChar; m: TMarkers): TUniStringArray;
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
    pc := WStrPosIgnoreCommentsI(ps, PWideChar(UniString(sep)), m);
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

function Split(s: UniString; sep: WideChar): TUniStringArray;
begin
  Result := SplitI(s, sep, NoEndCommandMarkers);
end;


//Extracts a field name (first word, possibly in [] brackets) from a CREATE TABLE
//field definition.
//Allows spaces in field name, treats comments inside of brackets as normal symbols,
//treats comments outside as separators.
function FieldNameFromDefinition(s: UniString): UniString;
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

function Marker(s, e: UniString; f: TMarkerFlags=[]): TMarker;
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
    mk[Length(mk)-Length(r)+i] := r[i];
end;

initialization
  SetLength(CommentMarkers, 3);
  CommentMarkers[0] := Marker('{', '}' , [mfEscaping]);
  CommentMarkers[1] := Marker('/*', '*/', [mfEscaping]);
  CommentMarkers[2] := Marker('--', #13#10, [mfKeepEd]);

  SetLength(StringLiteralMarkers, 2);
  StringLiteralMarkers[0] := Marker('''', '''', [mfEscaping]);
  StringLiteralMarkers[1] := Marker('"', '"', [mfEscaping]);

  NoEndCommandMarkers := Copy(CommentMarkers);
  Append(NoEndCommandMarkers, StringLiteralMarkers);
  SetLength(NoEndCommandMarkers, 7);
  NoEndCommandMarkers[5] := Marker('[', ']');
  NoEndCommandMarkers[6] := Marker('(', ')', [mfNesting]);

  EndLineMarker := Marker(#13, #10);
end.
