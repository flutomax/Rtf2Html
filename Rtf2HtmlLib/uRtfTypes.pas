unit uRtfTypes;

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Generics.Collections;

type

  TRtfFontKind = (rfkNil, rfkRoman, rfkSwiss, rfkModern, rfkScript, rfkDecor,
    rfkTech, rfkBidi);
  TRtfFontPitch = (rfpDefault, rfpFixed, rfpVariable);
  TRtfTextAlignment = (rtaLeft, rtaCenter, rtaRight, rtaJustify);
  TRtfPropertyKind = (rpkUnknown, rpkIntegerNumber, rpkRealNumber, rpkDate,
    rpkBoolean, rpkText);
  TRtfImageFormat = (rivNone, rifEmf, rifPng, rifJpg, rifWmf, rifBmp);

  ERtfColor = class(Exception);
  ERtfInvalidData = class(Exception);
  ERtfUndefinedFont = class(Exception);
  ERtfUndefinedColor = class(Exception);

  ERtfParse = class(Exception);
  ERtfStructure = class(ERtfParse);
  ERtfHexEncoding = class(ERtfParse);
  ERtfBraceNesting = class(ERtfParse);
  ERtfEmptyDocument = class(ERtfParse);
  ERtfUnicodeEncoding = class(ERtfParse);
  ERtfMultiByteEncoding = class(ERtfParse);
  ERtfUnsupportedStructure = class(ERtfParse);

  TBaseCollection<T: class> = class(TObjectList<T>)
  protected
    class function ItemsIsNotEqual(const a, b: T): Boolean;
    function HaveSameContents(obj: TObject): Boolean;
    function IsEqual(obj: TObject): Boolean; virtual;
    function GetHashCode: Integer; override;
    function ComputeHashCode: Integer; virtual;
  public
    procedure CopyTo(&Array: TArray<T>; Index: Integer); virtual;
    function Add(Item: T): integer; virtual;
    function Equals(obj: TObject): Boolean; override;
    function ToString: string; reintroduce; overload;
    function ToString(const delimiterText: string): string; overload;
    function ToString(const startText, endText, delimiterText,
      undefinedValueText: string): string; overload;
  end;

  TIntList = TList<Integer>;

implementation


uses
  System.RTLConsts, System.Rtti, System.Math, uRtfMessages;

{ TBaseCollection<T> }

function TBaseCollection<T>.Add(Item: T): integer;
begin
  if PPointer(@Item)^ = nil then
    raise EArgumentNilException.CreateRes(@SArgumentNil);
  result := inherited Add(Item);
end;

procedure TBaseCollection<T>.CopyTo(&Array: TArray<T>; Index: Integer);
begin
  if (Index < 0) or (Index + Count > Length(&Array)) then
    Error(@SListIndexError, Index);
  TArray.Copy<T>(ToArray, &array, 0, Index, Count);
end;

function TBaseCollection<T>.ComputeHashCode: Integer;
var
  item: T;
begin
  Result := 0;
  for item in List do
    Result := Result * 31 + IfThen(Assigned(item), item.GetHashCode);
end;

function TBaseCollection<T>.Equals(obj: TObject): Boolean;
begin
  if obj = Self then
    Exit(True);

  if (obj = nil) or (Self.ClassType <> obj.ClassType) then
    Exit(False);

  Result := IsEqual(obj);
end;

function TBaseCollection<T>.GetHashCode: Integer;
var
  hash: Integer;
begin
  hash := GetHashCode;
  Result := ComputeHashCode;
  if hash <> 0 then
    Result := Result + hash * 31;
end;

function TBaseCollection<T>.HaveSameContents(obj: TObject): Boolean;
var
  otherItems: TList<T>.TEnumerator;
  item, otherItem: T;
begin
  Result := false;
  if Assigned(obj) and (obj is TBaseCollection<T>) then
  begin
    otherItems := (obj as TBaseCollection<T>).GetEnumerator;
    Result := True;
    for item in List do
    begin
      if otherItems.MoveNext then
      begin
        otherItem := otherItems.Current;
        if ItemsIsNotEqual(item, otherItem) then
        begin
          Result := false;
          Break;
        end;
      end
      else
      begin
        Result := false;
        Break;
      end;
    end;
    if Result and otherItems.MoveNext then
      Result := false;
  end;
end;

function TBaseCollection<T>.IsEqual(obj: TObject): Boolean;
begin
  Result := self = obj;
  if not Result and (self <> nil) and (obj <> nil) and
  (self.ClassType = obj.ClassType) then
    Result := HaveSameContents(obj);
end;

class function TBaseCollection<T>.ItemsIsNotEqual(const a, b: T): Boolean;
begin
  result := (a <> b) and ((a = nil) or not a.Equals(b));
end;

function TBaseCollection<T>.ToString: string;
begin
  Result := ToString('[', ']', ',', 'nil');
end;

function TBaseCollection<T>.ToString(const delimiterText: string): string;
begin
  Result := ToString('', '', delimiterText, '');
end;

function TBaseCollection<T>.ToString(const startText, endText, delimiterText,
  undefinedValueText: string): string;
var
  str: TStringBuilder;
  first: Boolean;
  item: T;
begin
  str := TStringBuilder.Create(startText);
  try
    first := true;
    for item in List do
    begin
      if (item = nil) and (undefinedValueText = '') then
        continue;
      if first then
        first := false
      else
        str.Append(delimiterText);
      if item = nil then
        str.Append(undefinedValueText)
      else
        str.Append(item.ToString);
    end;
    str.Append(endText);
    Result := str.ToString;
  finally
    str.Free;
  end;
end;

end.
