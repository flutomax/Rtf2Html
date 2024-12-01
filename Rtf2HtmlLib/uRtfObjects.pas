unit uRtfObjects;

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Generics.Collections,
  Vcl.Graphics, uRtfTypes, uRtfNullable;

type

  TRtfObject = class(TPersistent);

  TRtfHrefData = record
    URL: string;
    Text: string;
    procedure Reset;
  end;

  TColorHelper = record helper for TColor
  private
    function GetColorRef: TColorRef; inline;
    function GetBlue: Integer; inline;
    function GetGreen: Integer; inline;
    function GetRed: Integer; inline;
  public
    class function FromRGB(const R, G, B: Integer): TColor; inline; static;
    function GetHashCode: Integer;
    function ToString: string;
    property Red: Integer read GetRed;
    property Green: Integer read GetGreen;
    property Blue: Integer read GetBlue;
    property ColorRef: TColorRef read GetColorRef;
  end;


  TRtfFont = class(TRtfObject)
  private
    fId: string;
    fKind: TRtfFontKind;
    fPitch: TRtfFontPitch;
    fCharSet: Integer;
    fCodePage: Integer;
    fName: string;
    function GetCodePage: Integer;
    function IsEqual(obj: TObject): Boolean;
    function ComputeHashCode: Integer;
  public
    constructor Create(const aId: string; aKind: TRtfFontKind;
      aPitch: TRtfFontPitch; aCharSet, aCodePage: Integer; const aName: string);
    procedure Assign(Source: TPersistent); override;
    function Equals(obj: TObject): Boolean; override;
    function GetHashCode: Integer; override;
    function ToString: string; override;
    function GetEncoding: TEncoding;
  published
    property Id: string read fId;
    property Kind: TRtfFontKind read fKind;
    property Pitch: TRtfFontPitch read fPitch;
    property CharSet: Integer read fCharSet;
    property CodePage: Integer read GetCodePage;
    property Name: string read fName;
  end;

  TRtfTextFormat = class(TRtfObject)
  private
    fFont: TRtfFont;
    fFontSize: Integer;
    fSuperScript: Integer;
    fBold: Boolean;
    fItalic: Boolean;
    fUnderline: Boolean;
    fStrikeThrough: Boolean;
    fHidden: Boolean;
    fBackgroundColor: TColor;
    fForegroundColor: TColor;
    fAlignment: TRtfTextAlignment;
    function GetIsNormal: Boolean;
    function GetFontDescriptionDebug: string;
    function ComputeHashCode: Integer;
    function IsEqual(obj: TObject): Boolean;
  public
    constructor Create; overload;
    constructor Create(AFont: TRtfFont; AFontSize: Integer); overload;
    constructor Create(ACopy: TRtfTextFormat); overload;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    function Equals(Obj: TObject): Boolean; override;
    function GetHashCode: Integer; override;
    function ToString: string; override;
    function Duplicate: TRtfTextFormat;
    function DeriveWithSuperScript(aDeviation: Integer): TRtfTextFormat; overload;
    function DeriveWithSuperScript(aSuper: Boolean): TRtfTextFormat; overload;
    function DeriveNormal: TRtfTextFormat;
    function DeriveWithAlignment(aDerivedAlignment: TRtfTextAlignment): TRtfTextFormat;
    function DeriveWithBold(aDerivedBold: Boolean): TRtfTextFormat;
    function DeriveWithItalic(aDerivedItalic: Boolean): TRtfTextFormat;
    function DeriveWithUnderline(aDerivedUnderline: Boolean): TRtfTextFormat;
    function DeriveWithStrikeThrough(aDerivedStrikeThrough: Boolean): TRtfTextFormat;
    function DeriveWithHidden(aDerivedHidden: Boolean): TRtfTextFormat;
    function DeriveWithBackgroundColor(aDerivedBackgroundColor: TColor): TRtfTextFormat;
    function DeriveWithForegroundColor(aDerivedForegroundColor: TColor): TRtfTextFormat;
    function DeriveWithFont(aFont: TRtfFont): TRtfTextFormat;
    function DeriveWithFontSize(aDerivedFontSize: Integer): TRtfTextFormat;
  published
    property Font: TRtfFont read fFont;
    property FontSize: Integer read fFontSize;
    property SuperScript: Integer read fSuperScript;
    property IsNormal: Boolean read GetIsNormal;
    property IsBold: Boolean read fBold;
    property IsItalic: Boolean read fItalic;
    property IsUnderline: Boolean read fUnderline;
    property IsStrikeThrough: Boolean read fStrikeThrough;
    property IsHidden: Boolean read fHidden;
    property FontDescriptionDebug: string read GetFontDescriptionDebug;
    property BackgroundColor: TColor read fBackgroundColor;
    property ForegroundColor: TColor read fForegroundColor;
    property Alignment: TRtfTextAlignment read fAlignment;
  end;

  TRtfDocumentProperty = class(TRtfObject)
  private
    fPropertyKindCode: Integer;
    fPropertyKind: TRtfPropertyKind;
    fName: string;
    fStaticValue: string;
    fLinkValue: string;
    function IsEqual(Obj: TObject): Boolean;
    function ComputeHashCode: Integer;
  public
    constructor Create(aPropertyKindCode: Integer;
      const aName, aStaticValue: string; const aLinkValue: string = '');
    function Equals(Obj: TObject): Boolean; override;
    function GetHashCode: Integer; override;
    function ToString: string; override;
  published
    property PropertyKindCode: Integer read fPropertyKindCode;
    property PropertyKind: TRtfPropertyKind read fPropertyKind;
    property Name: string read fName;
    property StaticValue: string read fStaticValue;
    property LinkValue: string read fLinkValue;
  end;

  TRtfColorCollection = class(TList<TColor>);

  TRtfFontCollection = class(TBaseCollection<TRtfFont>)
  private
    fFontByIdMap: TDictionary<string, TRtfFont>;
    function GetFontById(const id: string): TRtfFont;
  public
    constructor Create(AOwnsObjects: Boolean = True); overload;
    constructor Create(const Collection: TRtfFontCollection); overload;
    destructor Destroy; override;
    function ContainsFontWithId(const fontId: string): Boolean;
    function Add(item: TRtfFont): integer;
    procedure Clear;
  public
    property FontById[const id: string]: TRtfFont read GetFontById;
  end;

  TRtfTextFormatCollection = class(TBaseCollection<TRtfTextFormat>)
  public
    function IndexOf(item: TRtfTextFormat): Integer;
  end;

  TRtfDocumentPropertyCollection = class(TBaseCollection<TRtfDocumentProperty>)
  private
    function GetItemByName(Name: string): TRtfDocumentProperty;
  public
    property ItemsByName[Name: string]: TRtfDocumentProperty read GetItemByName;
  end;

  TRtfIndent = class(TRtfObject)
  private
    fFirstIndent: Integer;
    fLeftIndent: Integer;
    fRightIndent: Integer;
    fSpaceBefore: TIntNullable;
    fSpaceAfter: TIntNullable;
    fSpaceBetweenLines: TIntNullable;
    function GetIsEmpty: Boolean;
  public
    constructor Create;
    procedure Assign(Source: TPersistent); override;
    procedure Reset;
  published
    property IsEmpty: Boolean read GetIsEmpty;
    property FirstIndent: Integer read fFirstIndent write fFirstIndent;
    property LeftIndent: Integer read fLeftIndent write fLeftIndent;
    property RightIndent: Integer read fRightIndent write fRightIndent;
    property SpaceBefore: TIntNullable read fSpaceBefore write fSpaceBefore;
    property SpaceAfter: TIntNullable read fSpaceAfter write fSpaceAfter;
    property SpaceBetweenLines: TIntNullable read fSpaceBetweenLines
      write fSpaceBetweenLines;
  end;


  { Table }

  TRtfTableValign = (tvaTop, taBottom, taCenter);
  TRtfActiveBorder = (abNone, abLeft, abTop, abRight, abBottom);

  TRtfObjectList = class(TObjectList<TRtfObject>)
  public
    procedure Assign(Source: TRtfObjectList);
  end;

  TRtfTableCell = class(TRtfObject)
  private
    fRowspan: Integer;
    fDataList: TRtfObjectList;
  public
    constructor Create; overload;
    constructor Create(aCell: TRtfTableCell); overload;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    procedure AddVisualObject(aObject: TRtfObject);
    property Rowspan: Integer read fRowspan write fRowspan;
    property DataList: TRtfObjectList read fDataList;
  end;

  TRtfCellBorder = record
    Visible: Boolean;
    Color: TColor;
    Width: Integer;
  end;

  TRtfTableCellDef = class(TRtfObject)
  private
    fBorderTop,
    fBorderBottom,
    fBorderLeft,
    fBorderRight: TRtfCellBorder;
    fActiveBorder: TRtfActiveBorder;
    fRight,
    fLeft: Integer;
    fMerged,
    fFirstMerged: Boolean;
    fVAlign: TRtfTableValign;
    fBackgroundColor: TColor;
    procedure SetActiveBorder(const Value: TRtfActiveBorder);
    function GetBorderVisible(index: TRtfActiveBorder): Boolean;
    procedure SetBorderVisible(index: TRtfActiveBorder; const Value: Boolean);
    function GetBorderColor(index: TRtfActiveBorder): TColor;
    procedure SetBorderColor(index: TRtfActiveBorder; const Value: TColor);
    function GetBorderWidth(index: TRtfActiveBorder): Integer;
    procedure SetBorderWidth(index: TRtfActiveBorder; const Value: Integer);
  public
    constructor Create; overload;
    constructor Create(aCellDef: TRtfTableCellDef); overload;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    function RightEquals(x: Integer): Boolean;
    function LeftEquals(x: Integer): Boolean;
    function BordersAreEquals: Boolean;
    property BackgroundColor: TColor read fBackgroundColor write fBackgroundColor;
    property ActiveBorder: TRtfActiveBorder read fActiveBorder write SetActiveBorder;
    property BorderTop: TRtfCellBorder read fBorderTop write fBorderTop;
    property BorderBottom: TRtfCellBorder read fBorderBottom write fBorderBottom;
    property BorderLeft: TRtfCellBorder read fBorderLeft write fBorderLeft;
    property BorderRight: TRtfCellBorder read fBorderRight write fBorderRight;
    property Left: Integer read fLeft write fLeft;
    property Right: Integer read fRight write fRight;
    property Merged: boolean read fMerged write fMerged;
    property FirstMerged: boolean read fFirstMerged write fFirstMerged;
    property VAlign: TRtfTableValign read fVAlign write fVAlign;
    property BorderVisible[index: TRtfActiveBorder]: Boolean read GetBorderVisible write SetBorderVisible;
    property BorderColor[index: TRtfActiveBorder]: TColor read GetBorderColor write SetBorderColor;
    property BorderWidth[index: TRtfActiveBorder]: Integer read GetBorderWidth write SetBorderWidth;
  end;

  TRtfTableCells = class(TObjectList<TRtfTableCell>)
  public
    procedure Assign(Source: TRtfTableCells);
  end;

  TRtfTableCellDefs = class(TObjectList<TRtfTableCellDef>)
  public
    procedure Assign(Source: TRtfTableCellDefs);
  end;

  TRtfTableRow = class(TRtfObject)
  private
    fCells: TRtfTableCells;
    fCellDefs: TRtfTableCellDefs;
    fHeight: TIntNullable;
    fLeft: TIntNullable;
    procedure SetCellDefs(const Value: TRtfTableCellDefs);
  public
    constructor Create; overload;
    constructor Create(src: TRtfTableRow); overload;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    property Cells: TRtfTableCells read fCells;
    property CellDefs: TRtfTableCellDefs read fCellDefs write SetCellDefs;
    property Left: TIntNullable read fLeft write fLeft;
    property Height: TIntNullable read fHeight write fHeight;
  end;

  TRtfTableRows = class(TObjectList<TRtfTableRow>)
  public
    procedure Assign(Source: TRtfTableRows);
  end;

implementation

uses
  System.Math, System.TypInfo, System.StrUtils, uRtfSpec, uRtfHash,
  uRtfMessages, uRtfVisual;


{ TRtfHrefData }

procedure TRtfHrefData.Reset;
begin
  URL := '';
  Text := '';
end;

{ TColorHelper }


class function TColorHelper.FromRGB(const R, G, B: Integer): TColor;
begin
  if not InRange(R, 0, 255) then
    raise ERtfColor.CreateFmt(sInvalidColorValue, [R]);
  if not InRange(G, 0, 255)  then
    raise ERtfColor.CreateFmt(sInvalidColorValue, [G]);
  if not InRange(B, 0, 255) then
    raise ERtfColor.CreateFmt(sInvalidColorValue, [B]);
  result := RGB(R, G, B);
end;

function TColorHelper.GetColorRef: TColorRef;
begin
  if self < 0 then
    Result := GetSysColor(self and $000000FF)
  else
    Result := self;
end;

function TColorHelper.GetBlue: Integer;
begin
  Result := GetBValue(GetColorRef);
end;

function TColorHelper.GetGreen: Integer;
begin
  Result := GetGValue(GetColorRef);
end;

function TColorHelper.GetRed: Integer;
begin
  result := GetRValue(GetColorRef);
end;

function TColorHelper.ToString: string;
begin
  Result := Format('Color{%d,%d,%d}', [Red, Green, Blue]);
end;

function TColorHelper.GetHashCode: Integer;
begin
  Result := Red;
  Result := Result xor Green;
  Result := Result xor Blue;
end;

{ TRtfFont }

constructor TRtfFont.Create(const aId: string; aKind: TRtfFontKind;
  aPitch: TRtfFontPitch; aCharSet, aCodePage: Integer; const aName: string);
begin
  if aId = '' then
    raise EArgumentException.Create(sEmptyID);
  if aCharSet < 0 then
    raise EArgumentException.CreateFmt(sInvalidCharacterSet, [charSet]);
  if aCodePage < 0 then
    raise EArgumentException.CreateFmt(sInvalidCodePage, [codePage]);
  if aName = '' then
    raise EArgumentException.Create(sEmptyName);

  fId := aId;
  fKind := aKind;
  fPitch := aPitch;
  fCharSet := aCharSet;
  fCodePage := aCodePage;
  fName := aName;
end;

procedure TRtfFont.Assign(Source: TPersistent);
begin
  if Source is TRtfFont then
  begin
    fId := TRtfFont(Source).fId;
    fKind := TRtfFont(Source).fKind;
    fPitch := TRtfFont(Source).fPitch;
    fCharSet := TRtfFont(Source).fCharSet;
    fCodePage := TRtfFont(Source).fCodePage;
    fName := TRtfFont(Source).fName;
  end
  else
    inherited;
end;

function TRtfFont.GetCodePage: Integer;
begin
  if fCodePage = 0 then
    Result := uRtfSpec.GetCodePage(fCharSet)
  else
    Result := fCodePage;
end;

function TRtfFont.GetEncoding: TEncoding;
begin
  Result := TEncoding.GetEncoding(GetCodePage);
end;

function TRtfFont.Equals(obj: TObject): Boolean;
begin
  if obj = Self then
    Exit(true);
  if (obj = nil) or (Self.ClassType <> obj.ClassType) then
    Exit(false);
  Result := IsEqual(obj);
end;

function TRtfFont.GetHashCode: Integer;
begin
  Result := AddHashCode(inherited GetHashCode, ComputeHashCode);
end;

function TRtfFont.ToString: string;
begin
  Result := fId + ':' + fName;
end;

function TRtfFont.IsEqual(obj: TObject): Boolean;
var
  compare: TRtfFont;
begin
  compare := TRtfFont(obj); // guaranteed to be non-null
  Result :=
    (compare <> nil) and
    (fId = compare.fId) and
    (fKind = compare.fKind) and
    (fPitch = compare.fPitch) and
    (fCharSet = compare.fCharSet) and
    (fCodePage = compare.fCodePage) and
    (fName = compare.fName);
end;

function TRtfFont.ComputeHashCode: Integer;
var
  hash: Integer;
begin
  hash := fId.GetHashCode;
  hash := AddHashCode(hash, Ord(fKind));
  hash := AddHashCode(hash, Ord(fPitch));
  hash := AddHashCode(hash, fCharSet);
  hash := AddHashCode(hash, fCodePage);
  hash := AddHashCode(hash, fName);
  Result := hash;
end;

{ TRtfTextFormat }

constructor TRtfTextFormat.Create;
begin
  fFont := TRtfFont.Create('f0', rfkNil, rfpDefault, 0, 0, 'Times New Roman');
  fBackgroundColor := clNone;
  fForegroundColor := clBlack;
  fFontSize := DefaultFontSize;
  fSuperScript := 0;
  fBold := false;
  fItalic := false;
  fUnderline := false;
  fStrikeThrough := false;
  fHidden := false;
  fAlignment := rtaLeft;
end;

constructor TRtfTextFormat.Create(AFont: TRtfFont; AFontSize: Integer);
begin
  if AFont = nil then
    raise EArgumentNilException.Create(sNilFont);
  if (AFontSize <= 0) or (AFontSize > $FFFF) then
    raise EArgumentOutOfRangeException.CreateFmt(sFontSizeOutOfRange, [AFontSize]);
  Create;
  fFont.Assign(AFont);
  fFontSize := AFontSize;
end;

constructor TRtfTextFormat.Create(ACopy: TRtfTextFormat);
begin
  if ACopy = nil then
    raise EArgumentNilException.Create(sNilCopy);
  Create;
  Assign(ACopy);
end;

destructor TRtfTextFormat.Destroy;
begin
  fFont.Free;
  inherited Destroy;
end;

procedure TRtfTextFormat.Assign(Source: TPersistent);
begin
  if Source is TRtfTextFormat then
  begin
    fFont.Assign(TRtfTextFormat(Source).fFont);
    fFontSize := TRtfTextFormat(Source).fFontSize;
    fSuperScript := TRtfTextFormat(Source).fSuperScript;
    fBold := TRtfTextFormat(Source).fBold;
    fItalic := TRtfTextFormat(Source).fItalic;
    fUnderline := TRtfTextFormat(Source).fUnderline;
    fStrikeThrough := TRtfTextFormat(Source).fStrikeThrough;
    fHidden := TRtfTextFormat(Source).fHidden;
    fBackgroundColor := TRtfTextFormat(Source).fBackgroundColor;
    fForegroundColor := TRtfTextFormat(Source).fForegroundColor;
    fAlignment := TRtfTextFormat(Source).fAlignment;
  end
  else
    inherited;
end;

function TRtfTextFormat.Duplicate: TRtfTextFormat;
begin
  Result := TRtfTextFormat.Create(Self);
end;

function TRtfTextFormat.GetFontDescriptionDebug: string;
var
  buf: TStringList;
  combined: Boolean;
begin
  buf := TStringList.Create;
  try
    buf.Add(fFont.Name);
    buf.Add(IntToStr(fFontSize));
    buf.Add(IfThen(fSuperScript >= 0, '+') + IntToStr(fSuperScript));
    combined := False;

    if fBold or fItalic or fUnderline or fStrikeThrough then
    begin
      if fBold then
      begin
        buf.Add('bold');
        combined := True;
      end;
      if fItalic then
      begin
        buf.Add(IfThen(combined, '+') + 'italic');
        combined := True;
      end;
      if fUnderline then
      begin
        buf.Add(IfThen(combined, '+') + 'underline');
        combined := True;
      end;
      if fStrikeThrough then
        buf.Add(IfThen(combined, '+') + 'strikethrough');
    end
    else
      buf.Add('plain');

    if fHidden then
      buf.Add(', hidden');

    Result := buf.Text;
  finally
    buf.Free;
  end;
end;

function TRtfTextFormat.GetIsNormal: Boolean;
begin
  Result := not fBold and not fItalic and not fUnderline and not fStrikeThrough
    and not fHidden and (fFontSize = DefaultFontSize) and (fSuperScript = 0) and
    (fForegroundColor = clBlack) and (fBackgroundColor = clWhite);
end;

function TRtfTextFormat.DeriveWithSuperScript(aDeviation: Integer): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fSuperScript := aDeviation;
  if aDeviation = 0 then
    Result.fFontSize := (fFontSize div 2) * 3;
end;

function TRtfTextFormat.DeriveWithSuperScript(aSuper: Boolean): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fFontSize := Max(1, (fFontSize * 2) div 3);
  Result.fSuperScript := IfThen(aSuper, 1, -1) * Max(1, fFontSize div 2);
end;

function TRtfTextFormat.DeriveNormal: TRtfTextFormat;
begin
  Result := TRtfTextFormat.Create(fFont, DefaultFontSize);
  Result.fAlignment := fAlignment; // this is a paragraph property, keep it
end;

function TRtfTextFormat.DeriveWithBold(aDerivedBold: Boolean): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fBold := aDerivedBold;
end;

function TRtfTextFormat.DeriveWithItalic(aDerivedItalic: Boolean): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fItalic := aDerivedItalic;
end;

function TRtfTextFormat.DeriveWithUnderline(aDerivedUnderline: Boolean): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fUnderline := aDerivedUnderline;
end;

function TRtfTextFormat.DeriveWithStrikeThrough(aDerivedStrikeThrough: Boolean): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fStrikeThrough := aDerivedStrikeThrough;
end;

function TRtfTextFormat.DeriveWithHidden(aDerivedHidden: Boolean): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fHidden := aDerivedHidden;
end;

function TRtfTextFormat.DeriveWithBackgroundColor(aDerivedBackgroundColor: TColor): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fBackgroundColor := aDerivedBackgroundColor;
end;

function TRtfTextFormat.DeriveWithForegroundColor(aDerivedForegroundColor: TColor): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.fForegroundColor := aDerivedForegroundColor;
end;

function TRtfTextFormat.DeriveWithAlignment(aDerivedAlignment: TRtfTextAlignment): TRtfTextFormat;
begin
  Result := Duplicate;
  Result.FAlignment := ADerivedAlignment;
end;

function TRtfTextFormat.DeriveWithFont(aFont: TRtfFont): TRtfTextFormat;
begin
  if aFont = nil then
    raise EArgumentNilException.Create(sNilFont);
  Result := Duplicate;
  Result.fFont.Assign(aFont);
end;

function TRtfTextFormat.DeriveWithFontSize(aDerivedFontSize: Integer): TRtfTextFormat;
begin
  if (aDerivedFontSize < 0) or (aDerivedFontSize > $FFFF) then
    raise EArgumentException.CreateFmt(sFontSizeOutOfRange, [aDerivedFontSize]);
  Result := Duplicate;
  Result.fFontSize := aDerivedFontSize;
end;

function TRtfTextFormat.Equals(Obj: TObject): Boolean;
begin
  if Obj = Self then
    Exit(True);
  if (Obj = nil) or (Obj.ClassType <> ClassType) then
    Exit(False);
  Result := IsEqual(obj);
end;

function TRtfTextFormat.IsEqual(obj: TObject): Boolean;
var
  compare: TRtfTextFormat;
begin
  if not(obj is TRtfTextFormat) then
    exit(false);
  compare := obj as TRtfTextFormat;
  result := Assigned(compare) and
    				fFont.Equals(compare.fFont) and
    				(fFontSize = compare.fFontSize) and
    				(fSuperScript = compare.fSuperScript) and
    				(fBold = compare.fBold) and
    				(fItalic = compare.fItalic) and
    				(fUnderline = compare.fUnderline) and
    				(fStrikeThrough = compare.fStrikeThrough) and
    				(fHidden = compare.fHidden) and
    				(fBackgroundColor = compare.fBackgroundColor) and
    				(fForegroundColor = compare.fForegroundColor) and
    				(fAlignment = compare.fAlignment);
end;

function TRtfTextFormat.ComputeHashCode: Integer;
var
  hash: Integer;
begin
  hash := (fFont as TRtfFont).GetHashCode;
  hash := AddHashCode(hash, fFontSize);
  hash := AddHashCode(hash, fSuperScript);
  hash := AddHashCode(hash, Ord(fBold));
  hash := AddHashCode(hash, Ord(fItalic));
  hash := AddHashCode(hash, Ord(fUnderline));
  hash := AddHashCode(hash, Ord(fStrikeThrough));
  hash := AddHashCode(hash, Ord(fHidden));
  hash := AddHashCode(hash, fBackgroundColor);
  hash := AddHashCode(hash, fForegroundColor);
  hash := AddHashCode(hash, Ord(fAlignment));
  Result := hash;
end;

function TRtfTextFormat.GetHashCode: Integer;
begin
  Result := AddHashCode(inherited GetHashCode, self.ComputeHashCode);
end;

function TRtfTextFormat.ToString: string;
var
  buf: TStringList;
begin
  buf := TStringList.Create;
  try
    buf.Add('Font ' + FontDescriptionDebug);
    buf.Add(', ' + GetEnumName(TypeInfo(TRtfTextAlignment), Ord(FAlignment)));
    buf.Add(', ' + fForegroundColor.ToString + ' on ' + fBackgroundColor.ToString);
    Result := buf.Text;
  finally
    buf.Free;
  end;
end;

{ TRtfDocumentProperty }

constructor TRtfDocumentProperty.Create(aPropertyKindCode: Integer;
  const aName, aStaticValue: string; const aLinkValue: string);
begin
  if aName.IsEmpty then
    raise EArgumentException.Create(sEmptyName);
  if aStaticValue = '' then
    raise EArgumentException.Create(sEmptyStaticValue);

  fPropertyKindCode := aPropertyKindCode;
  case fPropertyKindCode of
    0: fPropertyKind := rpkIntegerNumber;
    1: fPropertyKind := rpkRealNumber;
    2: fPropertyKind := rpkDate;
    3: fPropertyKind := rpkBoolean;
    4: fPropertyKind := rpkText;
  else
    fPropertyKind := rpkUnknown;
  end;

  fName := Name;
  fStaticValue := StaticValue;
  fLinkValue := LinkValue;
end;

function TRtfDocumentProperty.Equals(Obj: TObject): Boolean;
begin
  if Obj = Self then
    Exit(True);
  if not (Obj is TRtfDocumentProperty) then
    Exit(False);
  Result := IsEqual(Obj);
end;

function TRtfDocumentProperty.GetHashCode: Integer;
begin
  Result := AddHashCode(ClassType.ClassName.GetHashCode, ComputeHashCode);
end;

function TRtfDocumentProperty.IsEqual(Obj: TObject): Boolean;
var
  Compare: TRtfDocumentProperty;
begin
  Compare := TRtfDocumentProperty(Obj);
  Result := (fPropertyKindCode = Compare.fPropertyKindCode) and
            (fPropertyKind = Compare.fPropertyKind) and
            (fName = Compare.fName) and
            (AnsiSameStr(fStaticValue, Compare.fStaticValue)) and
            (AnsiSameStr(fLinkValue, Compare.fLinkValue));
end;

function TRtfDocumentProperty.ComputeHashCode: Integer;
begin
  Result := fPropertyKindCode;
  Result := AddHashCode(Result, Integer(fPropertyKind));
  Result := AddHashCode(Result, fName.GetHashCode);
  Result := AddHashCode(Result, fStaticValue.GetHashCode);
  Result := AddHashCode(Result, fLinkValue.GetHashCode);
end;

function TRtfDocumentProperty.ToString: string;
var
  Buf: TStringBuilder;
begin
  Buf := TStringBuilder.Create(fName);
  try
    if fStaticValue <> '' then
    begin
      Buf.Append('=');
      Buf.Append(fStaticValue);
    end;
    if fLinkValue <> '' then
    begin
      Buf.Append('@');
      Buf.Append(fLinkValue);
    end;
    Result := Buf.ToString;
  finally
    Buf.Free;
  end;
end;

{ TRtfFontCollection }

constructor TRtfFontCollection.Create(AOwnsObjects: Boolean);
begin
  inherited Create(AOwnsObjects);
  fFontByIdMap := TDictionary<string, TRtfFont>.Create;
end;

constructor TRtfFontCollection.Create(const Collection: TRtfFontCollection);
begin
  inherited Create(Collection);
  fFontByIdMap := TDictionary<string, TRtfFont>.Create;
end;

destructor TRtfFontCollection.Destroy;
begin
  fFontByIdMap.Free;
  inherited Destroy;
end;

function TRtfFontCollection.Add(item: TRtfFont): integer;
begin
  if item = nil then
    raise EArgumentNilException.Create(sNilItem);
  result := inherited Add(item);
  fFontByIdMap.Add(item.Id, item);
end;

procedure TRtfFontCollection.Clear;
begin
  inherited Clear;
  fFontByIdMap.Clear;
end;

function TRtfFontCollection.ContainsFontWithId(const fontId: string): Boolean;
begin
  Result := fFontByIdMap.ContainsKey(fontId);
end;

function TRtfFontCollection.GetFontById(const id: string): TRtfFont;
begin
  fFontByIdMap.TryGetValue(id, Result);
end;


{ TRtfTextFormatCollection }

function TRtfTextFormatCollection.IndexOf(item: TRtfTextFormat): Integer;
var
  i: Integer;
begin
  Result := -1;
  if Assigned(item) then
    for i := 0 to Count - 1 do
      if item.Equals(Items[i]) then
        exit(i);
end;

{ TRtfDocumentPropertyCollection }

function TRtfDocumentPropertyCollection.GetItemByName(
  Name: string): TRtfDocumentProperty;
var
  Item: TRtfDocumentProperty;
begin
  Result := nil;
  if Name.IsEmpty then
    exit;
  for Item in List do
    if Item.Name = Name then
      exit(Item);
end;

{ TRtfIndent }

constructor TRtfIndent.Create;
begin
  Reset;
end;

function TRtfIndent.GetIsEmpty: Boolean;
begin
  result := (fFirstIndent = 0) and (fLeftIndent = 0) and (fRightIndent = 0);
end;

procedure TRtfIndent.Reset;
begin
  fFirstIndent := 0;
  fLeftIndent := 0;
  fRightIndent := 0;
  fSpaceBefore := nil;
  fSpaceAfter := nil;
  fSpaceBetweenLines := nil;
end;

procedure TRtfIndent.Assign(Source: TPersistent);
begin
  if Source is TRtfIndent then
  begin
    fFirstIndent := TRtfIndent(Source).fFirstIndent;
    fLeftIndent := TRtfIndent(Source).fLeftIndent;
    fRightIndent := TRtfIndent(Source).fRightIndent;
    fSpaceBefore := TRtfIndent(Source).fSpaceBefore;
    fSpaceAfter := TRtfIndent(Source).fSpaceAfter;
    fSpaceBetweenLines := TRtfIndent(Source).fSpaceBetweenLines;
  end
  else
    inherited;
end;


{ TRtfObjectList }

procedure TRtfObjectList.Assign(Source: TRtfObjectList);
var
  i: integer;
  obj: TRtfObject;
  cls: TRtfVisualClass;
begin
  Clear;
  for i := 0 to Source.Count - 1 do
  begin
    obj := nil;
    cls := TRtfVisualClass(Source[i].ClassType);
    if Source[i] is TRtfVisualText then
      obj := TRtfVisualText.Create(TRtfVisualText(Source[i]))
    else
    if Source[i] is TRtfVisualImage then
      obj := TRtfVisualImage.Create(TRtfVisualImage(Source[i]))
    else
    if Source[i] is TRtfVisualSpecialChar then
      obj := TRtfVisualSpecialChar.Create(TRtfVisualSpecialChar(Source[i]))
    else
    if Source[i] is TRtfVisualBreak then
      obj := TRtfVisualBreak.Create(TRtfVisualBreak(Source[i]));
    if obj = nil then
      continue;
    Add(obj);
  end;
end;


{ TRtfTableCell }

constructor TRtfTableCell.Create;
begin
  fRowspan := 0;
  fDataList := TRtfObjectList.Create;
end;

constructor TRtfTableCell.Create(aCell: TRtfTableCell);
begin
  Create;
  Assign(aCell);
end;

destructor TRtfTableCell.Destroy;
begin
  FreeAndNil(fDataList);
  inherited;
end;

procedure TRtfTableCell.AddVisualObject(aObject: TRtfObject);
begin
  if aObject is TRtfVisual then
    fDataList.Add(aObject)
  else
    raise EArgumentException.CreateFmt(sInvalidClassName, [aObject.ClassName]);
end;

procedure TRtfTableCell.Assign(Source: TPersistent);
var
  src: TRtfTableCell;
begin
  if Source is TRtfTableCell then
  begin
    src := Source as TRtfTableCell;
    fRowspan := src.fRowspan;
    fDataList.Assign(src.DataList);
  end
  else
    inherited;
end;


{ TRtfTableCellDef }


constructor TRtfTableCellDef.Create;
begin
  FillChar(fBorderLeft, sizeof(TRtfCellBorder), 0);
  FillChar(fBorderTop, sizeof(TRtfCellBorder), 0);
  FillChar(fBorderRight, sizeof(TRtfCellBorder), 0);
  FillChar(fBorderBottom, sizeof(TRtfCellBorder), 0);
  fActiveBorder := abNone;
  fRight := 0;
  fLeft := 0;
  fMerged := False;
  fFirstMerged := False;
  fVAlign := tvaTop;
  fBackgroundColor := clNone;
end;

constructor TRtfTableCellDef.Create(aCellDef: TRtfTableCellDef);
begin
  Create;
  Assign(aCellDef);
end;

destructor TRtfTableCellDef.Destroy;
begin
  inherited; // for debug only
end;

procedure TRtfTableCellDef.Assign(Source: TPersistent);
var
  src: TRtfTableCellDef;
begin
  if Source is TRtfTableCellDef then
  begin
    src := Source as TRtfTableCellDef;
    fBorderTop := src.fBorderTop;
    fBorderBottom := src.fBorderBottom;
    fBorderLeft := src.fBorderLeft;
    fBorderRight := src.fBorderRight;
    fActiveBorder := src.fActiveBorder;
    fLeft := src.fLeft;
    fRight := src.fRight;
    fMerged := src.fMerged;
    fFirstMerged := src.fFirstMerged;
    fVAlign := src.fVAlign;
    fBackgroundColor := src.fBackgroundColor;
  end
  else
    inherited;
end;

function TRtfTableCellDef.LeftEquals(x: Integer): Boolean;
begin
  Result := x = fLeft;
end;

function TRtfTableCellDef.RightEquals(x: Integer): Boolean;
begin
  Result := x = fRight;
end;

function TRtfTableCellDef.BordersAreEquals: Boolean;
begin
  result := CompareMem(@fBorderLeft, @fBorderTop, sizeof(TRtfCellBorder)) and
  CompareMem(@fBorderTop, @fBorderRight, sizeof(TRtfCellBorder)) and
  CompareMem(@fBorderRight, @fBorderBottom, sizeof(TRtfCellBorder)) and
  CompareMem(@fBorderBottom, @fBorderLeft, sizeof(TRtfCellBorder));
end;

procedure TRtfTableCellDef.SetActiveBorder(const Value: TRtfActiveBorder);
begin
  if Value = abNone then
    case fActiveBorder of
      abLeft: fBorderLeft.Visible := false;
      abTop: fBorderTop.Visible := false;
      abRight: fBorderRight.Visible := false;
      abBottom: fBorderBottom.Visible := false;
    end;
  fActiveBorder := Value;
end;

function TRtfTableCellDef.GetBorderColor(index: TRtfActiveBorder): TColor;
begin
  case index of
    abLeft: result := fBorderLeft.Color;
    abTop: result := fBorderTop.Color;
    abRight: result := fBorderRight.Color;
    abBottom: result := fBorderBottom.Color;
    else result := 0;
  end;
end;

function TRtfTableCellDef.GetBorderVisible(index: TRtfActiveBorder): Boolean;
begin
  case index of
    abLeft: result := fBorderLeft.Visible;
    abTop: result := fBorderTop.Visible;
    abRight: result := fBorderRight.Visible;
    abBottom: result := fBorderBottom.Visible;
    else result := false;
  end;
end;

function TRtfTableCellDef.GetBorderWidth(index: TRtfActiveBorder): Integer;
begin
  case index of
    abLeft: result := fBorderLeft.Width;
    abTop: result := fBorderTop.Width;
    abRight: result := fBorderRight.Width;
    abBottom: result := fBorderBottom.Width;
    else result := 0;
  end;
end;

procedure TRtfTableCellDef.SetBorderColor(index: TRtfActiveBorder;
  const Value: TColor);
begin
  case index of
    abLeft: fBorderLeft.Color := Value;
    abTop: fBorderTop.Color := Value;
    abRight: fBorderRight.Color := Value;
    abBottom: fBorderBottom.Color := Value;
  end;
end;

procedure TRtfTableCellDef.SetBorderVisible(index: TRtfActiveBorder;
  const Value: Boolean);
begin
  case index of
    abLeft: fBorderLeft.Visible := Value;
    abTop: fBorderTop.Visible := Value;
    abRight: fBorderRight.Visible := Value;
    abBottom: fBorderBottom.Visible := Value;
  end;
end;

procedure TRtfTableCellDef.SetBorderWidth(index: TRtfActiveBorder;
  const Value: Integer);
begin
  case index of
    abLeft: fBorderLeft.Width := Value;
    abTop: fBorderTop.Width := Value;
    abRight: fBorderRight.Width := Value;
    abBottom: fBorderBottom.Width := Value;
  end;
end;

{ TRtfTableRow }

constructor TRtfTableRow.Create;
begin
  fCells := TRtfTableCells.Create;
  fCellDefs := TRtfTableCellDefs.Create;
  fHeight := nil;
  fLeft := nil;
end;

constructor TRtfTableRow.Create(src: TRtfTableRow);
begin
  Create;
  Assign(src);
end;

destructor TRtfTableRow.Destroy;
begin
  FreeAndNil(fCells);
  FreeAndNil(fCellDefs);
  inherited;
end;

procedure TRtfTableRow.SetCellDefs(const Value: TRtfTableCellDefs);
begin
  FreeAndNil(fCellDefs);
  fCellDefs := Value;
end;

procedure TRtfTableRow.Assign(Source: TPersistent);
var
  src: TRtfTableRow;
begin
  if Source is TRtfTableRow then
  begin
    src := Source as TRtfTableRow;
    fCells.Assign(src.fCells);
    fCellDefs.Assign(src.fCellDefs);
    fHeight := src.fHeight;
    fLeft := src.fLeft;
  end
  else
    inherited;
end;

{ TRtfTableCells }

procedure TRtfTableCells.Assign(Source: TRtfTableCells);
var
  i: integer;
begin
  Clear;
  for i := 0 to Source.Count - 1 do
    Add(TRtfTableCell.Create(Source[i]));
end;

{ TRtfTableCellDefs }

procedure TRtfTableCellDefs.Assign(Source: TRtfTableCellDefs);
var
  i: integer;
begin
  Clear;
  for i := 0 to Source.Count - 1 do
    Add(TRtfTableCellDef.Create(Source[i]));
end;


{ TRtfTableRows }

procedure TRtfTableRows.Assign(Source: TRtfTableRows);
var
  i: integer;
begin
  Clear;
  for i := 0 to Source.Count - 1 do
    Add(TRtfTableRow.Create(Source[i]));
end;


end.
