unit uRtfVisual;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, Vcl.Graphics,
  uRtfTypes, uRtfObjects;

type
  TRtfVisualKind = (rvkNone, rvkText, rvkBreak, rvkSpecial, rvkImage, rvkTable);
  TRtfVisualBreakKind = (rvbLine, rvbPage, rvbParagraph, rvbSection);
  TRtfVisualSpecialCharKind = (rvsTabulator, rvsParagraphNumberBegin,
    rvsParagraphNumberEnd, rvsNonBreakingSpace, rvsEmDash, rvsEnDash,
    rvsEmSpace, rvsEnSpace, rvsQmSpace, rvsBullet, rvsLeftSingleQuote,
    rvsRightSingleQuote, rvsLeftDoubleQuote, rvsRightDoubleQuote,
    rvsOptionalHyphen, rvsNonBreakingHyphen);

  TRtfVisualTableKind = (rvtReset, rvtTrowd, rvtRow, rvtCell, rvtCellx,
    rvtRowLeft, rvtRowHeight, rvtCellFirstMerged, rvtCellMerged,
    rvtCellBorderBottom, rvtCellBorderTop, rvtCellBorderLeft,
    rvtCellBorderRight, rvtCellBorderNone, rvtCellBorderColor,
    rvtCellBorderWidth, rvtCellVerticalAlignTop, rvtCellVerticalAlignCenter,
    rvtCellVerticalAlignBottom, rvtTableInTable, rvtTableCellBackgroundColor,
    rvtTableNestingLevel);

  TRtfVisual = class;
  TRtfVisualText = class;
  TRtfVisualBreak = class;
  TRtfVisualImage = class;
  TRtfVisualTable = class;
  TRtfVisualSpecialChar = class;
  TRtfVisualClass = class of TRtfVisual;

  TRtfVisualVisitor = class(TObject)
  protected
    procedure DoVisitText(aVisualText: TRtfVisualText); virtual; abstract;
    procedure DoVisitBreak(aVisualBreak: TRtfVisualBreak); virtual; abstract;
    procedure DoVisitSpecial(aVisualSpecialChar: TRtfVisualSpecialChar); virtual; abstract;
    procedure DoVisitImage(aVisualImage: TRtfVisualImage); virtual; abstract;
    procedure DoVisitTable(aVisualTable: TRtfVisualTable); virtual; abstract;
  public
    procedure VisitText(aVisualText: TRtfVisualText);
    procedure VisitBreak(aVisualBreak: TRtfVisualBreak);
    procedure VisitSpecial(aVisualSpecialChar: TRtfVisualSpecialChar);
    procedure VisitImage(aVisualImage: TRtfVisualImage);
    procedure VisitTable(aVisualTable: TRtfVisualTable);
  end;

  TRtfVisual = class(TRtfObject)
  private
    fKind: TRtfVisualKind;
    fIndent: TRtfIndent;
    fIsInTable: Boolean;
  protected
    constructor Create(aKind: TRtfVisualKind; aInTable: Boolean);
    destructor Destroy; override;
    procedure DoVisit(visitor: TRtfVisualVisitor); virtual; abstract;
    function IsEqual(obj: TObject): Boolean; virtual;
    function ComputeHashCode: Integer; virtual;
  public
    procedure Visit(visitor: TRtfVisualVisitor);
    procedure Assign(Source: TPersistent); override;
    function Equals(obj: TObject): Boolean; override;
    function GetHashCode: Integer; override;
    property Kind: TRtfVisualKind read fKind;
    property Indent: TRtfIndent read fIndent write fIndent;
    property IsInTable: Boolean read fIsInTable;
  end;

  TRtfVisualText = class(TRtfVisual)
  private
    fText: string;
    fURL: string;
    fFormat: TRtfTextFormat;
  protected
    procedure DoVisit(visitor: TRtfVisualVisitor); override;
    function IsEqual(obj: TObject): Boolean; override;
    function ComputeHashCode: Integer; override;
  public
    constructor Create(src: TRtfVisualText); overload;
    constructor Create(const aText: string; aFormat: TRtfTextFormat;
      aIndent: TRtfIndent; aURL: string = ''; aInTable: Boolean = false); overload;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    function ToString: string; override;
    property Text: string read fText;
    property URL: string read fURL;
    property Format: TRtfTextFormat read fFormat write fFormat;
  end;

  TRtfVisualSpecialChar = class(TRtfVisual)
  private
    fCharKind: TRtfVisualSpecialCharKind;
  public
    constructor Create(src: TRtfVisualSpecialChar); overload;
    constructor Create(ACharKind: TRtfVisualSpecialCharKind;
      aInTable: Boolean = false); overload;
    procedure Assign(Source: TPersistent); override;
    procedure DoVisit(AVisitor: TRtfVisualVisitor); override;
    function IsEqual(AObj: TObject): Boolean; override;
    function ComputeHashCode: Integer; override;
    function ToString: string; override;
    property CharKind: TRtfVisualSpecialCharKind read fCharKind;
  end;

  TRtfVisualBreak = class(TRtfVisual)
  private
    fBreakKind: TRtfVisualBreakKind;
  public
    constructor Create(src: TRtfVisualBreak); overload;
    constructor Create(ABreakKind: TRtfVisualBreakKind;
      aInTable: Boolean = false); overload;
    procedure Assign(Source: TPersistent); override;
    procedure DoVisit(AVisitor: TRtfVisualVisitor); override;
    function ToString: string; override;
    function IsEqual(AObj: TObject): Boolean; override;
    function ComputeHashCode: Integer; override;
    property BreakKind: TRtfVisualBreakKind read fBreakKind;
  end;

  TRtfVisualImage = class(TRtfVisual)
  private
    fImgFormat: TRtfImageFormat;
    fAlignment: TRtfTextAlignment;
    fWidth: Integer;
    fHeight: Integer;
    fDesiredWidth: Integer;
    fDesiredHeight: Integer;
    fScaleWidthPercent: Integer;
    fScaleHeightPercent: Integer;
    fImageDataHex: string;
    fImageDataBinary: TBytes;
    function GetImageDataBinary: TBytes;
  protected
    procedure DoVisit(AVisitor: TRtfVisualVisitor); override;
    function IsEqual(AObj: TObject): Boolean; override;
    function ComputeHashCode: Integer; override;
  public
    constructor Create(src: TRtfVisualImage); overload;
    constructor Create(AFormat: TRtfImageFormat; AAlignment: TRtfTextAlignment;
      AWidth, AHeight, ADesiredWidth, ADesiredHeight, AScaleWidthPercent, AScaleHeightPercent: Integer;
      const AImageDataHex: string; aIndent: TRtfIndent; aInTable: Boolean = false); overload;
    procedure Assign(Source: TPersistent); override;
    function ToString: string; override;
  published
    property ImgFormat: TRtfImageFormat read fImgFormat;
    property Alignment: TRtfTextAlignment read fAlignment write fAlignment;
    property Width: Integer read fWidth;
    property Height: Integer read fHeight;
    property DesiredWidth: Integer read fDesiredWidth;
    property DesiredHeight: Integer read fDesiredHeight;
    property ScaleWidthPercent: Integer read fScaleWidthPercent;
    property ScaleHeightPercent: Integer read fScaleHeightPercent;
    property ImageDataHex: string read fImageDataHex;
    property ImageDataBinary: TBytes read GetImageDataBinary;
  end;

  TRtfVisualTable = class(TRtfVisual)
  private
    fRows: TRtfTableRows;
  protected
    procedure DoVisit(AVisitor: TRtfVisualVisitor); override;
    function IsEqual(AObj: TObject): Boolean; override;
    function ComputeHashCode: Integer; override;
  public
    constructor Create(aInTable: Boolean = false); overload;
    constructor Create(src: TRtfVisualTable); overload;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    function Add(aRow: TRtfTableRow): Integer;
    function ToString: string; override;
    function CellDefsExists(CellDefs: TRtfTableCellDefs): Boolean;
    property Rows: TRtfTableRows read fRows;
  end;

  TRtfVisualCollection = class(TBaseCollection<TRtfVisual>);

implementation

uses
  System.TypInfo, uRtfGraphics, uRtfHash, uRtfMessages, uRtfHtmlFunctions;

{ TRtfVisualVisitor }

procedure TRtfVisualVisitor.VisitBreak(aVisualBreak: TRtfVisualBreak);
begin
  if Assigned(aVisualBreak) then
    DoVisitBreak(aVisualBreak);
end;

procedure TRtfVisualVisitor.VisitImage(aVisualImage: TRtfVisualImage);
begin
  if Assigned(aVisualImage) then
    DoVisitImage(aVisualImage);
end;

procedure TRtfVisualVisitor.VisitSpecial(
  aVisualSpecialChar: TRtfVisualSpecialChar);
begin
  if Assigned(aVisualSpecialChar) then
    DoVisitSpecial(aVisualSpecialChar);
end;

procedure TRtfVisualVisitor.VisitTable(aVisualTable: TRtfVisualTable);
begin
  if Assigned(aVisualTable) then
    DoVisitTable(aVisualTable);
end;

procedure TRtfVisualVisitor.VisitText(aVisualText: TRtfVisualText);
begin
  if Assigned(aVisualText) then
    DoVisitText(aVisualText);
end;

{ TRtfVisual }

constructor TRtfVisual.Create(aKind: TRtfVisualKind; aInTable: Boolean);
begin
  inherited Create;
  fIndent := TRtfIndent.Create;
  fKind := aKind;
  fIsInTable := aInTable;
end;

destructor TRtfVisual.Destroy;
begin
  fIndent.Free;
  inherited;
end;

procedure TRtfVisual.Visit(visitor: TRtfVisualVisitor);
begin
  if visitor = nil then
    raise EArgumentNilException.Create(sNilVisitor);
  DoVisit(visitor);
end;

function TRtfVisual.Equals(obj: TObject): Boolean;
begin
  if obj = Self then
    Exit(True);
  if (obj = nil) or (Self.ClassType <> obj.ClassType) then
    Exit(False);
  Result := IsEqual(obj);
end;

function TRtfVisual.GetHashCode: Integer;
begin
  Result := AddHashCode(Self.ClassType.ClassName.GetHashCode, ComputeHashCode);
end;

function TRtfVisual.IsEqual(obj: TObject): Boolean;
begin
  Result := True;
end;

function TRtfVisual.ComputeHashCode: Integer;
begin
  Result := $0F00DEAD;
end;

procedure TRtfVisual.Assign(Source: TPersistent);
begin
  if Source is TRtfVisual then
  begin
    Assert(fKind = TRtfVisual(Source).fKind);
    fIsInTable := TRtfVisual(Source).fIsInTable;
    fIndent.Assign(TRtfVisual(Source).fIndent);
  end
  else
    inherited;
end;


{ TRtfVisualText }

constructor TRtfVisualText.Create(src: TRtfVisualText);
begin
  inherited Create(rvkText, src.IsInTable);
  if src = nil then
    raise EArgumentNilException.Create(sNilSrc);
  fFormat := TRtfTextFormat.Create;
  Assign(src);
end;

constructor TRtfVisualText.Create(const aText: string; aFormat: TRtfTextFormat;
  aIndent: TRtfIndent; aURL: string; aInTable: Boolean);
begin
  inherited Create(rvkText, aInTable);
  fIndent.Assign(aIndent);
  if aText.IsEmpty then
    raise EArgumentNilException.Create(sNilText);
  if aFormat = nil then
    raise EArgumentNilException.Create(sNilFormat);
  fText := aText;
  fURL := aURL;
  fFormat := TRtfTextFormat.Create(aFormat);
  fIsInTable := aInTable;
end;

destructor TRtfVisualText.Destroy;
begin
  fFormat.Free;
  inherited;
end;

procedure TRtfVisualText.DoVisit(visitor: TRtfVisualVisitor);
begin
  visitor.VisitText(Self);
end;

function TRtfVisualText.IsEqual(obj: TObject): Boolean;
var
  compare: TRtfVisualText;
begin
  compare := TRtfVisualText(obj); // guaranteed to be non-nil
  Result := (compare <> nil) and
    inherited IsEqual(compare) and
    (fText = compare.fText) and
    (fFormat = compare.fFormat);
end;

procedure TRtfVisualText.Assign(Source: TPersistent);
begin
  inherited;
  if Source is TRtfVisualText then
  begin
    fText := TRtfVisualText(Source).fText;
    fURL := TRtfVisualText(Source).fURL;
    fFormat.Assign(TRtfVisualText(Source).fFormat);
  end;
end;

function TRtfVisualText.ComputeHashCode: Integer;
var
  hash: Integer;
begin
  hash := inherited ComputeHashCode;
  hash := AddHashCode(hash, fText);
  hash := AddHashCode(hash, fFormat as TRtfTextFormat);
  Result := hash;
end;

function TRtfVisualText.ToString: string;
begin
  Result := '''' + fText + '''';
end;

{ TRtfVisualSpecialChar }

constructor TRtfVisualSpecialChar.Create(ACharKind: TRtfVisualSpecialCharKind;
  aInTable: Boolean);
begin
  inherited Create(rvkSpecial, aInTable);
  fCharKind := ACharKind;
end;

constructor TRtfVisualSpecialChar.Create(src: TRtfVisualSpecialChar);
begin
  inherited Create(rvkSpecial, src.IsInTable);
  Assign(src);
end;

procedure TRtfVisualSpecialChar.DoVisit(AVisitor: TRtfVisualVisitor);
begin
  AVisitor.VisitSpecial(Self);
end;

function TRtfVisualSpecialChar.IsEqual(AObj: TObject): Boolean;
var
  Compare: TRtfVisualSpecialChar;
begin
  Compare := TRtfVisualSpecialChar(AObj); // guaranteed to be non-null
  Result := (Compare <> nil) and inherited IsEqual(Compare) and
    (fCharKind = Compare.fCharKind);
end;

procedure TRtfVisualSpecialChar.Assign(Source: TPersistent);
begin
  inherited;
  if Source is TRtfVisualSpecialChar then
    fCharKind := TRtfVisualSpecialChar(Source).fCharKind;
end;

function TRtfVisualSpecialChar.ComputeHashCode: Integer;
begin
  Result := AddHashCode(inherited ComputeHashCode, Integer(fCharKind));
end;

function TRtfVisualSpecialChar.ToString: string;
begin
  Result := GetEnumName(TypeInfo(TRtfVisualSpecialCharKind), Integer(fCharKind));
end;

{ TRtfVisualBreak }

constructor TRtfVisualBreak.Create(src: TRtfVisualBreak);
begin
  inherited Create(rvkBreak, src.IsInTable);
  Assign(src);
end;

constructor TRtfVisualBreak.Create(ABreakKind: TRtfVisualBreakKind;
  aInTable: Boolean);
begin
  inherited Create(rvkBreak, aInTable);
  fBreakKind := ABreakKind;
end;

function TRtfVisualBreak.ToString: string;
begin
  Result := GetEnumName(TypeInfo(TRtfVisualBreakKind), Ord(fBreakKind));
end;

procedure TRtfVisualBreak.DoVisit(AVisitor: TRtfVisualVisitor);
begin
  AVisitor.VisitBreak(Self);
end;

function TRtfVisualBreak.IsEqual(AObj: TObject): Boolean;
var
  Compare: TRtfVisualBreak;
begin
  Compare := TRtfVisualBreak(AObj); // guaranteed to be non-null
  Result := (Compare <> nil) and inherited IsEqual(Compare)
    and (fBreakKind = Compare.fBreakKind);
end;

procedure TRtfVisualBreak.Assign(Source: TPersistent);
begin
  inherited;
  if Source is TRtfVisualBreak then
    fBreakKind := TRtfVisualBreak(Source).fBreakKind;
end;

function TRtfVisualBreak.ComputeHashCode: Integer;
begin
  Result := AddHashCode(inherited ComputeHashCode, Ord(fBreakKind));
end;

{ TRtfVisualImage }

constructor TRtfVisualImage.Create(src: TRtfVisualImage);
begin
  inherited Create(rvkImage, src.IsInTable);
  if src = nil then
    raise EArgumentNilException.Create(sNilSrc);
  Assign(src);
end;

constructor TRtfVisualImage.Create(AFormat: TRtfImageFormat; AAlignment: TRtfTextAlignment;
  AWidth, AHeight, ADesiredWidth, ADesiredHeight, AScaleWidthPercent, AScaleHeightPercent: Integer;
  const AImageDataHex: string; aIndent: TRtfIndent; aInTable: Boolean);
begin
  inherited Create(rvkImage, aInTable);
  if AWidth <= 0 then
    raise EArgumentException.CreateFmt(sInvalidImageWidth, [AWidth]);
  if AHeight <= 0 then
    raise EArgumentException.CreateFmt(sInvalidImageHeight, [AHeight]);
  if ADesiredWidth <= 0 then
    raise EArgumentException.CreateFmt(sInvalidImageDesiredWidth, [ADesiredWidth]);
  if ADesiredHeight <= 0 then
    raise EArgumentException.CreateFmt(sInvalidImageDesiredHeight, [ADesiredHeight]);
  if AScaleWidthPercent <= 0 then
    raise EArgumentException.CreateFmt(sInvalidImageScaleWidth, [AScaleWidthPercent]);
  if AScaleHeightPercent <= 0 then
    raise EArgumentException.CreateFmt(sInvalidImageScaleHeight, [AScaleHeightPercent]);
  if AImageDataHex = '' then
    raise EArgumentException.Create(sEmptyImageDataHex);

  fImgFormat := AFormat;
  fAlignment := AAlignment;
  fWidth := AWidth;
  fHeight := AHeight;
  fDesiredWidth := ADesiredWidth;
  fDesiredHeight := ADesiredHeight;
  fScaleWidthPercent := AScaleWidthPercent;
  fScaleHeightPercent := AScaleHeightPercent;
  fImageDataHex := AImageDataHex;
  fIndent.Assign(aIndent);
end;

procedure TRtfVisualImage.Assign(Source: TPersistent);
begin
  inherited;
  if Source is TRtfVisualImage then
  begin
    fImgFormat := TRtfVisualImage(Source).fImgFormat;
    fAlignment := TRtfVisualImage(Source).fAlignment;
    fWidth := TRtfVisualImage(Source).fWidth;
    fHeight := TRtfVisualImage(Source).fHeight;
    fDesiredWidth := TRtfVisualImage(Source).fDesiredWidth;
    fDesiredHeight := TRtfVisualImage(Source).fDesiredHeight;
    fScaleWidthPercent := TRtfVisualImage(Source).fScaleWidthPercent;
    fScaleHeightPercent := TRtfVisualImage(Source).fScaleHeightPercent;
    fImageDataHex := TRtfVisualImage(Source).fImageDataHex;
    fIndent.Assign(TRtfVisualImage(Source).fIndent);
  end;
end;

procedure TRtfVisualImage.DoVisit(AVisitor: TRtfVisualVisitor);
begin
  AVisitor.VisitImage(self);
end;

function TRtfVisualImage.GetImageDataBinary: TBytes;
begin
  if Length(fImageDataBinary) = 0 then
    fImageDataBinary := TextToBinary(fImageDataHex);
  Result := fImageDataBinary;
end;

function TRtfVisualImage.IsEqual(AObj: TObject): Boolean;
var
  Compare: TRtfVisualImage;
begin
  Compare := TRtfVisualImage(AObj);
  Result :=
    (Compare <> nil) and
    inherited IsEqual(Compare) and
    (fImgFormat = Compare.fImgFormat) and
    (fAlignment = Compare.fAlignment) and
    (fWidth = Compare.fWidth) and
    (fHeight = Compare.fHeight) and
    (fDesiredWidth = Compare.fDesiredWidth) and
    (fDesiredHeight = Compare.fDesiredHeight) and
    (fScaleWidthPercent = Compare.fScaleWidthPercent) and
    (fScaleHeightPercent = Compare.fScaleHeightPercent) and
    (fImageDataHex = Compare.fImageDataHex);
end;

function TRtfVisualImage.ToString: string;
begin
  Result := Format('[%s: %s, %d x %d (%d x %d) {%d%% x %d%%} : %d bytes]',
    [GetEnumName(TypeInfo(TRtfImageFormat), Ord(fImgFormat)),
    GetEnumName(TypeInfo(TRtfTextAlignment), Ord(fAlignment)),
    fWidth, fHeight, fDesiredWidth, fDesiredHeight,
    fScaleWidthPercent, fScaleHeightPercent,
    Length(fImageDataHex) div 2]);
end;

function TRtfVisualImage.ComputeHashCode: Integer;
var
  Hash: Integer;
begin
  Hash := inherited ComputeHashCode;
  Hash := AddHashCode(Hash, Ord(fImgFormat));
  Hash := AddHashCode(Hash, Ord(fAlignment));
  Hash := AddHashCode(Hash, fWidth);
  Hash := AddHashCode(Hash, fHeight);
  Hash := AddHashCode(Hash, fDesiredWidth);
  Hash := AddHashCode(Hash, fDesiredHeight);
  Hash := AddHashCode(Hash, fScaleWidthPercent);
  Hash := AddHashCode(Hash, fScaleHeightPercent);
  Hash := AddHashCode(Hash, fImageDataHex);
  Result := Hash;
end;

{ TRtfVisualTable }

constructor TRtfVisualTable.Create(aInTable: Boolean);
begin
  inherited Create(rvkTable, aInTable);
  fRows := TRtfTableRows.Create;
end;

constructor TRtfVisualTable.Create(src: TRtfVisualTable);
begin
  inherited Create(rvkTable, src.IsInTable);
  fRows := TRtfTableRows.Create;
  Assign(src);
end;

destructor TRtfVisualTable.Destroy;
begin
  FreeAndNil(fRows);
  inherited;
end;

procedure TRtfVisualTable.Assign(Source: TPersistent);
begin
  inherited;
  if Source is TRtfVisualTable then
    fRows.Assign(TRtfVisualTable(Source).fRows);
end;

procedure TRtfVisualTable.DoVisit(AVisitor: TRtfVisualVisitor);
begin
  AVisitor.VisitTable(self);
end;

function TRtfVisualTable.CellDefsExists(CellDefs: TRtfTableCellDefs): Boolean;
var
  row: TRtfTableRow;
begin
  result := false;
  for row in fRows do
    if row.CellDefs = CellDefs then
      exit(true);
end;

function TRtfVisualTable.ComputeHashCode: Integer;
var
  Hash: Integer;
begin
  Hash := inherited ComputeHashCode;
  Hash := AddHashCode(Hash, fRows.Count);
  Result := Hash;
end;

function TRtfVisualTable.IsEqual(AObj: TObject): Boolean;
var
  Compare: TRtfVisualTable;
begin
  Compare := TRtfVisualTable(AObj);
  Result := (Compare <> nil) and inherited IsEqual(Compare)
    and (fRows.Count = Compare.fRows.Count);
end;

function TRtfVisualTable.ToString: string;
var
  rows, cols: integer;
begin
  cols := 0;
  rows := fRows.Count;
  if rows > 0 then
    cols := fRows[0].Cells.Count;
  Result := Format('[%dx%d]', [cols, rows]);
end;

function TRtfVisualTable.Add(aRow: TRtfTableRow): Integer;
begin
  result := fRows.Add(aRow);
end;

end.
