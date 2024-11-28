unit uRtfBuilders;

interface

uses
  System.SysUtils, System.Classes, Generics.Collections, Vcl.Graphics,
  uRtfTypes, uRtfObjects, uRtfElement, uRtfDocumentInfo;

type

  TRtfFontBuilder = class(TRtfElementVisitor)
  private
    fFontId: string;
    fFontIndex: Integer;
    fFontCharset: Integer;
    fFontCodePage: Integer;
    fFontKind: TRtfFontKind;
    fFontPitch: TRtfFontPitch;
    fFontNameBuffer: TStringBuilder;
    function GetFontName: string;
  protected
    procedure DoVisitGroup(group: TRtfGroup); override;
    procedure DoVisitTag(tag: TRtfTag); override;
    procedure DoVisitText(text: TRtfText); override;
  public
    constructor Create;
    destructor Destroy; override;
    function CreateFont: TRtfFont;
    procedure Reset;
    property FontId: string read fFontId;
    property FontName: string read GetFontName;
    property FontIndex: Integer read fFontIndex;
    property FontCharset: Integer read fFontCharset;
    property FontCodePage: Integer read fFontCodePage;
    property FontKind: TRtfFontKind read fFontKind;
    property FontPitch: TRtfFontPitch read fFontPitch;
  end;

  ERtfFontTableFormat = class(Exception);

  TRtfFontTableBuilder = class(TRtfElementVisitor)
  private
    fFontBuilder: TRtfFontBuilder;
    fFontTable: TRtfFontCollection;
    fIgnoreDuplicatedFonts: Boolean;
    procedure BuildFontFromGroup(AGroup: TRtfGroup);
    procedure AddCurrentFont;
  protected
    procedure DoVisitGroup(AGroup: TRtfGroup); override;
  public
    constructor Create(AFontTable: TRtfFontCollection; AIgnoreDuplicatedFonts: Boolean = False);
    destructor Destroy; override;
    function IgnoreDuplicatedFonts: Boolean;
    procedure Reset;
  end;

  ERtfColorTableFormat = class(Exception);

  TRtfColorTableBuilder = class(TRtfElementVisitor)
  private
    fColorTable: TRtfColorCollection;
    fCurRed: Integer;
    fCurGreen: Integer;
    fCurBlue: Integer;
  protected
    procedure DoVisitGroup(AGroup: TRtfGroup); override;
    procedure DoVisitTag(ATag: TRtfTag); override;
    procedure DoVisitText(AText: TRtfText); override;
  public
    constructor Create(AColorTable: TRtfColorCollection);
    procedure Reset;
  end;

  TRtfTextBuilder = class(TRtfElementVisitor)
  private
    fBuffer: TStringList;
  protected
    procedure DoVisitText(text: TRtfText); override;
  public
    constructor Create;
    destructor Destroy; override;
    function CombinedText: string;
    procedure Reset;
  end;

  TRtfTimestampBuilder = class(TRtfElementVisitor)
  private
    fYear: Integer;
    fMonth: Integer;
    fDay: Integer;
    fHour: Integer;
    fMinutes: Integer;
    fSeconds: Integer;
  protected
    procedure DoVisitTag(ATag: TRtfTag); override;
  public
    constructor Create;
    procedure Reset;
    function CreateTimestamp: TDateTime;
  end;

  TRtfDocumentInfoBuilder = class(TRtfElementVisitor)
  private
    fInfo: TRtfDocumentInfo;
    fTextBuilder: TRtfTextBuilder;
    fTimestampBuilder: TRtfTimestampBuilder;
    function ExtractGroupText(AGroup: TRtfGroup): string;
    function ExtractTimestamp(AGroup: TRtfGroup): TDateTime;
  protected
    procedure DoVisitGroup(AGroup: TRtfGroup); override;
    procedure DoVisitTag(ATag: TRtfTag); override;
  public
    constructor Create(const AInfo: TRtfDocumentInfo);
    destructor Destroy; override;
    procedure Reset;
  end;

  TRtfUserPropertyBuilder = class(TRtfElementVisitor)
  private
    fCollectedProperties: TRtfDocumentPropertyCollection;
    fTextBuilder: TRtfTextBuilder;
    fPropertyTypeCode: Integer;
    fPropertyName: string;
    fStaticValue: string;
    fLinkValue: string;
  protected
    procedure DoVisitGroup(AGroup: TRtfGroup); override;
    procedure DoVisitTag(ATag: TRtfTag); override;
  public
    constructor Create(ACollectedProperties: TRtfDocumentPropertyCollection);
    destructor Destroy; override;
    function CreateProperty: TRtfDocumentProperty;
    procedure Reset;
  end;

  TRtfImageBuilder = class(TRtfElementVisitor)
  private
    fFormat: TRtfImageFormat;
    fWidth: Integer;
    fHeight: Integer;
    fDesiredWidth: Integer;
    fDesiredHeight: Integer;
    fScaleWidthPercent: Integer;
    fScaleHeightPercent: Integer;
    fImageDataHex: string;
    procedure Reset;
  protected
    procedure DoVisitGroup(group: TRtfGroup); override;
    procedure DoVisitTag(tag: TRtfTag); override;
    procedure DoVisitText(text: TRtfText); override;
  public
    constructor Create;
    property Format: TRtfImageFormat read fFormat;
    property Width: Integer read fWidth;
    property Height: Integer read fHeight;
    property DesiredWidth: Integer read fDesiredWidth;
    property DesiredHeight: Integer read fDesiredHeight;
    property ScaleWidthPercent: Integer read fScaleWidthPercent;
    property ScaleHeightPercent: Integer read fScaleHeightPercent;
    property ImageDataHex: string read fImageDataHex;
  end;


implementation

uses
  System.StrUtils, System.DateUtils, uRtfSpec, uRtfMessages;

{ TRtfFontBuilder }

constructor TRtfFontBuilder.Create;
begin
  inherited Create(rvNonRecursive);
  fFontNameBuffer := TStringBuilder.Create;
  Reset;
end;

destructor TRtfFontBuilder.Destroy;
begin
  fFontNameBuffer.Free;
  inherited Destroy;
end;

function TRtfFontBuilder.GetFontName: string;
var
  len: Integer;
begin
  Result := '';
  len := fFontNameBuffer.Length;
  if (len > 0) and (fFontNameBuffer[len - 1] = ';') then
  begin
    Result := fFontNameBuffer.ToString.Substring(0, len - 1).Trim;
    if Result = '' then
      Result := '';
  end;
end;

function TRtfFontBuilder.CreateFont: TRtfFont;
var
  fontName: string;
begin
  fontName := GetFontName;
  if fontName.IsEmpty then
    fontName := 'UnnamedFont_' + fFontId;
  Result := TRtfFont.Create(fFontId, fFontKind, fFontPitch, fFontCharset, fFontCodePage, fontName);
end;

procedure TRtfFontBuilder.Reset;
begin
  fFontIndex := 0;
  fFontCharset := 0;
  fFontCodePage := 0;
  fFontKind := rfkNil;
  fFontPitch := rfpDefault;
  fFontNameBuffer.Clear;
end;

procedure TRtfFontBuilder.DoVisitGroup(group: TRtfGroup);
begin
  if IndexStr(group.Destination, TagsFontSpecials) >= 0 then
    VisitGroupChildren(group);
end;

procedure TRtfFontBuilder.DoVisitTag(tag: TRtfTag);
begin
  case IndexStr(tag.Name, [TagFont, TagFontKindNil, TagFontKindRoman,
      TagFontKindSwiss, TagFontKindModern, TagFontKindScript, TagFontKindDecor,
      TagFontKindTech, TagFontKindBidi, TagFontCharset, TagCodePage, TagFontPitch]) of
    0:
    begin
      fFontId := tag.FullName;
      fFontIndex := tag.ValueAsNumber;
    end;
    1: fFontKind := rfkNil;
    2: fFontKind := rfkRoman;
    3: fFontKind := rfkSwiss;
    4: fFontKind := rfkModern;
    5: fFontKind := rfkScript;
    6: fFontKind := rfkDecor;
    7: fFontKind := rfkTech;
    8: fFontKind := rfkBidi;
    9: fFontCharset := tag.ValueAsNumber;
    10: fFontCodePage := tag.ValueAsNumber;
    11:
      case tag.ValueAsNumber of
        0: fFontPitch := rfpDefault;
        1: fFontPitch := rfpFixed;
        2: fFontPitch := rfpVariable;
      end;
  end;
end;

procedure TRtfFontBuilder.DoVisitText(text: TRtfText);
begin
  fFontNameBuffer.Append(text.Text);
end;

{ TRtfFontTableBuilder }

constructor TRtfFontTableBuilder.Create(AFontTable: TRtfFontCollection; AIgnoreDuplicatedFonts: Boolean = False);
begin
  inherited Create(rvNonRecursive);
  if AFontTable = nil then
    raise EArgumentNilException.Create(sNilFontTable);
  fFontBuilder := TRtfFontBuilder.Create;
  fFontTable := AFontTable;
  fIgnoreDuplicatedFonts := AIgnoreDuplicatedFonts;
end;

function TRtfFontTableBuilder.IgnoreDuplicatedFonts: Boolean;
begin
  Result := fIgnoreDuplicatedFonts;
end;

procedure TRtfFontTableBuilder.Reset;
begin
  fFontTable.Clear;
end;

destructor TRtfFontTableBuilder.Destroy;
begin
  fFontBuilder.Free;
  inherited;
end;

procedure TRtfFontTableBuilder.DoVisitGroup(AGroup: TRtfGroup);
var
  ChildCount, i: Integer;
begin
  if IndexStr(AGroup.Destination, TagsFontSpecials) >= 0 then
    BuildFontFromGroup(AGroup)
  else
  if AGroup.Destination = TagFontTable then
    if AGroup.Contents.Count > 1 then
      if AGroup.Contents[1].Kind = ekGroup then
        VisitGroupChildren(AGroup)
      else
      begin
        ChildCount := AGroup.Contents.Count;
        fFontBuilder.Reset;
        for I := 1 to ChildCount - 1 do // skip over the initial \fonttbl tag
        begin
          AGroup.Contents[I].Visit(fFontBuilder);
          if not fFontBuilder.FontName.IsEmpty then
          begin
            AddCurrentFont;
            fFontBuilder.Reset;
          end;
        end;
      end;
end;

procedure TRtfFontTableBuilder.BuildFontFromGroup(AGroup: TRtfGroup);
begin
  fFontBuilder.Reset;
  fFontBuilder.VisitGroup(AGroup);
  AddCurrentFont;
end;

procedure TRtfFontTableBuilder.AddCurrentFont;
begin
  if not fFontTable.ContainsFontWithId(fFontBuilder.FontId) then
    fFontTable.Add(fFontBuilder.CreateFont)
  else if not IgnoreDuplicatedFonts then
    raise ERtfFontTableFormat.CreateFmt(sDuplicateFont, [fFontBuilder.FontId]);
end;

{ TRtfColorTableBuilder }

constructor TRtfColorTableBuilder.Create(AColorTable: TRtfColorCollection);
begin
  inherited Create(rvNonRecursive);
  if AColorTable = nil then
    raise EArgumentNilException.Create(sNilColorTable);
  fColorTable := AColorTable;
end;

procedure TRtfColorTableBuilder.Reset;
begin
  fColorTable.Clear;
  fCurRed := 0;
  fCurGreen := 0;
  fCurBlue := 0;
end;

procedure TRtfColorTableBuilder.DoVisitGroup(AGroup: TRtfGroup);
begin
  if TagColorTable = AGroup.Destination then
    VisitGroupChildren(AGroup);
end;

procedure TRtfColorTableBuilder.DoVisitTag(ATag: TRtfTag);
begin
  case IndexStr(ATag.Name, [TagColorRed, TagColorGreen, TagColorBlue]) of
    0: fCurRed := ATag.ValueAsNumber;
    1: fCurGreen := ATag.ValueAsNumber;
    2: fCurBlue := ATag.ValueAsNumber;
  end;
end;

procedure TRtfColorTableBuilder.DoVisitText(AText: TRtfText);
begin
  if TagDelimiter = AText.Text then
  begin
    fColorTable.Add(TColor.FromRGB(fCurRed, fCurGreen, fCurBlue));
    fCurRed := 0;
    fCurGreen := 0;
    fCurBlue := 0;
  end
  else
    raise ERtfColorTableFormat.CreateFmt(sColorTableUnsupportedText, [AText.Text]);
end;

{ TRtfTextBuilder }

constructor TRtfTextBuilder.Create;
begin
  inherited Create(rvDepthFirst);
  fBuffer := TStringList.Create;
  Reset;
end;

destructor TRtfTextBuilder.Destroy;
begin
  fBuffer.Free;
  inherited;
end;

function TRtfTextBuilder.CombinedText: string;
begin
  Result := fBuffer.Text;
end;

procedure TRtfTextBuilder.Reset;
begin
  fBuffer.Clear;
end;

procedure TRtfTextBuilder.DoVisitText(text: TRtfText);
begin
  fBuffer.Add(text.Text);
end;

{ TRtfTimestampBuilder }

constructor TRtfTimestampBuilder.Create;
begin
  inherited Create(rvBreadthFirst);
  Reset;
end;

procedure TRtfTimestampBuilder.Reset;
begin
  fYear := 1970;
  fMonth := 1;
  fDay := 1;
  fHour := 0;
  fMinutes := 0;
  fSeconds := 0;
end;

function TRtfTimestampBuilder.CreateTimestamp: TDateTime;
begin
  Result := EncodeDateTime(fYear, fMonth, fDay, fHour, fMinutes, fSeconds, 0);
end;

procedure TRtfTimestampBuilder.DoVisitTag(ATag: TRtfTag);
begin
  case IndexStr(ATag.Name, [TagInfoYear, TagInfoMonth, TagInfoDay, TagInfoHour,
      TagInfoMinute, TagInfoSecond]) of
    0: fYear := ATag.ValueAsNumber;
    1: fMonth := ATag.ValueAsNumber;
    2: fDay := ATag.ValueAsNumber;
    3: fHour := ATag.ValueAsNumber;
    4: fMinutes := ATag.ValueAsNumber;
    5: fSeconds := ATag.ValueAsNumber;
  end;
end;

{ TRtfDocumentInfoBuilder }

constructor TRtfDocumentInfoBuilder.Create(const AInfo: TRtfDocumentInfo);
begin
  inherited Create(rvNonRecursive);
  fInfo := AInfo;
  fTextBuilder := TRtfTextBuilder.Create;
  fTimestampBuilder := TRtfTimestampBuilder.Create;
end;

destructor TRtfDocumentInfoBuilder.Destroy;
begin
  fTextBuilder.Free;
  fTimestampBuilder.Free;
  inherited;
end;

procedure TRtfDocumentInfoBuilder.Reset;
begin
  fInfo.Reset;
end;

procedure TRtfDocumentInfoBuilder.DoVisitGroup(AGroup: TRtfGroup);
begin
  case IndexStr(AGroup.Destination, [TagInfo, TagInfoTitle, TagInfoSubject,
      TagInfoAuthor, TagInfoManager, TagInfoCompany, TagInfoOperator,
      TagInfoCategory, TagInfoKeywords, TagInfoComment, TagInfoDocumentComment,
      TagInfoHyperLinkBase, TagInfoCreationTime, TagInfoRevisionTime,
      TagInfoPrintTime, TagInfoBackupTime]) of
    0: VisitGroupChildren(AGroup);
    1: fInfo.Title := ExtractGroupText(AGroup);
    2: fInfo.Subject := ExtractGroupText(AGroup);
    3: fInfo.Author := ExtractGroupText(AGroup);
    4: fInfo.Manager := ExtractGroupText(AGroup);
    5: fInfo.Company := ExtractGroupText(AGroup);
    6: fInfo.operator := ExtractGroupText(AGroup);
    7: fInfo.Category := ExtractGroupText(AGroup);
    8: fInfo.Keywords := ExtractGroupText(AGroup);
    9: fInfo.Comment := ExtractGroupText(AGroup);
    10: fInfo.DocumentComment := ExtractGroupText(AGroup);
    11: fInfo.HyperLinkbase := ExtractGroupText(AGroup);
    12: fInfo.CreationTime := ExtractTimestamp(AGroup);
    13: fInfo.RevisionTime := ExtractTimestamp(AGroup);
    14: fInfo.PrintTime := ExtractTimestamp(AGroup);
    15: fInfo.BackupTime := ExtractTimestamp(AGroup);
  end;
end;

procedure TRtfDocumentInfoBuilder.DoVisitTag(ATag: TRtfTag);
begin
  case IndexStr(ATag.Name, [TagInfoVersion, TagInfoRevision,
      TagInfoNumberOfPages, TagInfoNumberOfWords, TagInfoNumberOfChars,
      TagInfoId, TagInfoEditingTimeMinutes]) of
    0: fInfo.Version := ATag.ValueAsNumber;
    1: fInfo.Revision := ATag.ValueAsNumber;
    2: fInfo.NumberOfPages := ATag.ValueAsNumber;
    3: fInfo.NumberOfWords := ATag.ValueAsNumber;
    4: fInfo.NumberOfCharacters := ATag.ValueAsNumber;
    5: fInfo.Id := ATag.ValueAsNumber;
    6: fInfo.EditingTimeInMinutes := ATag.ValueAsNumber;
  end;
end;

function TRtfDocumentInfoBuilder.ExtractGroupText(AGroup: TRtfGroup): string;
begin
  fTextBuilder.Reset;
  fTextBuilder.VisitGroup(AGroup);
  Result := fTextBuilder.CombinedText;
end;

function TRtfDocumentInfoBuilder.ExtractTimestamp(AGroup: TRtfGroup): TDateTime;
begin
  fTimestampBuilder.Reset;
  fTimestampBuilder.VisitGroup(AGroup);
  Result := fTimestampBuilder.CreateTimestamp;
end;

{ TRtfUserPropertyBuilder }

constructor TRtfUserPropertyBuilder.Create(ACollectedProperties: TRtfDocumentPropertyCollection);
begin
  inherited Create(rvNonRecursive);
  if ACollectedProperties = nil then
    raise EArgumentNilException.Create(sNilCollectedProperties);
  fCollectedProperties := ACollectedProperties;
  fTextBuilder := TRtfTextBuilder.Create;
end;

destructor TRtfUserPropertyBuilder.Destroy;
begin
  fTextBuilder.Free;
  inherited;
end;

function TRtfUserPropertyBuilder.CreateProperty: TRtfDocumentProperty;
begin
  Result := TRtfDocumentProperty.Create(fPropertyTypeCode, fPropertyName, fStaticValue, fLinkValue);
end;

procedure TRtfUserPropertyBuilder.Reset;
begin
  fPropertyTypeCode := 0;
  fPropertyName := '';
  fStaticValue := '';
  fLinkValue := '';
end;

procedure TRtfUserPropertyBuilder.DoVisitGroup(AGroup: TRtfGroup);  var s: string;
begin
  case IndexStr(AGroup.Destination, [TagUserProperties, string.Empty,
      TagUserPropertyName, TagUserPropertyValue, TagUserPropertyLink]) of
    0: VisitGroupChildren(AGroup);
    1:
    begin
      Reset;
      VisitGroupChildren(AGroup);
      fCollectedProperties.Add(CreateProperty);
    end;
    2:
    begin
      fTextBuilder.Reset;
      fTextBuilder.VisitGroup(AGroup);
      fPropertyName := fTextBuilder.CombinedText;
    end;
    3:
    begin
      fTextBuilder.Reset;
      fTextBuilder.VisitGroup(AGroup);
      fStaticValue := fTextBuilder.CombinedText;
    end;
    4:
    begin
      fTextBuilder.Reset;
      fTextBuilder.VisitGroup(AGroup);
      fLinkValue := fTextBuilder.CombinedText;
    end;
  end;
end;

procedure TRtfUserPropertyBuilder.DoVisitTag(ATag: TRtfTag);
begin
  if ATag.Name = TagUserPropertyType then
    fPropertyTypeCode := ATag.ValueAsNumber;
end;

{ TRtfImageBuilder }

constructor TRtfImageBuilder.Create;
begin
  inherited Create(rvDepthFirst);
  Reset;
end;

procedure TRtfImageBuilder.Reset;
begin
  fFormat := rifBmp;
  fWidth := 0;
  fHeight := 0;
  fDesiredWidth := 0;
  fDesiredHeight := 0;
  fScaleWidthPercent := 100;
  fScaleHeightPercent := 100;
  fImageDataHex := '';
end;

procedure TRtfImageBuilder.DoVisitGroup(group: TRtfGroup);
begin
  if group.Destination = TagPicture then
  begin
    Reset;
    VisitGroupChildren(group);
  end;
end;

procedure TRtfImageBuilder.DoVisitTag(tag: TRtfTag);
begin
  case IndexStr(tag.Name, [TagPictureFormatWinDib, TagPictureFormatWinBmp,
      TagPictureFormatEmf, TagPictureFormatJpg, TagPictureFormatPng,
      TagPictureFormatWmf, TagPictureWidth, TagPictureHeight, TagPictureWidthGoal,
      TagPictureHeightGoal, TagPictureWidthScale, TagPictureHeightScale]) of
    0, 1: fFormat := rifBmp;
    2: fFormat := rifEmf;
    3: fFormat := rifJpg;
    4: fFormat := rifPng;
    5: fFormat := rifWmf;
    6:
    begin
      fWidth := Abs(tag.ValueAsNumber);
      fDesiredWidth := fWidth;
    end;
    7:
    begin
      fHeight := Abs(tag.ValueAsNumber);
      fDesiredHeight := fHeight;
    end;
    8:
    begin
      fDesiredWidth := Abs(tag.ValueAsNumber);
      if fWidth = 0 then
        fWidth := fDesiredWidth;
    end;
    9:
    begin
      fDesiredHeight := Abs(tag.ValueAsNumber);
      if fHeight = 0 then
        fHeight := fDesiredHeight;
    end;
    10: fScaleWidthPercent := Abs(tag.ValueAsNumber);
    11: fScaleHeightPercent := Abs(tag.ValueAsNumber);
  end;
end;

procedure TRtfImageBuilder.DoVisitText(text: TRtfText);
begin
  fImageDataHex := text.Text;
end;


end.
