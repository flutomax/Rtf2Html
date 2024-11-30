unit uRtfHtmlObjects;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, uRtfTypes,
  uRtfObjects, uRtfGraphics, uRtfVisual, uRtfHtmlWriter;

const

  hcsNone = $00000000;
  hcsDocument = $00000001;
  hcsHtml = $00000010;
  hcsHead = $00000100;
  hcsBody = $00001000;
  hcsContent = $00010000;
  hcsAll = hcsDocument or hcsHtml or hcsHead or hcsBody or hcsContent;

  DefaultDocumentHeader = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">';
  DefaultDocumentCharacterSet = 'UTF-8';
  DefaultVisualHyperlinkPattern = '[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,3}(:[a-zA-Z0-9]*)?/?([a-zA-Z0-9\-\._\?\,''/\\\+&%\$#\=~])*';
  DefaultHtmlFileExtension = '.html';
  DefaultFileNamePattern = '%d%s';

type

  TRtfHtmlStyle = class(TObject)
  private
    fForegroundColor: string;
    fBackgroundColor: string;
    fFontFamily: string;
    fFontSize: string;
    fDefault: boolean;
    function GetIsEmpty: Boolean;
    function ComputeHashCode: Integer;
  public
    class var Empty: TRtfHtmlStyle;
    constructor Create(aDefault: boolean = false);
    function Equals(Obj: TObject): Boolean; override;
    function GetHashCode: Integer; override;
  published
    property ForegroundColor: string read fForegroundColor write fForegroundColor;
    property BackgroundColor: string read fBackgroundColor write fBackgroundColor;
    property FontFamily: string read fFontFamily write fFontFamily;
    property FontSize: string read fFontSize write fFontSize;
    property IsEmpty: Boolean read GetIsEmpty;
    property Default: boolean read fDefault;
  end;

  TRtfHtmlCssStyle = class(TObject)
  private
    fProperties: TStringList;
    fSelectorName: string;
  public
    constructor Create(const ASelectorName: string;
      const AProperties: string = '');
    destructor Destroy; override;
    property Properties: TStringList read fProperties;
    property SelectorName: string read fSelectorName;
  end;

  TRtfHtmlCssStyleCollection = class(TBaseCollection<TRtfHtmlCssStyle>);

  TRtfHtmlConvertSettings = class
  private
    fGraphics: TRtfGraphics;
    fConvertScope: Integer;
    fStyles: TRtfHtmlCssStyleCollection;
    fStyleSheetLinks: TStrings;
    fDocumentHeader: string;
    fTitle: string;
    fCharacterSet: string;
    fImagesPath: string;
    fVisualHyperlinkPattern: string;
    fSpecialCharsRepresentation: string;
    fIsShowHiddenText: Boolean;
    fConvertVisualHyperlinks: Boolean;
    fUseNonBreakingSpaces: Boolean;
    function GetHasStyles: boolean;
    function GetHasStyleSheetLinks: boolean;
  public
    constructor Create; overload;
    constructor Create(aConvertScope: integer); overload;
    constructor Create(aGraphics: TRtfGraphics); overload;
    constructor Create(aGraphics: TRtfGraphics; aConvertScope: integer); overload;
    destructor Destroy; override;
    function GetImageUrl(index: Integer; ImageFormat: TRtfImageFormat): string;
  published
    property Graphics: TRtfGraphics read fGraphics;
    property ConvertScope: Integer read fConvertScope write fConvertScope;
    property HasStyles: Boolean read GetHasStyles;
    property Styles: TRtfHtmlCssStyleCollection read fStyles;
    property HasStyleSheetLinks: Boolean read GetHasStyleSheetLinks;
    property StyleSheetLinks: TStrings read fStyleSheetLinks;
    property DocumentHeader: string read fDocumentHeader write fDocumentHeader;
    property Title: string read fTitle write fTitle;
    property CharacterSet: string read fCharacterSet write fCharacterSet;
    property VisualHyperlinkPattern: string read fVisualHyperlinkPattern
      write fVisualHyperlinkPattern;
    property SpecialCharsRepresentation: string read fSpecialCharsRepresentation
      write fSpecialCharsRepresentation;
    property IsShowHiddenText: Boolean read fIsShowHiddenText write fIsShowHiddenText;
    property ConvertVisualHyperlinks: Boolean read fConvertVisualHyperlinks
      write fConvertVisualHyperlinks;
    property UseNonBreakingSpaces: Boolean read fUseNonBreakingSpaces
      write fUseNonBreakingSpaces;
    property ImagesPath: string read fImagesPath write fImagesPath;
  end;

  TRtfHtmlElementPath = class
  private
    fElements: TStack<THtmlTextWriterTag>;
  public
    constructor Create;
    destructor Destroy; override;
    function Count: Integer;
    function Current: THtmlTextWriterTag;
    function IsCurrent(tag: THtmlTextWriterTag): Boolean;
    function Contains(tag: THtmlTextWriterTag): Boolean;
    procedure Push(tag: THtmlTextWriterTag);
    function Pop: THtmlTextWriterTag;
    function ToString: string; override;
  end;

  TRtfHtmlSpecialCharCollection = class(TDictionary<TRtfVisualSpecialCharKind, string>)
  public
    constructor Create; overload;
    constructor Create(const settings: string); overload;
    procedure LoadSettings(const settings: string);
    function GetSettings: string;
  end;

implementation

uses
  System.Math, System.TypInfo, uRtfHash, uRtfMessages;


{ TRtfHtmlStyle }

constructor TRtfHtmlStyle.Create(aDefault: boolean = false);
begin
  fDefault := aDefault;
end;

function TRtfHtmlStyle.GetIsEmpty: Boolean;
begin
  Result := Equals(Empty);
end;

function TRtfHtmlStyle.Equals(Obj: TObject): Boolean;
var
  Compare: TRtfHtmlStyle;
begin
  if Obj = Self then
    Exit(True);
  if (Obj = nil) or (Self.ClassType <> Obj.ClassType) then
    Exit(False);
  Compare := TRtfHtmlStyle(Obj);
  Result := (fForegroundColor = Compare.fForegroundColor) and
    (fBackgroundColor = Compare.fBackgroundColor) and
    (fFontFamily = Compare.fFontFamily) and
    (fFontSize = Compare.fFontSize);
end;

function TRtfHtmlStyle.GetHashCode: Integer;
begin
  Result := AddHashCode(ClassType.ClassName.GetHashCode, ComputeHashCode);
end;

function TRtfHtmlStyle.ComputeHashCode: Integer;
var
  Hash: Integer;
begin
  Hash := fForegroundColor.GetHashCode;
  Hash := AddHashCode(Hash, fBackgroundColor);
  Hash := AddHashCode(Hash, fFontFamily);
  Hash := AddHashCode(Hash, fFontSize);
  Result := Hash;
end;


{ TRtfHtmlCssStyle }

constructor TRtfHtmlCssStyle.Create(const ASelectorName: string;
  const AProperties: string);
begin
  if ASelectorName.IsEmpty then
    raise EArgumentException.Create(sEmptySelectorName);
  fSelectorName := ASelectorName;
  fProperties := TStringList.Create;
  if not AProperties.IsEmpty then
    fProperties.Text := AProperties;
end;

destructor TRtfHtmlCssStyle.Destroy;
begin
  fProperties.Free;
  inherited;
end;

{ TRtfHtmlConvertSettings }

constructor TRtfHtmlConvertSettings.Create;
begin
  Create(TRtfGraphics.Create, hcsAll);
end;

constructor TRtfHtmlConvertSettings.Create(aConvertScope: integer);
begin
  Create(TRtfGraphics.Create, aConvertScope);
end;

constructor TRtfHtmlConvertSettings.Create(aGraphics: TRtfGraphics);
begin
  Create(aGraphics, hcsAll);
end;

constructor TRtfHtmlConvertSettings.Create(aGraphics: TRtfGraphics;
  aConvertScope: integer);
begin
  if aGraphics = nil then
    raise EArgumentNilException.Create(sNilGraphics);
  fStyles := TRtfHtmlCssStyleCollection.Create;
  fStyleSheetLinks := TStringList.Create;
  fGraphics := aGraphics;
  fConvertScope := aConvertScope;
  fVisualHyperlinkPattern := DefaultVisualHyperlinkPattern;
  fDocumentHeader := DefaultDocumentHeader;
  fCharacterSet := DefaultDocumentCharacterSet;
  fStyles.Add(TRtfHtmlCssStyle.Create('p', 'margin=0'));
end;

destructor TRtfHtmlConvertSettings.Destroy;
begin
  fStyles.Free;
  fStyleSheetLinks.Free;
end;

function TRtfHtmlConvertSettings.GetHasStyles: boolean;
begin
  result := fStyles.Count > 0;
end;

function TRtfHtmlConvertSettings.GetHasStyleSheetLinks: boolean;
begin
  result := fStyleSheetLinks.Count > 0;
end;

function TRtfHtmlConvertSettings.GetImageUrl(index: Integer;
  ImageFormat: TRtfImageFormat): string;
var
  FileName: string;
begin
  FileName := fGraphics.ResolveFileName(index, ImageFormat);
  Result := StringReplace(FileName, '\', '/', [rfReplaceAll]);
end;

{ TRtfHtmlElementPath }

constructor TRtfHtmlElementPath.Create;
begin
  fElements := TStack<THtmlTextWriterTag>.Create;
end;

destructor TRtfHtmlElementPath.Destroy;
begin
  fElements.Free;
  inherited;
end;

function TRtfHtmlElementPath.Count: Integer;
begin
  Result := fElements.Count;
end;

function TRtfHtmlElementPath.Current: THtmlTextWriterTag;
begin
  Result := fElements.Peek;
end;

function TRtfHtmlElementPath.IsCurrent(tag: THtmlTextWriterTag): Boolean;
begin
  Result := Current = tag;
end;

function TRtfHtmlElementPath.Contains(tag: THtmlTextWriterTag): Boolean;
var
  a: TArray<THtmlTextWriterTag>;
  i: integer;
begin
  a := fElements.ToArray;
  TArray.Sort<THtmlTextWriterTag>(a);
  Result := TArray.BinarySearch<THtmlTextWriterTag>(a, tag, i);
end;

procedure TRtfHtmlElementPath.Push(tag: THtmlTextWriterTag);
begin
  fElements.Push(tag);
end;

function TRtfHtmlElementPath.Pop: THtmlTextWriterTag;
begin
  result := fElements.Pop;
end;

function TRtfHtmlElementPath.ToString: string;
var
  sb: TStringBuilder;
  first: Boolean;
  element: THtmlTextWriterTag;
begin
  if fElements.Count = 0 then
    Exit(inherited ToString);
  sb := TStringBuilder.Create;
  try
    first := true;
    for element in fElements do
    begin
      if not first then
        sb.Insert(0, ' > ');
      sb.Insert(0, GetEnumName(TypeInfo(THtmlTextWriterTag), Ord(element)));
      first := false;
    end;
    Result := sb.ToString;
  finally
    sb.Free;
  end;
end;


{ TRtfHtmlSpecialCharCollection }

constructor TRtfHtmlSpecialCharCollection.Create;
begin
  inherited Create;
end;

constructor TRtfHtmlSpecialCharCollection.Create(const settings: string);
begin
  Create;
  LoadSettings(settings);
end;

procedure TRtfHtmlSpecialCharCollection.LoadSettings(const settings: string);
var
  items, tokens: TArray<string>;
  item: string;
  kind: TRtfVisualSpecialCharKind;
begin
  Clear;
  if settings = '' then
    exit;

  items := settings.Split([',']);
  for item in items do
  begin
    tokens := item.Split(['=']);
    if Length(tokens) <> 2 then
      continue;

    kind := TRtfVisualSpecialCharKind(GetEnumValue(TypeInfo(
      TRtfVisualSpecialCharKind), tokens[0]));
    Add(kind, tokens[1]);
  end;
end;

function TRtfHtmlSpecialCharCollection.GetSettings: string;
var
  sb: TStringList;
  kind: TRtfVisualSpecialCharKind;
begin
  if Count = 0 then
    exit('');

  sb := TStringList.Create;
  try
    for kind in Keys do
      sb.Add(Format('%s=%s', [GetEnumName(TypeInfo(TRtfVisualSpecialCharKind),
        Ord(kind)), Items[kind]]));
    Result := sb.DelimitedText;
  finally
    sb.Free;
  end;
end;

initialization
  TRtfHtmlStyle.Empty := TRtfHtmlStyle.Create(true);

finalization
  FreeAndNil(TRtfHtmlStyle.Empty);

end.
