unit uRtfDocument;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, uRtfTypes,
  uRtfObjects, uRtfDocumentInfo, uRtfVisual, uRtfInterpreterContext;

type

  TRtfDocument = class(TObject)
  private
    fRTFVersion: Integer;
    fDefaultFont: TRtfFont;
    fDefaultTextFormat: TRtfTextFormat;
    fFontTable: TRtfFontCollection;
    fColorTable: TRtfColorCollection;
    fGenerator: string;
    fUniqueTextFormats: TRtfTextFormatCollection;
    fDocumentInfo: TRtfDocumentInfo;
    fUserProperties: TRtfDocumentPropertyCollection;
    fVisualContent: TRtfVisualCollection;
  public
    constructor Create(AContext: TRtfInterpreterContext;
      AVisualContent: TRtfVisualCollection); overload;
    constructor Create(ARTFVersion: Integer; ADefaultFont: TRtfFont;
      AFontTable: TRtfFontCollection;
      AColorTable: TRtfColorCollection; const AGenerator: string;
      AUniqueTextFormats: TRtfTextFormatCollection;
      const ADocumentInfo: TRtfDocumentInfo;
      AUserProperties: TRtfDocumentPropertyCollection;
      AVisualContent: TRtfVisualCollection); overload;
    destructor Destroy; override;
    function ToString: string; override;
  published
    property VisualContent: TRtfVisualCollection read fVisualContent;
  end;

implementation

uses uRtfSpec, uRtfMessages;

{ TRtfDocument }

constructor TRtfDocument.Create(AContext: TRtfInterpreterContext;
  AVisualContent: TRtfVisualCollection);
begin
  Create(AContext.RtfVersion, AContext.DefaultFont, AContext.FontTable, AContext.ColorTable,
    AContext.Generator, AContext.UniqueTextFormats, AContext.DocumentInfo, AContext.UserProperties,
    AVisualContent);
end;

constructor TRtfDocument.Create(ARTFVersion: Integer; ADefaultFont: TRtfFont;
  AFontTable: TRtfFontCollection; AColorTable: TRtfColorCollection;
  const AGenerator: string; AUniqueTextFormats: TRtfTextFormatCollection;
  const ADocumentInfo: TRtfDocumentInfo;
  AUserProperties: TRtfDocumentPropertyCollection;
  AVisualContent: TRtfVisualCollection);
begin
  if ARTFVersion <> RtfVersion1 then
    raise EArgumentException.CreateFmt(sUnsupportedRtfVersion, [ARTFVersion]);
  if ADefaultFont = nil then
    raise EArgumentNilException.Create(sNilDefaultFont);
  if AFontTable = nil then
    raise EArgumentNilException.Create(sNilFontTable);
  if AColorTable = nil then
    raise EArgumentNilException.Create(sNilColorTable);
  if AUniqueTextFormats = nil then
    raise EArgumentNilException.Create(sNilUniqueTextFormats);
  if AUserProperties = nil then
    raise EArgumentNilException.Create(sNilUserProperties);
  if AVisualContent = nil then
    raise EArgumentNilException.Create(sNilVisualContent);

  fRTFVersion := ARTFVersion;
  fDefaultFont := ADefaultFont;
  fDefaultTextFormat := TRtfTextFormat.Create(fDefaultFont, DefaultFontSize);
  fFontTable := TRtfFontCollection.Create(AFontTable);
  fColorTable := TRtfColorCollection.Create(AColorTable);
  fGenerator := AGenerator;
  fUniqueTextFormats := TRtfTextFormatCollection.Create(AUniqueTextFormats);
  fDocumentInfo := ADocumentInfo;
  fUserProperties := TRtfDocumentPropertyCollection.Create(AUserProperties);
  fVisualContent := TRtfVisualCollection.Create(AVisualContent);
end;

destructor TRtfDocument.Destroy;
begin
  fDefaultTextFormat.Free;
  fVisualContent.Free;
  fColorTable.Free;
  fFontTable.Free;
  fUserProperties.Free;
  fUniqueTextFormats.Free;
  inherited;
end;

function TRtfDocument.ToString: string;
begin
  Result := 'RTFv' + IntToStr(fRTFVersion);
end;

end.


