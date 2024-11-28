unit uRtfInterpreterContext;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, uRtfTypes,
  uRtfObjects, uRtfDocumentInfo;

type

  TRtfInterpreterState = (risInit, risInHeader, risInDocument, risEnded);

  TRtfInterpreterContext = class(TObject)
  private
    fState: TRtfInterpreterState;
    fRtfVersion: Integer;
    fDefaultFontId: string;
    fFontTable: TRtfFontCollection;
    fColorTable: TRtfColorCollection;
    fGenerator: string;
    fUniqueTextFormats: TRtfTextFormatCollection;
    fTextFormatStack: TStack<TRtfTextFormat>;
    fCurrentTextFormat: TRtfTextFormat;
    fDocumentInfo: TRtfDocumentInfo;
    fUserProperties: TRtfDocumentPropertyCollection;
    fIndent: TRtfIndent;
    function GetDefaultFont: TRtfFont;
    function GetCurrentTextFormat: TRtfTextFormat;
    procedure SetCurrentTextFormat(Value: TRtfTextFormat);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Reset;
    procedure PushCurrentTextFormat;
    procedure PopCurrentTextFormat;
    procedure ApplyCurrentTextFormat(Value: TRtfTextFormat);
    function GetSafeCurrentTextFormat: TRtfTextFormat;
    function GetUniqueTextFormatInstance(templateFormat: TRtfTextFormat): TRtfTextFormat;
  published
    property FontTable: TRtfFontCollection read fFontTable;
    property ColorTable: TRtfColorCollection read fColorTable;
    property State: TRtfInterpreterState read fState write fState;
    property RtfVersion: Integer read fRtfVersion write fRtfVersion;
    property DefaultFontId: string read fDefaultFontId write fDefaultFontId;
    property DefaultFont: TRtfFont read GetDefaultFont;
    property Generator: string read fGenerator write fGenerator;
    property UniqueTextFormats: TRtfTextFormatCollection read fUniqueTextFormats;
    property CurrentTextFormat: TRtfTextFormat read GetCurrentTextFormat write SetCurrentTextFormat;
    property DocumentInfo: TRtfDocumentInfo read fDocumentInfo;
    property Indent: TRtfIndent read fIndent;
    property UserProperties: TRtfDocumentPropertyCollection read fUserProperties;
  end;

implementation

uses uRtfSpec, uRtfParserListener, uRtfMessages;

constructor TRtfInterpreterContext.Create;
begin
  inherited Create;
  fFontTable := TRtfFontCollection.Create(false);
  fColorTable := TRtfColorCollection.Create;
  fUniqueTextFormats := TRtfTextFormatCollection.Create(false);
  fTextFormatStack := TStack<TRtfTextFormat>.Create;
  fUserProperties := TRtfDocumentPropertyCollection.Create(false);
  fIndent := TRtfIndent.Create;
  fDocumentInfo.Reset;
  fCurrentTextFormat := nil;
end;

destructor TRtfInterpreterContext.Destroy;
begin
  fIndent.Free;
  fFontTable.Free;
  fColorTable.Free;
  fUniqueTextFormats.Free;
  fTextFormatStack.Free;
  fUserProperties.Free;
  fCurrentTextFormat.Free;
  inherited Destroy;
end;

procedure TRtfInterpreterContext.Reset;
begin
  fState := risInit;
  fRtfVersion := RtfVersion1;
  fDefaultFontId := 'f0';
  fFontTable.Clear;
  fColorTable.Clear;
  fGenerator := '';
  fUniqueTextFormats.Clear;
  fTextFormatStack.Clear;
  FreeAndNil(fCurrentTextFormat);
  fDocumentInfo.Reset;
  fIndent.Reset;
  fUserProperties.Clear;
end;

function TRtfInterpreterContext.GetDefaultFont: TRtfFont;
var
  DefaultFont: TRtfFont;
begin
  DefaultFont := fFontTable.FontById[fDefaultFontId];
  if Assigned(DefaultFont) then
    Result := DefaultFont
  else
    raise ERtfUndefinedFont.CreateFmt(sInvalidDefaultFont, [fDefaultFontId]);
end;

function TRtfInterpreterContext.GetSafeCurrentTextFormat: TRtfTextFormat;
begin
  if Assigned(fCurrentTextFormat) then
    result := fCurrentTextFormat
  else
    result := GetCurrentTextFormat;
end;

function TRtfInterpreterContext.GetCurrentTextFormat: TRtfTextFormat;
begin
  if fCurrentTextFormat = nil then
    fCurrentTextFormat := TRtfTextFormat.Create(DefaultFont, DefaultFontSize);
  Result := fCurrentTextFormat;
end;

function TRtfInterpreterContext.GetUniqueTextFormatInstance(templateFormat: TRtfTextFormat): TRtfTextFormat;
var
  UniqueInstance: TRtfTextFormat;
  ExistingEquivalentPos: Integer;
begin
  if not Assigned(templateFormat) then
    raise EArgumentNilException.Create(sNilTemplateFormat);

  ExistingEquivalentPos := fUniqueTextFormats.IndexOf(templateFormat);
  if ExistingEquivalentPos >= 0 then
    UniqueInstance := fUniqueTextFormats[ExistingEquivalentPos]
  else
  begin
    fUniqueTextFormats.Add(templateFormat);
    UniqueInstance := templateFormat;
  end;
  Result := UniqueInstance;
end;

procedure TRtfInterpreterContext.ApplyCurrentTextFormat(Value: TRtfTextFormat);
var
  UniqueInstance: TRtfTextFormat;
  ExistingEquivalentPos: Integer;
begin
  if not Assigned(Value) then
    raise EArgumentNilException.Create(sNilTemplateFormat);
  ExistingEquivalentPos := fUniqueTextFormats.IndexOf(Value);
  if ExistingEquivalentPos >= 0 then
  begin
    UniqueInstance := fUniqueTextFormats[ExistingEquivalentPos];
    Value.Free;
  end
  else
  begin
    fUniqueTextFormats.Add(Value);
    UniqueInstance := Value;
  end;
  fCurrentTextFormat.Assign(UniqueInstance);
end;

procedure TRtfInterpreterContext.SetCurrentTextFormat(Value: TRtfTextFormat);
begin
  fCurrentTextFormat.Assign(GetUniqueTextFormatInstance(Value));
end;

procedure TRtfInterpreterContext.PopCurrentTextFormat;
var
  TextFormat: TRtfTextFormat;
begin
  if fTextFormatStack.Count = 0 then
    raise ERtfStructure.Create(sInvalidTextContextState);
  TextFormat := fTextFormatStack.Pop;
  fCurrentTextFormat.Assign(TextFormat);
  TextFormat.Free;
end;

procedure TRtfInterpreterContext.PushCurrentTextFormat;
var
  TextFormat: TRtfTextFormat;
begin
  TextFormat := TRtfTextFormat.Create(CurrentTextFormat);
  fTextFormatStack.Push(TextFormat);
end;

end.
