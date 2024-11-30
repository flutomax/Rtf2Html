unit uRtf2Html;

interface

uses
  System.SysUtils, System.Classes, uRtfTypes, uRtfElement, uRtfGraphics,
  uRtfDocument, uRtfHtmlObjects;

type

  TLogEvent = procedure(Sender: TObject; const Text: string) of object;

  TRtf2Html = class(TComponent)
  private
    fOutputFolder: string;
    fTitle: string;
    fImagesPath: string;
    fStyleSheets: string;
    fStringStream: TStringStream;
    fOnLogMessage: TLogEvent;
    fShowHiddenText: boolean;
    fUseNonBreakingSpaces: boolean;
    fConvertScope: integer;
    fConvertVisualHyperlinks: boolean;
    fVisualHyperlinkPattern: string;
    fSpecialCharsRepresentation: string;
    fOutputFileName: string;
    fGenerator: string;
    procedure Add2Log(const Text: string);
    procedure StructureBuilderParseBegin(Sender: TObject);
    procedure StructureBuilderParseEnd(Sender: TObject);
    procedure StructureBuilderParseSuccess(Sender: TObject);
    procedure StructureBuilderParseFail(const s: string);
    function ParseRtf(out RtfGroup: TRtfGroup): boolean;
    function InterpretRtf(RtfGroup: TRtfGroup;
      Graphics: TRtfGraphics; out RtfDocument: TRtfDocument): boolean;
    function ConvertHmtl(RtfDocument: TRtfDocument; Graphics: TRtfGraphics;
      out html: string): boolean;
    procedure SetOutputFileName(const Value: string);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(Stream: TStream);
    procedure LoadFromString(const data: string);
    procedure Convert;
  published
    property ConvertScope: integer read fConvertScope write fConvertScope;
    property ConvertVisualHyperlinks: boolean read fConvertVisualHyperlinks
      write fConvertVisualHyperlinks;
    property Generator: string read fGenerator write fGenerator;
    property ImagesPath: string read fImagesPath write fImagesPath;
    property Title: string read fTitle write fTitle;
    property OutputFolder: string read fOutputFolder write fOutputFolder;
    property OutputFileName: string read fOutputFileName write SetOutputFileName;
    property ShowHiddenText: boolean read fShowHiddenText write fShowHiddenText;
    property SpecialCharsRepresentation: string read fSpecialCharsRepresentation
      write fSpecialCharsRepresentation;
    property StyleSheets: string read fStyleSheets write fStyleSheets;
    property UseNonBreakingSpaces: boolean read fUseNonBreakingSpaces
      write fUseNonBreakingSpaces;
    property VisualHyperlinkPattern: string read fVisualHyperlinkPattern
      write fVisualHyperlinkPattern;
    property OnLogMessage: TLogEvent read fOnLogMessage write fOnLogMessage;
  end;


implementation

uses
  System.StrUtils, System.IOUtils, uRtfParser, uRtfMessages, uRtfParserListener,
  uRtfInterpreter, uRtfHtmlConverter;

{$IFDEF DUMP_DOCUMENT}
procedure DumpDocument(RtfDocument: TRtfDocument);
var
  i: integer;
  sl: TStringList;
begin
  sl := TStringList.Create;
  try
    for i := 0 to RtfDocument.VisualContent.Count - 1 do
      sl.Add(RtfDocument.VisualContent[i].ToString);
    sl.SaveToFile('C:\dump.txt');
  finally
    sl.Free;
  end;
end;
{$ENDIF}

{ TRtf2Html }

constructor TRtf2Html.Create(AOwner: TComponent);
begin
  inherited;
  fStringStream := TStringStream.Create;
  fConvertVisualHyperlinks := true;
  fUseNonBreakingSpaces := false;
  fShowHiddenText := false;
  fVisualHyperlinkPattern := DefaultVisualHyperlinkPattern;
  fConvertScope := hcsAll;
  fTitle := 'Untitled';
  fGenerator := 'Rtf2Html';
end;

destructor TRtf2Html.Destroy;
begin
  fStringStream.Free;
  inherited;
end;

procedure TRtf2Html.Add2Log(const Text: string);
begin
  if Assigned(fOnLogMessage) then
    fOnLogMessage(self, Text);
end;

procedure TRtf2Html.SetOutputFileName(const Value: string);
begin
  fOutputFileName := Value;
  if fOutputFolder.IsEmpty then
    fOutputFolder := ExtractFilePath(fOutputFileName);
end;

procedure TRtf2Html.StructureBuilderParseBegin(Sender: TObject);
begin
  Add2Log(sInfParseBegin);
end;

procedure TRtf2Html.StructureBuilderParseEnd(Sender: TObject);
begin
  Add2Log(sInfParseEnd);
end;

procedure TRtf2Html.StructureBuilderParseFail(const s: string);
begin
  Add2Log(sInfParseFail + s);
end;

procedure TRtf2Html.StructureBuilderParseSuccess(Sender: TObject);
begin
  Add2Log(sInfParseSucc);
end;

procedure TRtf2Html.LoadFromFile(const FileName: string);
var
  fs: TFileStream;
begin
  fs := TFileStream.Create(FileName, fmOpenRead or fmShareDenyNone);
  try
    LoadFromStream(fs);
  finally
    fs.Free;
  end;
end;

procedure TRtf2Html.LoadFromStream(Stream: TStream);
var
  StringStream: TStringStream;
begin
  Stream.Seek(0, 0);
  fStringStream.Clear;
  fStringStream.LoadFromStream(Stream);
end;

procedure TRtf2Html.LoadFromString(const data: string);
begin
  fStringStream.Clear;
  fStringStream.WriteString(data);
end;


function TRtf2Html.ParseRtf(out RtfGroup: TRtfGroup): boolean;
var
  RtfParser: TRtfParser;
  StructureBuilder: TRtfParserListenerStructureBuilder;
  Reader: TStringReader;
begin
  RtfGroup := nil;
  result := false;
  StructureBuilder := TRtfParserListenerStructureBuilder.Create;
  StructureBuilder.OnParseBegin := StructureBuilderParseBegin;
  StructureBuilder.OnParseEnd := StructureBuilderParseEnd;
  StructureBuilder.OnParseSuccess := StructureBuilderParseSuccess;
  StructureBuilder.OnParseFail := StructureBuilderParseFail;
  RtfParser := TRtfParser.Create([StructureBuilder]);
  Reader := TStringReader.Create(fStringStream.DataString);
  try
    try
      RtfParser.Parse(Reader);
    except
      on E: Exception do
      begin
        Add2Log(sErrInterpreting + E.Message);
        exit;
      end;
    end;
    RtfGroup := (StructureBuilder as TRtfParserListenerStructureBuilder).StructureRoot;
  finally
    RtfParser.Free;
    Reader.Free;
    StructureBuilder.Free;
    fStringStream.Clear;
  end;
  result := true;
end;

function TRtf2Html.InterpretRtf(RtfGroup: TRtfGroup;
  Graphics: TRtfGraphics; out RtfDocument: TRtfDocument): boolean;
var
  Converter: TRtfGraphicsConverter;
  ConverterSettings: TRtfGraphicsConvertSettings;
  InterpreterSettings: TRtfInterpreterSettings;
begin
  result := false;
  Add2Log(sInfInterprBegin);
  ConverterSettings := TRtfGraphicsConvertSettings.Create(Graphics);
  ConverterSettings.ImagesPath := IfThen(fImagesPath.IsEmpty, fOutputFolder, fImagesPath);
  Converter := TRtfGraphicsConverter.Create(ConverterSettings);
  try
    try
      RtfDocument := BuildDoc(RtfGroup, InterpreterSettings, [Converter]);
    except
      on E: Exception do
      begin
        Add2Log(sErrInterpreting + E.Message);
        exit;
      end;
    end;
  finally
    Converter.Free;
  end;
  Add2Log(sInfInterprEnd);
  result := true;
end;

procedure TRtf2Html.Convert;
var
  RtfGroup: TRtfGroup;
  Graphics: TRtfGraphics;
  RtfDocument: TRtfDocument;
  OutStream: TStringStream;
  html: string;
begin
  if fOutputFileName.IsEmpty then
    raise EArgumentException.Create(sEmptyFileName);
  if not DirectoryExists(fOutputFolder) then
    raise EDirectoryNotFoundException.CreateFmt(sDirectoryNotFound, [fOutputFolder]);
  RtfGroup := nil;
  RtfDocument := nil;
  if not ParseRtf(RtfGroup) then
    exit;
  Graphics := TRtfGraphics.Create(
    TPath.GetFileNameWithoutExtension(fOutputFileName) + DefaultFileNamePattern);
  try
    if not InterpretRtf(RtfGroup, Graphics, RtfDocument) then
      exit;
    if not ConvertHmtl(RtfDocument, Graphics, html) then
      exit;
{$IFDEF DUMP_DOCUMENT}
    DumpDocument(RtfDocument);
{$ENDIF}
  finally
    RtfGroup.Free;
    Graphics.Free;
    RtfDocument.Free;
  end;
  TFile.WriteAllText(fOutputFileName, html);
end;

function TRtf2Html.ConvertHmtl(RtfDocument: TRtfDocument;
  Graphics: TRtfGraphics; out html: string): boolean;
var
  ConvertSettings: TRtfHtmlConvertSettings;
  RtfHtmlConverter: TRtfHtmlConverter;
  Sheets: TArray<string>;
  i: integer;
begin
  result := false;
  Add2Log(sInfConvBegin);
  ConvertSettings := TRtfHtmlConvertSettings.Create(Graphics);
  try
    ConvertSettings.CharacterSet := DefaultDocumentCharacterSet;
    ConvertSettings.Title := fTitle;
    ConvertSettings.ImagesPath := IfThen(fImagesPath.IsEmpty, fOutputFolder, fImagesPath);
    ConvertSettings.IsShowHiddenText := fShowHiddenText;
    ConvertSettings.UseNonBreakingSpaces := fUseNonBreakingSpaces;
    if fConvertScope <> hcsNone then
      ConvertSettings.ConvertScope := fConvertScope;
    if not fStyleSheets.IsEmpty then
    begin
      Sheets := fStyleSheets.Split([',']);
      for i := Low(Sheets) to High(Sheets) do
        ConvertSettings.StyleSheetLinks.Add(Sheets[i]);
    end;
    ConvertSettings.ConvertVisualHyperlinks := fConvertVisualHyperlinks;
    if not fVisualHyperlinkPattern.IsEmpty then
      ConvertSettings.VisualHyperlinkPattern := fVisualHyperlinkPattern;
    ConvertSettings.SpecialCharsRepresentation := fSpecialCharsRepresentation;
    RtfHtmlConverter := TRtfHtmlConverter.Create(RtfDocument, ConvertSettings);
    try
      RtfHtmlConverter.Generator := fGenerator;
      try
        html := RtfHtmlConverter.Convert;
        Add2Log(sInfConvEnd);
      except
        on E: Exception do
        begin
          Add2Log(sErrConverting + E.Message);
          exit;
        end;
      end;
    finally
      RtfHtmlConverter.Free;
    end;
  finally
    ConvertSettings.Free;
  end;
  result := true;
  Add2Log(sInfConvSucc);
end;

end.
