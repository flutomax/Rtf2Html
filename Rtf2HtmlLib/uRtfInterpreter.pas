unit uRtfInterpreter;

interface

uses
  System.SysUtils, System.Classes, Generics.Collections, Vcl.Graphics,
  uRtfTypes, uRtfVisual, uRtfObjects, uRtfDocumentInfo, uRtfDocument,
  uRtfElement, uRtfBuilders, uRtfInterpreterContext, uRtfInterpreterListener;

type

  TRtfInterpreterSettings = record
    IgnoreDuplicatedFonts: boolean;
    IgnoreUnknownFonts: boolean;
  end;

  TRtfInterpreterBase = class(TRtfElementVisitor)
  private
    fContext: TRtfInterpreterContext;
    fSettings: TRtfInterpreterSettings;
    fListeners: TList<TRtfInterpreterListener>;
    function GetSettings: TRtfInterpreterSettings;
  protected
    procedure DoInterpret(rtfDocument: TRtfGroup); virtual; abstract;
    procedure NotifyBeginDocument;
    procedure NotifyInsertText(const text: string);
    procedure NotifyInsertSpecialChar(kind: TRtfVisualSpecialCharKind);
    procedure NotifyInsertBreak(kind: TRtfVisualBreakKind);
    procedure NotifyInsertImage(format: TRtfImageFormat; width, height,
      desiredWidth, desiredHeight, scaleWidthPercent, scaleHeightPercent: Integer;
      const imageDataHex: string);
    procedure NotifyHandleTable(kind: TRtfVisualTableKind; value: Integer = 0);
    procedure NotifyEndDocument;
    property Context: TRtfInterpreterContext read fContext;
  public
    constructor Create(settings: TRtfInterpreterSettings; listeners: array of TRtfInterpreterListener);
    destructor Destroy; override;
    procedure AddInterpreterListener(listener: TRtfInterpreterListener);
    procedure RemoveInterpreterListener(listener: TRtfInterpreterListener);
    procedure Interpret(rtfDocument: TRtfGroup);
    property Settings: TRtfInterpreterSettings read GetSettings;
  end;

  ERtfInterpreter = class(Exception);

  TRtfInterpreter = class(TRtfInterpreterBase)
  private
    fFontTableBuilder: TRtfFontTableBuilder;
    fColorTableBuilder: TRtfColorTableBuilder;
    fDocumentInfoBuilder: TRtfDocumentInfoBuilder;
    fUserPropertyBuilder: TRtfUserPropertyBuilder;
    fImageBuilder: TRtfImageBuilder;
    fLastGroupWasPictureWrapper: Boolean;
  protected
    procedure InterpretContents(RtfDocument: TRtfGroup);
    procedure VisitChildrenOf(Group: TRtfGroup);
    procedure VisitTag(tag: TRtfTag); override;
    procedure VisitGroup(group: TRtfGroup); override;
    procedure VisitText(text: TRtfText); override;
    procedure DoInterpret(RtfDocument: TRtfGroup); override;
  public
    constructor Create(Settings: TRtfInterpreterSettings;
      Listeners: array of TRtfInterpreterListener);
    destructor Destroy; override;
    class function IsSupportedDocument(RtfDocument: TRtfGroup): Boolean; static;
    class function GetSupportedDocument(RtfDocument: TRtfGroup): TRtfGroup; static;
  end;

function BuildDoc(rtfDocument: TRtfGroup; const settings: TRtfInterpreterSettings;
  const listeners: array of TRtfInterpreterListener): TRtfDocument;

implementation

uses
  System.Math, System.StrUtils, uRtfParserListener, uRtfSpec, uRtfMessages;


function BuildDoc(rtfDocument: TRtfGroup; const settings: TRtfInterpreterSettings;
  const listeners: array of TRtfInterpreterListener): TRtfDocument;
var
  docBuilder: TRtfInterpreterListenerDocumentBuilder;
  allListeners: array of TRtfInterpreterListener;
  interpreter: TRtfInterpreter;
  i: Integer;
begin
  docBuilder := TRtfInterpreterListenerDocumentBuilder.Create;
  try
    if Length(listeners) = 0 then
    begin
      SetLength(allListeners, 1);
      allListeners[0] := docBuilder;
    end
    else
    begin
      SetLength(allListeners, Length(listeners) + 1);
      allListeners[0] := docBuilder;
      for i := 0 to High(listeners) do
        allListeners[i + 1] := listeners[i];
    end;

    interpreter := TRtfInterpreter.Create(settings, allListeners);
    try
      interpreter.Interpret(rtfDocument);
    finally
      interpreter.Free;
    end;
    Result := docBuilder.Document;
  finally
    docBuilder.Free;
  end;
end;


{ TRtfInterpreterBase }

constructor TRtfInterpreterBase.Create(settings: TRtfInterpreterSettings;
  listeners: array of TRtfInterpreterListener);
var
  listener: TRtfInterpreterListener;
begin
  fSettings := settings;
  fContext := TRtfInterpreterContext.Create;
  fListeners := TList<TRtfInterpreterListener>.Create;

  for listener in listeners do
    AddInterpreterListener(listener);
end;

destructor TRtfInterpreterBase.Destroy;
begin
  fListeners.Free;
  fContext.Free;
  inherited;
end;

function TRtfInterpreterBase.GetSettings: TRtfInterpreterSettings;
begin
  Result := fSettings;
end;

procedure TRtfInterpreterBase.AddInterpreterListener(listener: TRtfInterpreterListener);
begin
  if listener = nil then
    raise EArgumentNilException.Create(sNilListener);

  if fListeners.IndexOf(listener) = -1 then
    fListeners.Add(listener);
end;

procedure TRtfInterpreterBase.RemoveInterpreterListener(listener: TRtfInterpreterListener);
begin
  if listener = nil then
    raise EArgumentNilException.Create(sNilListener);

  fListeners.Remove(listener);
end;

procedure TRtfInterpreterBase.Interpret(rtfDocument: TRtfGroup);
begin
  if rtfDocument = nil then
    raise EArgumentNilException.Create(sNilDocument);

  DoInterpret(rtfDocument);
end;

procedure TRtfInterpreterBase.NotifyBeginDocument;
var
  i: Integer;
begin
  for i := 0 to fListeners.Count - 1 do
    fListeners[i].BeginDocument(fContext);
end;

procedure TRtfInterpreterBase.NotifyInsertText(const text: string);
var
  i: Integer;
begin
  for i := 0 to fListeners.Count - 1 do
    fListeners[i].InsertText(fContext, text);
end;

procedure TRtfInterpreterBase.NotifyInsertSpecialChar(kind: TRtfVisualSpecialCharKind);
var
  i: Integer;
begin
  for i := 0 to fListeners.Count - 1 do
    fListeners[i].InsertSpecialChar(fContext, kind);
end;

procedure TRtfInterpreterBase.NotifyInsertBreak(kind: TRtfVisualBreakKind);
var
  i: Integer;
begin
  for i := 0 to fListeners.Count - 1 do
    fListeners[i].InsertBreak(fContext, kind);
end;

procedure TRtfInterpreterBase.NotifyHandleTable(kind: TRtfVisualTableKind;
  value: Integer);
var
  i: Integer;
begin
  for i := 0 to fListeners.Count - 1 do
    fListeners[i].HandleTable(fContext, kind, value);
end;


procedure TRtfInterpreterBase.NotifyInsertImage(format: TRtfImageFormat; width,
  height, desiredWidth, desiredHeight, scaleWidthPercent, scaleHeightPercent: Integer;
  const imageDataHex: string);
var
  i: Integer;
begin
  for i := 0 to fListeners.Count - 1 do
    fListeners[i].InsertImage(fContext, format, width,
      height, desiredWidth, desiredHeight, scaleWidthPercent, scaleHeightPercent,
      imageDataHex);
end;

procedure TRtfInterpreterBase.NotifyEndDocument;
var
  i: Integer;
begin
  for i := 0 to fListeners.Count - 1 do
    fListeners[i].EndDocument(fContext);
end;

{ TRtfInterpreter }

constructor TRtfInterpreter.Create(Settings: TRtfInterpreterSettings;
  Listeners: array of TRtfInterpreterListener);
begin
  inherited Create(Settings, Listeners);
  fFontTableBuilder := TRtfFontTableBuilder.Create(Context.FontTable,
    Settings.IgnoreDuplicatedFonts);
  fColorTableBuilder := TRtfColorTableBuilder.Create(Context.ColorTable);
  fDocumentInfoBuilder := TRtfDocumentInfoBuilder.Create(Context.DocumentInfo);
  fUserPropertyBuilder := TRtfUserPropertyBuilder.Create(Context.UserProperties);
  fImageBuilder := TRtfImageBuilder.Create;
end;

destructor TRtfInterpreter.Destroy;
begin
  fFontTableBuilder.Free;
  fColorTableBuilder.Free;
  fDocumentInfoBuilder.Free;
  fUserPropertyBuilder.Free;
  fImageBuilder.Free;
  inherited;
end;

class function TRtfInterpreter.IsSupportedDocument(RtfDocument: TRtfGroup): Boolean;
begin
  try
    GetSupportedDocument(RtfDocument);
  except
    on E: ERtfInterpreter do
      Exit(False);
  end;
  Result := True;
end;

class function TRtfInterpreter.GetSupportedDocument(RtfDocument: TRtfGroup): TRtfGroup;
var
  FirstElement: TRtfElement;
  FirstTag: TRtfTag;
begin
  if RtfDocument = nil then
    raise EArgumentNilException.Create(sNilDocument);
  if RtfDocument.Contents.Count = 0 then
    raise ERtfEmptyDocument.Create(sEmptyDocument);
  FirstElement := RtfDocument.Contents[0];
  if FirstElement.Kind <> ekTag then
    raise ERtfStructure.Create(sMissingDocumentStartTag);
  FirstTag := TRtfTag(FirstElement);
  if TagRtf <> FirstTag.Name then
    raise ERtfStructure.CreateFmt(sInvalidDocumentStartTag, [TagRtf]);
  if not FirstTag.HasValue then
    raise ERtfUnsupportedStructure.Create(sMissingRtfVersion);
  if FirstTag.ValueAsNumber <> RtfVersion1 then
    raise ERtfUnsupportedStructure.CreateFmt(sUnsupportedRtfVersion,
      [FirstTag.ValueAsNumber]);
  Result := RtfDocument;
end;

procedure TRtfInterpreter.DoInterpret(RtfDocument: TRtfGroup);
begin
  InterpretContents(GetSupportedDocument(RtfDocument));
end;

procedure TRtfInterpreter.InterpretContents(RtfDocument: TRtfGroup);
begin
  fContext.Reset;
  fLastGroupWasPictureWrapper := false;
  NotifyBeginDocument;
  VisitChildrenOf(RtfDocument);
  fContext.State := risEnded;
  NotifyEndDocument;
end;

procedure TRtfInterpreter.VisitChildrenOf(Group: TRtfGroup);
var
  PushedTextFormat: Boolean;
  i: integer;
begin
  PushedTextFormat := false;
  if Context.State = risInDocument then
  begin
    Context.PushCurrentTextFormat;
    PushedTextFormat := True;
  end;
  try
    for i := 0 to Group.Contents.Count - 1 do
      Group.Contents[i].Visit(Self);
  finally
    if PushedTextFormat then
      Context.PopCurrentTextFormat;
  end;
end;

procedure TRtfInterpreter.VisitTag(tag: TRtfTag);
var
  btmp: boolean;
  stmp: string;
  itmp: integer;
  ctmp: TColor;
begin
  if Context.State <> risInDocument then
    if Context.FontTable.Count > 0 then
      if (Context.ColorTable.Count > 0) or (TagViewKind = tag.Name) then
        Context.State := risInDocument;
  case (Context.State) of
    risInit:
      if TagRtf = tag.Name then
      begin
        Context.State := risInHeader;
        Context.RtfVersion := tag.ValueAsNumber;
      end
      else
        raise ERtfStructure.CreateFmt(sInvalidInitTagState, [tag.ToString]);
    risInHeader:
      if tag.Name = TagDefaultFont then
        Context.DefaultFontId := TagFont + tag.ValueAsNumber.ToString;
    risInDocument:
      case IndexStr(tag.Name, [TagPlain, TagParagraphDefaults, TagSectionDefaults,
          TagBold, TagItalic, TagUnderLine, TagUnderLineNone, TagStrikeThrough,
          TagHidden, TagFont, TagFontSize, TagFontSubscript, TagFontSuperscript,
          TagFontNoSuperSub, TagFontDown, TagFontUp, TagAlignLeft, TagAlignCenter,
          TagAlignRight, TagAlignJustify, TagColorBackground,
          TagColorBackgroundWord, TagColorHighlight, TagColorForeground,
          TagSection, TagParagraph, TagLine, TagPage, TagTabulator, TagTilde,
          TagEmDash, TagEnDash, TagEmSpace, TagEnSpace, TagQmSpace, TagBulltet,
          TagLeftSingleQuote, TagRightSingleQuote, TagLeftDoubleQuote,
          TagRightDoubleQuote, TagHyphen, TagUnderscore, TagFirstLineIndent,
          TagLeftIndent, TagRightIndent, TagTableRowDefaults, TagTableRowBreak,
          TagTableCellBreak, TagTableRightCellBoundary, TagTableRowLeft,
          TagTableRowHeight, TagTableCellFirstMerged, TagTableCellMerged,
          TagTableCellBorderBottom, TagTableCellBorderTop, TagTableCellBorderLeft,
          TagTableCellBorderRight, TagBorderNone, TagBorderColor, TagBorderWidth,
          TagTableCellVerticalAlignTop, TagTableCellVerticalAlignCenter,
          TagTableCellVerticalAlignBottom, TagTableInTable,
          TagTableCellBackgroundColor, TagTableNestingLevel]) of

        0: Context.ApplyCurrentTextFormat(Context.CurrentTextFormat.DeriveNormal);
        1, 2: // TagParagraphDefaults, TagSectionDefaults
        begin
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithAlignment(rtaLeft));
          Context.Indent.Reset;
          NotifyHandleTable(rvtReset);
        end;
        3: // TagBold
        begin
          btmp := (not tag.HasValue) or (tag.ValueAsNumber <> 0);
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithBold(btmp));
        end;
        4: // TagItalic
        begin
          btmp := (not tag.HasValue) or (tag.ValueAsNumber <> 0);
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithItalic(btmp));
        end;
        5: // TagUnderLine
        begin
          btmp := (not tag.HasValue) or (tag.ValueAsNumber <> 0);
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithUnderline(btmp));
        end;
        6: // TagUnderLineNone
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithUnderline(false));
        7: // TagStrikeThrough
        begin
          btmp := (not tag.HasValue) or (tag.ValueAsNumber <> 0);
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithStrikeThrough(btmp));
        end;
        8: // TagHidden
        begin
          btmp := (not tag.HasValue) or (tag.ValueAsNumber <> 0);
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithHidden(btmp));
        end;
        9: // TagFont
        begin
          stmp := tag.FullName;
          if (Context.FontTable.ContainsFontWithId(stmp)) then
            Context.ApplyCurrentTextFormat(
              Context.CurrentTextFormat.DeriveWithFont(
              Context.FontTable.FontById[stmp]))
          else
          if Settings.IgnoreUnknownFonts and (Context.FontTable.Count > 0) then
            Context.ApplyCurrentTextFormat(
              Context.CurrentTextFormat.DeriveWithFont(Context.FontTable[0]))
          else
            raise ERtfUndefinedFont.CreateFmt(sUndefinedFont, [stmp]);
        end;
        10: // TagFontSize
        begin
          itmp := tag.ValueAsNumber;
          if itmp >= 0 then
            Context.ApplyCurrentTextFormat(
              Context.CurrentTextFormat.DeriveWithFontSize(itmp))
          else
            raise ERtfInvalidData.CreateFmt(sInvalidFontSize, [itmp]);
        end;
        11: // TagFontSubscript
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithSuperScript(false));
        12: // TagFontSuperscript
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithSuperScript(true));
        13: // TagFontNoSuperSub
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithSuperScript(0));
        14: // TagFontDown
        begin
          itmp := tag.ValueAsNumber;
          if itmp = 0 then
            itmp := 6;
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithSuperScript(-itmp));
        end;
        15: // TagFontUp
        begin
          itmp := tag.ValueAsNumber;
          if itmp = 0 then
            itmp := 6;
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithSuperScript(itmp));
        end;
        16: // TagAlignLeft
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithAlignment(rtaLeft));
        17: // TagAlignCenter
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithAlignment(rtaCenter));
        18: // TagAlignRight
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithAlignment(rtaRight));
        19: // TagAlignJustify
          Context.ApplyCurrentTextFormat(
            Context.CurrentTextFormat.DeriveWithAlignment(rtaJustify));
        20..23: // TagColorBackground .. TagColorForeground
        begin
          itmp := tag.ValueAsNumber;
          if (itmp >= 0) and (itmp < Context.ColorTable.Count) then
          begin
            btmp := TagColorForeground = tag.Name;
            ctmp := Context.ColorTable[itmp];
            if btmp then
              Context.ApplyCurrentTextFormat(
                Context.CurrentTextFormat.DeriveWithForegroundColor(ctmp))
            else
            begin
              if itmp = 0 then
                ctmp := clNone;
              Context.ApplyCurrentTextFormat(
                Context.CurrentTextFormat.DeriveWithBackgroundColor(ctmp));
            end;
          end
          else
            raise ERtfUndefinedColor.CreateFmt(sUndefinedColor, [itmp]);
        end;
        24: NotifyInsertBreak(rvbSection); // TagSection
        25: NotifyInsertBreak(rvbParagraph); // TagParagraph
        26: NotifyInsertBreak(rvbLine); // TagLine
        27: NotifyInsertBreak(rvbPage); // TagPage
        28: NotifyInsertSpecialChar(rvsTabulator); // TagTabulator
        29: NotifyInsertSpecialChar(rvsNonBreakingSpace); // TagTilde
        30: NotifyInsertSpecialChar(rvsEmDash); // TagEmDash
        31: NotifyInsertSpecialChar(rvsEnDash); // TagEnDash
        32: NotifyInsertSpecialChar(rvsEmSpace); // TagEmSpace
        33: NotifyInsertSpecialChar(rvsEnSpace); // TagEnSpace
        34: NotifyInsertSpecialChar(rvsQmSpace); // TagQmSpace
        35: NotifyInsertSpecialChar(rvsBullet);  // TagBullet
        36: NotifyInsertSpecialChar(rvsLeftSingleQuote); // TagLeftSingleQuote
        37: NotifyInsertSpecialChar(rvsRightSingleQuote); // TagRightSingleQuote)
        38: NotifyInsertSpecialChar(rvsLeftDoubleQuote); // TagLeftDoubleQuote
        39: NotifyInsertSpecialChar(rvsRightDoubleQuote); // TagRightDoubleQuote
        40: NotifyInsertSpecialChar(rvsOptionalHyphen); // TagOptionalHyphen
        41: NotifyInsertSpecialChar(rvsNonBreakingHyphen); // TagNonBreakingHyphen
        42: Context.Indent.FirstIndent := tag.ValueAsNumber; // TagFirstLineIndent
        43: Context.Indent.LeftIndent := tag.ValueAsNumber; // TagLeftIndent
        44: Context.Indent.RightIndent := tag.ValueAsNumber; // TagRightIndent
        45: NotifyHandleTable(rvtTrowd); // TagTableRowDefaults,
        46: NotifyHandleTable(rvtRow); // TagTableRowBreak
        47: NotifyHandleTable(rvtCell); // TagTableCellBreak
        48: NotifyHandleTable(rvtCellx, tag.ValueAsNumber); // TagTableRightCellBoundary
        49: NotifyHandleTable(rvtRowLeft, tag.ValueAsNumber); // TagTableRowLeft
        50: NotifyHandleTable(rvtRowHeight, tag.ValueAsNumber); // TagTableRowHeight
        51: NotifyHandleTable(rvtCellFirstMerged); // TagTableCellFirstMerged
        52: NotifyHandleTable(rvtCellMerged); // TagTableCellMerged
        53: NotifyHandleTable(rvtCellBorderBottom); // TagTableCellBorderBottom
        54: NotifyHandleTable(rvtCellBorderTop); // TagTableCellBorderTop
        55: NotifyHandleTable(rvtCellBorderLeft); // TagTableCellBorderLeft
        56: NotifyHandleTable(rvtCellBorderRight); // TagTableCellBorderRight
        57: NotifyHandleTable(rvtCellBorderNone); // TagBorderNone
        58: NotifyHandleTable(rvtCellBorderColor, tag.ValueAsNumber); // TagBorderColor
        59: NotifyHandleTable(rvtCellBorderWidth, tag.ValueAsNumber); // TagBorderWidth
        60: NotifyHandleTable(rvtCellVerticalAlignTop); // TagTableCellVerticalAlignTop
        61: NotifyHandleTable(rvtCellVerticalAlignCenter); // TagTableCellVerticalAlignCenter
        62: NotifyHandleTable(rvtCellVerticalAlignBottom); // TagTableCellVerticalAlignBottom
        63: NotifyHandleTable(rvtTableInTable); // TagTableInTable
        64: NotifyHandleTable(rvtTableCellBackgroundColor, tag.ValueAsNumber); // TagTagTableCellBackgroundColor
        65: NotifyHandleTable(rvtTableNestingLevel, tag.ValueAsNumber); // TagTableNestingLevel
      end;
  end;
end;

procedure TRtfInterpreter.VisitGroup(group: TRtfGroup);
var
  destination, stmp: string;
  generator: TRtfText;
  alt, awu: TRtfGroup;
begin
  destination := group.Destination;
  case Context.State of
    risInit:
      if TagRtf = destination then
        VisitChildrenOf(group)
      else
        raise ERtfStructure.CreateFmt (sInvalidInitGroupState, [destination]);

    risInHeader:
      case IndexStr(destination, [TagFontTable, TagColorTable, TagGenerator,
          TagPlain, TagParagraphDefaults, TagSectionDefaults, TagUnderLineNone,
          string.Empty]) of
        0: fFontTableBuilder.VisitGroup(group);
        1: fColorTableBuilder.VisitGroup(group);
        2:
        begin
          Context.State := risInDocument;
          if group.Contents.Count = 3  then
            generator := group.Contents[2] as TRtfText
          else
            generator := nil;
          if Assigned(generator) then
          begin
            stmp := generator.Text;
            Context.Generator := IfThen(stmp.EndsWith(';'),
              stmp.Substring(0, stmp.Length - 1), stmp);
          end
          else
            raise ERtfInvalidData.CreateFmt(sInvalidGeneratorGroup, [group.ToString]);
        end;
        3..7:
        begin
          Context.State := risInDocument;
          if not group.IsExtensionDestination then
            VisitChildrenOf(group);
        end;
      end;
    risInDocument:
      case IndexStr(destination, [TagUserProperties, TagInfo,
          TagUnicodeAlternativeChoices, TagHeader, TagHeaderFirst, TagHeaderLeft,
          TagHeaderRight, TagFooter, TagFooterFirst, TagFooterLeft, TagFooterRight,
          TagFootnote, TagStyleSheet, TagPictureWrapper,
          TagPictureWrapperAlternative, TagPicture, TagParagraphNumberText,
          TagListNumberText]) of
        0: fUserPropertyBuilder.VisitGroup(group);
        1: fDocumentInfoBuilder.VisitGroup(group);
        2:
        begin
          alt := group.SelectChildGroupWithDestination(TagUnicodeAlternativeUnicode);
          if Assigned(alt) then
            VisitChildrenOf(alt)
          else
          begin
            if group.Contents.Count > 2  then
              awu := group.Contents[2] as TRtfGroup
            else
              awu := nil;
            if Assigned(awu) then
              VisitChildrenOf(awu);
          end;
        end;
        3..12: ;
        // groups we currently ignore, so their content doesn't intermix with
        // the actual document content
        13:
        begin
          VisitChildrenOf(group);
          fLastGroupWasPictureWrapper := true;
        end;
        14:
        begin
          if not fLastGroupWasPictureWrapper then
            VisitChildrenOf(group);
          fLastGroupWasPictureWrapper := false;
        end;
        15:
          with fImageBuilder do
          begin
            VisitGroup(group);
            NotifyInsertImage(Format, Width, Height, DesiredWidth, DesiredHeight,
              ScaleWidthPercent, ScaleHeightPercent, ImageDataHex);
          end;
        16, 17:
        begin
          NotifyInsertSpecialChar(rvsParagraphNumberBegin);
          VisitChildrenOf(group);
          NotifyInsertSpecialChar(rvsParagraphNumberEnd);
        end;
      else
        if not group.IsExtensionDestination then
          VisitChildrenOf(group);
      end;
  end;
end;

procedure TRtfInterpreter.VisitText(text: TRtfText);
begin
  case Context.State of
    risInit:
      raise ERtfStructure.CreateFmt(sInvalidInitTextState, [text.Text]);
    risInHeader:
      // allow spaces in between header tables
      if not string.IsNullOrEmpty(text.Text.Trim) then
        Context.State := risInDocument;
    risInDocument: ; // nothing
  end;
  NotifyInsertText(text.Text);
end;

end.
