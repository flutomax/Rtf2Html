unit uRtfInterpreterListener;

interface

uses
  System.SysUtils, System.Classes, uRtfTypes, uRtfVisual, uRtfObjects,
  uRtfDocument, uRtfInterpreterContext, uRtfNullable;

type

  TRtfInterpreterListener = class(TObject)
  protected
    procedure DoBeginDocument(context: TRtfInterpreterContext); virtual;
    procedure DoInsertText(context: TRtfInterpreterContext; const text: string); virtual;
    procedure DoInsertHyperlink(context: TRtfInterpreterContext; const Data: TRtfHrefData); virtual;
    procedure DoInsertSpecialChar(context: TRtfInterpreterContext; kind: TRtfVisualSpecialCharKind); virtual;
    procedure DoInsertBreak(context: TRtfInterpreterContext; kind: TRtfVisualBreakKind); virtual;
    procedure DoInsertImage(AContext: TRtfInterpreterContext; AFormat: TRtfImageFormat;
      AWidth, AHeight, ADesiredWidth, ADesiredHeight, AScaleWidthPercent,
      AScaleHeightPercent: Integer; AImageDataHex: string); virtual;
    procedure DoHandleTable(context: TRtfInterpreterContext; kind: TRtfVisualTableKind;
      value: Integer); virtual;
    procedure DoEndDocument(context: TRtfInterpreterContext); virtual;
  public
    procedure BeginDocument(context: TRtfInterpreterContext);
    procedure InsertText(context: TRtfInterpreterContext; const text: string);
    procedure InsertHyperlink(context: TRtfInterpreterContext; const Data: TRtfHrefData);
    procedure InsertSpecialChar(context: TRtfInterpreterContext; kind: TRtfVisualSpecialCharKind);
    procedure InsertBreak(context: TRtfInterpreterContext; kind: TRtfVisualBreakKind);
    procedure InsertImage(context: TRtfInterpreterContext; format: TRtfImageFormat;
      width, height, desiredWidth, desiredHeight: Integer;
      scaleWidthPercent, scaleHeightPercent: Integer; const imageDataHex: string);
    procedure HandleTable(context: TRtfInterpreterContext; kind: TRtfVisualTableKind; value: Integer);
    procedure EndDocument(context: TRtfInterpreterContext);
  end;

  TRtfInterpreterListenerDocumentBuilder = class(TRtfInterpreterListener)
  private
    fDocument: TRtfDocument;
    fCombineTextWithSameFormat: Boolean;
    fVisualDocumentContent: TRtfVisualCollection;
    fPendingParagraphContent: TRtfVisualCollection;
    fPendingTextFormat: TRtfTextFormat;
    fPendingText: TStringBuilder;
    fPendingIndent: TRtfIndent;
    fPendingHrefData: TRtfHrefData;
    fCurCellDefs: TRtfTableCellDefs;
    fCurCellDef: TRtfTableCellDef;
    fCurCell: TRtfTableCell;
    fCurRow: TRtfTableRow;
    fCurTable: TRtfVisualTable;
    fCurCellObjList: TList;
    fLastRowLeft: TIntNullable;
    fLastRowHeight: TIntNullable;
    fInTable: Boolean;
    fPapInTbl: Boolean;
    fTableNesting: Boolean;
    procedure FlushPendingText;
    procedure AppendAlignedVisual(AVisual: TRtfVisual);
    procedure EndParagraph(AContext: TRtfInterpreterContext);
  protected
    procedure DoBeginDocument(AContext: TRtfInterpreterContext); override;
    procedure DoInsertText(AContext: TRtfInterpreterContext; const AText: string); override;
    procedure DoInsertHyperlink(context: TRtfInterpreterContext; const Data: TRtfHrefData); override;
    procedure DoInsertSpecialChar(AContext: TRtfInterpreterContext; AKind: TRtfVisualSpecialCharKind); override;
    procedure DoInsertBreak(AContext: TRtfInterpreterContext; AKind: TRtfVisualBreakKind); override;
    procedure DoInsertImage(AContext: TRtfInterpreterContext; AFormat: TRtfImageFormat;
      AWidth, AHeight, ADesiredWidth, ADesiredHeight, AScaleWidthPercent,
      AScaleHeightPercent: Integer; AImageDataHex: string); override;
    procedure DoHandleTable(AContext: TRtfInterpreterContext; kind: TRtfVisualTableKind;
      value: Integer); override;
    procedure DoEndDocument(AContext: TRtfInterpreterContext); override;
  public
    constructor Create;
    destructor Destroy; override;
    property CombineTextWithSameFormat: Boolean read fCombineTextWithSameFormat write fCombineTextWithSameFormat;
    property Document: TRtfDocument read fDocument;
  end;

implementation

uses
  System.Math;

{ TRtfInterpreterListener }

procedure TRtfInterpreterListener.BeginDocument(context: TRtfInterpreterContext);
begin
  if Assigned(context) then
    DoBeginDocument(context);
end;

procedure TRtfInterpreterListener.HandleTable(context: TRtfInterpreterContext;
  kind: TRtfVisualTableKind; value: Integer);
begin
  if Assigned(context) then
    DoHandleTable(context, kind, value);
end;

procedure TRtfInterpreterListener.InsertText(context: TRtfInterpreterContext;
  const text: string);
begin
  if Assigned(context) then
    DoInsertText(context, text);
end;

procedure TRtfInterpreterListener.InsertHyperlink(context: TRtfInterpreterContext;
  const Data: TRtfHrefData);
begin
  if Assigned(context) then
    DoInsertHyperlink(context, Data);
end;

procedure TRtfInterpreterListener.InsertSpecialChar(
  context: TRtfInterpreterContext; kind: TRtfVisualSpecialCharKind);
begin
  if Assigned(context) then
    DoInsertSpecialChar(context, kind);
end;

procedure TRtfInterpreterListener.InsertBreak(context: TRtfInterpreterContext;
  kind: TRtfVisualBreakKind);
begin
  if Assigned(context) then
    DoInsertBreak(context, kind);
end;

procedure TRtfInterpreterListener.InsertImage(context: TRtfInterpreterContext;
  format: TRtfImageFormat; width, height, desiredWidth, desiredHeight: Integer;
  scaleWidthPercent, scaleHeightPercent: Integer; const imageDataHex: string);
begin
  if Assigned(context) then
    DoInsertImage(context, format, width, height, desiredWidth, desiredHeight,
      scaleWidthPercent, scaleHeightPercent, imageDataHex);
end;

procedure TRtfInterpreterListener.EndDocument(context: TRtfInterpreterContext);
begin
  if Assigned(context) then
    DoEndDocument(context);
end;

procedure TRtfInterpreterListener.DoBeginDocument(context: TRtfInterpreterContext);
begin
end;

procedure TRtfInterpreterListener.DoHandleTable(context: TRtfInterpreterContext;
  kind: TRtfVisualTableKind; value: Integer);
begin
end;

procedure TRtfInterpreterListener.DoInsertText(context: TRtfInterpreterContext;
  const text: string);
begin
end;

procedure TRtfInterpreterListener.DoInsertHyperlink(context: TRtfInterpreterContext;
  const Data: TRtfHrefData);
begin
end;

procedure TRtfInterpreterListener.DoInsertSpecialChar(context: TRtfInterpreterContext;
  kind: TRtfVisualSpecialCharKind);
begin
end;

procedure TRtfInterpreterListener.DoInsertBreak(context: TRtfInterpreterContext;
  kind: TRtfVisualBreakKind);
begin
end;

procedure TRtfInterpreterListener.DoInsertImage(AContext: TRtfInterpreterContext;
  AFormat: TRtfImageFormat; AWidth, AHeight, ADesiredWidth, ADesiredHeight,
  AScaleWidthPercent, AScaleHeightPercent: Integer; AImageDataHex: string);
begin
end;

procedure TRtfInterpreterListener.DoEndDocument(context: TRtfInterpreterContext);
begin
end;

{ TRtfInterpreterListenerDocumentBuilder }

constructor TRtfInterpreterListenerDocumentBuilder.Create;
begin
  inherited Create;
  fVisualDocumentContent := nil;
  fCombineTextWithSameFormat := true;
  fPendingParagraphContent := TRtfVisualCollection.Create(false);
  fPendingTextFormat := nil;
  fPendingText := TStringBuilder.Create;
  fPendingIndent := TRtfIndent.Create;
  fPendingHrefData.Reset;
  fCurCell := TRtfTableCell.Create;
  fCurCellDef := TRtfTableCellDef.Create;
  fCurCellDefs := nil;
  fCurRow := TRtfTableRow.Create;
  fCurTable := TRtfVisualTable.Create;
  fCurCellObjList := TList.Create;
  fLastRowLeft := 0;
  fLastRowHeight := nil;
  fInTable := false;
  fPapInTbl := false;
  fTableNesting := false;
end;

destructor TRtfInterpreterListenerDocumentBuilder.Destroy;
begin
  fPendingText.Free;
  fPendingIndent.Free;
  fPendingTextFormat.Free;
  fVisualDocumentContent.Free;
  fPendingParagraphContent.Free;
  fCurCellObjList.Free;
  fCurCell.Free;
  fCurCellDef.Free;
  fCurRow.Free;
  fCurTable.Free;
  inherited Destroy;
end;

procedure TRtfInterpreterListenerDocumentBuilder.DoBeginDocument(
  AContext: TRtfInterpreterContext);
begin
  fDocument := nil;
  fVisualDocumentContent := TRtfVisualCollection.Create(false);
end;

procedure TRtfInterpreterListenerDocumentBuilder.DoHandleTable(
  AContext: TRtfInterpreterContext; kind: TRtfVisualTableKind; value: Integer);

  procedure HandleRow;
  begin
    if fCurRow.Cells.Count > 0 then
    begin
      fCurRow.CellDefs := fCurCellDefs;
      if not fCurRow.Left.HasValue then
        fCurRow.Left := fLastRowLeft;
      if not fCurRow.Height.HasValue then
        fCurRow.Height := fLastRowHeight;
      fCurTable.Add(fCurRow);
      fCurRow := TRtfTableRow.Create;
    end;
    fInTable := true;
  end;

begin

  case kind of
    rvtReset:
      fPapInTbl := false;

    rvtTrowd:
    begin
      if not fCurTable.CellDefsExists(fCurCellDefs) then
        FreeAndNil(fCurCellDefs);
      fCurCellDefs := TRtfTableCellDefs.Create;
      HandleRow;
    end;

    rvtRow:
      HandleRow;

    rvtCell:
    begin
      fCurRow.Cells.Add(TRtfTableCell.Create(fCurCell));
      FreeAndNil(fCurCell);
      fCurCell := TRtfTableCell.Create;
    end;

    rvtCellx:
    begin
      fCurCellDef.Right := value;
      fCurCellDefs.Add(fCurCellDef);
      fCurCellDef := TRtfTableCellDef.Create;
    end;

    rvtRowLeft:
    begin
      fCurRow.Left := value;
      fLastRowLeft := value;
    end;

    rvtRowHeight:
    begin
      fCurRow.Height := value;
      fLastRowHeight := value;
    end;

    rvtCellFirstMerged: fCurCellDef.FirstMerged := true;
    rvtCellMerged: fCurCellDef.Merged := true;

    rvtCellBorderBottom:
    begin
      fCurCellDef.BorderVisible[abBottom] := true;
      fCurCellDef.ActiveBorder := abBottom;
    end;

    rvtCellBorderTop:
    begin
      fCurCellDef.BorderVisible[abTop] := true;
      fCurCellDef.ActiveBorder := abTop;
    end;

    rvtCellBorderLeft:
    begin
      fCurCellDef.BorderVisible[abLeft] := true;
      fCurCellDef.ActiveBorder := abLeft;
    end;

    rvtCellBorderRight:
    begin
      fCurCellDef.BorderVisible[abRight] := true;
      fCurCellDef.ActiveBorder := abRight;
    end;

    rvtCellBorderColor:
      if InRange(value, 0, AContext.ColorTable.Count - 1) then
        fCurCellDef.BorderColor[fCurCellDef.ActiveBorder] :=
          AContext.ColorTable[value];

    rvtCellBorderWidth:
      fCurCellDef.BorderWidth[fCurCellDef.ActiveBorder] := value;

    rvtCellBorderNone: fCurCellDef.ActiveBorder := abNone;
    rvtCellVerticalAlignTop: fCurCellDef.VAlign := tvaTop;
    rvtCellVerticalAlignCenter: fCurCellDef.VAlign := taCenter;
    rvtCellVerticalAlignBottom: fCurCellDef.VAlign := taBottom;
    rvtTableInTable: fPapInTbl := true;
    rvtTableCellBackgroundColor:
      if InRange(value, 0, AContext.ColorTable.Count - 1) then
        fCurCellDef.BackgroundColor := AContext.ColorTable[value];
    rvtTableNestingLevel: fTableNesting := true;
  end;
end;

procedure TRtfInterpreterListenerDocumentBuilder.DoInsertText(
  AContext: TRtfInterpreterContext; const AText: string);
var
  NewFormat: TRtfTextFormat;
  TableText: TRtfVisualText;
begin
  if fInTable then
  begin
    TableText := TRtfVisualText.Create(Atext, AContext.GetSafeCurrentTextFormat,
      AContext.Indent, AContext.HrefData, true);
    if fTableNesting then
      fCurCellObjList.Add(TableText)
    else
      fCurCell.AddVisualObject(TableText);
  end
  else
  if fCombineTextWithSameFormat then
  begin
    NewFormat := AContext.GetSafeCurrentTextFormat;
    if not NewFormat.Equals(fPendingTextFormat) then
      FlushPendingText;
    if fPendingTextFormat = nil then
      fPendingTextFormat := TRtfTextFormat.Create(NewFormat)
    else
      fPendingTextFormat.Assign(NewFormat);
    fPendingText.Append(AText);
    fPendingIndent.Assign(AContext.Indent);
  end
  else
    AppendAlignedVisual(TRtfVisualText.Create(AText, AContext.CurrentTextFormat,
      AContext.Indent, AContext.HrefData));
end;

procedure TRtfInterpreterListenerDocumentBuilder.DoInsertSpecialChar(
  AContext: TRtfInterpreterContext; AKind: TRtfVisualSpecialCharKind);
var
  ch: TRtfVisualSpecialChar;
begin
  FlushPendingText;
  ch := TRtfVisualSpecialChar.Create(AKind, fInTable);
  if fInTable then
  begin
    if fTableNesting then
      fCurCellObjList.Add(ch)
    else
      fCurCell.AddVisualObject(ch);
  end
  else
    fVisualDocumentContent.Add(ch);
end;

procedure TRtfInterpreterListenerDocumentBuilder.DoInsertBreak(
  AContext: TRtfInterpreterContext; AKind: TRtfVisualBreakKind);
var
  Brk: TRtfVisualBreak;
begin
  FlushPendingText;
  Brk := TRtfVisualBreak.Create(AKind, fInTable);
  if fInTable and not (AKind in [rvbParagraph, rvbSection]) then
  begin
    if fTableNesting then
      fCurCellObjList.Add(Brk)
    else
      fCurCell.AddVisualObject(Brk);
  end
  else
  begin
    fVisualDocumentContent.Add(Brk);
    if AKind in [rvbParagraph, rvbSection] then
      EndParagraph(AContext);
  end;
end;

procedure TRtfInterpreterListenerDocumentBuilder.DoInsertHyperlink(
  context: TRtfInterpreterContext; const Data: TRtfHrefData);
begin
  if fInTable then
    context.HrefData := Data
  else
    fPendingHrefData := Data;
end;

procedure TRtfInterpreterListenerDocumentBuilder.DoInsertImage(AContext: TRtfInterpreterContext;
  AFormat: TRtfImageFormat; AWidth, AHeight, ADesiredWidth, ADesiredHeight,
  AScaleWidthPercent, AScaleHeightPercent: Integer; AImageDataHex: string);
var
  Image: TRtfVisualImage;
begin
  Image := TRtfVisualImage.Create(AFormat,
    AContext.CurrentTextFormat.Alignment, AWidth, AHeight,
    ADesiredWidth, ADesiredHeight, AScaleWidthPercent,
    AScaleHeightPercent, AImageDataHex, AContext.Indent, fInTable);
  FlushPendingText;
  if fInTable then
  begin
    if fTableNesting then
      fCurCellObjList.Add(Image)
    else
      fCurCell.AddVisualObject(Image);
  end
  else
    AppendAlignedVisual(Image);
end;

procedure TRtfInterpreterListenerDocumentBuilder.DoEndDocument(AContext: TRtfInterpreterContext);
begin
  FlushPendingText;
  EndParagraph(AContext);
  fDocument := TRtfDocument.Create(AContext, fVisualDocumentContent);
  FreeAndNil(fVisualDocumentContent);
end;

procedure TRtfInterpreterListenerDocumentBuilder.EndParagraph(AContext: TRtfInterpreterContext);
var
  FinalParagraphAlignment: TRtfTextAlignment;
  AlignedVisual: TRtfVisual;
  Image: TRtfVisualImage;
  Text: TRtfVisualText;
  CorrectedFormat, CorrectedUniqueFormat: TRtfTextFormat;
  i: integer;
begin
  FinalParagraphAlignment := AContext.CurrentTextFormat.Alignment;
  if fInTable then
  begin
    if fPapInTbl then
    begin
      fCurCell.AddVisualObject(TRtfVisualBreak.Create(rvbParagraph, true));
      for i := 0 to fCurCellObjList.Count - 1 do
        fCurCell.AddVisualObject(fCurCellObjList[i]);
      fCurCellObjList.Clear;
      fPapInTbl := false;
    end
    else
    begin
      fInTable := false;
      AppendAlignedVisual(fCurTable);
      fCurTable := TRtfVisualTable.Create;
    end;
    fTableNesting := false;
  end
  else
    for AlignedVisual in fPendingParagraphContent do
      case AlignedVisual.Kind of
        rvkImage:
        begin
          Image := TRtfVisualImage(AlignedVisual);
          if Image.Alignment <> FinalParagraphAlignment then
            Image.Alignment := FinalParagraphAlignment;
        end;
        rvkText:
        begin
          Text := TRtfVisualText(AlignedVisual);
          if Text.Format.Alignment <> FinalParagraphAlignment then
          begin
            CorrectedFormat := Text.Format.DeriveWithAlignment(FinalParagraphAlignment);
            CorrectedUniqueFormat := AContext.GetUniqueTextFormatInstance(CorrectedFormat);
            Text.Format.Assign(CorrectedUniqueFormat);
          end;
        end;
      end;
  fPendingParagraphContent.Clear;
end;

procedure TRtfInterpreterListenerDocumentBuilder.FlushPendingText;
begin
  if Assigned(fPendingTextFormat) and (fPendingText.Length > 0) then
  begin
    AppendAlignedVisual(TRtfVisualText.Create(fPendingText.ToString,
      fPendingTextFormat, fPendingIndent, fPendingHrefData));
    FreeAndNil(fPendingTextFormat);
    fPendingText.Clear;
    fPendingIndent.Reset;
    fPendingHrefData.Reset;
  end;
end;

procedure TRtfInterpreterListenerDocumentBuilder.AppendAlignedVisual(AVisual: TRtfVisual);
begin
  fVisualDocumentContent.Add(AVisual);
  fPendingParagraphContent.Add(AVisual);
end;

end.
