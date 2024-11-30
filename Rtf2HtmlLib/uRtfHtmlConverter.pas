unit uRtfHtmlConverter;

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.RegularExpressions,
  Vcl.Graphics, uRtfVisual, uRtfHtmlObjects, uRtfDocument, uRtfGraphics,
  uRtfObjects, uRtfHtmlWriter, uRtfTypes, uRtfCssTextWriter;

type

  ERtfHtmlConverter = class(Exception);

  TRtfHtmlConverter = class(TRtfVisualVisitor)
  private
    fDocumentImages: TRtfGraphicsConvertInfoCollection;
    fElementPath: TRtfHtmlElementPath;
    fRtfDocument: TRtfDocument;
    fSettings: TRtfHtmlConvertSettings;
    fSpecialCharacters: TRtfHtmlSpecialCharCollection;
    fWriter: THtmlTextWriter;
    fLastVisual: TRtfVisual;
    fIsInParagraphNumber: Boolean;
    fHyperlinkRegEx: TRegEx;
    fGenerator: string;
    function ConvertVisualHyperlink(text: string): string;
    function EnterVisual(visual: TRtfVisual): boolean;
    function EnsureOpenList(visual: TRtfVisual): boolean;
    procedure LeaveVisual(visual: TRtfVisual);
    procedure EnsureClosedList; overload;
    procedure EnsureClosedList(visual: TRtfVisual); overload;
  protected
    function GetWriter: THtmlTextWriter;
    function GetElementPath: TRtfHtmlElementPath;
    function GetHtmlStyle(visual: TRtfVisual): TRtfHtmlStyle; virtual;
    function FormatHtmlText(Text: string): string; virtual;
    function OnEnterVisual(visual: TRtfVisual): Boolean; virtual;
    function IsInParagraph: Boolean;
    function IsInList: Boolean;
    function IsInListItem: Boolean;
    function GetGenerator: string; virtual;
    function IsCurrentElement(ATag: THtmlTextWriterTag): Boolean;
    function IsInElement(ATag: THtmlTextWriterTag): Boolean;
    procedure OnLeaveVisual(visual: TRtfVisual); virtual;
    procedure BeginParagraph(Indent: TRtfIndent); virtual;
    procedure EndParagraph; virtual;
{$REGION 'Rendering'}
    procedure RenderDocumentSection;
    procedure RenderHtmlSection;
    procedure RenderHeadSection;
    procedure RenderBodySection;
    procedure RenderDocumentHeader; virtual;
    procedure RenderHeadAttributes; virtual;
    procedure RenderMetaContentType; virtual;
    procedure RenderMetaGenerator; virtual;
    procedure RenderLinkStyleSheets; virtual;
    procedure RenderTitle; virtual;
    procedure RenderStyles; virtual;
    procedure RenderRtfContent; virtual;
    procedure RenderAssertTag(ATag: THtmlTextWriterTag);
    procedure RenderBeginTag(ATag: THtmlTextWriterTag);
    procedure RenderEndTag; overload;
    procedure RenderEndTag(ALineBreak: Boolean); overload; virtual;
    procedure RenderTitleTag; virtual;
    procedure RenderMetaTag; virtual;
    procedure RenderHtmlTag; virtual;
    procedure RenderLinkTag; virtual;
    procedure RenderHeadTag; virtual;
    procedure RenderBodyTag; virtual;
    procedure RenderStyleTag; virtual;
    procedure RenderLineBreak; virtual;
    procedure RenderATag; virtual;
    procedure RenderPTag; virtual;
    procedure RenderBTag; virtual;
    procedure RenderITag; virtual;
    procedure RenderUTag; virtual;
    procedure RenderSTag; virtual;
    procedure RenderSubTag; virtual;
    procedure RenderSupTag; virtual;
    procedure RenderSpanTag; virtual;
    procedure RenderImgTag; virtual;
    procedure RenderUlTag; virtual;
    procedure RenderOlTag; virtual;
    procedure RenderLiTag; virtual;
{$ENDREGION}
{$REGION 'RtfVisuals'}
    procedure DoVisitText(aVisualText: TRtfVisualText); override;
    procedure DoVisitImage(aVisualImage: TRtfVisualImage); override;
    procedure DoVisitSpecial(aVisualSpecialChar: TRtfVisualSpecialChar); override;
    procedure DoVisitBreak(aVisualBreak: TRtfVisualBreak); override;
    procedure DoVisitTable(aVisualTable: TRtfVisualTable); override;
{$ENDREGION}
  public
    constructor Create(ARtfDocument: TRtfDocument; ASettings: TRtfHtmlConvertSettings);
    destructor Destroy; override;
    function Convert: string;
  published
    property ElementPath: TRtfHtmlElementPath read fElementPath;
    property RtfDocument: TRtfDocument read fRtfDocument;
    property Settings: TRtfHtmlConvertSettings read fSettings;
    property SpecialCharacters: TRtfHtmlSpecialCharCollection read fSpecialCharacters;
    property DocumentImages: TRtfGraphicsConvertInfoCollection read fDocumentImages;
    property Generator: string read fGenerator write fGenerator;
  end;

implementation

uses
  System.Math, uRtfMessages, uRtfHtmlFunctions;

const
  GeneratorName = 'Rtf2Html';
  NonBreakingSpace = '&nbsp;';
  UnsortedListValue = '·';


{ TRtfHtmlConverter }

constructor TRtfHtmlConverter.Create(ARtfDocument: TRtfDocument; ASettings: TRtfHtmlConvertSettings);
begin
  if ARtfDocument = nil then
    raise EArgumentNilException.Create(sNilDocument);
  if ASettings = nil then
    raise EArgumentNilException.Create(sNilSettings);
  fRtfDocument := ARtfDocument;
  fSettings := ASettings;
  fGenerator := GeneratorName;
  if not fSettings.VisualHyperlinkPattern.IsEmpty then
    fHyperlinkRegEx := TRegEx.Create(fSettings.VisualHyperlinkPattern);
  fSpecialCharacters := TRtfHtmlSpecialCharCollection.Create(fSettings.SpecialCharsRepresentation);
  fDocumentImages := TRtfGraphicsConvertInfoCollection.Create;
  fElementPath := TRtfHtmlElementPath.Create;
end;

destructor TRtfHtmlConverter.Destroy;
begin
  fElementPath.Free;
  fDocumentImages.Free;
  fSpecialCharacters.Free;
  inherited;
end;

procedure TRtfHtmlConverter.EnsureClosedList(visual: TRtfVisual);
var
  prevpar: TRtfVisualBreak;
  cpec: TRtfVisualSpecialChar;
begin
  if not IsInList then
    exit;
  if not (fLastVisual is TRtfVisualBreak) then
    exit;

  prevpar := fLastVisual as TRtfVisualBreak;
  if prevpar.BreakKind <> rvbParagraph then
    exit;

  cpec := nil;
  if (visual is TRtfVisualSpecialChar) then
    cpec := visual as TRtfVisualSpecialChar;

  if (cpec = nil) or (cpec.CharKind <> rvsParagraphNumberBegin) then
    RenderEndTag(true); // close ul/ol list
end;

procedure TRtfHtmlConverter.EnsureClosedList;
begin
  if fLastVisual = nil then
    exit;
  EnsureClosedList(fLastVisual);
end;

procedure TRtfHtmlConverter.OnLeaveVisual(visual: TRtfVisual);
begin
  // nothin
end;

function TRtfHtmlConverter.EnsureOpenList(visual: TRtfVisual): boolean;
var
  Text: TRtfVisualText;
  unsorted: Boolean;
begin
  result := false;
  if (not (visual is TRtfVisualText)) or (not fIsInParagraphNumber) then
    exit;
  Text := visual as TRtfVisualText;
  if not IsInList then
  begin
    unsorted := UnsortedListValue = Text.Text;
    if unsorted then
      RenderUlTag
    else
      RenderOlTag;
  end;
  RenderLiTag;
  result := true;
end;

function TRtfHtmlConverter.ConvertVisualHyperlink(text: string): string;
begin
  result := '';
  if text.IsEmpty or fSettings.VisualHyperlinkPattern.IsEmpty then
    exit;
  if fHyperlinkRegEx.IsMatch(text) then
    result := text;
end;

function TRtfHtmlConverter.EnterVisual(visual: TRtfVisual): boolean;
begin
  result := false;
  if EnsureOpenList(visual) then
    exit;
  EnsureClosedList(visual);
  result := OnEnterVisual(visual);
end;

function TRtfHtmlConverter.GetHtmlStyle(visual: TRtfVisual): TRtfHtmlStyle;
begin
  if visual.Kind = rvkText then
    result := TextToHtml(visual as TRtfVisualText)
  else
    result := TRtfHtmlStyle.Empty;
end;

function TRtfHtmlConverter.OnEnterVisual(visual: TRtfVisual): Boolean;
begin
  result := true;
end;

function TRtfHtmlConverter.GetWriter: THtmlTextWriter;
begin
  Result := fWriter;
end;

function TRtfHtmlConverter.GetElementPath: TRtfHtmlElementPath;
begin
  Result := fElementPath;
end;

function TRtfHtmlConverter.IsInParagraph: Boolean;
begin
  Result := IsInElement(twtP);
end;

function TRtfHtmlConverter.IsInList: Boolean;
begin
  Result := IsInElement(twtUl) or IsInElement(twtOl);
end;

function TRtfHtmlConverter.IsInListItem: Boolean;
begin
  Result := IsInElement(twtLi);
end;

function TRtfHtmlConverter.GetGenerator: string;
begin
  Result := GeneratorName;
end;


function TRtfHtmlConverter.Convert: string;
begin
  fDocumentImages.Clear;

  fWriter := THtmlTextWriter.Create;
  try
    RenderDocumentSection;
    RenderHtmlSection;
    Result := fWriter.ToString;
  finally
    fWriter.Free;
  end;
  {
  if fElementPath.Count <> 0 then
    raise ERtfHtmlConverter.Create('unbalanced element structure');
  }
end;

function TRtfHtmlConverter.IsCurrentElement(ATag: THtmlTextWriterTag): Boolean;
begin
  Result := fElementPath.IsCurrent(ATag);
end;

function TRtfHtmlConverter.IsInElement(ATag: THtmlTextWriterTag): Boolean;
begin
  Result := fElementPath.Contains(ATag);
end;

procedure TRtfHtmlConverter.RenderAssertTag(ATag: THtmlTextWriterTag);
begin
  Assert(fElementPath.IsCurrent(ATag), 'no valid tag');
end;

procedure TRtfHtmlConverter.RenderBeginTag(ATag: THtmlTextWriterTag);
begin
  fWriter.RenderBeginTag(ATag);
  fElementPath.Push(ATag);
end;

procedure TRtfHtmlConverter.RenderEndTag;
begin
  RenderEndTag(False);
end;

procedure TRtfHtmlConverter.RenderEndTag(ALineBreak: Boolean);
begin
  fWriter.RenderEndTag;
  if ALineBreak then
    fWriter.WriteLine;
  fElementPath.Pop;
end;

procedure TRtfHtmlConverter.RenderTitleTag;
begin
  RenderBeginTag(twtTitle);
end;

procedure TRtfHtmlConverter.RenderMetaContentType;
var
  content: string;
begin
  fWriter.AddAttribute('http-equiv', 'content-type');
  content := 'text/html';
  if not fSettings.CharacterSet.IsEmpty then
    content := Concat(content, '; charset=', fSettings.CharacterSet);
  fWriter.AddAttribute(twaContent, content);
  RenderMetaTag;
  RenderEndTag;
end;

procedure TRtfHtmlConverter.RenderMetaGenerator;
begin
  if fGenerator.IsEmpty then
    exit;
  fWriter.WriteLine;
  fWriter.AddAttribute(twaName, 'generator');
  fWriter.AddAttribute(twaContent, fGenerator);
  RenderMetaTag;
  RenderEndTag;
end;

procedure TRtfHtmlConverter.RenderMetaTag;
begin
  RenderBeginTag(twtMeta);
end;

procedure TRtfHtmlConverter.RenderHtmlTag;
begin
  RenderBeginTag(twtHtml);
end;

procedure TRtfHtmlConverter.RenderLinkStyleSheets;
var
  s: string;
begin
  if not fSettings.HasStyleSheetLinks then
    exit;
  for s in fSettings.StyleSheetLinks do
  begin
    if s.IsEmpty then
      continue;
    fWriter.WriteLine;
    fWriter.AddAttribute(twaHref, s);
    fWriter.AddAttribute(twaType, 'text/css');
    fWriter.AddAttribute(twaRel, 'stylesheet');
    RenderLinkTag;
    RenderEndTag;
  end;
end;

procedure TRtfHtmlConverter.RenderTitle;
begin
  if fSettings.Title.IsEmpty then
    exit;
  fWriter.WriteLine;
  RenderTitleTag;
  fWriter.Write(fSettings.Title);
  RenderEndTag;
end;

procedure TRtfHtmlConverter.RenderStyles;
var
  first: Boolean;
  style: TRtfHtmlCssStyle;
  i: integer;
begin
  if not fSettings.HasStyles then
    exit;
  fWriter.WriteLine;
  RenderStyleTag;

  first := true;
  for style in fSettings.Styles do
  begin
    if style.Properties.Count = 0 then
      continue;
    if not first then
      fWriter.WriteLine;

    fWriter.WriteLine(style.SelectorName);
    fWriter.WriteLine('{');
    with style.Properties do
      for i := 0 to Count - 1 do
        fWriter.WriteLine(Format('  %s: %s;', [Names[i], Values[Names[i]]]));
    fWriter.Write('}');
    first := false;
  end;
  RenderEndTag;
end;

procedure TRtfHtmlConverter.RenderStyleTag;
begin
  RenderBeginTag(twtStyle);
end;

procedure TRtfHtmlConverter.RenderLinkTag;
begin
  RenderBeginTag(twtLink);
end;

procedure TRtfHtmlConverter.RenderHeadTag;
begin
  RenderBeginTag(twtHead);
end;

procedure TRtfHtmlConverter.RenderBodyTag;
begin
  RenderBeginTag(twtBody);
end;

procedure TRtfHtmlConverter.RenderLineBreak;
begin
  fWriter.WriteBreak;
  fWriter.WriteLine;
end;

procedure TRtfHtmlConverter.RenderATag;
begin
  RenderBeginTag(twtA);
end;

procedure TRtfHtmlConverter.RenderPTag;
begin
  RenderBeginTag(twtP);
end;

procedure TRtfHtmlConverter.RenderBTag;
begin
  RenderBeginTag(twtB);
end;

procedure TRtfHtmlConverter.RenderDocumentHeader;
begin
  if fSettings.DocumentHeader.IsEmpty then
    exit;
  fWriter.WriteLine(fSettings.DocumentHeader);
end;

procedure TRtfHtmlConverter.RenderDocumentSection;
begin
  if (fSettings.ConvertScope and hcsDocument) <> hcsDocument then
    exit;
  RenderDocumentHeader;
end;

procedure TRtfHtmlConverter.RenderHtmlSection;
begin
  if (fSettings.ConvertScope and hcsHtml) = hcsHtml then
    RenderHtmlTag;
  RenderHeadSection;
  RenderBodySection;
  if (fSettings.ConvertScope and hcsHtml) = hcsHtml then
    RenderEndTag(true);
end;

procedure TRtfHtmlConverter.RenderHeadAttributes;
begin
  RenderMetaContentType;
  RenderMetaGenerator;
  RenderLinkStyleSheets;
end;

procedure TRtfHtmlConverter.RenderHeadSection;
begin
  if (fSettings.ConvertScope and hcsHead) <> hcsHead then
    exit;
  RenderHeadTag;
  RenderHeadAttributes;
  RenderTitle;
  RenderStyles;
  RenderEndTag(true);
end;

procedure TRtfHtmlConverter.RenderBodySection;
begin
  if (fSettings.ConvertScope and hcsBody) = hcsBody then
    RenderBodyTag;

  if (fSettings.ConvertScope and hcsContent) = hcsContent then
    RenderRtfContent;

  if (fSettings.ConvertScope and hcsBody) = hcsBody then
    RenderEndTag;
end;

procedure TRtfHtmlConverter.RenderRtfContent;
var
  visual: TRtfVisual;
begin
  for visual in fRtfDocument.VisualContent do
    visual.Visit(self);
  EnsureClosedList;
end;

procedure TRtfHtmlConverter.RenderImgTag;
begin
  RenderBeginTag(twtImg);
end;

procedure TRtfHtmlConverter.RenderITag;
begin
  RenderBeginTag(twtI);
end;

procedure TRtfHtmlConverter.RenderUTag;
begin
  RenderBeginTag(twtU);
end;

procedure TRtfHtmlConverter.RenderSTag;
begin
  RenderBeginTag(twtS);
end;

procedure TRtfHtmlConverter.RenderSubTag;
begin
  RenderBeginTag(twtSub);
end;

procedure TRtfHtmlConverter.RenderSupTag;
begin
  RenderBeginTag(twtSup);
end;

procedure TRtfHtmlConverter.RenderSpanTag;
begin
  RenderBeginTag(twtSpan);
end;

procedure TRtfHtmlConverter.RenderUlTag;
begin
  RenderBeginTag(twtUl);
end;

procedure TRtfHtmlConverter.RenderOlTag;
begin
  RenderBeginTag(twtOl);
end;

procedure TRtfHtmlConverter.RenderLiTag;
begin
  RenderBeginTag(twtLi);
end;

procedure TRtfHtmlConverter.BeginParagraph(Indent: TRtfIndent);
begin
  if IsInParagraph then
    exit;
  with Indent do
  begin
    if FirstIndent <> 0 then
      fWriter.AddStyleAttribute(twsTextIndent, Format('%dpx', [TwipToPixel(FirstIndent)]));
    if LeftIndent <> 0 then
      fWriter.AddStyleAttribute(twsMarginLeft, Format('%dpx', [TwipToPixel(LeftIndent)]));
    if RightIndent <> 0 then
      fWriter.AddStyleAttribute(twsMarginRight, Format('%dpx', [TwipToPixel(RightIndent)]));
  end;
  RenderPTag;
end;

procedure TRtfHtmlConverter.EndParagraph;
begin
  if not IsInParagraph then
    exit;
  RenderEndTag(true);
end;

function TRtfHtmlConverter.FormatHtmlText(Text: string): string;
begin
  result := HtmlEncodeAttr(Text, fSettings.UseNonBreakingSpaces);
end;

procedure TRtfHtmlConverter.LeaveVisual(visual: TRtfVisual);
begin
  OnLeaveVisual(visual);
  fLastVisual := visual;
end;

procedure TRtfHtmlConverter.DoVisitText(aVisualText: TRtfVisualText);
var
  HtmlStyle: TRtfHtmlStyle;
  TextFormat: TRtfTextFormat;
  IsHyperlink: Boolean;
  href, htmltext: string;
begin
  if not EnterVisual(aVisualText) then
    exit;
  // suppress hidden text
  if (aVisualText.Format.IsHidden and (not settings.IsShowHiddenText)) then
    exit;

  TextFormat := aVisualText.Format;
  if not aVisualText.IsInTable then
    case TextFormat.Alignment of
      rtaLeft: ;   //fWriter.AddStyleAttribute(twsTextAlign, 'left');
      rtaCenter:
        fWriter.AddStyleAttribute(twsTextAlign, 'center');
      rtaRight:
        fWriter.AddStyleAttribute(twsTextAlign, 'right');
      rtaJustify:
        fWriter.AddStyleAttribute(twsTextAlign, 'justify');
    end;

  if (not IsInListItem) and (not aVisualText.IsInTable) then
    BeginParagraph(aVisualText.Indent);
  // visual hyperlink
  IsHyperlink := false;
  if fSettings.ConvertVisualHyperlinks then
  begin
    href := ConvertVisualHyperlink(aVisualText.Text);
    if not href.IsEmpty then
    begin
      IsHyperlink := true;
      fWriter.AddAttribute(twaHref, href);
      RenderATag;
    end;
  end
  else
  // make hyperlink
  if not aVisualText.URL.IsEmpty then
  begin
    IsHyperlink := true;
    fWriter.AddAttribute(twaHref, aVisualText.URL, false);
    RenderATag;
  end;

  // format tags
  if TextFormat.IsBold then
    RenderBTag;

  if TextFormat.IsItalic then
    RenderITag;

  if TextFormat.IsUnderline then
    RenderUTag;

  if TextFormat.IsStrikeThrough then
    RenderSTag;

  // span with style
  HtmlStyle := GetHtmlStyle(aVisualText);
  try
    if not HtmlStyle.IsEmpty then
    begin
      if not HtmlStyle.ForegroundColor.IsEmpty then
        fWriter.AddStyleAttribute(twsColor, HtmlStyle.ForegroundColor);

      if not HtmlStyle.BackgroundColor.IsEmpty then
        fWriter.AddStyleAttribute(twsBackgroundColor, HtmlStyle.BackgroundColor);

      if not HtmlStyle.FontFamily.IsEmpty then
        fWriter.AddStyleAttribute(twsFontFamily, HtmlStyle.FontFamily);

      if not HtmlStyle.FontSize.IsEmpty then
        fWriter.AddStyleAttribute(twsFontSize, HtmlStyle.FontSize);
      RenderSpanTag;
    end;
    // subscript and superscript
    if TextFormat.SuperScript < 0 then
      RenderSubTag
    else if TextFormat.SuperScript > 0 then
      RenderSupTag;
    htmltext := FormatHtmlText(aVisualText.Text);
    fWriter.Write(htmltext);
    // subscript and superscript
    if TextFormat.SuperScript < 0 then
      RenderEndTag // sub
    else if TextFormat.SuperScript > 0 then
      RenderEndTag; // sup
    // span with style
    if not HtmlStyle.IsEmpty then
      RenderEndTag;
    // format tags
    if TextFormat.IsStrikeThrough then
      RenderEndTag; // s
    if TextFormat.IsUnderline then
      RenderEndTag; // u
    if TextFormat.IsItalic then
      RenderEndTag; // i
    if TextFormat.IsBold then
      RenderEndTag; // b
    // visual hyperlink
    if IsHyperlink then
      RenderEndTag; // a
  finally
    if not HtmlStyle.Default then
      FreeAndNil(HtmlStyle);
  end;
  LeaveVisual(aVisualText);
end;

procedure TRtfHtmlConverter.DoVisitImage(aVisualImage: TRtfVisualImage);
var
  ImageIndex, Width, Height: integer;
  FileName, HtmlFileName: string;
begin
  if not EnterVisual(aVisualImage) then
    exit;
  if not aVisualImage.IsInTable then
  begin
    case aVisualImage.Alignment of
      rtaLeft:;
      rtaCenter: fWriter.AddStyleAttribute(twsTextAlign, 'center');
      rtaRight: fWriter.AddStyleAttribute(twsTextAlign, 'right');
      rtajustify: fWriter.AddStyleAttribute(twsTextAlign, 'justify');
    end;

    BeginParagraph(aVisualImage.Indent);
  end;

  ImageIndex := fDocumentImages.Count + 1;
  FileName := fSettings.GetImageUrl(ImageIndex, aVisualImage.ImgFormat);
  Width := fSettings.Graphics.CalcImageWidth(aVisualImage.ImgFormat,
    aVisualImage.Width, aVisualImage.DesiredWidth, aVisualImage.ScaleWidthPercent);
  Height := fSettings.Graphics.CalcImageHeight(aVisualImage.ImgFormat,
    aVisualImage.Height, aVisualImage.DesiredHeight, aVisualImage.ScaleHeightPercent);

  fWriter.AddAttribute(twaWidth, Width.ToString);
  fWriter.AddAttribute(twaHeight, Height.ToString);
  HtmlFileName := EncodeUrl(FileName);
  fWriter.AddAttribute(twaSrc, HtmlFileName, false);
  RenderImgTag;
  RenderEndTag;

  DocumentImages.Add(TRtfGraphicsConvertInfo.Create(HtmlFileName,
    aVisualImage.ImgFormat, TSize.Create(Width, Height)));

  LeaveVisual(aVisualImage);
end;

procedure TRtfHtmlConverter.DoVisitSpecial(aVisualSpecialChar: TRtfVisualSpecialChar);
begin
  if not EnterVisual(aVisualSpecialChar) then
    exit;

  case aVisualSpecialChar.CharKind of
    rvsParagraphNumberBegin: fIsInParagraphNumber := true;
    rvsParagraphNumberEnd: fIsInParagraphNumber := false;
  else
    if fSpecialCharacters.ContainsKey(aVisualSpecialChar.CharKind) then
      fWriter.Write(fSpecialCharacters[aVisualSpecialChar.CharKind]);
  end;

  LeaveVisual(aVisualSpecialChar);
end;

procedure TRtfHtmlConverter.DoVisitBreak(aVisualBreak: TRtfVisualBreak);
begin
  if not EnterVisual(aVisualBreak) then
    exit;
  case aVisualBreak.BreakKind of
    rvbLine: RenderLineBreak;
    rvbParagraph:
      if IsInParagraph then
        EndParagraph
      else if IsInListItem then
      begin
        EndParagraph;
        RenderEndTag(true);
      end
      else
      begin
        BeginParagraph(aVisualBreak.Indent);
        // Uncomment if the height is not enough
        //fWriter.Write(NonBreakingSpace);
        EndParagraph;
      end;
  end;
  LeaveVisual(aVisualBreak);
end;

procedure TRtfHtmlConverter.DoVisitTable(aVisualTable: TRtfVisualTable);

  function FindIf(cells: TRtfTableCellDefs; const value: integer;
    const right_eq: boolean; out cell: TRtfTableCellDef): boolean;
  var
    i: integer;
  begin
    result := false;
    for i := 0 to cells.Count - 1 do
    begin
      cell := cells[i];
      if right_eq then
      begin
        if cell.RightEquals(value) then
          Exit(true);
      end
      else
      if cell.LeftEquals(value) then
        Exit(true);
    end;
  end;

  function FormatBorder(const Name: string; const B: TRtfCellBorder): string;
  begin
    result := Format('%s:%dpx solid %s;',
      [Name, Max(TwipToPixel(B.Width), 1), ColorToHtml(B.Color)]);
  end;

  function SetAlignment(Alignment: TRtfTextAlignment): boolean;
  begin
    result := false;
    case Alignment of
      rtaLeft: exit;
      rtaCenter:
        fWriter.AddAttribute(twaAlign, 'center');
      rtaRight:
        fWriter.AddAttribute(twaAlign, 'right');
      rtaJustify:
        fWriter.AddAttribute(twaAlign, 'justify');
    end;
    result := true;
  end;

var
  i, j, k, m, n, mw, left, right, colspan: integer;
  row, span_row, row2: TRtfTableRow;
  pts: TIntList;
  s: string;
  cell_def, prev_cell_def, cell_def_2: TRtfTableCellDef;
  cell: TRtfTableCell;
  btop, bbottom, bleft, bright, found, celldiv: boolean;
  vt: TRtfVisualText;
  vi: TRtfVisualImage;
begin
  celldiv := false;
  // remove empty rows
  for i := aVisualTable.Rows.Count - 1 downto 0 do
  begin
    row := aVisualTable.Rows[i];
    if row.Cells.Count = 0 then
      aVisualTable.Rows.Delete(i);
  end;

  pts := TIntList.Create;
  try
    for i := 0 to aVisualTable.Rows.Count - 1 do
    begin
      row := aVisualTable.Rows[i];
      if row.Left.HasValue and (pts.IndexOf(row.Left.Value) < 0) then
        pts.Add(row.Left.Value);
      for j := 0 to row.CellDefs.Count - 1 do
        if pts.IndexOf(row.CellDefs[j].Right) < 0 then
          pts.Add(row.CellDefs[j].Right);
    end;
    if pts.Count = 0 then
      raise ERtfHtmlConverter.Create(sNoCellDefs);

    pts.Sort;

    mw := TwipToPixel(pts.Last - pts.First);
    s := Format('margin-left:%d; border-collapse: collapse;', [TwipToPixel(pts.First)]);
    fWriter.AddAttribute(twaBorder, '0');
    fWriter.AddAttribute(twaWidth, mw.ToString);
    fWriter.AddAttribute(twaStyle, s);
    RenderBeginTag(twtTable);
{$IFDEF OLDTABLESTYLE}
    fWriter.AddAttribute(twaHeight, '0');
    RenderBeginTag(twtTr);
    for i := 1 to pts.Count - 1 do
    begin
      mw := TwipToPixel(pts[i] - pts[i - 1]);
      fWriter.AddAttribute(twaWidth, mw.ToString);
      RenderBeginTag(twtTd);
      RenderEndTag; // twtTd
    end;
    RenderEndTag; // twtTr
{$ELSE}
    RenderBeginTag(twtColgroup);
    for i := 1 to pts.Count - 1 do
    begin
      mw := TwipToPixel(pts[i] - pts[i - 1]);
      s := Format('width:%dpx;', [mw]);
      fWriter.AddAttribute(twaStyle, s);
      RenderBeginTag(twtCol);
      RenderAssertTag(twtCol);
      RenderEndTag; // twtCol
    end;
    RenderAssertTag(twtColgroup);
    RenderEndTag; // twtColgroup
{$ENDIF}
    // first, we'll determine all the rowspans and leftsides
    for i := 0 to aVisualTable.Rows.Count - 1 do
    begin
      row := aVisualTable.Rows[i];
      if row.CellDefs.Count <> row.Cells.Count then
        raise ERtfHtmlConverter.Create(sUnequalCells);

      for j := 0 to row.CellDefs.Count - 1 do
      begin
        cell := row.Cells[j];
        cell_def := row.CellDefs[j];
        cell_def.Left := IfThen((j = 0) and row.Left.HasValue, row.Left.Value,
          prev_cell_def.Right);
        if cell_def.FirstMerged then
        begin
          n := i + 1;
          if n <= aVisualTable.Rows.Count - 1 then
            for k := n to aVisualTable.Rows.Count - 1 do
            begin
              span_row := aVisualTable.Rows[k];
              found := FindIf(span_row.CellDefs, cell_def.Right, true, cell_def_2);
              if not found then
                break;
              if not cell_def_2.Merged then
                break;
            end;
          cell.Rowspan := k - i;
        end;
        prev_cell_def := cell_def;
      end;
    end;

    for i := 0 to aVisualTable.Rows.Count - 1 do
    begin
      row := aVisualTable.Rows[i];
      if row.Height.HasValue then
        fWriter.AddAttribute(twaHeight, IntToStr(TwipToPixel(row.Height.Value)));
      RenderBeginTag(twtTr);
      if row.Left.HasValue then
        n := pts.IndexOf(row.Left.Value);
      if n < 0 then
        raise ERtfHtmlConverter.Create(sNoRowLeftPoint);
      if n > 0 then
      begin
        fWriter.AddAttribute(twaColspan, n.ToString);
        RenderBeginTag(twtTd);
        RenderAssertTag(twtTd);
        RenderEndTag; // twtTd
      end;
      for j := 0 to row.CellDefs.Count - 1 do
      begin
        cell := row.Cells[j];
        cell_def := row.CellDefs[j];
        m := pts.IndexOf(cell_def.Right);
        if m < 0 then
          raise ERtfHtmlConverter.Create(sNoCelldefRightPoint);
        colspan := Abs(n - m);
        n := m;
        if not cell_def.Merged then
        begin
          // analyzing borders
          left := cell_def.Left;
          right := cell_def.Right;
          bbottom := cell_def.BorderBottom.Visible;
          btop := cell_def.BorderTop.Visible;
          bleft := cell_def.BorderLeft.Visible;
          bright := cell_def.BorderRight.Visible;
          m := i;
          if cell_def.FirstMerged then
            m := i + cell.Rowspan - 1;
          for k := i to m - 1 do
          begin
            row2 := aVisualTable.Rows[k];
            found := FindIf(span_row.CellDefs, left, true, cell_def_2);
            if found then
              bleft := bleft and cell_def_2.BorderRight.Visible;
            found := FindIf(span_row.CellDefs, right, false, cell_def_2);
            if found then
              bright := bright and cell_def_2.BorderLeft.Visible;
          end;

          if cell_def.BordersAreEquals then
            s := FormatBorder('border', cell_def.BorderLeft)
          else
          begin
            s := '';
            if bbottom then
              s := s + FormatBorder('border-bottom', cell_def.BorderBottom);
            if btop then
              s := s + FormatBorder('border-top', cell_def.BorderTop);
            if bleft then
              s := s + FormatBorder('border-left', cell_def.BorderLeft);
            if bright then
              s := s + FormatBorder('border-right', cell_def.BorderRight);
          end;
          if not s.IsEmpty then
            fWriter.AddAttribute(twaStyle, s);
          if colspan > 1 then
            fWriter.AddAttribute(twaColspan, colspan.ToString);
          if cell_def.FirstMerged then
            fWriter.AddAttribute(twaRowspan, cell.Rowspan.ToString);
          case cell_def.VAlign of
            tvaTop: fWriter.AddAttribute(twaValign, 'top');
            taBottom: fWriter.AddAttribute(twaValign, 'bottom');
          end;
          if cell_def.BackgroundColor <> clNone then
            fWriter.AddAttribute(twaBgcolor, ColorToHtml(cell_def.BackgroundColor));
          RenderBeginTag(twtTd);
          if cell.DataList.Count = 0 then
            fWriter.Write(NonBreakingSpace)
          else
          if cell.DataList.Count > 0 then
          begin
            if cell.DataList[0] is TRtfVisualText then
            begin
              vt := cell.DataList[0] as TRtfVisualText;
              celldiv := SetAlignment(vt.Format.Alignment);
            end;
            if cell.DataList[0] is TRtfVisualImage then
            begin
              vi := cell.DataList[0] as TRtfVisualImage;
              celldiv := SetAlignment(vi.Alignment);
            end;
            if celldiv then
              RenderBeginTag(twtDiv);

            for k := 0 to cell.DataList.Count - 1 do
              case TRtfVisual(cell.DataList[k]).Kind of
                rvkText: DoVisitText(cell.DataList[k] as TRtfVisualText);
                rvkBreak: DoVisitBreak(cell.DataList[k] as TRtfVisualBreak);
                rvkSpecial: DoVisitSpecial(cell.DataList[k] as TRtfVisualSpecialChar);
                rvkImage: DoVisitImage(cell.DataList[k] as TRtfVisualImage);
              end;
            if celldiv then
            begin
              RenderAssertTag(twtDiv);
              RenderEndTag; // twtDiv
            end;
          end;
          RenderAssertTag(twtTd);
          RenderEndTag; // twtTd
        end;
      end;
      RenderAssertTag(twtTr);
      RenderEndTag; // twtTr
    end;
    RenderAssertTag(twtTable);
    RenderEndTag; // twtTable
  finally
    pts.Free;
  end;

  LeaveVisual(aVisualTable);
end;

end.
