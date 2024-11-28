unit uRtfHtmlWriter;

interface

uses
  System.Types, System.SysUtils, System.Classes, System.Generics.Collections,
  uRtfCssTextWriter;

type

  THtmlTextWriterTag = (
    twtUnknown, twtA, twtAcronym, twtAddress, twtArea, twtB, twtBase,
    twtBasefont, twtBdo, twtBgsound, twtBig, twtBlockquote, twtBody, twtBr,
    twtButton, twtCaption, twtCenter, twtCite, twtCode, twtCol, twtColgroup,
    twtDd, twtDel, twtDfn, twtDir, twtDiv, twtDl, twtDt, twtEm, twtEmbed,
    twtFieldset, twtFont, twtForm, twtFrame, twtFrameset, twtH1, twtH2, twtH3,
    twtH4, twtH5, twtH6, twtHead, twtHr, twtHtml, twtI, twtIframe, twtImg,
    twtInput, twtIns, twtIsindex, twtKbd, twtLabel, twtLegend, twtLi, twtLink,
    twtMap, twtMarquee, twtMenu, twtMeta, twtNobr, twtNoframes, twtNoscript,
    twtObject, twtOl, twtOption, twtP, twtParam, twtPre, twtQ, twtRt, twtRuby,
    twtS, twtSamp, twtScript, twtSelect, twtSmall, twtSpan, twtStrike,
    twtStrong, twtStyle, twtSub, twtSup, twtTable, twtTbody, twtTd, twtTextarea,
    twtTfoot, twtTh, twtThead, twtTitle, twtTr, twtTt, twtU, twtUl, twtVar,
    twtWbr, twtXml
    );

  THtmlTextWriterAttribute = (
    twaAccesskey, twaAlign, twaAlt, twaBackground, twaBgcolor, twaBorder,
    twaBordercolor, twaCellpadding, twaCellspacing, twaChecked, twaClass,
    twaCols, twaColspan, twaDisabled, twaFor, twaHeight, twaHref, twaId,
    twaMaxlength, twaMultiple, twaName, twaNowrap, twaOnchange, twaOnclick,
    twaReadOnly, twaRows, twaRowspan, twaRules, twaSelected, twaSize, twaSrc,
    twaStyle, twaTabindex, twaTarget, twaTitle, twaType, twaValign, twaValue,
    twaWidth, twaWrap, twaAbbr, twaAutoComplete, twaAxis, twaContent, twaCoords,
    twaDesignerRegion, twaDir, twaHeaders, twaLongdesc, twaRel, twaScope,
    twaShape, twaUsemap, twaVCardName
    );

  TTagType = (ttInline, ttNonClosing, ttOther);

  TTagInformation = record
    Name: string;
    TagType: TTagType;
    ClosingTag: string;
    constructor Create(const aName: string; aTagType: TTagType; aClosingTag: string);
  end;

  TTagStackEntry = record
    TagKey: THtmlTextWriterTag;
    EndTagText: string;
  end;

  TAttributeInformation = record
    Name: string;
    Encode: Boolean;
    IsUrl: Boolean;
    constructor Create(const aName: string; aEncode: Boolean; aIsUrl: Boolean);
  end;

  TRenderAttribute = record
    Name: string;
    Value: string;
    Key: THtmlTextWriterAttribute;
    Encode: Boolean;
    IsUrl: Boolean;
    constructor Create(const aName, aValue: string;
        aKey: THtmlTextWriterAttribute; aEncode, aIsUrl: Boolean);
  end;

  THTMLTextWriter = class(TObject)
  private
    fWriter: TStringWriter;
    fIsDescendant: Boolean;
    fTabsPending: Boolean;
    fIndentLevel: Integer;
    fInlineCount: Integer;
    fTagKeyLookupTable: TDictionary<string, THtmlTextWriterTag>;
    fTagNameLookupArray: array[THtmlTextWriterTag] of TTagInformation;
    fAttrKeyLookupTable: TDictionary<string, THtmlTextWriterAttribute>;
    fAttrNameLookupArray: array[THtmlTextWriterAttribute] of TAttributeInformation;
    fAttrList: TList<TRenderAttribute>;
    fTagName: string;
    fTagKey: THtmlTextWriterTag;
    fStyleList: TList<TRenderStyle>;
    fEndTags: TStack<TTagStackEntry>;
    fCssTextWriter: TCssTextWriter;
    procedure RegisterTag(const Name: string; Key: THtmlTextWriterTag;
      tagType: TTagType);
    procedure RegisterAttribute(const Name: string; Key: THtmlTextWriterAttribute;
      Encode: Boolean; isUrl: Boolean = false);
    procedure SetTagName(const aTagName: string);
    procedure SetTagKey(const aKey: THtmlTextWriterTag);
    procedure AddAttribute(const aName, aValue: string; aKey:
      THtmlTextWriterAttribute; aEncode, aIsUrl: Boolean); overload;
    procedure AddStyleAttribute(const aName, aValue: string;
      aKey: THtmlTextWriterStyle); overload;
    procedure WriteHtmlAttributeEncode(const s: string);
  protected
    function CheckTagNameLookupArray(key: THtmlTextWriterTag): Boolean;
    function CheckAttrNameLookupArray(key: THtmlTextWriterAttribute): Boolean;
    function GetTagKey(const aTagName: string): THtmlTextWriterTag;
    function GetAttributeKey(const aName: string): THtmlTextWriterAttribute;
    function EncodeAttributeValue(AttrKey: THtmlTextWriterAttribute;
      const Value: string): string; overload;
    function EncodeAttributeValue(const Value: string;
      aEncode: Boolean): string; overload;
    function RenderBeforeContent: string; virtual;
    function RenderAfterContent: string; virtual;
    function RenderAfterTag: string; virtual;
    function PopEndTag: string;
    procedure PushEndTag(const EndTag: string);
    procedure OutputTabs;
    property TagName: string read fTagName write SetTagName;
    property TagKey: THtmlTextWriterTag read fTagKey write SetTagKey;
  public
    constructor Create;
    destructor Destroy; override;
    function ToString: string;
    procedure AddAttribute(const aName, aValue: string); overload;
    procedure AddAttribute(aKey: THtmlTextWriterAttribute;
      const aValue: string); overload;
    procedure AddAttribute(aKey: THtmlTextWriterAttribute;
      const aValue: string; aEncode: boolean); overload;
    procedure AddStyleAttribute(aKey: THtmlTextWriterStyle;
      const aValue: string); overload;
    procedure Write(const s: string);
    procedure WriteBreak;
    procedure WriteLine; overload;
    procedure WriteLine(const s: string); overload;
    procedure RenderBeginTag(const aTagName: string); overload;
    procedure RenderBeginTag(aTagKey: THtmlTextWriterTag); overload;
    procedure RenderEndTag;
  end;


implementation

uses
  System.Math, System.StrUtils, uRtfMessages, uRtfHtmlFunctions;

const

  TagLeftChar = '<';
  TagRightChar = '>';
  SelfClosingChars = ' /';
  SelfClosingTagEnd = ' />';
  EndTagLeftChars = '</';
  DoubleQuoteChar = '"';
  SingleQuoteChar = '''';
  SpaceChar = ' ';
  EqualsChar = '=';
  SlashChar = '/';
  EqualsDoubleQuoteString = '="';
  SemicolonChar = ';';
  StyleEqualsChar = ':';
  DefaultTabString = #9;
  DesignerRegionAttributeName = '_designerRegion';


{ TTagInformation }

constructor TTagInformation.Create(const aName: string; aTagType: TTagType;
  aClosingTag: string);
begin
  Name := aName;
  TagType := aTagType;
  ClosingTag := aClosingTag;
end;


{ TAttributeInformation }

constructor TAttributeInformation.Create(const aName: string; aEncode,
  aIsUrl: Boolean);
begin
  Name := aName;
  Encode := aEncode;
  IsUrl := aIsUrl;
end;


{ TRenderAttribute }

constructor TRenderAttribute.Create(const aName, aValue: string;
  aKey: THtmlTextWriterAttribute; aEncode, aIsUrl: Boolean);
begin
  Name := aName;
  Value := aValue;
  Key := aKey;
  Encode := aEncode;
  IsUrl := aIsUrl;
end;

{ THTMLTextWriter }

constructor THTMLTextWriter.Create;
begin
  fIsDescendant := false;
  fTabsPending := false;
  fIndentLevel := 0;
  fInlineCount := 0;
  fWriter := TStringWriter.Create;
  fCssTextWriter := TCssTextWriter.Create(fWriter);
  fEndTags := TStack<TTagStackEntry>.Create;
  fAttrList := TList<TRenderAttribute>.Create;
  fStyleList := TList<TRenderStyle>.Create;
  fTagKeyLookupTable := TDictionary<string, THtmlTextWriterTag>.Create;
  fAttrKeyLookupTable := TDictionary<string, THtmlTextWriterAttribute>.Create;

  RegisterTag('',           twtUnknown,        ttOther);
  RegisterTag('a',          twtA,              ttInline);
  RegisterTag('acronym',    twtAcronym,        ttInline);
  RegisterTag('address',    twtAddress,        ttOther);
  RegisterTag('area',       twtArea,           ttNonClosing);
  RegisterTag('b',          twtB,              ttInline);
  RegisterTag('base',       twtBase,           ttNonClosing);
  RegisterTag('basefont',   twtBasefont,       ttNonClosing);
  RegisterTag('bdo',        twtBdo,            ttInline);
  RegisterTag('bgsound',    twtBgsound,        ttNonClosing);
  RegisterTag('big',        twtBig,            ttInline);
  RegisterTag('blockquote', twtBlockquote,     ttOther);
  RegisterTag('body',       twtBody,           ttOther);
  RegisterTag('br',         twtBr,             ttOther);
  RegisterTag('button',     twtButton,         ttInline);
  RegisterTag('caption',    twtCaption,        ttOther);
  RegisterTag('center',     twtCenter,         ttOther);
  RegisterTag('cite',       twtCite,           ttInline);
  RegisterTag('code',       twtCode,           ttInline);
  RegisterTag('col',        twtCol,            ttNonClosing);
  RegisterTag('colgroup',   twtColgroup,       ttOther);
  RegisterTag('del',        twtDel,            ttInline);
  RegisterTag('dd',         twtDd,             ttInline);
  RegisterTag('dfn',        twtDfn,            ttInline);
  RegisterTag('dir',        twtDir,            ttOther);
  RegisterTag('div',        twtDiv,            ttOther);
  RegisterTag('dl',         twtDl,             ttOther);
  RegisterTag('dt',         twtDt,             ttInline);
  RegisterTag('em',         twtEm,             ttInline);
  RegisterTag('embed',      twtEmbed,          ttNonClosing);
  RegisterTag('fieldset',   twtFieldset,       ttOther);
  RegisterTag('font',       twtFont,           ttInline);
  RegisterTag('form',       twtForm,           ttOther);
  RegisterTag('frame',      twtFrame,          ttNonClosing);
  RegisterTag('frameset',   twtFrameset,       ttOther);
  RegisterTag('h1',         twtH1,             ttOther);
  RegisterTag('h2',         twtH2,             ttOther);
  RegisterTag('h3',         twtH3,             ttOther);
  RegisterTag('h4',         twtH4,             ttOther);
  RegisterTag('h5',         twtH5,             ttOther);
  RegisterTag('h6',         twtH6,             ttOther);
  RegisterTag('head',       twtHead,           ttOther);
  RegisterTag('hr',         twtHr,             ttNonClosing);
  RegisterTag('html',       twtHtml,           ttOther);
  RegisterTag('i',          twtI,              ttInline);
  RegisterTag('iframe',     twtIframe,         ttOther);
  RegisterTag('img',        twtImg,            ttNonClosing);
  RegisterTag('input',      twtInput,          ttNonClosing);
  RegisterTag('ins',        twtIns,            ttInline);
  RegisterTag('isindex',    twtIsindex,        ttNonClosing);
  RegisterTag('kbd',        twtKbd,            ttInline);
  RegisterTag('label',      twtLabel,          ttInline);
  RegisterTag('legend',     twtLegend,         ttOther);
  RegisterTag('li',         twtLi,             ttInline);
  RegisterTag('link',       twtLink,           ttNonClosing);
  RegisterTag('map',        twtMap,            ttOther);
  RegisterTag('marquee',    twtMarquee,        ttOther);
  RegisterTag('menu',       twtMenu,           ttOther);
  RegisterTag('meta',       twtMeta,           ttNonClosing);
  RegisterTag('nobr',       twtNobr,           ttInline);
  RegisterTag('noframes',   twtNoframes,       ttOther);
  RegisterTag('noscript',   twtNoscript,       ttOther);
  RegisterTag('object',     twtObject,         ttOther);
  RegisterTag('ol',         twtOl,             ttOther);
  RegisterTag('option',     twtOption,         ttOther);
  RegisterTag('p',          twtP,              ttInline);
  RegisterTag('param',      twtParam,          ttOther);
  RegisterTag('pre',        twtPre,            ttOther);
  RegisterTag('ruby',       twtRuby,           ttOther);
  RegisterTag('rt',         twtRt,             ttOther);
  RegisterTag('q',          twtQ,              ttInline);
  RegisterTag('s',          twtS,              ttInline);
  RegisterTag('samp',       twtSamp,           ttInline);
  RegisterTag('script',     twtScript,         ttOther);
  RegisterTag('select',     twtSelect,         ttOther);
  RegisterTag('small',      twtSmall,          ttOther);
  RegisterTag('span',       twtSpan,           ttInline);
  RegisterTag('strike',     twtStrike,         ttInline);
  RegisterTag('strong',     twtStrong,         ttInline);
  RegisterTag('style',      twtStyle,          ttOther);
  RegisterTag('sub',        twtSub,            ttInline);
  RegisterTag('sup',        twtSup,            ttInline);
  RegisterTag('table',      twtTable,          ttOther);
  RegisterTag('tbody',      twtTbody,          ttOther);
  RegisterTag('td',         twtTd,             ttInline);
  RegisterTag('textarea',   twtTextarea,       ttInline);
  RegisterTag('tfoot',      twtTfoot,          ttOther);
  RegisterTag('th',         twtTh,             ttInline);
  RegisterTag('thead',      twtThead,          ttOther);
  RegisterTag('title',      twtTitle,          ttOther);
  RegisterTag('tr',         twtTr,             ttOther);
  RegisterTag('tt',         twtTt,             ttInline);
  RegisterTag('u',          twtU,              ttInline);
  RegisterTag('ul',         twtUl,             ttOther);
  RegisterTag('var',        twtVar,            ttInline);
  RegisterTag('wbr',        twtWbr,            ttNonClosing);
  RegisterTag('xml',        twtXml,            ttOther);

  RegisterAttribute('abbr',            twaAbbr,           true);
  RegisterAttribute('accesskey',       twaAccesskey,      true);
  RegisterAttribute('align',           twaAlign,          false);
  RegisterAttribute('alt',             twaAlt,            true);
  RegisterAttribute('autocomplete',    twaAutoComplete,   false);
  RegisterAttribute('axis',            twaAxis,           true);
  RegisterAttribute('background',      twaBackground,     true,     true);
  RegisterAttribute('bgcolor',         twaBgcolor,        false);
  RegisterAttribute('border',          twaBorder,         false);
  RegisterAttribute('bordercolor',     twaBordercolor,    false);
  RegisterAttribute('cellpadding',     twaCellpadding,    false);
  RegisterAttribute('cellspacing',     twaCellspacing,    false);
  RegisterAttribute('checked',         twaChecked,        false);
  RegisterAttribute('class',           twaClass,          true);
  RegisterAttribute('cols',            twaCols,           false);
  RegisterAttribute('colspan',         twaColspan,        false);
  RegisterAttribute('content',         twaContent,        true);
  RegisterAttribute('coords',          twaCoords,         false);
  RegisterAttribute('dir',             twaDir,            false);
  RegisterAttribute('disabled',        twaDisabled,       false);
  RegisterAttribute('for',             twaFor,            false);
  RegisterAttribute('headers',         twaHeaders,        true);
  RegisterAttribute('height',          twaHeight,         false);
  RegisterAttribute('href',            twaHref,           true,      true);
  RegisterAttribute('id',              twaId,             false);
  RegisterAttribute('longdesc',        twaLongdesc,       true,      true);
  RegisterAttribute('maxlength',       twaMaxlength,      false);
  RegisterAttribute('multiple',        twaMultiple,       false);
  RegisterAttribute('name',            twaName,           false);
  RegisterAttribute('nowrap',          twaNowrap,         false);
  RegisterAttribute('onclick',         twaOnclick,        true);
  RegisterAttribute('onchange',        twaOnchange,       true);
  RegisterAttribute('readonly',        twaReadOnly,       false);
  RegisterAttribute('rel',             twaRel,            false);
  RegisterAttribute('rows',            twaRows,           false);
  RegisterAttribute('rowspan',         twaRowspan,        false);
  RegisterAttribute('rules',           twaRules,          false);
  RegisterAttribute('scope',           twaScope,          false);
  RegisterAttribute('selected',        twaSelected,       false);
  RegisterAttribute('shape',           twaShape,          false);
  RegisterAttribute('size',            twaSize,           false);
  RegisterAttribute('src',             twaSrc,            true,      true);
  RegisterAttribute('style',           twaStyle,          false);
  RegisterAttribute('tabindex',        twaTabindex,       false);
  RegisterAttribute('target',          twaTarget,         false);
  RegisterAttribute('title',           twaTitle,          true);
  RegisterAttribute('type',            twaType,           false);
  RegisterAttribute('usemap',          twaUsemap,         false);
  RegisterAttribute('valign',          twaValign,         false);
  RegisterAttribute('value',           twaValue,          true);
  RegisterAttribute('vcard_name',      twaVCardName,      false);
  RegisterAttribute('width',           twaWidth,          false);
  RegisterAttribute('wrap',            twaWrap,           false);
  RegisterAttribute(DesignerRegionAttributeName, twaDesignerRegion, false);
end;

destructor THTMLTextWriter.Destroy;
begin
  fWriter.Free;
  fEndTags.Free;
  fAttrList.Free;
  fStyleList.Free;
  fCssTextWriter.Free;
  fTagKeyLookupTable.Free;
  fAttrKeyLookupTable.Free;
  inherited;
end;


function THTMLTextWriter.CheckAttrNameLookupArray(
  key: THtmlTextWriterAttribute): Boolean;
begin
  result := InRange(Ord(Key), Ord(Low(fAttrNameLookupArray)),
    Ord(High(fAttrNameLookupArray)));
end;

function THTMLTextWriter.CheckTagNameLookupArray(
  key: THtmlTextWriterTag): Boolean;
begin
  result := InRange(Ord(Key), Ord(Low(fTagNameLookupArray)),
    Ord(High(fTagNameLookupArray)));
end;

procedure THTMLTextWriter.RegisterAttribute(const Name: string;
  Key: THtmlTextWriterAttribute; Encode, isUrl: Boolean);
var
  NameLCase: string;
begin
  NameLCase := LowerCase(Name);
  fAttrKeyLookupTable.Add(NameLCase, Key);
  if CheckAttrNameLookupArray(Key) then
    fAttrNameLookupArray[Key] := TAttributeInformation.Create(Name, Encode, isUrl);
end;

procedure THTMLTextWriter.RegisterTag(const Name: string;
  Key: THtmlTextWriterTag; TagType: TTagType);
var
  NameLCase: string;
  EndTag: string;
begin
  NameLCase := Name.ToLower;
  fTagKeyLookupTable.Add(NameLCase, Key);
  // Pre-resolve the end tag
  EndTag := '';
  if (TagType <> ttNonClosing) and (Key <> twtUnknown) then
    EndTag := EndTagLeftChars + NameLCase + TagRightChar;

  if CheckTagNameLookupArray(Key) then
    fTagNameLookupArray[Key] := TTagInformation.Create(Name, TagType, EndTag);
end;

function THTMLTextWriter.GetAttributeKey(
  const aName: string): THtmlTextWriterAttribute;
begin
  if not aName.IsEmpty then
    if fAttrKeyLookupTable.TryGetValue(aName.ToLower, result) then
      exit;
  result := THtmlTextWriterAttribute(-1);
end;

function THTMLTextWriter.GetTagKey(const aTagName: string): THtmlTextWriterTag;
begin
  if not aTagName.IsEmpty then
    if fTagKeyLookupTable.TryGetValue(aTagName.ToLower, result) then
      exit;
  result := twtUnknown;
end;

function THTMLTextWriter.EncodeAttributeValue(AttrKey: THtmlTextWriterAttribute;
  const Value: string): string;
var
  Encode: boolean;
begin
  Encode := true;
  if CheckAttrNameLookupArray(AttrKey) then
    Encode := fAttrNameLookupArray[AttrKey].Encode;
  result := EncodeAttributeValue(Value, Encode);
end;

function THTMLTextWriter.EncodeAttributeValue(const Value: string;
  aEncode: Boolean): string;
begin
  if value.IsEmpty then
    exit('');
  result := IfThen(aEncode, HtmlEncodeString(Value), Value);
end;

procedure THTMLTextWriter.OutputTabs;
var
  i: integer;
begin
  if fTabsPending then
  begin
    for i := 0 to fIndentLevel - 1 do
      fWriter.Write(DefaultTabString);
    fTabsPending := false;
  end;
end;

function THTMLTextWriter.PopEndTag: string;
var
  e: TTagStackEntry;
begin
  e := fEndTags.Pop;
  TagKey := e.TagKey;
  result := e.EndTagText;
end;

procedure THTMLTextWriter.PushEndTag(const EndTag: string);
var
  e: TTagStackEntry;
begin
  e.TagKey := fTagKey;
  e.EndTagText := EndTag;
  fEndTags.Push(e);
end;

procedure THTMLTextWriter.SetTagKey(const aKey: THtmlTextWriterTag);
var
  s: string;
begin
  if not CheckTagNameLookupArray(aKey) then
    raise EArgumentException.Create(sOutOfRangeHtmlTextWriterTag);
  fTagKey := aKey;
  if fTagKey <> twtUnknown then
    for s in fTagKeyLookupTable.Keys do
      if fTagKeyLookupTable[s] = fTagKey then
      begin
        fTagName := s;
        break;
      end;
end;

procedure THTMLTextWriter.SetTagName(const aTagName: string);
begin
  fTagName := aTagName;
  fTagKey := GetTagKey(fTagName);
end;

function THTMLTextWriter.ToString: string;
begin
  result := fWriter.ToString;
end;

procedure THTMLTextWriter.AddAttribute(const aName, aValue: string;
  aKey: THtmlTextWriterAttribute; aEncode, aIsUrl: Boolean);
begin
  fAttrList.Add(TRenderAttribute.Create(aName, aValue, aKey, aEncode, aIsUrl));
end;

procedure THTMLTextWriter.AddAttribute(const aName, aValue: string);
var
  AttributeKey: THtmlTextWriterAttribute;
  AttributeValue: string;
begin
  AttributeKey := GetAttributeKey(aName);
  AttributeValue := EncodeAttributeValue(AttributeKey, aValue);
  AddAttribute(aName, AttributeValue, AttributeKey, false, false);
end;

procedure THTMLTextWriter.AddAttribute(aKey: THtmlTextWriterAttribute;
  const aValue: string);
var
  info: TAttributeInformation;
begin
  if CheckAttrNameLookupArray(aKey) then
  begin
    info := fAttrNameLookupArray[aKey];
    AddAttribute(info.Name, aValue, aKey, info.Encode, info.IsUrl);
  end;
end;

procedure THTMLTextWriter.AddAttribute(aKey: THtmlTextWriterAttribute;
  const aValue: string; aEncode: boolean);
var
  info: TAttributeInformation;
begin
  if CheckAttrNameLookupArray(aKey) then
  begin
    info := fAttrNameLookupArray[aKey];
    AddAttribute(info.Name, aValue, aKey, aEncode, info.IsUrl);
  end;
end;

procedure THTMLTextWriter.AddStyleAttribute(aKey: THtmlTextWriterStyle;
  const aValue: string);
begin
  AddStyleAttribute(fCssTextWriter.GetStyleName(aKey), aValue, aKey);
end;

procedure THTMLTextWriter.AddStyleAttribute(const aName, aValue: string;
  aKey: THtmlTextWriterStyle);
var
  Style: TRenderStyle;
  AttributeValue: string;
begin
  style.Name := aName;
  Style.Key := aKey;
  AttributeValue := aValue;
  if fCssTextWriter.IsStyleEncoded(aKey) then
    AttributeValue := HtmlEncodeString(aValue);
  Style.Value := AttributeValue;
  fStyleList.Add(Style);
end;

procedure THTMLTextWriter.WriteHtmlAttributeEncode(const s: string);
begin
  fWriter.Write(HtmlEncodeAttr(s));
end;

procedure THTMLTextWriter.Write(const s: string);
begin
  if fTabsPending then
    OutputTabs;
  fWriter.Write(s);
end;

procedure THTMLTextWriter.WriteBreak;
begin
  fWriter.Write('<br />');
end;

procedure THTMLTextWriter.WriteLine;
begin
  fWriter.WriteLine;
  fTabsPending := true;
end;

procedure THTMLTextWriter.WriteLine(const s: string);
begin
  if fTabsPending then
    OutputTabs;
  fWriter.WriteLine(s);
  fTabsPending := true;
end;

procedure THTMLTextWriter.RenderBeginTag(const aTagName: string);
begin
  TagName := aTagName;
  RenderBeginTag(fTagKey);
end;

function THTMLTextWriter.RenderBeforeContent: string;
begin
  result := '';
end;

function THTMLTextWriter.RenderAfterContent: string;
begin
  result := '';
end;

function THTMLTextWriter.RenderAfterTag: string;
begin
  result := '';
end;

procedure THTMLTextWriter.RenderBeginTag(aTagKey: THtmlTextWriterTag);
var
  TagInfo: TTagInformation;
  TagType: TTagType;
  EndTag, StyleValue, AttrValue,
  TextBeforeContent, TextAfterTag,
  TextAfterContent: string;
  RenderTag, RenderEndTag: Boolean;
  i: integer;
  attr: TRenderAttribute;
begin
  TagKey := aTagKey;
  RenderTag := true;
  TagInfo := fTagNameLookupArray[fTagKey];
  TagType := TagInfo.TagType;
  RenderEndTag := RenderTag and (TagType <> ttNonClosing);
  EndTag := IfThen(RenderEndTag, TagInfo.ClosingTag);
  // write the begin tag
  if RenderTag then
  begin
    if fTabsPending then
      OutputTabs;
    fWriter.Write(TagLeftChar);
    fWriter.Write(fTagName);
    StyleValue := '';

    for i := 0 to fAttrList.Count - 1 do
    begin
      attr := fAttrList[i];
      if attr.Key = twaStyle then
        // append style attribute in with other styles
        StyleValue := attr.Value
      else
      begin
        fWriter.Write(SpaceChar);
        fWriter.Write(attr.Name);
        if not attr.Value.IsEmpty then
        begin
          fWriter.Write(EqualsDoubleQuoteString);
          AttrValue := attr.Value;
          if attr.IsUrl then
            if (attr.Key <> twaHref) or (not AttrValue.StartsWith('javascript:', true)) then
              AttrValue := EncodeUrl(attrValue);
          if attr.Encode then
            WriteHtmlAttributeEncode(AttrValue)
          else
            fWriter.Write(AttrValue);
          fWriter.Write(DoubleQuoteChar);
        end;
      end;
    end;

    if (fStyleList.Count > 0) or (not StyleValue.IsEmpty) then
    begin
      fWriter.Write(SpaceChar);
      fWriter.Write('style');
      fWriter.Write(EqualsDoubleQuoteString);

      fCssTextWriter.WriteAttributes(fStyleList.ToArray, fStyleList.Count);
      if not StyleValue.IsEmpty then
        fWriter.Write(StyleValue);

      fWriter.Write(DoubleQuoteChar);
    end;

    if TagType = ttNonClosing then
      fWriter.Write(SelfClosingTagEnd)
    else
      fWriter.Write(TagRightChar);
  end;

  TextBeforeContent := RenderBeforeContent;
  if not TextBeforeContent.IsEmpty then
  begin
    if fTabsPending then
      OutputTabs;
    fWriter.Write(textBeforeContent);
  end;

  // write text before the content
  if RenderEndTag then
  begin
    if TagType = ttInline then
      inc(fInlineCount)
    else
    begin
      // writeline and indent before rendering content
      WriteLine;
      Inc(fIndentLevel);
    end;
    // Manually build end tags for unknown tag types.
    if EndTag.IsEmpty then
      EndTag := EndTagLeftChars + fTagName + TagRightChar;
  end;

  if fIsDescendant then
  begin
    // append text after the tag
    TextAfterTag := RenderAfterTag;
    if not TextAfterTag.IsEmpty then
      EndTag := IfThen(EndTag.IsEmpty, TextAfterTag, TextAfterTag + EndTag);
    // build end content and push it on stack to write in RenderEndTag
    // prepend text after the content
    TextAfterContent := RenderAfterContent;
    if not TextAfterContent.IsEmpty then
      EndTag := IfThen(EndTag.IsEmpty, TextAfterContent, TextAfterContent + EndTag);
  end;

  // push end tag onto stack
  PushEndTag(EndTag);

  // flush attribute and style lists for next tag
  fAttrList.Clear;
  fStyleList.Clear;
end;

procedure THTMLTextWriter.RenderEndTag;
var
  EndTag: string;
  EndChar: char;
begin
  EndTag := PopEndTag;

  if not EndTag.IsEmpty then
    if fTagNameLookupArray[fTagKey].tagType = ttInline then
    begin
      dec(fInlineCount);
      // Never inject crlfs at end of inline tags.
      Write(endTag);
    end
    else
    begin
      // unindent if not an inline tag
      // fixed: exclude empty line
      EndChar := fWriter.ToString[fWriter.ToString.Length];
      if EndChar <> #10 then
        WriteLine;
      dec(fIndentLevel);
      fIndentLevel := Max(fIndentLevel, 0);
      Write(EndTag);
    end;
end;

end.
