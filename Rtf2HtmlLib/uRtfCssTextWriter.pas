unit uRtfCssTextWriter;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections;

type

  THtmlTextWriterStyle = (
    twsBackgroundColor, twsBackgroundImage, twsBorderCollapse, twsBorderColor,
    twsBorderStyle, twsBorderWidth, twsColor, twsFontFamily, twsFontSize,
    twsFontStyle, twsFontWeight, twsHeight, twsTextDecoration, twsWidth,
    twsListStyleImage, twsListStyleType, twsCursor, twsDirection, twsDisplay,
    twsFilter, twsFontVariant, twsLeft, twsLineHeight, twsMargin, twsMarginBottom,
    twsMarginLeft, twsMarginRight, twsMarginTop, twsOverflow, twsOverflowX,
    twsOverflowY, twsPadding, twsPaddingBottom, twsPaddingLeft, twsPaddingRight,
    twsPaddingTop, twsPosition, twsTextAlign, twsVerticalAlign, twsTextOverflow,
    twsTextIndent, twsTop, twsVisibility, twsWhiteSpace, twsZIndex
    );

  TAttributeInformation = record
    Name: string;
    IsUrl: Boolean;
    Encode: Boolean;
    constructor Create(const aName: string; aEncode, aIsUrl: Boolean);
  end;

  TRenderStyle = record
    Name: string;
    Value: string;
    Key: THtmlTextWriterStyle;
    constructor Create(const aName, aValue: string; aKey: THtmlTextWriterStyle);
  end;

  TCssTextWriter = class(TObject)
  private
    fWriter: TTextWriter;
    fAttrKeyLookupTable: TDictionary<string, THtmlTextWriterStyle>;
    fAttrNameLookupArray: array[THtmlTextWriterStyle] of TAttributeInformation;
    procedure WriteAttribute(Key: THtmlTextWriterStyle; const Name, Value: string); overload;
  protected
    function CheckAttrNameLookupArray(Key: THtmlTextWriterStyle): Boolean; inline;
  public
    constructor Create(aWriter: TTextWriter);
    destructor Destroy; override;
    function GetStyleKey(const StyleName: string): THtmlTextWriterStyle;
    function GetStyleName(StyleKey: THtmlTextWriterStyle): string;
    function IsStyleEncoded(StyleKey: THtmlTextWriterStyle): Boolean;
    procedure RegisterAttribute(const Name: string; Key: THtmlTextWriterStyle;
      Encode: Boolean = false; IsUrl: Boolean = false); overload;
    procedure WriteAttribute(const Name, Value: string); overload;
    procedure WriteAttribute(Key: THtmlTextWriterStyle; const Value: string); overload;
    procedure WriteAttributes(const Styles: TArray<TRenderStyle>; Count: Integer);
    procedure WriteUrlAttribute(const Url: string);
    procedure WriteBeginCssRule(const Selector: string);
    procedure WriteEndCssRule;
  end;

implementation

uses
  System.Math, System.StrUtils, uRtfHtmlFunctions;

{ TAttributeInformation }

constructor TAttributeInformation.Create(const aName: string; aEncode,
  aIsUrl: Boolean);
begin
  Name := aName;
  Encode := aEncode;
  IsUrl := aIsUrl;
end;


{ TRenderStyle }

constructor TRenderStyle.Create(const aName, aValue: string;
  aKey: THtmlTextWriterStyle);
begin
  Name := aName;
  Value := aValue;
  Key := aKey;
end;

{ TCssTextWriter }

constructor TCssTextWriter.Create(aWriter: TTextWriter);
begin
  fWriter := aWriter;
  fAttrKeyLookupTable := TDictionary<string, THtmlTextWriterStyle>.Create;
  RegisterAttribute('background-color', twsBackgroundColor);
  RegisterAttribute('background-image', twsBackgroundImage, true, true);
  RegisterAttribute('border-collapse', twsBorderCollapse);
  RegisterAttribute('border-color', twsBorderColor);
  RegisterAttribute('border-style', twsBorderStyle);
  RegisterAttribute('border-width', twsBorderWidth);
  RegisterAttribute('color', twsColor);
  RegisterAttribute('cursor', twsCursor);
  RegisterAttribute('direction', twsDirection);
  RegisterAttribute('display', twsDisplay);
  RegisterAttribute('filter', twsFilter);
  RegisterAttribute('font-family', twsFontFamily, true);
  RegisterAttribute('font-size', twsFontSize);
  RegisterAttribute('font-style', twsFontStyle);
  RegisterAttribute('font-variant', twsFontVariant);
  RegisterAttribute('font-weight', twsFontWeight);
  RegisterAttribute('height', twsHeight);
  RegisterAttribute('left', twsLeft);
  RegisterAttribute('line-height', twsLineHeight);
  RegisterAttribute('list-style-image', twsListStyleImage, true, true);
  RegisterAttribute('list-style-type', twsListStyleType);
  RegisterAttribute('margin', twsMargin);
  RegisterAttribute('margin-bottom', twsMarginBottom);
  RegisterAttribute('margin-left', twsMarginLeft);
  RegisterAttribute('margin-right', twsMarginRight);
  RegisterAttribute('margin-top', twsMarginTop);
  RegisterAttribute('overflow-x', twsOverflowX);
  RegisterAttribute('overflow-y', twsOverflowY);
  RegisterAttribute('overflow', twsOverflow);
  RegisterAttribute('padding', twsPadding);
  RegisterAttribute('padding-bottom', twsPaddingBottom);
  RegisterAttribute('padding-left', twsPaddingLeft);
  RegisterAttribute('padding-right', twsPaddingRight);
  RegisterAttribute('padding-top', twsPaddingTop);
  RegisterAttribute('position', twsPosition);
  RegisterAttribute('text-align', twsTextAlign);
  RegisterAttribute('text-decoration', twsTextDecoration);
  RegisterAttribute('text-indent', twsTextIndent);
  RegisterAttribute('text-overflow', twsTextOverflow);
  RegisterAttribute('top', twsTop);
  RegisterAttribute('vertical-align', twsVerticalAlign);
  RegisterAttribute('visibility', twsVisibility);
  RegisterAttribute('width', twsWidth);
  RegisterAttribute('white-space', twsWhiteSpace);
  RegisterAttribute('z-index', twsZIndex);
end;

destructor TCssTextWriter.Destroy;
begin
  fWriter := nil;
  fAttrKeyLookupTable.Free;
  inherited;
end;

function TCssTextWriter.CheckAttrNameLookupArray(Key: THtmlTextWriterStyle): Boolean;
begin
  Result := InRange(Ord(Key), Ord(Low(fAttrNameLookupArray)), Ord(High(fAttrNameLookupArray)));
end;

function TCssTextWriter.GetStyleKey(const StyleName: string): THtmlTextWriterStyle;
var
  key: THtmlTextWriterStyle;
begin
  if not StyleName.IsEmpty then
    if fAttrKeyLookupTable.TryGetValue(LowerCase(StyleName), key) then
      Exit(key);
  Result := THtmlTextWriterStyle(-1);
end;

function TCssTextWriter.GetStyleName(StyleKey: THtmlTextWriterStyle): string;
begin
  Result := IfThen(CheckAttrNameLookupArray(StyleKey),
    fAttrNameLookupArray[StyleKey].Name);
end;

function TCssTextWriter.IsStyleEncoded(StyleKey: THtmlTextWriterStyle): Boolean;
begin
  if CheckAttrNameLookupArray(StyleKey) then
    Result := fAttrNameLookupArray[StyleKey].Encode
  else
    Result := true;
end;

procedure TCssTextWriter.RegisterAttribute(const Name: string;
  Key: THtmlTextWriterStyle; Encode, IsUrl: Boolean);
begin
  fAttrKeyLookupTable.Add(Name.ToLower, key);
  if CheckAttrNameLookupArray(Key) then
    fAttrNameLookupArray[Key] := TAttributeInformation.Create(Name, Encode, IsUrl);
end;

procedure TCssTextWriter.WriteAttribute(Key: THtmlTextWriterStyle; const Name,
  Value: string);
var
  IsUrl: Boolean;
begin
  fWriter.Write(Name);
  fWriter.Write(':');
  IsUrl := false;
  if CheckAttrNameLookupArray(Key) then
    IsUrl := fAttrNameLookupArray[Key].IsUrl;
  if not IsUrl then
    fWriter.Write(Value)
  else
    WriteUrlAttribute(Value);
  fWriter.Write(';');
end;

procedure TCssTextWriter.WriteAttribute(const Name, Value: string);
begin
  WriteAttribute(GetStyleKey(Name), Name, Value);
end;

procedure TCssTextWriter.WriteAttribute(Key: THtmlTextWriterStyle; const Value: string);
begin
  WriteAttribute(Key, GetStyleName(key), value);
end;

procedure TCssTextWriter.WriteAttributes(const Styles: TArray<TRenderStyle>; Count: Integer);
var
  i: Integer;
begin
  for i := 0 to Count - 1 do
    WriteAttribute(Styles[i].Key, Styles[i].Name, styles[i].Value);
end;

procedure TCssTextWriter.WriteBeginCssRule(const Selector: string);
begin
  fWriter.Write(Selector);
  fWriter.Write(' { ');
end;

procedure TCssTextWriter.WriteEndCssRule;
begin
  fWriter.WriteLine(' }');
end;

procedure TCssTextWriter.WriteUrlAttribute(const Url: string);
const
  Quotes: array[0..1] of Char = (#39, '"');
var
  UrlValue: string;
  Quote: Char;
  SurroundingQuote: Char;
  UrlIndex, UrlLength: Integer;
begin
  UrlValue := Url;
  surroundingQuote := #0;
  if StartsStr('url(', url) then
  begin
    UrlIndex := 4;
    UrlLength := Length(Url) - 4;
    if EndsStr(')', Url) then
      Dec(UrlLength);
    UrlValue := Trim(Copy(Url, UrlIndex + 1, UrlLength));
  end;

  for Quote in Quotes do
    if (StartsStr(Quote, UrlValue)) and (EndsStr(Quote, UrlValue)) then
    begin
      UrlValue := UrlValue.Trim([Quote]);
      SurroundingQuote := Quote;
      break;
    end;

  fWriter.Write('url(');
  if SurroundingQuote <> #0 then
    fWriter.Write(SurroundingQuote);
  fWriter.Write(HtmlEncodeString(UrlValue));
  if surroundingQuote <> #0 then
    fWriter.Write(surroundingQuote);
  fWriter.Write(')');
end;


end.
