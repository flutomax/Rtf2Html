unit uRtfHtmlFunctions;

interface

uses
  Winapi.Windows, System.SysUtils, Vcl.Graphics, uRtfObjects, uRtfVisual,
  uRtfHtmlObjects;

  function EncodeUrl(const url: string): string;
  function HtmlEncodeString(const s: string): string;
  function HtmlEncodeAttr(const s: string; NonBreakingSpaces: boolean = false): string;
  function ColorToHtml(Color: TColor): string;
  function TextToHtml(VisualText: TRtfVisualText): TRtfHtmlStyle;
  function TextToBinary(HexStr: string): TBytes;
  function TwipToPixel(const x: integer): integer;

implementation

uses
  System.NetEncoding, uRtfMessages;

type

  TColorPair = record
    Name: string;
    Value: Integer;
  end;

const

  HTMLColorTable: array[0..140] of TColorPair = (
    (Name: 'aliceblue'; Value: $F0F8FF),
    (Name: 'antiquewhite'; Value: $FAEBD7),
    (Name: 'aqua'; Value: $00FFFF),
    (Name: 'aquamarine'; Value: $7FFFD4),
    (Name: 'azure'; Value: $F0FFFF),
    (Name: 'beige'; Value: $F5F5DC),
    (Name: 'bisque'; Value: $FFE4C4),
    (Name: 'black'; Value: $000000),
    (Name: 'blanchedalmond'; Value: $FFFFCD),
    (Name: 'blue'; Value: $0000FF),
    (Name: 'blueviolet'; Value: $8A2BE2),
    (Name: 'brown'; Value: $A52A2A),
    (Name: 'burlywood'; Value: $DEB887),
    (Name: 'cadetblue'; Value: $5F9EA0),
    (Name: 'chartreuse'; Value: $7FFF00),
    (Name: 'chocolate'; Value: $D2691E),
    (Name: 'coral'; Value: $FF7F50),
    (Name: 'cornflowerblue'; Value: $6495ED),
    (Name: 'cornsilk'; Value: $FFF8DC),
    (Name: 'crimson'; Value: $DC143C),
    (Name: 'cyan'; Value: $00FFFF),
    (Name: 'darkblue'; Value: $00008B),
    (Name: 'darkcyan'; Value: $008B8B),
    (Name: 'darkgoldenrod'; Value: $B8860B),
    (Name: 'darkgray'; Value: $A9A9A9),
    (Name: 'darkgreen'; Value: $006400),
    (Name: 'darkkhaki'; Value: $BDB76B),
    (Name: 'darkmagenta'; Value: $8B008B),
    (Name: 'darkolivegreen'; Value: $556B2F),
    (Name: 'darkorange'; Value: $FF8C00),
    (Name: 'darkorchid'; Value: $9932CC),
    (Name: 'darkred'; Value: $8B0000),
    (Name: 'darksalmon'; Value: $E9967A),
    (Name: 'darkseagreen'; Value: $8FBC8F),
    (Name: 'darkslateblue'; Value: $483D8B),
    (Name: 'darkslategray'; Value: $2F4F4F),
    (Name: 'darkturquoise'; Value: $00CED1),
    (Name: 'darkviolet'; Value: $9400D3),
    (Name: 'deeppink'; Value: $FF1493),
    (Name: 'deepskyblue'; Value: $00BFFF),
    (Name: 'dimgray'; Value: $696969),
    (Name: 'dodgerblue'; Value: $1E90FF),
    (Name: 'firebrick'; Value: $B22222),
    (Name: 'floralwhite'; Value: $FFFAF0),
    (Name: 'forestgreen'; Value: $228B22),
    (Name: 'fuchsia'; Value: $FF00FF),
    (Name: 'gainsboro'; Value: $DCDCDC),
    (Name: 'ghostwhite'; Value: $F8F8FF),
    (Name: 'gold'; Value: $FFD700),
    (Name: 'goldenrod'; Value: $DAA520),
    (Name: 'gray'; Value: $808080),
    (Name: 'green'; Value: $008000),
    (Name: 'greenyellow'; Value: $ADFF2F),
    (Name: 'honeydew'; Value: $F0FFF0),
    (Name: 'hotpink'; Value: $FF69B4),
    (Name: 'indianred'; Value: $CD5C5C),
    (Name: 'indigo'; Value: $4B0082),
    (Name: 'ivory'; Value: $FFF0F0),
    (Name: 'khaki'; Value: $F0E68C),
    (Name: 'lavender'; Value: $E6E6FA),
    (Name: 'lavenderblush'; Value: $FFF0F5),
    (Name: 'lawngreen'; Value: $7CFC00),
    (Name: 'lemonchiffon'; Value: $FFFACD),
    (Name: 'lightblue'; Value: $ADD8E6),
    (Name: 'lightcoral'; Value: $F08080),
    (Name: 'lightcyan'; Value: $E0FFFF),
    (Name: 'lightgoldenrodyellow'; Value: $FAFAD2),
    (Name: 'lightgreen'; Value: $90EE90),
    (Name: 'lightgrey'; Value: $D3D3D3),
    (Name: 'lightpink'; Value: $FFB6C1),
    (Name: 'lightsalmon'; Value: $FFA07A),
    (Name: 'lightseagreen'; Value: $20B2AA),
    (Name: 'lightskyblue'; Value: $87CEFA),
    (Name: 'lightslategray'; Value: $778899),
    (Name: 'lightsteelblue'; Value: $B0C4DE),
    (Name: 'lightyellow'; Value: $FFFFE0),
    (Name: 'lime'; Value: $00FF00),
    (Name: 'limegreen'; Value: $32CD32),
    (Name: 'linen'; Value: $FAF0E6),
    (Name: 'magenta'; Value: $FF00FF),
    (Name: 'maroon'; Value: $800000),
    (Name: 'mediumaquamarine'; Value: $66CDAA),
    (Name: 'mediumblue'; Value: $0000CD),
    (Name: 'mediumorchid'; Value: $BA55D3),
    (Name: 'mediumpurple'; Value: $9370DB),
    (Name: 'mediumseagreen'; Value: $3CB371),
    (Name: 'mediumpurple'; Value: $9370DB),
    (Name: 'mediumslateblue'; Value: $7B68EE),
    (Name: 'mediumspringgreen'; Value: $00FA9A),
    (Name: 'mediumturquoise'; Value: $48D1CC),
    (Name: 'mediumvioletred'; Value: $C71585),
    (Name: 'midnightblue'; Value: $191970),
    (Name: 'mintcream'; Value: $F5FFFA),
    (Name: 'mistyrose'; Value: $FFE4E1),
    (Name: 'moccasin'; Value: $FFE4B5),
    (Name: 'navajowhite'; Value: $FFDEAD),
    (Name: 'navy'; Value: $000080),
    (Name: 'oldlace'; Value: $FDF5E6),
    (Name: 'olive'; Value: $808000),
    (Name: 'olivedrab'; Value: $6B8E23),
    (Name: 'orange'; Value: $FFA500),
    (Name: 'orangered'; Value: $FF4500),
    (Name: 'orchid'; Value: $DA70D6),
    (Name: 'palegoldenrod'; Value: $EEE8AA),
    (Name: 'palegreen'; Value: $98FB98),
    (Name: 'paleturquoise'; Value: $AFEEEE),
    (Name: 'palevioletred'; Value: $DB7093),
    (Name: 'papayawhip'; Value: $FFEFD5),
    (Name: 'peachpuff'; Value: $FFDBBD),
    (Name: 'peru'; Value: $CD853F),
    (Name: 'pink'; Value: $FFC0CB),
    (Name: 'plum'; Value: $DDA0DD),
    (Name: 'powderblue'; Value: $B0E0E6),
    (Name: 'purple'; Value: $800080),
    (Name: 'red'; Value: $FF0000),
    (Name: 'rosybrown'; Value: $BC8F8F),
    (Name: 'royalblue'; Value: $4169E1),
    (Name: 'saddlebrown'; Value: $8B4513),
    (Name: 'salmon'; Value: $FA8072),
    (Name: 'sandybrown'; Value: $F4A460),
    (Name: 'seagreen'; Value: $2E8B57),
    (Name: 'seashell'; Value: $FFF5EE),
    (Name: 'sienna'; Value: $A0522D),
    (Name: 'silver'; Value: $C0C0C0),
    (Name: 'skyblue'; Value: $87CEEB),
    (Name: 'slateblue'; Value: $6A5ACD),
    (Name: 'slategray'; Value: $708090),
    (Name: 'snow'; Value: $FFFAFA),
    (Name: 'springgreen'; Value: $00FF7F),
    (Name: 'steelblue'; Value: $4682B4),
    (Name: 'tan'; Value: $D2B48C),
    (Name: 'teal'; Value: $008080),
    (Name: 'thistle'; Value: $D8BFD8),
    (Name: 'tomato'; Value: $FD6347),
    (Name: 'turquoise'; Value: $40E0D0),
    (Name: 'violet'; Value: $EE82EE),
    (Name: 'wheat'; Value: $F5DEB3),
    (Name: 'white'; Value: $FFFFFF),
    (Name: 'whitesmoke'; Value: $F5F5F5),
    (Name: 'yellow'; Value: $FFFF00),
    (Name: 'yellowgreen'; Value: $9ACD32)
  );

function HtmlEncodeString(const s: string): string;
begin
  result := TNetEncoding.HTML.Encode(s);
end;

function HtmlEncodeAttr(const s: string; NonBreakingSpaces: boolean): string;
begin
  result := TNetEncoding.HTML.Encode(s);
  result := StringReplace(result, '''', '&apos;', [rfReplaceAll]);
  if NonBreakingSpaces then
    result := StringReplace(result, ' ', '&nbsp;', [rfReplaceAll]);
end;

function IsUncSharePath(const path: string): Boolean;
begin
  // e.g \\server\share\foo or //server/share/foo
  Result := (Length(path) > 2) and (path[1] in ['\', '/']) and (path[2] in ['\', '/']);
end;

function EncodeUrl(const url: string): string;
begin
  if not IsUNCSharePath(url) then
    Result := TNetEncoding.URL.Encode(url)
  else
    Result := url;
end;

function ColorToHtml(Color: TColor): string;
var
  c: TColorRef;
  i, v: integer;
begin
  c := ColorToRGB(Color);
  v := c and $FF0000 shr 16 + c and $00FF00 + c and $0000FF shl 16;
  for i := Low(HTMLColorTable) to High(HTMLColorTable) do
    if HTMLColorTable[i].Value = v then
      exit(HTMLColorTable[i].Name);
  result := Format('#%.2x%.2x%.2x', [GetRValue(c), GetGValue(c), GetBValue(c)]);
end;

function TextToHtml(VisualText: TRtfVisualText): TRtfHtmlStyle;
var
  HtmlStyle: TRtfHtmlStyle;
  TextFormat: TRtfTextFormat;
  BackgroundColor, ForegroundColor: TColor;
begin
  if VisualText = nil then
    raise EArgumentNilException.Create(sNilVisualText);

  HtmlStyle := TRtfHtmlStyle.Create;
  TextFormat := VisualText.Format;

  // background color
  BackgroundColor := TextFormat.BackgroundColor;
  if BackgroundColor <> clNone then
    HtmlStyle.BackgroundColor := ColorToHtml(BackgroundColor);

  // foreground color
  ForegroundColor := TextFormat.ForegroundColor;
  if ForegroundColor <> clBlack then
    HtmlStyle.ForegroundColor := ColorToHtml(ForegroundColor);

  // font
  HtmlStyle.FontFamily := TextFormat.Font.Name;
  if TextFormat.FontSize > 0 then
    HtmlStyle.FontSize := FloatToStr(TextFormat.FontSize / 2) + 'pt';

  Result := HtmlStyle;
end;

function TextToBinary(HexStr: string): TBytes;
var
  HexDigits, DataSize, DataPos, i, e: Integer;
  Hex: string;
  C: Char;
  B: byte;
begin
  if HexStr = '' then
    raise EArgumentException.Create(sEmptyHexStr);

  HexDigits := Length(HexStr);
  DataSize := HexDigits div 2;
  SetLength(Result, DataSize);
  Hex := '';
  DataPos := 0;

  for i := 1 to HexDigits do
  begin
    C := HexStr[i];
    if CharInSet(C, [' ', #9, #10, #13]) then
      Continue;

    Hex := Hex + C;
    if Length(Hex) = 2 then
    begin
      Val('$' + Hex, B, e);
      if e <> 0 then
        raise EArgumentException.CreateFmt(sInvalidHexStr, [Hex]);
      Result[DataPos] := B;
      Inc(DataPos);
      Hex := '';
    end;
  end;
end;

function TwipToPixel(const x: integer): integer;
begin
  result := Round(x * 96.0 / 1440);
end;

end.
