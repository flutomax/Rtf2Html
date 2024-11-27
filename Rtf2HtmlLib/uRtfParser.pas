unit uRtfParser;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, uRtfTypes,
  uRtfElement, uRtfParserListener;

type

  TRtfParser = class(TObject)
  private
    fCurText: TStringBuilder;
    fEncoding: TEncoding;
    fLevel: integer;
    fGroupCount: integer;
    fTagCount: integer;
    fTagCountAtLastGroupStart: integer;
    fUnicodeSkipCount: integer;
    fFontTableStartLevel: integer;
    fTargetFont: string;
    fExpectingThemeFont: boolean;
    fListeners: TList<TRtfParserListener>;
    fUnicodeSkipCountStack: TStack<Integer>;
    fCodePageStack: TStack<Cardinal>;
    fFontToCodePageMapping: TDictionary<string, Integer>;
    fHexDecodingBuffer: TMemoryStream;
    fByteDecodingBuffer: TBytes;
    fCharDecodingBuffer: TCharArray;
    fIgnoreContentAfterRootGroup: boolean;
    function PeekNextChar(reader: TStringReader; mandatory: Boolean): Integer;
    function ReadOneByte(reader: TStringReader): byte;
    function ReadOneChar(reader: TStringReader): char;
    function HandleTag(reader: TStringReader; tag: TRtfTag): boolean;
    function GetIgnoreContentAfterRootGroup: boolean;
    procedure SetIgnoreContentAfterRootGroup(value: boolean);
    procedure ParseTag(reader: TStringReader);
    procedure FlushText;
    procedure DecodeCurrentHexBuffer;
    procedure DoParse(reader: TStringReader);
    procedure NotifyTagFound(tag: TRtfTag);
    procedure NotifyTextFound(text: TRtfText);
    procedure NotifyGroupBegin;
    procedure NotifyGroupEnd;
    procedure NotifyParseBegin;
    procedure NotifyParseSuccess;
    procedure NotifyParseEnd;
    procedure NotifyParseFail(E: Exception);
    procedure UpdateEncoding(codepage: Cardinal); overload;
    procedure UpdateEncoding(tag: TRtfTag); overload;
    procedure ByteConvert(const Bytes: TBytes; ByteIndex, ByteCount: Integer;
      const Chars: TCharArray; CharIndex, CharCount: Integer;
      out BytesUsed, CharsUsed: Integer; out Completed: Boolean);
  public
    constructor Create; overload;
    constructor Create(listeners: TArray<TRtfParserListener>); overload;
    destructor Destroy; override;
    procedure AddParserListener(listener: TRtfParserListener);
    procedure Parse(reader: TStringReader);
    property IgnoreContentAfterRootGroup: boolean read GetIgnoreContentAfterRootGroup
      write SetIgnoreContentAfterRootGroup;
  end;

implementation

uses
  System.RTLConsts, System.Math, System.Character, uRtfSpec, uRtfMessages;

function IsHexDigit(b: Byte): Boolean;
const
  HexDigits = '0123456789abcdefABCDEF';
begin
  Result := Pos(Chr(b), HexDigits) <> 0;
end;

function IsASCIILetter(i: Integer): Boolean;
var
  c: char;
begin
  c := char(i);
  result := ((c >= 'a') and (c <= 'z')) or ((c >= 'A') and (c <= 'Z'));
end;

{ TRtfParser }

constructor TRtfParser.Create;
begin
  inherited Create;
  fFontToCodePageMapping := TDictionary<string, Integer>.Create;
  fHexDecodingBuffer := TMemoryStream.Create;
  fCurText := TStringBuilder.Create;
  fListeners := TList<TRtfParserListener>.Create;
  fUnicodeSkipCountStack := TStack<Integer>.Create;
  fCodePageStack := TStack<Cardinal>.Create;
  fEncoding := TEncoding.Default;
  SetLength(fByteDecodingBuffer, 8);
  SetLength(fCharDecodingBuffer, 1);
  fLevel := 0;
  fIgnoreContentAfterRootGroup := true;
end;

constructor TRtfParser.Create(listeners: TArray<TRtfParserListener>);
var
  i: integer;
begin
  Create;
  for i := Low(listeners) to High(listeners)  do
    AddParserListener(listeners[i]);
end;

destructor TRtfParser.Destroy;
begin
  fCodePageStack.Free;
  fUnicodeSkipCountStack.Free;
  fCurText.Free;
  fListeners.Free;
  fFontToCodePageMapping.Free;
  fHexDecodingBuffer.Free;
  if Assigned(fEncoding) and not TEncoding.IsStandardEncoding(fEncoding) then
    fEncoding.Free;
  inherited;
end;

procedure TRtfParser.ByteConvert(const Bytes: TBytes; ByteIndex, ByteCount: Integer;
  const Chars: TCharArray; CharIndex, CharCount: Integer;
  out BytesUsed, CharsUsed: Integer; out Completed: Boolean);
begin
  // Get ready to do it
  BytesUsed := ByteCount;

  // Its easy to do if it won't overrun our buffer.
  while BytesUsed > 0 do
  begin
    if fEncoding.GetCharCount(Bytes, ByteIndex, BytesUsed) <= CharCount then
    begin
      CharsUsed := fEncoding.GetChars(Bytes, ByteIndex, BytesUsed, Chars, CharIndex);
      Completed := (CharsUsed = ByteCount);
      exit;
    end;
      // Try again with 1/2 the count, won't flush then 'cause won't read it all
      BytesUsed := BytesUsed div 2;
  end;

  // Oops, we didn't have anything, we'll have to throw an overflow
  raise EArgumentOutOfRangeException.Create(sArgumentConversionOverflow);
end;

procedure TRtfParser.AddParserListener(listener: TRtfParserListener);
begin
  if listener = nil then
    raise EArgumentNilException.CreateRes(@SArgumentNil);
  if not fListeners.Contains(listener) then
    fListeners.Add(listener);
end;

procedure TRtfParser.FlushText;
begin
  if fCurText.Length > 0 then
  begin
    if fLevel = 0 then
      raise ERtfStructure.CreateFmt(sTextOnRootLevel, [fCurText.ToString]);
    NotifyTextFound(TRtfText.Create(fCurText.ToString));
    fCurText.Clear;
  end;
end;

function TRtfParser.HandleTag(reader: TStringReader; tag: TRtfTag): boolean;

function HandleTagUnicodeCode(skippedcontent: boolean): boolean;
var
  i: integer;
  uchar, nextchar, secondchar: char;
begin
  result := skippedcontent;
	uchar := char(tag.ValueAsNumber);
	fCurText.Append(uchar);
  // skip over the indicated number of 'alternative representation' text
  i := 0;
  while i < fUnicodeSkipCount do
  begin
    nextchar := char(PeekNextChar(reader, true));
    case nextchar of
      #10,#13,#32:
      begin
        reader.Read; // consume peeked char
        result := true;
        if i = 0 then
          Dec(i);
          // the first whitespace after the tag
					// -> only a delimiter, doesn't count for skipping ...
      end;
      '\':
      begin
        reader.Read; // consume peeked char
        result := true;
        secondchar := char(ReadOneByte(reader)); // mandatory
        if secondchar = #39 then
        begin
          // ok, this is a hex-encoded 'byte' -> need to consume both
					// hex digits too
					ReadOneByte(reader); // high nibble
					ReadOneByte(reader); // low nibble
        end;
      end;
      '{', '}':
        // don't consume peeked char and abort skipping
        i := fUnicodeSkipCount;
      else
      begin
        reader.Read; // consume peeked char
        result := true;
      end;
    end;
    Inc(i);
  end;
end;

var
  detectfontname, skippedcontent: boolean;
  tagname: string;
  charset, codepage, newskipcount: integer;
begin
  if fLevel = 0 then
    raise ERtfStructure.CreateFmt(sTagOnRootLevel, [tag.ToString]);
  tagname := tag.Name;
  // this only handles the initial encoding tag in the header section
  if fTagCount < 4 then
    UpdateEncoding(tag);
  detectfontname := fExpectingThemeFont;

  if fTagCountAtLastGroupStart = fTagCount then
  begin
    fExpectingThemeFont := (tagname = TagThemeFontLoMajor) or
      (tagname = TagThemeFontHiMajor) or (tagname = TagThemeFontDbMajor) or
      (tagname = TagThemeFontBiMajor) or (tagname = TagThemeFontLoMinor) or
      (tagname = TagThemeFontHiMinor) or (tagname = TagThemeFontDbMinor) or
      (tagname = TagThemeFontBiMinor);
    detectfontname := true;
  end;

  if detectfontname then
  begin
    if tagname = TagFont then
    begin
      if fFontTableStartLevel > 0 then
      begin
        // in the font-table definition:
        //-> remember the target font for charset mapping
        fTargetFont := tag.FullName;
        fExpectingThemeFont := false;
      end;
    end
    else
    if tagname = TagFontTable then
    begin
      // -> remember we're in the font-table definition
      fFontTableStartLevel := fLevel;
    end;
  end;

  if not fTargetFont.IsEmpty then
  begin
    if TagFontCharset.Equals(tagname) then
    begin
			charset := tag.ValueAsNumber;
			codepage := GetCodePage(charset);
      if fFontToCodePageMapping.ContainsKey(fTargetFont) then
        fFontToCodePageMapping[fTargetFont] := codepage
      else
        fFontToCodePageMapping.Add(fTargetFont, codepage);
      UpdateEncoding(codepage);
    end;
  end;

  if (fFontToCodePageMapping.Count > 0) and (tagname = TagFont) then
  begin
    if fFontToCodePageMapping.TryGetValue(tag.FullName, codepage) then
      UpdateEncoding(codePage);
  end;

  skippedcontent := false;
  if tagname = TagUnicodeCode then
    skippedcontent := HandleTagUnicodeCode(skippedcontent)
  else
  if tagname = TagUnicodeSkipCount then
  begin
    newskipcount := tag.ValueAsNumber;
    if (newskipcount < 0) or (newskipcount > 10) then
      raise ERtfUnicodeEncoding.CreateFmt(sInvalidUnicodeSkipCount, [tag.ToString]);
    fUnicodeSkipCount := newskipcount;
  end
  else
    FlushText;
  NotifyTagFound(tag);
  Inc(fTagCount);
  result := skippedcontent;
end;

procedure TRtfParser.NotifyGroupBegin;
var
  listener: TRtfParserListener;
begin
  for listener in FListeners do
    listener.GroupBegin;
end;

procedure TRtfParser.NotifyGroupEnd;
var
  listener: TRtfParserListener;
begin
  for listener in fListeners do
    listener.GroupEnd;
end;

procedure TRtfParser.NotifyParseBegin;
var
  listener: TRtfParserListener;
begin
  for listener in fListeners do
    listener.ParseBegin;
end;

procedure TRtfParser.NotifyParseEnd;
var
  listener: TRtfParserListener;
begin
  for listener in fListeners do
    listener.ParseEnd;
end;

procedure TRtfParser.NotifyParseFail(E: Exception);
var
  listener: TRtfParserListener;
begin
  for listener in fListeners do
    listener.ParseFail(E);
end;

procedure TRtfParser.NotifyParseSuccess;
var
  listener: TRtfParserListener;
begin
  for listener in fListeners do
    listener.ParseSuccess;
end;

procedure TRtfParser.NotifyTagFound(tag: TRtfTag);
var
  listener: TRtfParserListener;
begin
  for listener in fListeners do
    listener.TagFound(tag);
end;

procedure TRtfParser.NotifyTextFound(text: TRtfText);
var
  listener: TRtfParserListener;
begin
  for listener in fListeners do
    listener.TextFound(text);
end;

procedure TRtfParser.ParseTag(reader: TStringReader);
var
  tagname, tagvalue: string;
  readingname, delimreached,
  skippedcontent: boolean;
  nextchar: Integer;
  newtag: TRtfTag;
begin
  tagname := '';
  tagvalue := '';
  readingname := true;
  delimreached := false;
  nextchar := PeekNextChar(reader, true);
  while not delimreached do
  begin
    if readingname and IsASCIILetter(nextchar) then
      tagname := tagname + ReadOneChar(reader) // must still consume the 'peek'ed char
    else
    if IsDigit(char(nextchar)) or ((nextchar = Ord('-')) and (tagvalue = '')) then
    begin
      readingname := false;
      tagvalue := tagvalue + ReadOneChar(reader); // must still consume the 'peek'ed char
    end
    else
    begin
      delimreached := true;
      if tagvalue.Length > 0 then
        newtag := TRtfTag.Create(tagname, tagvalue)
      else
        newtag := TRtfTag.Create(tagName);

      skippedcontent := HandleTag(reader, newtag);
      if (nextchar = 32) and not skippedcontent then
        reader.Read; // must still consume the 'peek'ed char
    end;
    if not delimreached then
      nextchar := PeekNextChar(reader, true);
  end;
end;

function TRtfParser.PeekNextChar(reader: TStringReader;
  mandatory: Boolean): Integer;
begin
  result := reader.Peek;
  if mandatory and (result = -1) then
    raise ERtfMultiByteEncoding.Create(sEndOfFileInvalidCharacter);
end;

function TRtfParser.ReadOneByte(reader: TStringReader): byte;
var
  bytevalue: integer;
begin
  bytevalue := reader.Read;
  if bytevalue = -1 then
    raise ERtfMultiByteEncoding.Create(sEndOfFileInvalidCharacter);
  result := byte(bytevalue);
end;

function TRtfParser.ReadOneChar(reader: TStringReader): char;
var
  completed: boolean;
  byteindex, usedbytes, usedchars: integer;
begin
  // NOTE: the handling of multi-byte encodings is probably not the most efficient here ...

	completed := false;
	byteindex := 0;
	while not completed do
  begin
		fByteDecodingBuffer[byteIndex] := ReadOneByte(reader);
		Inc(byteindex);
    ByteConvert(fByteDecodingBuffer, 0, byteindex, fCharDecodingBuffer, 0, 1,
			usedbytes, usedchars, completed);
		if (completed and ((usedbytes <> byteindex) or (usedChars <> 1))) then
      raise ERtfMultiByteEncoding.Create(
      InvalidMultiByteEncoding(fByteDecodingBuffer, byteindex, fEncoding));
	end;
	result := fCharDecodingBuffer[0];
end;

procedure TRtfParser.DecodeCurrentHexBuffer;
var
  bytecount: Int64;
  bytes: TBytes;
  chars: TCharArray;
  startindex,
  usedbytes, usedchars: Integer;
  completed: Boolean;
  s: string;
begin
  bytecount := fHexDecodingBuffer.Size;
  if bytecount > 0 then
  begin
    SetLength(bytes, bytecount);
    fHexDecodingBuffer.Position := 0;
    fHexDecodingBuffer.ReadBuffer(bytes[0], bytecount);
    SetLength(chars, bytecount); // should be enough

    startindex := 0;
    completed := false;
    while (not completed) and (startindex < Length(bytes)) do
    begin
      usedbytes := 0;
      usedchars := 0;
      ByteConvert(bytes, startindex, Length(bytes) - startindex,
        chars, 0, Length(chars), usedbytes, usedchars, completed);
      SetString(s, PChar(chars), usedchars);
      fCurText.Append(s, 0, usedchars);
      Inc(startindex, usedbytes);
    end;

    fHexDecodingBuffer.Clear;
  end;
end;

procedure TRtfParser.UpdateEncoding(tag: TRtfTag);
begin
  if tag.Name = TagEncodingAnsi then
    UpdateEncoding(AnsiCodePage)
  else
  if tag.Name = TagEncodingMac then
		UpdateEncoding(10000)
  else
  if tag.Name = TagEncodingPc then
    UpdateEncoding(437)
  else
  if tag.Name = TagEncodingPca then
    UpdateEncoding(850)
  else
  if tag.Name = TagEncodingAnsiCodePage then
    UpdateEncoding(tag.ValueAsNumber);
end;

procedure TRtfParser.UpdateEncoding(codepage: Cardinal);
begin
  if codepage <> fEncoding.CodePage then
  begin
    if Assigned(fEncoding) and not TEncoding.IsStandardEncoding(fEncoding) then
      fEncoding.Free;
    if codepage = SymbolFakeCodePage then
      fEncoding := TEncoding.Default
    else
      fEncoding := TEncoding.GetEncoding(codepage);
  end;
end;

procedure TRtfParser.DoParse(reader: TStringReader);
const
  eof = -1;
var
  peekchar, nextchar, secondchar,
  decodedbyte,
  hex1, hex2: byte;
  backslash, peekcharvalid,
  mustflush: boolean;
begin
  fUnicodeSkipCountStack.Clear;
  fCodePageStack.Clear;
	fUnicodeSkipCount := 1;
	fLevel := 0;
	fTagCountAtLastGroupStart := 0;
	fTagCount := 0;
	fFontTableStartLevel := -1;
	fTargetFont := '';
	fExpectingThemeFont := false;
	fFontToCodePageMapping.Clear;
	fHexDecodingBuffer.Clear;
	UpdateEncoding(AnsiCodePage);
	fGroupCount := 0;
  backslash := false;
  nextchar := PeekNextChar(reader, false);
  while nextchar <> eof do
  begin
    peekchar := 0;
		peekcharvalid := false;
    case Char(nextchar) of
      '\':
      begin
        if not backslash then
          reader.Read; // must still consume the 'peek'ed char
        secondchar := PeekNextChar(reader, true);
        case Char(secondchar) of
          '\', '{', '}':
            begin
              fCurText.Append(ReadOneChar(reader)); // must still consume the 'peek'ed char
            end;

          #10, #13:
            begin
              reader.Read; // must still consume the 'peek'ed char
              // must be treated as a 'par' tag if preceded by a backslash
              // (see RTF spec page 144)
              HandleTag(reader, TRtfTag.Create(TagParagraph));
            end;

          #39:
            begin
              reader.Read; // must still consume the 'peek'ed char
              hex1 := ReadOneByte(reader);
              hex2 := ReadOneByte(reader);
              if not IsHexDigit(hex1) then
                raise ERtfHexEncoding.CreateFmt('Invalid first hex digit: %s', [hex1]);
              if not IsHexDigit(hex2) then
                raise ERtfHexEncoding.CreateFmt('Invalid second hex digit: %s', [hex2]);
              decodedbyte := StrToInt('$' + Chr(hex1) + Chr(hex2));
              fHexDecodingBuffer.WriteData(decodedbyte);
              peekchar := PeekNextChar(reader, false);
              peekcharvalid := true;
              mustflush := true;
              if peekchar = Ord('\') then
              begin
                reader.Read;
                backslash := true;
                if PeekNextChar(reader, false) = 39 then
                  mustflush := false;
              end;
              if mustflush then
              begin
                // we may _NOT_ handle hex content in a character-by-character way as
                // this results in invalid text for japanese/chinese content ...
                // -> we wait until the following content is non-hex and then flush the
                //    pending data. ugly but necessary with our decoding model.
                DecodeCurrentHexBuffer;
              end;
            end;

          '|', '~', '-', '_', ':', '*':
            begin
              HandleTag(reader, TRtfTag.Create(ReadOneChar(reader))); // must still consume the 'peek'ed char
            end;

          else
            ParseTag(reader);
        end;
      end;

    #10, #13:
      begin
        reader.Read; // must still consume the 'peek'ed char
      end;

    #9:
      begin
        reader.Read; // must still consume the 'peek'ed char
        // should be treated as a 'tab' tag (see RTF spec page 144)
        HandleTag(reader, TRtfTag.Create(TagTabulator));
      end;
      '{':
      begin
        reader.Read;
        FlushText;
				NotifyGroupBegin;
				fTagCountAtLastGroupStart := fTagCount;
				fUnicodeSkipCountStack.Push(fUnicodeSkipCount);
				fCodePageStack.Push(IfThen(Assigned(fEncoding), fEncoding.CodePage));
				Inc(fLevel);
      end;
      '}':
      begin
        reader.Read;
        FlushText;
				if fLevel > 0 then
        begin
					fUnicodeSkipCount := fUnicodeSkipCountStack.Pop;
					if fFontTableStartLevel = fLevel then
					begin
						fFontTableStartLevel := -1;
						fTargetFont := '';
						fExpectingThemeFont := false;
          end;
					UpdateEncoding(fCodePageStack.Pop);
					Dec(fLevel);
					NotifyGroupEnd;
					Inc(fGroupCount);
				end
				else
          raise ERtfBraceNesting.Create(sToManyBraces);
      end;
      else
        fCurText.Append(ReadOneChar(reader));
    end; // case
    if (fLevel = 0) and IgnoreContentAfterRootGroup then
			 break;

    if peekcharvalid then
			nextchar := peekchar
		else
		begin
      nextChar := PeekNextChar(reader, false);
      backslash := false;
		end;
  end; // while
  FlushText;
  fCurText.Clear;
  reader.Close;
  if fLevel > 0 then
    raise ERtfBraceNesting.Create(sToFewBraces);
  if fGroupCount = 0 then
    raise ERtfEmptyDocument.Create(sNoRtfContent);
end;

function TRtfParser.GetIgnoreContentAfterRootGroup: boolean;
begin
  result := fIgnoreContentAfterRootGroup;
end;

procedure TRtfParser.SetIgnoreContentAfterRootGroup(value: boolean);
begin
  fIgnoreContentAfterRootGroup := value;
end;

procedure TRtfParser.Parse(reader: TStringReader);
begin
  NotifyParseBegin;
  try
    try
      DoParse(Reader);
      NotifyParseSuccess;
    except
      on E: Exception do
        NotifyParseFail(E);
    end;
  finally
    NotifyParseEnd;
  end;
end;


end.
