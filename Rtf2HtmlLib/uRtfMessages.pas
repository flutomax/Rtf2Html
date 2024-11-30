unit uRtfMessages;

interface

uses
  System.SysUtils;

resourcestring

  sErrOpenFile = 'can''t open file "%s"';
  sNoRtfContent = 'no rtf content';
  sTextOnRootLevel = 'a text cannot appear on root level, must be child of a group: "%s"';
  sMultipleRootLevelGroups = 'invalid state: multiple root level groups';
  sMissingGroupForNewText = 'invalid state: no group available yet for adding a text';
  sToManyBraces = 'improper nesting of braces: too many';
  sToFewBraces = 'improper nesting of braces: too few';
  sUnclosedGroups = 'invalid state: unclosed groups';
  sArgumentConversionOverflow = 'argument conversion overflow';
  sEndOfFileInvalidCharacter = 'unexpected end of file while examining next character';
  sTagOnRootLevel = 'a tag cannot appear on root level, must be child of a group: "%s"';
  sMissingGroupForNewTag = 'invalid state: no group available yet for adding a tag';
  sCollectionToolInvalidEnum = '"%s" is not a valid value for %s. must be one of %s.';
  sDuplicateFont = 'duplicate font id "%s"';
  sColorTableUnsupportedText = 'unsupported text in color table: "%s"';
  sEmptyDocument = 'document has not contents';
  sMissingDocumentStartTag = 'first element in document is not a tag';
  sMissingRtfVersion = 'unspecified RTF version';
  sUnsupportedRtfVersion = 'unsupported RTF version: %d';
  sUndefinedFont = 'undefined font: "%s"';
  sUndefinedColor = 'undefined color index: %d';
  sFontSizeOutOfRange = 'invalid font size, must be in the range [1..0xFFFF], but is %d';
  sEmptyDestination = 'destination cannot be empty';
  sEmptyID = 'id cannot be empty';
  sEmptyText = 'text cannot be empty';
  sEmptyName = 'name cannot be empty';
  sEmptyList = 'list cannot be empty';
  sEmptyValue = 'value cannot be empty';
  sEmptyHexStr = 'hex string cannot be empty';
  sEmptyFileName = 'file name cannot be empty';
  sEmptySelectorName = 'selector name cannot be empty';
  sEmptyStaticValue = 'static value cannot be empty';
  sEmptyImageDataHex = 'image data hex cannot be empty';
  sEmptyFileNamePattern = 'file name pattern cannot be empty';
  sInvalidHexStr = 'invalid characters "%s" in image data hex';
  sInvalidInitGroupState = 'init: illegal state for group: "%s"';
  sInvalidGeneratorGroup = 'invalid generator group: "%s"';
  sInvalidInitTextState = 'init: illegal state for text: "%s"';
  sInvalidCharacterSet = 'invalid character set: %d';
  sInvalidFontSize = 'invalid font size: %d';
  sInvalidCodePage = 'invalid code page: %d';
  sInvalidClassName = 'invalid class name: %s';
  sInvalidColorValue = 'invalid color value: %d';
  sInvalidDocumentStartTag = 'first tag in document is not %s';
  sInvalidInitTagState = 'init: illegal state for tag "%s"';
  sInvalidMultiByteEncoding = 'could not decode bytes 0x%s to character with encoding %s (from codepage %d)';
  sInvalidUnicodeSkipCount = 'invalid unicode skip count: "%s"';
  sInvalidImageWidth = 'invalid image width: %d';
  sInvalidImageHeight = 'invalid image height: %d';
  sInvalidImageDesiredWidth = 'invalid image desired width: %d';
  sInvalidImageDesiredHeight = 'invalid image desired height: %d';
  sInvalidImageScaleWidth = 'invalid image scale width: %d';
  sInvalidImageScaleHeight = 'invalid image scale height: %d';
  sInvalidGraphicFormat = 'invalid graphic format';
  sInvalidTextContextState = 'invalid text context state';
  sInvalidDefaultFont = 'invalid default font: %s';
  sNilDocument = 'document cannot be nil';
  sNilFormat = 'format cannot be nil';
  sNilFont = 'font cannot be nil';
  sNilCopy = 'copy cannot be nil';
  sNilInfo = 'info cannot be nil';
  sNilItem = 'item cannot be nil';
  sNilText = 'text cannot be nil';
  sNilSrc = 'src cannot be nil';
  sNilVisitor = 'visitor cannot be nil';
  sNilListener = 'listener cannot be nil';
  sNilGraphics = 'graphics cannot be nil';
  sNilSettings = 'settings cannot be nil';
  sNilDefaultFont = 'default font cannot be nil';
  sNilDerivedBackgroundColor = 'derived background color cannot be nil';
  sNilDerivedForegroundColor = 'derived foreground color cannot be nil';
  sNilCollectedProperties = 'collected properties cannot be nil';
  sNilFontTable = 'font table cannot be nil';
  sNilColorTable = 'color table cannot be nil';
  sNilTemplateFormat = 'template format cannot be nil';
  sNilUniqueTextFormats = 'unique text formats cannot be nil';
  sNilUserProperties = 'user properties cannot be nil';
  sNilVisualContent = 'visual content cannot be nil';
  sNilVisualText = 'visual text cannot be nil';
  sNoValueVar = 'variable has no value';
  sNoPointerAllowed = 'pointer value not allowed';
  sNoCellDefs = 'cells are not defined';
  sNoRowLeftPoint = 'no row.left point';
  sNoCelldefRightPoint = 'no celldef.right point';
  sOutOfRangeHtmlTextWriterTag = 'html text writer tag out of range';
  sDirectoryNotFound = 'directory "%s" not found';
  sErrInterpreting = 'error while interpreting rtf: ';
  sErrConverting = 'error while converting to html: ';
  sErrConvEncoder = 'failed to get encoder';
  sErrConvPicture = 'failed converting picture: %s';
  sInfParseBegin = 'parse begin';
  sInfParseEnd = 'parse end';
  sInfParseFail = 'parse fail: ';
  sInfParseSucc = 'parse successfully completed';
  sInfInterprBegin = 'interpreting rtf begin';
  sInfInterprEnd = 'interpreting rtf end';
  sInfConvBegin = 'converting rtf to html begin';
  sInfConvEnd = 'converting rtf to html end';
  sInfConvSucc = 'converting rtf to html successfully completed';
  sUnequalCells = 'number of cells and number of defined cell are unequal';


function InvalidMultiByteEncoding(const buffer: TBytes; index: integer;
  encoding: TEncoding): string;

implementation

function InvalidMultiByteEncoding(const buffer: TBytes; index: integer;
  encoding: TEncoding): string;
var
  s: string;
  i: integer;
begin
  s := '';
  for i := 0 to index - 1 do
    s := s + (IntToHex(buffer[i], 2));
  with encoding do
    Result := Format(sInvalidMultiByteEncoding, [s, EncodingName, CodePage]);
end;

end.
