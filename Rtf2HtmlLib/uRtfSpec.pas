unit uRtfSpec;

interface

uses
  System.SysUtils;

const
  // --- rtf general ----
  TagRtf = 'rtf';
  RtfVersion1 = 1;

  TagGenerator = 'generator';
  TagViewKind = 'viewkind';

  // --- encoding ----
  TagEncodingAnsi = 'ansi';
  TagEncodingMac = 'mac';
  TagEncodingPc = 'pc';
  TagEncodingPca = 'pca';
  TagEncodingAnsiCodePage = 'ansicpg';
  AnsiCodePage = 1252;
  SymbolFakeCodePage = 42; // a windows legacy hack ...

  TagUnicodeSkipCount = 'uc';
  TagUnicodeCode = 'u';
  TagUnicodeAlternativeChoices = 'upr';
  TagUnicodeAlternativeUnicode = 'ud';

  // --- font ----
  TagFontTable = 'fonttbl';
  TagDefaultFont = 'deff';
  TagFont = 'f';
  TagFontKindNil = 'fnil';
  TagFontKindRoman = 'froman';
  TagFontKindSwiss = 'fswiss';
  TagFontKindModern = 'fmodern';
  TagFontKindScript = 'fscript';
  TagFontKindDecor = 'fdecor';
  TagFontKindTech = 'ftech';
  TagFontKindBidi = 'fbidi';
  TagFontCharset = 'fcharset';
  TagFontPitch = 'fprq';
  TagFontSize = 'fs';
  TagFontDown = 'dn';
  TagFontUp = 'up';
  TagFontSubscript = 'sub';
  TagFontSuperscript = 'super';
  TagFontNoSuperSub = 'nosupersub';

  TagThemeFontLoMajor = 'flomajor'; // these are 'theme' fonts
  TagThemeFontHiMajor = 'fhimajor'; // used in new font tables
  TagThemeFontDbMajor = 'fdbmajor';
  TagThemeFontBiMajor = 'fbimajor';
  TagThemeFontLoMinor = 'flominor';
  TagThemeFontHiMinor = 'fhiminor';
  TagThemeFontDbMinor = 'fdbminor';
  TagThemeFontBiMinor = 'fbiminor';

  TagsFontSpecials: array[0..8] of string = (TagFont, TagThemeFontLoMajor,
    TagThemeFontHiMajor, TagThemeFontDbMajor, TagThemeFontBiMajor,
    TagThemeFontLoMinor, TagThemeFontHiMinor, TagThemeFontDbMinor,
    TagThemeFontBiMinor);

  DefaultFontSize = 24;

  TagCodePage = 'cpg';

  // --- color ----
  TagColorTable = 'colortbl';
  TagColorRed = 'red';
  TagColorGreen = 'green';
  TagColorBlue = 'blue';
  TagColorForeground = 'cf';
  TagColorBackground = 'cb';
  TagColorBackgroundWord = 'chcbpat';
  TagColorHighlight = 'highlight';

  // --- header/footer ----
  TagHeader = 'header';
  TagHeaderFirst = 'headerf';
  TagHeaderLeft = 'headerl';
  TagHeaderRight = 'headerr';
  TagFooter = 'footer';
  TagFooterFirst = 'footerf';
  TagFooterLeft = 'footerl';
  TagFooterRight = 'footerr';
  TagFootnote = 'footnote';

  // --- character ----
  TagDelimiter = ';';
  TagExtensionDestination = '*';
  TagTilde = '~';
  TagHyphen = '-';
  TagUnderscore = '_';

  // --- indent ---
  TagFirstLineIndent = 'fi';
  TagLeftIndent = 'li';
  TagRightIndent = 'ri';

  // --- special character ----
  TagPage = 'page';
  TagSection = 'sect';
  TagParagraph = 'par';
  TagLine = 'line';
  TagTabulator = 'tab';
  TagEmDash = 'emdash';
  TagEnDash = 'endash';

  TagEmSpace = 'emspace';
	TagEnSpace = 'enspace';
	TagQmSpace = 'qmspace';
	TagBulltet = 'bullet';
	TagLeftSingleQuote = 'lquote';
	TagRightSingleQuote = 'rquote';
	TagLeftDoubleQuote = 'ldblquote';
	TagRightDoubleQuote = 'rdblquote';

	// --- format ----
	TagPlain = 'plain';
	TagParagraphDefaults = 'pard';
	TagSectionDefaults = 'sectd';

	TagBold = 'b';
	TagItalic = 'i';
	TagUnderLine = 'ul';
	TagUnderLineNone = 'ulnone';
	TagStrikeThrough = 'strike';
	TagHidden = 'v';
	TagAlignLeft = 'ql';
	TagAlignCenter = 'qc';
	TagAlignRight = 'qr';
	TagAlignJustify = 'qj';

	TagStyleSheet = 'stylesheet';

	// --- info ----
	TagInfo = 'info';
	TagInfoVersion = 'version';
	TagInfoRevision = 'vern';
	TagInfoNumberOfPages = 'nofpages';
	TagInfoNumberOfWords = 'nofwords';
	TagInfoNumberOfChars = 'nofchars';
	TagInfoId = 'id';
	TagInfoTitle = 'title';
	TagInfoSubject = 'subject';
	TagInfoAuthor = 'author';
	TagInfoManager = 'manager';
	TagInfoCompany = 'company';
	TagInfoOperator = 'operator';
	TagInfoCategory = 'category';
	TagInfoKeywords = 'keywords';
	TagInfoComment = 'comment';
	TagInfoDocumentComment = 'doccomm';
	TagInfoHyperLinkBase = 'hlinkbase';
	TagInfoCreationTime = 'creatim';
	TagInfoRevisionTime = 'revtim';
	TagInfoPrintTime = 'printim';
	TagInfoBackupTime = 'buptim';
	TagInfoYear = 'yr';
	TagInfoMonth = 'mo';
	TagInfoDay = 'dy';
	TagInfoHour = 'hr';
	TagInfoMinute = 'min';
	TagInfoSecond = 'sec';
	TagInfoEditingTimeMinutes = 'edmins';

	// --- user properties ----
	TagUserProperties = 'userprops';
	TagUserPropertyType = 'proptype';
	TagUserPropertyName = 'propname';
	TagUserPropertyValue = 'staticval';
	TagUserPropertyLink = 'linkval';

	// this table is from the RTF specification 1.9.1, page 40
	PropertyTypeInteger = 3;
	PropertyTypeRealNumber = 5;
	PropertyTypeDate = 64;
	PropertyTypeBoolean = 11;
	PropertyTypeText = 30;

	// --- picture ----
	TagPicture = 'pict';
	TagPictureWrapper = 'shppict';
	TagPictureWrapperAlternative = 'nonshppict';
	TagPictureFormatEmf = 'emfblip';
	TagPictureFormatPng = 'pngblip';
	TagPictureFormatJpg = 'jpegblip';
	TagPictureFormatPict = 'macpict';
	TagPictureFormatOs2Metafile = 'pmmetafile';
	TagPictureFormatWmf = 'wmetafile';
	TagPictureFormatWinDib = 'dibitmap';
	TagPictureFormatWinBmp = 'wbitmap';
	TagPictureWidth = 'picw';
	TagPictureHeight = 'pich';
	TagPictureWidthGoal = 'picwgoal';
	TagPictureHeightGoal = 'pichgoal';
	TagPictureWidthScale = 'picscalex';
	TagPictureHeightScale = 'picscaley';

	// --- bullets/numbering ----
	TagParagraphNumberText = 'pntext';
	TagListNumberText = 'listtext';

  // --- tables ---
  TagTableRowDefaults = 'trowd';
  TagTableRowBreak = 'row';
  TagTableCellDefaults = 'tcelld';
  TagTableCellBreak = 'cell';
  TagTableHalfCellPadding = 'trgaph';
  TagTableBottomCellSpacing = 'trspdb';
  TagTableLeftCellSpacing = 'trspdl';
  TagTableRightCellSpacing = 'trspdr';
  TagTableTopCellSpacing = 'trspdt';
  TagTableRightCellBoundary = 'cellx';
  TagTableRowAutoFit = 'trautofit';
  TagTableRowLeft = 'trleft';
  TagTableRowHeight = 'trrh';
  TagTableRowTextAlign = 'trq';
  TagTableTableBorderSide = 'trbrdr';
  TagTableParagraphBorderSide = 'brdr';
  TagTableParagraphBorderBetween = 'brdrbtw';
  TagTableCellBorderSide = 'clbrdr';
  TagTableTablePaddingTop = 'trpaddt';
  TagTableTablePaddingRight = 'trpaddr';
  TagTableTablePaddingBottom = 'trpaddb';
  TagTableTablePaddingLeft = 'trpaddl';
  TagTableCellWidthType = 'clftsWidth';
  TagTableCellWidth = 'clwWidth';
  TagTableCellFirstMerged = 'clvmgf';
  TagTableCellMerged = 'clvmrg';
  TagTableCellMergePrevious = 'clmrg';
  TagTableCellBorderBottom = 'clbrdrb';
  TagTableCellBorderTop = 'clbrdrt';
  TagTableCellBorderLeft = 'clbrdrl';
  TagTableCellBorderRight = 'clbrdrr';
  TagTableCellVerticalAlign = 'clvertal';
  TagTableCellVerticalAlignTop = 'clvertalt';
  TagTableCellVerticalAlignCenter = 'clvertalc';
  TagTableCellVerticalAlignBottom = 'clvertalb';
  TagTableCellPaddingLeft = 'clpadl';
  TagTableCellPaddingTop = 'clpadt';
  TagTableCellPaddingBottom = 'clpadb';
  TagTableCellPaddingRight = 'clpadr';
  TagTableCellPaddingUnit = 'clpadf';
  TagTableInTable = 'intbl';
  TagTableCellBackgroundColor = 'clcbpat';
  TagTableHeaderRow = 'trhdr';
  TagTableNestedTableProperties = 'nesttableprops';
  TagTableNoNestedTables = 'nonesttables';
  TagTableNestedCellBreak = 'nestcell';
  TagTableNestedRowBreak = 'nestrow';
  TagTableNestingLevel = 'itap';

  // borders
  TagBorderNone = 'brdrnone';
  TagBorderColor = 'brdrcf';
  TagBorderWidth = 'brdrw';


  function GetCodePage(charset: Integer): Integer;

implementation

function GetCodePage(charset: Integer): Integer;
begin
  case charSet of
    0: Result := 1252; // ANSI
    1: Result := 0; // Default
    2: Result := 42; // Symbol
    77: Result := 10000; // Mac Roman
    78: Result := 10001; // Mac Shift Jis
    79: Result := 10003; // Mac Hangul
    80: Result := 10008; // Mac GB2312
    81: Result := 10002; // Mac Big5
    82: Result := 0; // Mac Johab (old)
    83: Result := 10005; // Mac Hebrew
    84: Result := 10004; // Mac Arabic
    85: Result := 10006; // Mac Greek
    86: Result := 10081; // Mac Turkish
    87: Result := 10021; // Mac Thai
    88: Result := 10029; // Mac East Europe
    89: Result := 10007; // Mac Russian
    128: Result := 932; // Shift JIS
    129: Result := 949; // Hangul
    130: Result := 1361; // Johab
    134: Result := 936; // GB2312
    136: Result := 950; // Big5
    161: Result := 1253; // Greek
    162: Result := 1254; // Turkish
    163: Result := 1258; // Vietnamese
    177: Result := 1255; // Hebrew
    178: Result := 1256; // Arabic
    179: Result := 0; // Arabic Traditional (old)
    180: Result := 0; // Arabic user (old)
    181: Result := 0; // Hebrew user (old)
    186: Result := 1257; // Baltic
    204: Result := 1251; // Russian
    222: Result := 874; // Thai
    238: Result := 1250; // Eastern European
    254: Result := 437; // PC 437
    255: Result := 850; // OEM
  else
    Result := 0;
  end;
end;


end.
