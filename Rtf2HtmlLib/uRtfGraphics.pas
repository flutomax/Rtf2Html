unit uRtfGraphics;

interface

uses
  System.Types, System.SysUtils, System.Classes, Vcl.Graphics, uRtfTypes,
  uRtfObjects, uRtfNullable, uRtfInterpreterListener, uRtfInterpreterContext;

type

  ERtfGraphics = class(Exception);

  TRtfGraphics = class(TObject)
  private
    fFileNamePattern: string;
    fDpiX: Double;
    fDpiY: Double;
  public
    constructor Create; overload;
    constructor Create(aFileNamePattern: string); overload;
    constructor Create(aFileNamePattern: string; aDpiX, aDpiY: Double); overload;
    function ResolveFileName(index: Integer; ImageFormat: TRtfImageFormat): string;
    function CalcImageWidth(format: TRtfImageFormat; width, desiredwidth,
      scalewidthpercent: Integer): Integer;
    function CalcImageHeight(format: TRtfImageFormat; height,
      desiredheight, scaleheightpercent: Integer): Integer;
  published
    property FileNamePattern: string read fFileNamePattern;
    property DpiX: Double read fDpiX;
    property DpiY: Double read fDpiY;
  end;

  TRtfGraphicsConvertSettings = record
    Graphics: TRtfGraphics;
    BackgroundColor: TNullable<TColor>;
    ImagesPath: string;
    ScaleImage: Boolean;
    ScaleOffset: Single;
    ScaleExtension: Single;
    constructor Create(AGraphics: TRtfGraphics);
    function GetImageFileName(Index: Integer;
      ARtfVisualImageFormat: TRtfImageFormat): string;
  end;

  TRtfGraphicsConvertInfo = class
  private
    fFileName: string;
    fFormat: TRtfImageFormat;
    fSize: TSize;
  public
    constructor Create(const AFileName: string; AFormat: TRtfImageFormat; ASize: TSize);
    function ToString: string;
    property FileName: string read fFileName;
    property Format: TRtfImageFormat read fFormat;
    property Size: TSize read fSize;
  end;

  TRtfGraphicsConvertInfoCollection = class(TBaseCollection<TRtfGraphicsConvertInfo>);

  TRtfGraphicsConverter = class(TRtfInterpreterListener)
  private
    fConvertedImages: TRtfGraphicsConvertInfoCollection;
    fSettings: TRtfGraphicsConvertSettings;
    procedure SaveImage(Stream: TStream; AFileName: string; ASize: TSize); overload;
  protected
    procedure DoBeginDocument(AContext: TRtfInterpreterContext); override;
    procedure DoInsertImage(AContext: TRtfInterpreterContext;
      AFormat: TRtfImageFormat;
      AWidth, AHeight, ADesiredWidth, ADesiredHeight,
      AScaleWidthPercent, AScaleHeightPercent: Integer;
      AImageDataHex: string); override;
    procedure SaveImage(AImageBuffer: TBytes; AFormat: TRtfImageFormat;
      AFileName: string; ASize: TSize); overload;
    procedure EnsureImagesPath(AImageFileName: string);
  public
    constructor Create(const ASettings: TRtfGraphicsConvertSettings);
    destructor Destroy; override;
    function GetSettings: TRtfGraphicsConvertSettings;
    function GetConvertedImages: TRtfGraphicsConvertInfoCollection;
  end;


implementation

uses
  System.TypInfo, System.StrUtils, Winapi.ActiveX, Winapi.GDIPAPI,
  Winapi.GDIPOBJ, Winapi.GDIPUTIL, uRtfMessages, uRtfHtmlFunctions,
  uRtfHtmlObjects;

const
  DefaultDpi = 96.0;
  TwipsPerInch = 1440;
  ExtFromRtfImageFormat: array[TRtfImageFormat] of string =
    ('', '.png', '.png', '.jpg', '.png', '.png');


{ TRtfGraphics }

constructor TRtfGraphics.Create;
begin
  Create(DefaultFileNamePattern, DefaultDpi, DefaultDpi);
end;

constructor TRtfGraphics.Create(aFileNamePattern: string);
begin
  Create(aFileNamePattern, DefaultDpi, DefaultDpi);
end;

constructor TRtfGraphics.Create(aFileNamePattern: string;
  aDpiX, aDpiY: Double);
begin
  if aFileNamePattern = '' then
    raise EArgumentException.Create(sEmptyFileNamePattern);

  fFileNamePattern := aFileNamePattern;
  fDpiX := aDpiX;
  fDpiY := aDpiY;
end;

function TRtfGraphics.ResolveFileName(index: Integer;
  ImageFormat: TRtfImageFormat): string;
begin
  Result := Format(fFileNamePattern, [index, ExtFromRtfImageFormat[ImageFormat]]);
end;

function TRtfGraphics.CalcImageWidth(format: TRtfImageFormat;
  width, desiredwidth, scalewidthpercent: Integer): Integer;
var
  ScaleX: Double;
begin
  ScaleX := scalewidthpercent / 100.0;
  Result := Round(desiredwidth * ScaleX / TwipsPerInch * dpiX);
end;

function TRtfGraphics.CalcImageHeight(format: TRtfImageFormat;
  height, desiredheight, scaleheightpercent: Integer): Integer;
var
  ScaleY: Double;
begin
  ScaleY := scaleheightpercent / 100.0;
  Result := Round(desiredheight * ScaleY / TwipsPerInch * dpiY);
end;

{ TRtfGraphicsConvertSettings }

constructor TRtfGraphicsConvertSettings.Create(AGraphics: TRtfGraphics);
begin
  if AGraphics = nil then
    raise EArgumentNilException.Create(sNilGraphics);
  Graphics := AGraphics;
  ScaleImage := true;
  ScaleOffset := 0;
  ScaleExtension := 0;
  BackgroundColor := nil;
end;

function TRtfGraphicsConvertSettings.GetImageFileName(Index: Integer;
  ARtfVisualImageFormat: TRtfImageFormat): string;
var
  ImageFileName: string;
begin
  ImageFileName := Graphics.ResolveFileName(Index, ARtfVisualImageFormat);
  if ImagesPath <> '' then
    ImageFileName := IncludeTrailingPathDelimiter(ImagesPath) + ImageFileName;
  Result := ImageFileName;
end;

{ TRtfGraphicsConvertInfo }

constructor TRtfGraphicsConvertInfo.Create(const AFileName: string;
  AFormat: TRtfImageFormat; ASize: TSize);
begin
  if AFileName.IsEmpty then
    raise EArgumentException.Create(sEmptyFileName);
  fFileName := AFileName;
  fFormat := AFormat;
  fSize := ASize;
end;

function TRtfGraphicsConvertInfo.ToString: string;
begin
  Result := FileName + ' ' + GetEnumName(TypeInfo(TRtfImageFormat), Ord(Format)) +
    ' ' + IntToStr(Size.cx) + 'x' + IntToStr(Size.cy);
end;

{ TRtfImageConverter }

constructor TRtfGraphicsConverter.Create(const ASettings: TRtfGraphicsConvertSettings);
begin
  fSettings := ASettings;
  fConvertedImages := TRtfGraphicsConvertInfoCollection.Create;
end;

destructor TRtfGraphicsConverter.Destroy;
begin
  fConvertedImages.Free;
  inherited;
end;

procedure TRtfGraphicsConverter.DoBeginDocument(AContext: TRtfInterpreterContext);
begin
  inherited;
  fConvertedImages.Clear;
end;

procedure TRtfGraphicsConverter.DoInsertImage(AContext: TRtfInterpreterContext;
  AFormat: TRtfImageFormat; AWidth, AHeight, ADesiredWidth, ADesiredHeight,
  AScaleWidthPercent, AScaleHeightPercent: Integer; AImageDataHex: string);
var
  ImageIndex: Integer;
  FileName: string;
  ImageBuffer: TBytes;
  ImageSize: TSize;
begin
  ImageIndex := fConvertedImages.Count + 1;
  FileName := fSettings.GetImageFileName(ImageIndex, AFormat);
  EnsureImagesPath(FileName);
  ImageBuffer := TextToBinary(AImageDataHex);

  if fSettings.ScaleImage then
  begin
    ImageSize := TSize.Create(
      fSettings.Graphics.CalcImageWidth(AFormat, AWidth, ADesiredWidth, AScaleWidthPercent),
      fSettings.Graphics.CalcImageHeight(AFormat, AHeight, ADesiredHeight, AScaleHeightPercent));
  end
  else
  begin
    ImageSize := TSize.Create(AWidth, AHeight);
  end;
  SaveImage(ImageBuffer, AFormat, FileName, ImageSize);
  fConvertedImages.Add(TRtfGraphicsConvertInfo.Create(FileName, AFormat, ImageSize));
end;

procedure TRtfGraphicsConverter.EnsureImagesPath(AImageFileName: string);
var
  DirectoryName: string;
begin
  DirectoryName := ExtractFileDir(AImageFileName);
  if (DirectoryName <> '') and not DirectoryExists(DirectoryName) then
    CreateDir(DirectoryName);
end;

function TRtfGraphicsConverter.GetConvertedImages: TRtfGraphicsConvertInfoCollection;
begin
  Result := fConvertedImages;
end;

function TRtfGraphicsConverter.GetSettings: TRtfGraphicsConvertSettings;
begin
  Result := fSettings;
end;

procedure TRtfGraphicsConverter.SaveImage(Stream: TStream;
  AFileName: string; ASize: TSize);
const
  JpegQuality: Cardinal = 80;
var
  ScaleOffset, ScaleExtension: single;
  Rect: TGPRectF;
  Input: TGPImage;
  Output: TGPBitmap;
  Status: TStatus;
  IsJpeg: boolean;
  Encoder: TGUID;
  EncoderStr: string;
  EncoderParameters: TEncoderParameters;
  EncoderParametersPtr: PEncoderParameters;
  Graphics: TGPGraphics;
  StreamAdapter: IStream;
begin
  Stream.Seek(0, 0);
  StreamAdapter := TStreamAdapter.Create(Stream);
  Input := TGPImage.Create(StreamAdapter);
  try
    Output := TGPBitmap.Create(ASize.Width, ASize.Height, Input.GetPixelFormat);
    try
      Graphics := TGPGraphics.Create(Output);
      try
         // set the composition mode to copy
        Graphics.SetCompositingMode(CompositingModeSourceCopy);
        // set high quality rendering modes
        Graphics.SetInterpolationMode(InterpolationModeHighQualityBicubic);
        Graphics.SetPixelOffsetMode(PixelOffsetModeHighQuality);
        Graphics.SetSmoothingMode(SmoothingModeHighQuality);
        if fSettings.BackgroundColor.HasValue then
          Graphics.Clear(ColorRefToARGB(ColorToRGB(fSettings.BackgroundColor.Value)));
        ScaleOffset := fSettings.ScaleOffset;
        ScaleExtension := fSettings.ScaleExtension;
        Rect := MakeRect(ScaleOffset,	ScaleOffset, ASize.Width + ScaleExtension,
		      ASize.Height + ScaleExtension);
        // draw the input image on the output in modified size
        Graphics.DrawImage(Input, Rect);
      finally
        Graphics.Free;
      end;
      IsJpeg := ExtractFileExt(AFileName).ToLower = '.jpg';
      EncoderStr := IfThen(IsJpeg, 'image/jpeg', 'image/png');
      EncoderParametersPtr := nil;
      if IsJpeg then
      begin
        EncoderParameters.Count := 1;
        EncoderParameters.Parameter[0].Guid := EncoderQuality;
        EncoderParameters.Parameter[0].Type_ := EncoderParameterValueTypeLong;
        EncoderParameters.Parameter[0].NumberOfValues := 1;
        EncoderParameters.Parameter[0].Value := @JpegQuality;
        EncoderParametersPtr := @EncoderParameters;
      end;
      if GetEncoderClsid(EncoderStr, Encoder) <> -1 then
      begin
        Status := Output.Save(AFileName, Encoder, EncoderParametersPtr);
        if Status <> Ok then
          raise ERtfGraphics.CreateFmt(sErrConvPicture, [GetStatus(Status)]);
      end
      else
        raise ERtfGraphics.Create(sErrConvEncoder);
    finally
      Output.Free;
    end;
  finally
    Input.Free;
  end;
end;

procedure TRtfGraphicsConverter.SaveImage(AImageBuffer: TBytes;
  AFormat: TRtfImageFormat; AFileName: string; ASize: TSize);
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    Stream.WriteBuffer(AImageBuffer[0], Length(AImageBuffer));
    if AFormat <> rifPng then
      fSettings.BackgroundColor.Value := clWhite;
    SaveImage(Stream, AFileName, ASize);
  finally
    Stream.Free;
  end;
end;

end.
