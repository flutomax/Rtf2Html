program Rtf2HtmlTest;

uses
  Vcl.Forms,
  uFrmTest in 'uFrmTest.pas' {FrmTest},
  uRtf2Html in '..\Rtf2HtmlLib\uRtf2Html.pas',
  uRtfTypes in '..\Rtf2HtmlLib\uRtfTypes.pas',
  uRtfParser in '..\Rtf2HtmlLib\uRtfParser.pas',
  uRtfSpec in '..\Rtf2HtmlLib\uRtfSpec.pas',
  uRtfParserListener in '..\Rtf2HtmlLib\uRtfParserListener.pas',
  uRtfMessages in '..\Rtf2HtmlLib\uRtfMessages.pas',
  uRtfElement in '..\Rtf2HtmlLib\uRtfElement.pas',
  uRtfHash in '..\Rtf2HtmlLib\uRtfHash.pas',
  uRtfObjects in '..\Rtf2HtmlLib\uRtfObjects.pas',
  uRtfNullable in '..\Rtf2HtmlLib\uRtfNullable.pas',
  uRtfDocument in '..\Rtf2HtmlLib\uRtfDocument.pas',
  uRtfVisual in '..\Rtf2HtmlLib\uRtfVisual.pas',
  uRtfDocumentInfo in '..\Rtf2HtmlLib\uRtfDocumentInfo.pas',
  uRtfBuilders in '..\Rtf2HtmlLib\uRtfBuilders.pas',
  uRtfGraphics in '..\Rtf2HtmlLib\uRtfGraphics.pas',
  uRtfInterpreterContext in '..\Rtf2HtmlLib\uRtfInterpreterContext.pas',
  uRtfInterpreterListener in '..\Rtf2HtmlLib\uRtfInterpreterListener.pas',
  uRtfInterpreter in '..\Rtf2HtmlLib\uRtfInterpreter.pas',
  uRtfHtmlObjects in '..\Rtf2HtmlLib\uRtfHtmlObjects.pas',
  uRtfHtmlConverter in '..\Rtf2HtmlLib\uRtfHtmlConverter.pas',
  uRtfHtmlWriter in '..\Rtf2HtmlLib\uRtfHtmlWriter.pas',
  uRtfCssTextWriter in '..\Rtf2HtmlLib\uRtfCssTextWriter.pas',
  uRtfHtmlFunctions in '..\Rtf2HtmlLib\uRtfHtmlFunctions.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmTest, FrmTest);
  Application.Run;
end.
