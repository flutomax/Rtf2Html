unit uFrmTest;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  System.ImageList, Vcl.ImgList, System.Actions, Vcl.ActnList, Vcl.ComCtrls;

type

  TRichEdit = class(Vcl.ComCtrls.TRichEdit)
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  end;

  TFrmTest = class(TForm)
    LbLog: TListBox;
    ActionList1: TActionList;
    ImageList1: TImageList;
    DlgOpenRTF: TOpenDialog;
    DlgSaveHtml: TSaveDialog;
    Button2: TButton;
    CmdConvert: TAction;
    PnlToolbar: TPanel;
    CmdFileOpen: TAction;
    Button3: TButton;
    Splitter1: TSplitter;
    Editor: TRichEdit;
    procedure FormCreate(Sender: TObject);
    procedure CmdConvertExecute(Sender: TObject);
    procedure CmdConvertUpdate(Sender: TObject);
    procedure CmdFileOpenExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    procedure Add2Log(Sender: TObject; const Text: string);
  public
    { Public declarations }
  end;

var
  FrmTest: TFrmTest;

implementation

{$R *.dfm}

uses
  uRtf2Html;

var
  RichEditVersion: integer;
  GLibHandle: THandle = 0;

const
  RichEdit10ModuleName = 'RICHED32.DLL';
  RichEdit20ModuleName = 'RICHED20.DLL';
  RichEdit40ModuleName = 'MSFTEDIT.DLL';

  {$IFDEF UNICODE}
  MSFTEDIT_CLASS = 'RICHEDIT50W';
  RICHEDIT_CLASS = 'RICHEDIT20W';
  {$ELSE}
  MSFTEDIT_CLASS = 'RICHEDIT50A';
  RICHEDIT_CLASS = 'RICHEDIT20A';
  {$ENDIF SUPPORTS_UNICODE}
  RICHEDIT_CLASS10A = 'RICHEDIT';

procedure InitRichEditDll;
begin
  RichEditVersion := 1;

  // RichEdit 4.1 (XP SP1), 6 (Office 2007), 7.5 (Win8), 8.5 (Win10/Win11)
  GLibHandle := SafeLoadLibrary(RichEdit40ModuleName);
  if GLibHandle <> 0 then
    RichEditVersion := 4; // at least version 4

  if GLibHandle = 0 then
  begin
    // RichEdit 2.0 (Win98), 3.0 (Win2000), 3.1 (Win2003)
    GLibHandle := SafeLoadLibrary(RichEdit20ModuleName);
    if GLibHandle <> 0 then
      RichEditVersion := 2; // at least version 2
  end;

  if GLibHandle = 0 then
  begin
    // RichEdit 1.0 (Win95)
    RichEditVersion := 1; // fall back to version 1
    GLibHandle := SafeLoadLibrary(RichEdit10ModuleName);
  end;
end;

procedure FinalRichEditDll;
begin
  if GLibHandle > 0 then
  begin
    FreeLibrary(GLibHandle);
    GLibHandle := 0;
  end;
end;

{ TRichEdit }

procedure TRichEdit.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  if RichEditVersion >= 4 then
    CreateSubClass(Params, MSFTEDIT_CLASS)
  else
  begin
    case RichEditVersion of
      2: CreateSubClass(Params, RICHEDIT_CLASS);
      1: CreateSubClass(Params, RICHEDIT_CLASS10A);
    else
      CreateSubClass(Params, RICHEDIT_CLASS);
    end;
  end;
end;

{ TFrmTest }

procedure TFrmTest.FormCreate(Sender: TObject);
begin
  ReportMemoryLeaksOnShutdown := true;
end;

procedure TFrmTest.FormShow(Sender: TObject);
begin
  LbLog.Items.Add('Wellcome to ' + Caption);
  LbLog.Items.Add('RichEdit Version: ' + RichEditVersion.ToString);
end;

procedure TFrmTest.Add2Log(Sender: TObject; const Text: string);
begin
  LbLog.ItemIndex := LbLog.Items.Add(Text);
end;

procedure TFrmTest.CmdConvertUpdate(Sender: TObject);
begin
  CmdConvert.Enabled := Editor.Lines.Count > 0;
end;

procedure TFrmTest.CmdFileOpenExecute(Sender: TObject);
begin
  if not DlgOpenRTF.Execute then
    exit;
  Editor.Lines.LoadFromFile(DlgOpenRTF.FileName);
  LbLog.Clear;
end;

procedure TFrmTest.CmdConvertExecute(Sender: TObject);
var
  Rtf2Html: TRtf2Html;
  Stream: TMemoryStream;
begin
  DlgSaveHtml.FileName := ChangeFileExt(ExtractFileName(DlgOpenRTF.FileName), '.html');
  if not DlgSaveHtml.Execute then
    exit;
  LbLog.Clear;
  Rtf2Html := TRtf2Html.Create(nil);
  Stream := TMemoryStream.Create;
  try
    Editor.Lines.SaveToStream(Stream);
    Stream.Position := 0;
    Rtf2Html.OnLogMessage := Add2Log;
    Rtf2Html.OutputFileName := DlgSaveHtml.FileName;
    Rtf2Html.LoadFromStream(Stream);
    Rtf2Html.Convert;
  finally
    Rtf2Html.Free;
    Stream.Free;
  end;
end;


initialization
  InitRichEditDll;

finalization
  FinalRichEditDll;

end.
