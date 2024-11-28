unit uRtfParserListener;

interface

uses
  System.Classes, System.SysUtils, System.Generics.Collections, uRtfTypes,
  uRtfElement;

type


  TRtfParserListener = class(TObject)
  protected
    fLevel: integer;
    fOnParseBegin: TNotifyEvent;
    fOnParseEnd: TNotifyEvent;
    fOnParseSuccess: TNotifyEvent;
    fOnParseFail: TGetStrProc;
    procedure DoGroupBegin; virtual;
    procedure DoGroupEnd; virtual;
    procedure DoTagFound(tag: TRtfTag); virtual;
    procedure DoTextFound(text: TRtfText); virtual;
    procedure DoParseBegin; virtual;
    procedure DoParseSuccess; virtual;
    procedure DoParseFail(E: Exception); virtual;
    procedure DoParseEnd; virtual;
  public
    procedure ParseBegin;
    procedure GroupBegin;
    procedure TagFound(tag: TRtfTag);
    procedure TextFound(text: TRtfText);
    procedure GroupEnd;
    procedure ParseSuccess;
    procedure ParseFail(E: Exception);
    procedure ParseEnd;
  published
    property Level: integer read fLevel;
    property OnParseBegin: TNotifyEvent read fOnParseBegin write fOnParseBegin;
    property OnParseEnd: TNotifyEvent read fOnParseEnd write fOnParseEnd;
    property OnParseSuccess: TNotifyEvent read fOnParseSuccess write fOnParseSuccess;
    property OnParseFail: TGetStrProc read fOnParseFail write fOnParseFail;
  end;

  TRtfParserListenerStructureBuilder = class(TRtfParserListener)
  private
    fCurGroup: TRtfGroup;
    fStructureRoot: TRtfGroup;
    fOpenGroupStack: TStack<TRtfGroup>;
  protected
    procedure DoGroupBegin; override;
    procedure DoGroupEnd; override;
    procedure DoTextFound(text: TRtfText); override;
    procedure DoTagFound(tag: TRtfTag); override;
    procedure DoParseBegin; override;
    procedure DoParseEnd; override;
  public
    constructor Create;
    destructor Destroy; override;
    property StructureRoot: TRtfGroup read fStructureRoot;
  end;

implementation

uses
  uRtfMessages;

{ TRtfParserListenerBase }

procedure TRtfParserListener.DoParseBegin;
begin
  if Assigned(fOnParseBegin) then
    fOnParseBegin(self);
end;

procedure TRtfParserListener.DoParseEnd;
begin
  if Assigned(fOnParseEnd) then
    fOnParseEnd(self);
end;

procedure TRtfParserListener.DoParseFail;
var
  s: string;
begin
  s := Format('Error while parsing rtf %s: %s.', [E.ClassName, E.Message]);
  if Assigned(fOnParseFail) then
    fOnParseFail(s);
end;

procedure TRtfParserListener.DoParseSuccess;
begin
  if Assigned(fOnParseSuccess) then
    fOnParseSuccess(self);
end;

procedure TRtfParserListener.DoGroupBegin;
begin
end;

procedure TRtfParserListener.DoTextFound(text: TRtfText);
begin
end;

procedure TRtfParserListener.DoTagFound(tag: TRtfTag);
begin
end;

procedure TRtfParserListener.DoGroupEnd;
begin
end;

procedure TRtfParserListener.ParseBegin;
begin
  fLevel := 0;
  DoParseBegin;
end;

procedure TRtfParserListener.GroupBegin;
begin
  DoGroupBegin;
  inc(fLevel);
end;

procedure TRtfParserListener.TagFound(tag: TRtfTag);
begin
  if Assigned(tag) then
    DoTagFound(tag);
end;

procedure TRtfParserListener.TextFound(text: TRtfText);
begin
  if Assigned(text) then
    DoTextFound(text);
end;

procedure TRtfParserListener.GroupEnd;
begin
  dec(fLevel);
  DoGroupEnd;
end;

procedure TRtfParserListener.ParseEnd;
begin
  DoParseEnd;
end;

procedure TRtfParserListener.ParseFail(E: Exception);
begin
  DoParseFail(E);
end;

procedure TRtfParserListener.ParseSuccess;
begin
  DoParseSuccess;
end;

{ TRtfParserListenerStructureBuilder }

constructor TRtfParserListenerStructureBuilder.Create;
begin
  inherited Create;
  fStructureRoot := nil;
  fCurGroup := nil;
  fOpenGroupStack := TStack<TRtfGroup>.Create;
end;

destructor TRtfParserListenerStructureBuilder.Destroy;
begin
  fOpenGroupStack.Free;
  inherited;
end;

procedure TRtfParserListenerStructureBuilder.DoGroupBegin;
var
  newgroup: TRtfGroup;
begin
  newgroup := TRtfGroup.Create;
  if Assigned(fCurGroup) then
  begin
    fOpenGroupStack.Push(fCurGroup);
    fCurGroup.Contents.Add(newgroup);
  end;
  fCurGroup := newGroup;
end;

procedure TRtfParserListenerStructureBuilder.DoGroupEnd;
begin
  if fOpenGroupStack.Count > 0 then
    fCurGroup := fOpenGroupStack.Pop
  else
  begin
    if Assigned(fStructureRoot) then
      raise ERtfStructure.Create(sMultipleRootLevelGroups);
    fStructureRoot := fCurGroup;
    fCurGroup := nil;
  end;
end;

procedure TRtfParserListenerStructureBuilder.DoParseBegin;
begin
  fOpenGroupStack.Clear;
  fStructureRoot := nil;
  fCurGroup := nil;
  inherited DoParseBegin;
end;

procedure TRtfParserListenerStructureBuilder.DoParseEnd;
begin
  if fOpenGroupStack.Count > 0 then
    raise ERtfBraceNesting.Create(sUnclosedGroups);
  inherited DoParseEnd;
end;

procedure TRtfParserListenerStructureBuilder.DoTagFound(tag: TRtfTag);
begin
  if fCurGroup = nil then
    raise ERtfStructure.Create(sMissingGroupForNewTag);
  fCurGroup.Contents.Add(tag);
end;

procedure TRtfParserListenerStructureBuilder.DoTextFound(text: TRtfText);
begin
  if fCurGroup = nil then
    raise ERtfStructure.Create(sMissingGroupForNewText);
  fCurGroup.Contents.Add(text);
end;

end.
