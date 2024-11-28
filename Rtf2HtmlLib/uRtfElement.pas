unit uRtfElement;

interface

uses
  System.SysUtils, System.Generics.Collections, uRtfTypes;

type

  TRtfTag = class;
  TRtfGroup = class;
  TRtfText = class;

  TRtfElementKind = (ekTag, ekGroup, ekText);
  TRtfVisitorOrder = (rvNonRecursive, rvDepthFirst, rvBreadthFirst);

  TRtfElementVisitor = class(TObject)
  private
    fOrder: TRtfVisitorOrder;
  protected
    procedure DoVisitTag(tag: TRtfTag); virtual;
    procedure DoVisitGroup(group: TRtfGroup); virtual;
    procedure DoVisitText(text: TRtfText); virtual;
    procedure VisitGroupChildren(group: TRtfGroup); virtual;
  public
    constructor Create(aOrder: TRtfVisitorOrder); virtual;
    procedure VisitTag(tag: TRtfTag); virtual;
    procedure VisitGroup(group: TRtfGroup); virtual;
    procedure VisitText(text: TRtfText); virtual;
  end;

  TRtfElement = class(TObject)
  private
    fKind: TRtfElementKind;
    function GetKind: TRtfElementKind;
  protected
    function IsEqual(obj: TObject): boolean; virtual;
    function ComputeHashCode: integer; virtual;
    procedure DoVisit(visitor: TRtfElementVisitor); virtual; abstract;
  public
    constructor Create(aKind: TRtfElementKind); virtual;
    function Equals(obj: TObject): boolean;
    function ToString: string; virtual; abstract;
    procedure Visit(visitor: TRtfElementVisitor);
  published
    property Kind: TRtfElementKind read GetKind;
  end;

  TRtfElementCollection = class(TBaseCollection<TRtfElement>);

  TRtfGroup = class(TRtfElement)
  private
    fContents: TRtfElementCollection;
    function GetDestination: string;
    function GetIsExtensionDestination: Boolean;
  protected
    procedure DoVisit(visitor: TRtfElementVisitor); override;
    function IsEqual(obj: TObject): boolean; override;
  public
    constructor Create;
    destructor Destroy; override;
    function ToString: string; override;
    function SelectChildGroupWithDestination(const destination: string): TRtfGroup;
  published
    property Contents: TRtfElementCollection read fContents;
    property Destination: string read GetDestination;
    property IsExtensionDestination: Boolean read GetIsExtensionDestination;
  end;

  TRtfText = class(TRtfElement)
  private
    fText: string;
  protected
    procedure DoVisit(visitor: TRtfElementVisitor); override;
    function IsEqual(obj: TObject): boolean; override;
  public
    constructor Create(const aText: string);
    function ToString: string; override;
  published
    property Text: string read fText;
  end;

  TRtfTag = class(TRtfElement)
  private
    fName: string;
    fFullName: string;
    fValueAsText: string;
    fValueAsNumber: Integer;
    function GetHasValue: Boolean;
  protected
    procedure DoVisit(visitor: TRtfElementVisitor); override;
    function IsEqual(obj: TObject): boolean; override;
  public
    constructor Create(const aName: string); overload;
    constructor Create(const aName, aValue: string); overload;
    destructor Destroy;override;
    function ToString: string; override;
    function ComputeHashCode: integer; override;
  published
    property FullName: string read fFullName;
    property Name: string read fName;
    property ValueAsText: string read fValueAsText;
    property ValueAsNumber: Integer read fValueAsNumber;
    property HasValue: Boolean read GetHasValue;
  end;

implementation

uses
  System.RTLConsts, uRtfSpec, uRtfMessages, uRtfHash;


{ TRtfElementVisitorBase }

constructor TRtfElementVisitor.Create(aOrder: TRtfVisitorOrder);
begin
  fOrder := aOrder;
end;

procedure TRtfElementVisitor.VisitTag(tag: TRtfTag);
begin
  if Assigned(tag) then
    DoVisitTag(tag);
end;

procedure TRtfElementVisitor.DoVisitGroup(group: TRtfGroup);
begin

end;

procedure TRtfElementVisitor.DoVisitTag(tag: TRtfTag);
begin

end;

procedure TRtfElementVisitor.DoVisitText(text: TRtfText);
begin

end;

procedure TRtfElementVisitor.VisitGroup(group: TRtfGroup);
begin
  if Assigned(group) then
  begin
    if fOrder = rvDepthFirst then
      VisitGroupChildren(group);
    DoVisitGroup(group);
    if fOrder = rvBreadthFirst then
      VisitGroupChildren(group);
  end;
end;

procedure TRtfElementVisitor.VisitText(text: TRtfText);
begin
  if Assigned(text) then
    DoVisitText(text);
end;

procedure TRtfElementVisitor.VisitGroupChildren(group: TRtfGroup);
var
  child: TRtfElement;
  i: Integer;
begin
  for i := 0 to group.Contents.Count - 1 do
  begin
    child := group.Contents[i];
    child.Visit(Self);
  end;
end;


{ TRtfElement }

constructor TRtfElement.Create(aKind: TRtfElementKind);
begin
  inherited Create;
  fKind := aKind;
end;

function TRtfElement.GetKind: TRtfElementKind;
begin
  result := fKind;
end;

function TRtfElement.Equals(obj: TObject): boolean;
begin
  if obj = self then
    exit(true);
  if (obj = nil) or (ClassType <> obj.ClassType) then
    exit(false);
  result := IsEqual(obj);
end;

function TRtfElement.IsEqual(obj: TObject): boolean;
begin
  result := true;
end;

procedure TRtfElement.Visit(visitor: TRtfElementVisitor);
begin
  if visitor = nil then
    raise EArgumentNilException.Create(sNilVisitor);
  DoVisit(visitor);
end;

function TRtfElement.ComputeHashCode: integer;
begin
  result := $0f00dead;
end;

{ TRtfGroup }

constructor TRtfGroup.Create;
begin
  inherited Create(ekGroup);
  fContents := TRtfElementCollection.Create;
end;

destructor TRtfGroup.Destroy;
begin
  fContents.Free;
  inherited;
end;

procedure TRtfGroup.DoVisit(visitor: TRtfElementVisitor);
begin
  visitor.VisitGroup(self);
end;

function TRtfGroup.IsEqual(obj: TObject): boolean;
var
  compare: TRtfGroup;
begin
  if not (obj is TRtfGroup) then
    exit(false);
  compare := obj as TRtfGroup; // guaranteed to be non-null
  result := Assigned(compare) and inherited IsEqual(obj) and
    contents.Equals(compare.contents);
end;

function TRtfGroup.GetDestination: string;
var
  first_el, second_el: TRtfElement;
  first_tag, second_tag: TRtfTag;
begin
  result := '';
  if fContents.Count > 0 then
  begin
    first_el := fContents[0];
    if first_el.Kind = ekTag then
    begin
      first_tag := first_el as TRtfTag;
      if TagExtensionDestination = first_tag.Name then
        if fContents.Count > 1 then
        begin
          second_el := fContents[1];
          if second_el.Kind = ekTag then
          begin
            second_tag := second_el as TRtfTag;
            exit(second_tag.Name);
          end;
        end;
      result := first_tag.Name;
    end;
  end;
end;

function TRtfGroup.GetIsExtensionDestination: Boolean;
var
  first_el: TRtfElement;
  first_tag: TRtfTag;
begin
  result := false;
  if fContents.Count > 0 then
  begin
    first_el := fContents[0];
    if first_el.Kind = ekTag then
    begin
      first_tag := first_el as TRtfTag;
      result := TagExtensionDestination = first_tag.Name;
    end;
  end;
end;

function TRtfGroup.SelectChildGroupWithDestination(
  const destination: string): TRtfGroup;
var
  child: TRtfElement;
  group: TRtfGroup;
  i: integer;
begin
  result := nil;
  if destination.IsEmpty then
    raise EArgumentException.Create(sEmptyDestination);

  for i := 0 to fContents.Count - 1 do
  begin
    child := fContents[i];
    if child.Kind = ekGroup then
    begin
      group := child as TRtfGroup;
      if destination = group.Destination then
        Exit(group);
    end;
  end;
end;

function TRtfGroup.ToString: string;
var
  i: integer;
begin
  result := Format('{%d', [fContents.Count]);
  if fContents.Count > 0 then
  begin
    result := Format('%s: [%s', [result, fContents[0].ToString]);
    if fContents.Count > 1 then
    begin
      result := Format('%s, %s', [result, fContents[1].ToString]);
      if fContents.Count > 2 then
      begin

        if fContents.Count > 3 then
          result := Format('%s..., ', [result]);
        result := Format('%s %s', [result, fContents[fContents.Count - 1].ToString]);
      end;
    end;
    result := Format('%s]', [result]);
  end;
  result := Format('%s}', [result]);
end;

{ TRtfText }

constructor TRtfText.Create(const aText: string);
begin
  inherited Create(ekText);
  if aText.IsEmpty then
    raise EArgumentException.Create(sEmptyText);
  fText := aText;
end;

procedure TRtfText.DoVisit(visitor: TRtfElementVisitor);
begin
  visitor.VisitText(self);
end;

function TRtfText.ToString: string;
begin
  result := fText;
end;

function TRtfText.IsEqual(obj: TObject): boolean;
var
  compare: TRtfText;
begin
  result := false;
  if not (obj is TRtfText) then
    exit;
  compare := obj as TRtfText;
  result := Assigned(compare) and (inherited IsEqual(obj)) and
    fText.Equals(compare.text);
end;


{ TRtfTag }

constructor TRtfTag.Create(const aName: string);
begin
  inherited Create(ekTag);
  if aName.IsEmpty then
    raise EArgumentException.Create(sEmptyName);
  fName := aName;
  fFullName := aName;
  fValueAsText := '';
  fValueAsNumber := -1;
end;

constructor TRtfTag.Create(const aName, aValue: string);
begin
  Create(aName);
  if aValue.IsEmpty then
    raise EArgumentException.Create(sEmptyValue);
  fFullName := aName + aValue;
  fValueAsText := aValue;
  fValueAsNumber := StrToIntDef(aValue, -1);
end;

destructor TRtfTag.Destroy;
begin

  inherited;
end;
function TRtfTag.GetHasValue: Boolean;
begin
  result := not fValueAsText.IsEmpty;
end;

function TRtfTag.ToString: string;
begin
  result := '\' + fFullName;
end;

procedure TRtfTag.DoVisit(visitor: TRtfElementVisitor);
begin
  visitor.VisitTag(self);
end;

function TRtfTag.IsEqual(obj: TObject): boolean;
var
  compare: TRtfTag;
begin
  result := false;
  if not (obj is TRtfTag) then
    exit;
  compare := obj as TRtfTag; // guaranteed to be non-null
  result := Assigned(compare) and inherited IsEqual(obj) and
    fFullName.Equals(compare.fFullName);
end;

function TRtfTag.ComputeHashCode: integer;
begin
  result := inherited ComputeHashCode;
  result := AddHashCode(result, fFullName);
end;


end.
