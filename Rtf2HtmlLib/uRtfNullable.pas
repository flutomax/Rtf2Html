unit uRtfNullable;

interface

uses
  System.SysUtils, System.Classes;

type

  ENullable = class(Exception);

  TNullable<T> = record
  private
    fHasValue: Boolean;
    fValue: T;
    function GetValue: T;
    procedure SetValue(AValue: T);
  public
    procedure Clear;
    property HasValue: Boolean read fHasValue;
    property Value: T read GetValue write SetValue;
    class operator Implicit(A: T): TNullable<T>;
    class operator Implicit(A: Pointer): TNullable<T>;
  end;

  TIntNullable = TNullable<Integer>;
  TDateNullable = TNullable<TDateTime>;

implementation

uses
  uRtfMessages;

{ TNullable }

function TNullable<T>.GetValue: T;
begin
  if fHasValue then
     Result := fValue
  else
    raise ENullable.Create(sNoValueVar);
end;

procedure TNullable<T>.SetValue(AValue: T);
begin
  fValue := AValue;
  fHasValue := True;
end;

procedure TNullable<T>.Clear;
begin
  fHasValue := False;
end;

class operator TNullable<T>.Implicit(A: T): TNullable<T>;
begin
  Result.Value := A;
end;

class operator TNullable<T>.Implicit(A: Pointer): TNullable<T>;
begin
  if A = nil then
    Result.Clear
  else
    raise ENullable.Create(sNoPointerAllowed);
end;

end.
