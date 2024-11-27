unit uRtfHash;

interface

uses
  System.SysUtils, System.Generics.Collections;

function AddHashCode(hash: Integer; obj: TObject): Integer; overload;
function AddHashCode(hash: Integer; objHash: Integer): Integer; overload;
function AddHashCode(hash: Integer; objHash: string): Integer; overload;
function ComputeHashCode(enumerable: TEnumerable<TObject>): Integer;

implementation

uses
  System.Math;

function AddHashCode(hash: Integer; obj: TObject): Integer;
var
  combinedHash: Integer;
begin
  if Assigned(obj) then
    combinedHash := obj.GetHashCode
  else
    combinedHash := 0;
  if hash <> 0 then
    combinedHash := combinedHash + hash * 31;
  Result := combinedHash;
end;

function AddHashCode(hash: Integer; objHash: Integer): Integer;
var
  combinedHash: Integer;
begin
  combinedHash := objHash;
  if hash <> 0 then
    combinedHash := combinedHash + hash * 31;
  Result := combinedHash;
end;

function AddHashCode(hash: Integer; objHash: string): Integer;
var
  combinedHash: Integer;
begin
  combinedHash := objHash.GetHashCode;
  if hash <> 0 then
    combinedHash := combinedHash + hash * 31;
  Result := combinedHash;
end;

function ComputeHashCode(enumerable: TEnumerable<TObject>): Integer;
var
  hash: Integer;
  item: TObject;
begin
  hash := 1;
  if enumerable = nil then
    raise EArgumentNilException.Create('enumerable cannot be nil');
  for item in enumerable do
  begin
    hash := hash * 31 + IfThen(Assigned(item), item.GetHashCode);
  end;
  Result := hash;
end;

end.
