unit uRtfDocumentInfo;

interface

uses
  System.SysUtils, System.DateUtils, uRtfNullable, uRtfTypes;

type

  TRtfDocumentInfo = record
    Id: TIntNullable;
    Version: TIntNullable;
    Revision: TIntNullable;
    Title: string;
    Subject: string;
    Author: string;
    Manager: string;
    Company: string;
    &Operator: string;
    Category: string;
    Keywords: string;
    Comment: string;
    DocumentComment: string;
    HyperLinkbase: string;
    CreationTime: TDateNullable;
    RevisionTime: TDateNullable;
    PrintTime: TDateNullable;
    BackupTime: TDateNullable;
    NumberOfPages: TIntNullable;
    NumberOfWords: TIntNullable;
    NumberOfCharacters: TIntNullable;
    EditingTimeInMinutes: TIntNullable;
    class function ToString: string; inline; static;
    procedure Reset;
  end;

implementation

{ TRtfDocumentInfo }

procedure TRtfDocumentInfo.Reset;
begin
  Id := nil;
  Version := nil;
  Revision := nil;
  Title := '';
  Subject := '';
  Author := '';
  Manager := '';
  Company := '';
  &Operator := '';
  Category := '';
  Keywords := '';
  Comment := '';
  DocumentComment := '';
  HyperLinkbase := '';
  CreationTime := nil;
  RevisionTime := nil;
  PrintTime := nil;
  BackupTime := nil;
  NumberOfPages := nil;
  NumberOfWords := nil;
  NumberOfCharacters := nil;
  EditingTimeInMinutes := nil;
end;

class function TRtfDocumentInfo.ToString: string;
begin
  result := 'RTFDocInfo';
end;

end.
