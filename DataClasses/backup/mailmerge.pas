{
 Class that holds the information for a report to be generated
 Author : Danny Van Geyte
 Created: 23/03/2023
}
unit MailMerge;

{$mode ObjFPC}{$H+}


interface

uses
  Classes, SysUtils;

type
  TMailMerge = class
  private
    FName: string;
    FFirstName: string;
    FFullName: string;
    FAddress: string;
    FCity: string;
    FZipCode: string;
    FPhone: string;
    FEmail: string;
    FSocialSecurity: string;
    FSubject: string;
    FSalutation: string;
    FSignature: string;
    function GetFullName:string;
  public
    property Name: string read FName write FName;
    property FirstName: string read FFirstName write FFirstName;
    property FullName: string read GetFullName;
    property Address: string read FAddress write FAddress;
    property City: string read FCity write FCity;
    property ZipCode: string read FZipCode write FZipCode;
    property Phone: string read FPhone write FPhone;
    property EMail: string read FEmail write FEmail;
    property SocialSecurity: string read FSocialSecurity write FSocialSecurity;
    property Subject: string read FSubject write FSubject;
    property Salutation: string read FSalutation write FSalutation;
    property Signature: string read FSignature write FSignature;
    constructor Create(); overload;
    constructor Create(const AName, AFirstName, AAddress, ACity,
      AZipCode, APhone, AEmail, ASocialSecurity, ASubject, ASalutation,
      ASignature: string);
      overload;

  end;



implementation

constructor TMailMerge.Create;
begin
  // Empty contructor
end;

{ Class constructor }
constructor TMailMerge.Create(
  const AName, AFirstName, AAddress, ACity, AZipCode, APhone,
  AEmail, ASocialSecurity, ASubject, ASalutation, ASignature: string);
begin
  FName := AName;
  FFirstName := AFirstName;
  FAddress := AAddress;
  FCity := ACity;
  FZipCode := AZipCode;
  FPhone := APhone;
  FEmail := AEmail;
  FSocialSecurity := ASocialSecurity;
  FSubject := ASubject;
  FSalutation := ASalutation;
  FSignature := ASignature;
end;

function TMailMerge.GetFullName: string;
begin
  Result := FName + ' ' + FFirstName;
end;

end.
