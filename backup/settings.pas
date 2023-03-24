{
 Class that holds the Settings for this application
 Author : Danny Van Geyte
 Created: 24/03/2023
}
unit Settings;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, IniFiles;

type
  TAppSettings = class
    FDbPath: string;
    FTemplatePath: string;
    FNewDocStorage: string;
  public
    property dbPath: string read FDbPath write FDbPath;
    property TemplatePath: string read FTemplatePath write FTemplatePath;
    property NewDocStorage: string read FNewDocStorage write FNewDocStorage;
    procedure ReadSettings;
    procedure WriteSettings;
  end;


implementation

procedure TAppSettings.ReadSettings;
var
  ini: TIniFile;
begin
  a := 0;
end;

procedure TAppSettings.WriteSettings;
var
  ini: TIniFile;
begin
end;

end.
