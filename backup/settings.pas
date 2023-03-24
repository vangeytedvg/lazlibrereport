{
 Settings manager for this application,
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
  private
    FDbPath: string;
    FTemplatePath: string;
    FNewDocStorage: string;
    FIniFile: string;
  public
    property dbPath: string read FDbPath write FDbPath;
    property TemplatePath: string read FTemplatePath write FTemplatePath;
    property NewDocStorage: string read FNewDocStorage write FNewDocStorage;
    property IniFile: string read FIniFile write FIniFile;
    constructor Create(AIniFile: string); overload;
    procedure WriteSettings;
    procedure ReadSettings;
  end;


implementation

constructor TAppSettings.Create(AIniFile: string);
begin
  IniFile := AIniFile;
end;

procedure TAppSettings.WriteSettings;
{ Write settins }
var
  myIni: TIniFile;
begin
  myIni := TIniFile.Create('Twister.ini');
  try
    myIni.WriteString('Settings', 'DataBasePath', dbPath);
    myIni.WriteString('Settings', 'TemplatePath', TemplatePath);
    myIni.WriteString('Settings', 'DocPath', NewDocStorage);
  finally
    myIni.Free;
  end;
end;

procedure TAppSettings.ReadSettings;
{ Read settings }
var
  myIni: TIniFile;
begin
  myIni := TIniFile.Create(IniFile);
  try
    dbPath := myIni.ReadString('Settings', 'DataBasePath', '');
    TemplatePath := myIni.ReadString('Settings', 'TemplatePath', '');
    NewDocStorage := myIni.ReadString('Settings', 'DocPath', '');
  finally
    myIni.Free;
  end;
end;

end.
