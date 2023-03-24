{
   frmSettings
   A simple settings form for the Letter Wizard
   Author : Danny Van Geyte
   Created : 24/03/2023
}
unit frmsettings;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Buttons,
  ExtCtrls, IniFiles, LCLIntf, SpinEx;

type

  { TFormSettings }

  TFormSettings = class(TForm)
    btnCLoseM: TButton;
    btnOutputPath: TSpeedButton;
    btnSave: TButton;
    btnTemplate: TSpeedButton;
    btnTwisterDB: TSpeedButton;
    editDBPath: TEdit;
    editTemplatePath: TEdit;
    editNewDocPath: TEdit;
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    SelFile: TOpenDialog;
    SelDir: TSelectDirectoryDialog;
    btnEditTemplate: TSpeedButton;
    spinStartLineNr: TSpinEditEx;
    procedure btnCLoseMClick(Sender: TObject);
    procedure btnEditTemplateClick(Sender: TObject);
    procedure btnOutputPathClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnTemplateClick(Sender: TObject);
    procedure btnTwisterDBClick(Sender: TObject);
    procedure PathChanged(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    IsDirty: boolean;
    procedure WriteSettings;
    procedure ReadSettings;

  public

  end;

var
  FormSettings: TFormSettings;

implementation

uses
  Settings;

{$R *.lfm}

{ TFormSettings }

procedure TFormSettings.btnCLoseMClick(Sender: TObject);
begin
  if IsDirty then
    if MessageDlg('Instellingen', 'Wilt U de wijzigingen opslaan?',
      mtConfirmation, [mbYes, mbNo, mbIgnore], 0) = mrYes then
      WriteSettings;
  Close;
end;

procedure TFormSettings.btnEditTemplateClick(Sender: TObject);
begin
  OpenDocument(editTemplatePath.Text);
end;

procedure TFormSettings.btnOutputPathClick(Sender: TObject);
begin
  if SelDir.Execute then
  begin
    editNewDocPath.Text := Seldir.FileName;
  end;
end;

procedure TFormSettings.btnSaveClick(Sender: TObject);
begin
  WriteSettings;
  close;
end;

procedure TFormSettings.btnTemplateClick(Sender: TObject);
{ Select the template file }
begin
  SelFile.Filter := 'LibreOffice file (*.odt)|*.odt';
  SelFile.DefaultExt := '*.odt';
  if Selfile.Execute then
  begin
    editTemplatePath.Text := SelFile.FileName;
  end;
end;

procedure TFormSettings.btnTwisterDBClick(Sender: TObject);
{ Select the database file }
begin
  SelFile.Filter := 'Sql DB file (*.db)|*.db';
  SelFile.DefaultExt := '*.db';
  if Selfile.Execute then
  begin
    editDBPath.Text := SelFile.FileName;
  end;
end;

procedure TFormSettings.PathChanged(Sender: TObject);
begin
  IsDirty := True;
end;

procedure TFormSettings.FormCreate(Sender: TObject);
begin
  ReadSettings;
  IsDirty := False;
end;

{ Read Write Settings }
procedure TFormSettings.WriteSettings;
var
   AppSettings : TAppSettings;
begin
  AppSettings := TAppSettings.Create('Twister.ini');
  try
    AppSettings.dbPath := editDBPath.Text;
    AppSettings.TemplatePath := editTemplatePath.Text;
    AppSettings.NewDocStorage := editNewDocPath.Text;
    AppSettings.StartLine:= spinStartLineNr.Value;
    AppSettings.WriteSettings;
  finally
    AppSettings.Free;
  end;
end;

procedure TFormSettings.ReadSettings;
var
   AppSettings : TAppSettings;
begin
  AppSettings := TAppSettings.Create('Twister.ini');
  try
    AppSettings.ReadSettings;
    editDBPath.Text := AppSettings.dbPath;
    editTemplatePath.Text := AppSettings.TemplatePath;
    editNewDocPath.Text := AppSettings.NewDocStorage;
    spinStartLineNr.Value:= AppSettings.StartLine;
  finally
    AppSettings.Free;
  end;
end;

end.
