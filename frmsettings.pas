unit frmsettings;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Buttons,
  ExtCtrls;

type

  { TFormSettings }

  TFormSettings = class(TForm)
    btnCLoseM: TButton;
    btnOutputPath: TSpeedButton;
    btnSave: TButton;
    btnTemplate: TSpeedButton;
    btnTwisterDB: TSpeedButton;
    editDBPath: TEdit;
    editDBPath1: TEdit;
    editDBPath2: TEdit;
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    procedure btnCLoseMClick(Sender: TObject);
  private

  public

  end;

var
  FormSettings: TFormSettings;

implementation

{$R *.lfm}

{ TFormSettings }

procedure TFormSettings.btnCLoseMClick(Sender: TObject);
begin

end;

end.

