unit frmSettings;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, StdCtrls,
  Buttons;

type

  { TForm1 }

  TFormSettings = class(TForm)
    btnSave: TButton;
    btnClose: TButton;
    editDBPath: TEdit;
    editDBPath1: TEdit;
    editDBPath2: TEdit;
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnTwisterDB: TSpeedButton;
    btnTemplate: TSpeedButton;
    btnOutputPath: TSpeedButton;
    procedure btnCloseClick(Sender: TObject);
  private

  public

  end;

var
  FormSettings: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TFormSettings.btnCloseClick(Sender: TObject);
begin
  Close;
end;

end.

