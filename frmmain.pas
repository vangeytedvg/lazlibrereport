{
  Main Module for the Letter Wizard, Lazarus Edition.
  Created : 23/03/2023
  Author  : Danny Van Geyte
}

unit frmMain;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, DBGrids,
  DBCtrls, ExtCtrls, ComCtrls, PopupNotifier, Buttons, CheckBoxThemed, ComObj,
  SQLite3Conn, SQLDB, DB, variants, mailmerge;

type

  { TmainForm }

  TmainForm = class(TForm)
    btnSettings: TBitBtn;
    btnGenerateDocument: TButton;
    Button1: TButton;
    cbSocialSecurityInclude: TCheckBox;
    cmbSenders: TComboBox;
    DS_Senders: TDataSource;
    editSubject: TEdit;
    groupSignature: TRadioGroup;
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    MEMOFrom: TMemo;
    groupSalutation: TRadioGroup;
    MEMOTo: TMemo;
    SQLConnection: TSQLite3Connection;
    SQLSenders: TSQLQuery;
    SQLGeneralPurpose: TSQLQuery;
    SQLTransaction: TSQLTransaction;
    StatusBar: TStatusBar;
    procedure btnGenerateDocumentClick(Sender: TObject);
    procedure btnSettingsClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure cbSocialSecurityIncludeChange(Sender: TObject);
    procedure cmbSendersChange(Sender: TObject);
    procedure cmbSendersSelect(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    Reporter: TMailMerge;
  public

  end;

var
  mainForm: TmainForm;

implementation

uses
  CustomComboItem, frmSettings;

{$R *.lfm}

{ TmainForm }
procedure TmainForm.FormCreate(Sender: TObject);
{ Initial settings }
var
  dbPath: string;
begin
  dbPath := 'C:\Development\Lazarus\LetterWizard\Twister.db';
  StatusBar.Panels[0].Text := 'Opening database';
  SQLConnection.DatabaseName := dbPath;
  SQLConnection.Connected := True;
  SQLSenders.Active := True;
  Reporter := TMailMerge.Create;
  with SQLSenders do
  begin
    Close;
    Open;
  end;
  // Fill The ComboBox
  while not SQLSenders.EOF do
  begin
    cmbSenders.items.AddObject(SqlSenders.FieldByName('Name').AsString +
      ' ' + SqlSenders.FieldByName('firstname').AsString,
      TCustomComboBoxItem.Create(SQLSenders.FieldByName('id').AsInteger,
      SQLSenders.FieldByName('Name').AsString));
    SQLSenders.Next;
  end;
  StatusBar.Panels[0].Text := 'Ready...';
end;

procedure TmainForm.btnGenerateDocumentClick(Sender: TObject);
{
 Load a template document and generate a new document based on that
 template.
}
var
  LOInstance, LOComponent, Text_, Cursor_, LoadParams, TextBody, ovc: variant;
  newFileName, templateFileName: string;
begin

  // First Fill the remaining data of the Reporter object
  Reporter.Subject := editSubject.Text;
  Reporter.Destination := MemoTo.Text;

  // The Greeting
  if groupSalutation.ItemIndex <> -1 then
    Reporter.Salutation := groupSalutation.Items[groupSalutation.ItemIndex];

  // The signature
  if groupSignature.ItemIndex <> -1 then
    Reporter.Signature := groupSignature.Items[groupSignature.ItemIndex];

  templateFileName := 'file:///C:/temp/template.odt';
  newFileName := 'file:///C:/temp/Brief.odt';

  // Create instance of LibreOffice
  LoadParams := VarArrayCreate([0, -1], varVariant);
  LOInstance := CreateOleObject('com.sun.star.ServiceManager');
  LOComponent := LOInstance.createInstance('com.sun.star.frame.Desktop');

  // Load the template document
  TextBody := LOComponent.loadComponentFromURL(templateFileName,
    '_blank', 0, LoadParams);

  // Replace the {SENDER} tag
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{SENDER}');
  Text_.setReplaceString(MEMOFrom.Text);
  TextBody.ReplaceAll(Text_);
  Text_ := nil;

  // Replace the {SENDDATE} tag
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{SENDDATE}');
  Text_.setReplaceString('DAADa');
  TextBody.ReplaceAll(Text_);


  // Save the new file
  TextBody.storeToURL(newFileName, VarArrayCreate([0, 0], varVariant));
  TextBody.Close(True);
  // Now load the newly created document
  TextBody := LOComponent.loadComponentFromURL(newFileName, '_blank', 0, LoadParams);
  Text_ := TextBody.getText();

  // Move the cursor to a new location
  oVC := TextBody.getCurrentController.getViewCursor;
  Cursor_ := Text_.createTextCursorByRange(oVC);
  oVC.JumpToStartOfPage;
  oVC.goDown(8, False);

end;

procedure TmainForm.btnSettingsClick(Sender: TObject);
var
  FSetting : TFormSettings;
begin
  FSetting := TFormSettings.Create(self);
  FSetting.ShowModal;
end;

procedure TmainForm.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TmainForm.cbSocialSecurityIncludeChange(Sender: TObject);
begin
end;

procedure TmainForm.cmbSendersChange(Sender: TObject);
begin

end;

procedure TmainForm.cmbSendersSelect(Sender: TObject);
var
  myItem: TCustomComboBoxItem;
  qryString: string;
  SenderFullName: string;
begin

  MEMOFrom.Clear;
  myItem := TCustomComboBoxItem(cmbSenders.items.Objects[cmbSenders.ItemIndex]);
  qryString := 'SELECT * FROM senders WHERE id=' + myItem.Value.ToString;

  // Execute the sql statement and store the details of the sender.
  with SQLGeneralPurpose do
  begin
    Close;
    SQL.Text := qryString;
    Open;
  end;
  if not (SQLGeneralPurpose.EOF) then
  begin
    // Add the address to the FROM box
    try
      Reporter.Name := SQLGeneralPurpose.FieldByName('name').AsString;
      Reporter.FirstName := SQLGeneralPurpose.FieldByName('FirstName').AsString;
      Reporter.Address := SQLGeneralPurpose.FieldByName('Address').AsString;
      Reporter.ZipCode := SQLGeneralPurpose.FieldByName('ZipCode').AsString;
      Reporter.City := SQLGeneralPurpose.FieldByName('City').AsString;
      Reporter.Phone := SQLGeneralPurpose.FieldByName('Phone').AsString;
      Reporter.EMail := SQLGeneralPurpose.FieldByName('Email').AsString;
      Reporter.SocialSecurity :=
        SQLGeneralPurpose.FieldByName('socialsecurity').AsString;

      MEMOFrom.Lines.Add(Reporter.FullName);
      MemoFrom.Lines.Add(Reporter.Address);
      MemoFrom.Lines.Add(Reporter.ZipCode + '  ' + Reporter.City);
      MemoFrom.Lines.Add('GSM : ' + Reporter.Phone);
      MemoFrom.Lines.Add('Email ' + Reporter.Email);
    except
      ShowMessage('Error occured');
    end;
  end;

end;

procedure TmainForm.FormDestroy(Sender: TObject);
begin
  Reporter.FreeInstance;
end;


end.
