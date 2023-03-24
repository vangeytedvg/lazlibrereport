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
  SQLite3Conn, SQLDB, DB, variants, mailmerge, IniFiles;

type

  { TmainForm }

  TmainForm = class(TForm)
    btnSettings: TBitBtn;
    btnGenerateDocument: TButton;
    Button1: TButton;
    cbSocialSecurityInclude: TCheckBox;
    cmbSenders: TComboBox;
    DS_Senders: TDataSource;
    edtDocName: TEdit;
    editSubject: TEdit;
    groupSignature: TRadioGroup;
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    MEMOFrom: TMemo;
    groupSalutation: TRadioGroup;
    MEMOTo: TMemo;
    SQLConnection: TSQLite3Connection;
    SQLSenders: TSQLQuery;
    SQLGeneralPurpose: TSQLQuery;
    SQLTransaction: TSQLTransaction;
    StatusBar: TStatusBar;
    procedure btnGenerateDocumentClick(Sender: TObject);
    procedure GenerateDocument;
    procedure btnSettingsClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure cbSocialSecurityIncludeChange(Sender: TObject);
    procedure cmbSendersChange(Sender: TObject);
    procedure cmbSendersSelect(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure groupSalutationClick(Sender: TObject);
    procedure groupSignatureClick(Sender: TObject);
  private
    FDbPath: string;
    FTemplate: string;
    FNewPath: string;
    Reporter: TMailMerge;
    procedure GetIniSettings;
    property DbPath: string read FDbPath write FDbPath;
    property Template: string read FTemplate write FTemplate;
    property NewPath: string read FNewPath write FNewPath;
  public

  end;

var
  mainForm: TmainForm;

implementation

uses
  CustomComboItem, frmSettings, Settings;

var
  IniSettings: TAppSettings;

{$R *.lfm}

{ TmainForm }
procedure TmainForm.FormCreate(Sender: TObject);
{ Initial settings }

begin
  // Get the settings from the Ini file
  IniSettings := TAppSettings.Create('Twister.ini');
  IniSettings.ReadSettings;

  // Set the path for the sqlite database
  dbPath := IniSettings.dbPath;
  Template := IniSettings.TemplatePath.Replace('\', '/');
  NewPath := IniSettings.NewDocStorage.Replace('\', '/');

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

procedure TmainForm.GetIniSettings;
var
  ini: TIniFile;
begin

end;

procedure TmainForm.btnGenerateDocumentClick(Sender: TObject);
{ Check if there is a name for the document }
begin
  // Check if a document name was entered
  if edtDocName.Text = '' then
  begin
    ShowMessage('Geef een naam voor deze brief aub!');
    edtDocName.SetFocus;
    exit;
  end;
  if MEMOTo.Lines.Count = 0 then
  begin
    ShowMessage('Een bestemmeling is verplicht!');
    MEMOTo.SetFocus;
    exit;
  end;
  if editSubject.Text = '' then
  begin
    ShowMessage('Een onderwerp is verplicht!');
    editSubject.SetFocus;
    exit;
  end;
  // Check if a sender was selected
  if cmbSenders.ItemIndex <> -1 then
  begin
    // Do something with the selected item
    GenerateDocument;
  end
  else
  begin
    // No item is selected
    ShowMessage('U moet een afzender kiezen!.');
    cmbSenders.SetFocus;
  end;

end;

procedure TmainForm.GenerateDocument;
{
 Load a template document and generate a new document based on that
 template.
}
var
  LOInstance, LOComponent, Text_, Cursor_, LoadParams, TextBody, ovc: variant;
  newFileName, templateFileName: string;
  DocDate: TDateTime;
  FormattedDateCity: string;
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


  //  templateFileName := 'file:///C:/temp/template.odt';
  templateFileName := 'file:///' + Template;
  newFileName := 'file:///' + NewPath + '/' + edtDocName.Text + '.odt';

  // Create instance of LibreOffice
  LoadParams := VarArrayCreate([0, -1], varVariant);
  LOInstance := CreateOleObject('com.sun.star.ServiceManager');
  LOComponent := LOInstance.createInstance('com.sun.star.frame.Desktop');

  StatusBar.Panels[0].Text := 'Template laden...';
  Application.ProcessMessages;

  // Load the template document
  TextBody := LOComponent.loadComponentFromURL(templateFileName,
    '_blank', 0, LoadParams);

  StatusBar.Panels[0].Text := 'Afzender invoegen...';
  Application.ProcessMessages;

  // Replace the {SENDER} tag
  If cbSocialSecurityInclude.Checked then
  MEMOFrom.Text := MEMOFrom.Text + 'RijksRegister : ' + Reporter.SocialSecurity;
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{SENDER}');
  Text_.setReplaceString(MEMOFrom.Text);
  TextBody.ReplaceAll(Text_);

  StatusBar.Panels[0].Text := 'Datum invoegen...';
  Application.ProcessMessages;

  // Replace the {TO} tag
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{TO}');
  Text_.setReplaceString(MEMOTo.Text);
  TextBody.ReplaceAll(Text_);

  // Replace the {SENDDATECITY} tag
  DocDate := Now;
  FormattedDateCity := Reporter.City + ', ' + FormatDateTime('dd mmmm yyyy', DocDate);
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{SENDDATECITY}');
  Text_.setReplaceString(FormattedDateCity);
  TextBody.ReplaceAll(Text_);

  // Replace the {SUBJECT} tag
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{SUBJECT}');
  Text_.setReplaceString(Reporter.Subject);
  TextBody.ReplaceAll(Text_);

  // Replace the {GREETING} tag
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{GREETING}');
  Text_.setReplaceString(Reporter.Salutation);
  TextBody.ReplaceAll(Text_);

  // Replace the {CLOSURE} tag
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{CLOSURE}');
  Text_.setReplaceString(Reporter.Signature);
  TextBody.ReplaceAll(Text_);


  // Replace the {SIGNATURE} tag
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{SIGNATURE}');
  Text_.setReplaceString(Reporter.FullName);
  TextBody.ReplaceAll(Text_);

  StatusBar.Panels[0].Text := 'Gegenereerd document opslaan...';
  Application.ProcessMessages;

  // ----- Save the new file ----
  TextBody.storeToURL(newFileName, VarArrayCreate([0, 0], varVariant));
  TextBody.Close(True);

  StatusBar.Panels[0].Text := 'Laden gegenereerd document...';
  Application.ProcessMessages;

  // Now load the newly created document
  TextBody := LOComponent.loadComponentFromURL(newFileName, '_blank', 0, LoadParams);
  Text_ := TextBody.getText();

  StatusBar.Panels[0].Text := 'Klaar...';
  Application.ProcessMessages;

  // Move the cursor to the place where the user can start typing
  oVC := TextBody.getCurrentController.getViewCursor;
  Cursor_ := Text_.createTextCursorByRange(oVC);
  oVC.JumpToStartOfPage;
  oVC.goDown(27, False);

end;

procedure TmainForm.btnSettingsClick(Sender: TObject);
var
  FSetting: TFormSettings;
begin
  FSetting := TFormSettings.Create(self);
  FSetting.ShowModal;
  // Adapt the settings
  IniSettings.ReadSettings;
  // Set the path for the sqlite database
  dbPath := IniSettings.dbPath;
  Template := IniSettings.TemplatePath.Replace('\', '/');
  NewPath := IniSettings.NewDocStorage.Replace('\', '/');

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

procedure TmainForm.groupSalutationClick(Sender: TObject);
begin

end;

procedure TmainForm.groupSignatureClick(Sender: TObject);
begin

end;


end.
