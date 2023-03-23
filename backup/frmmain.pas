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
  DBCtrls, ExtCtrls, ComObj, SQLite3Conn, SQLDB, DB, variants;

type

  { TmainForm }

  TmainForm = class(TForm)
    btnGenerateDocument: TButton;
    Button1: TButton;
    cmbSenders: TComboBox;
    DS_Senders: TDataSource;
    editSubject: TEdit;
    groupSignature: TRadioGroup;
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MEMOFrom: TMemo;
    Panel1: TPanel;
    groupSalutation: TRadioGroup;
    SQLConnection: TSQLite3Connection;
    SQLSenders: TSQLQuery;
    SQLGeneralPurpose: TSQLQuery;
    SQLTransaction: TSQLTransaction;
    procedure btnGenerateDocumentClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure cmbSendersChange(Sender: TObject);

    procedure cmbSendersSelect(Sender: TObject);
    procedure DS_SendersDataChange(Sender: TObject; Field: TField);
    procedure FormCreate(Sender: TObject);
    procedure groupSalutationClick(Sender: TObject);
    procedure SQLConnectionAfterConnect(Sender: TObject);
  private

  public

  end;

var
  mainForm: TmainForm;

implementation
uses
    CustomComboItem, mailmerge;

{$R *.lfm}

{ TmainForm }

procedure TmainForm.btnGenerateDocumentClick(Sender: TObject);
{
 Load a template document and generate a new document based on that
 template.
}
var
  LOInstance, LOComponent, Text_, Cursor_, LoadParams, TextBody, ovc: variant;
  newFileName, templateFileName: string;
begin

  templateFileName := 'file:///C:/temp/template.odt';
  newFileName := 'file:///C:/temp/Brief.odt';

  // Create instance of LibreOffice
  LoadParams := VarArrayCreate([0, -1], varVariant);
  LOInstance := CreateOleObject('com.sun.star.ServiceManager');
  LOComponent := LOInstance.createInstance('com.sun.star.frame.Desktop');

  // Load the template document
  TextBody := LOComponent.loadComponentFromURL(templateFileName,
    '_blank', 0, LoadParams);


  // Get the body of the document
  Text_ := TextBody.createReplaceDescriptor;
  Text_.setSearchString('{NAME}');
  Text_.setReplaceString('Danny Van Geyte');
  TextBody.ReplaceAll(Text_);

  // Save the new file
  TextBody.storeToURL(newFileName, VarArrayCreate([0, 0], varVariant));
  TextBody.Close(True);
  // Now load the newly created document
  TextBody := LOComponent.loadComponentFromURL(newFileName, '_blank', 0, LoadParams);
  Text_ := TextBody.getText();

  oVC := TextBody.getCurrentController.getViewCursor;

  Cursor_ := Text_.createTextCursorByRange(oVC);
  oVC.JumpToStartOfPage;
  //  Cursor_.gotoRange(oVC, true);
  oVC.goDown(8, False);
  //  Cursor_.gotoRange(ovc, false);
  //  Cursor_.gotoPreviousWord(true);
  Text_.insertString(oVC.getStart(), 'JWAAAAJ', False);
end;

procedure TmainForm.Button1Click(Sender: TObject);
begin
  Close;
end;

procedure TmainForm.cmbSendersChange(Sender: TObject);
begin

end;

procedure TmainForm.cmbSendersSelect(Sender: TObject);
var
  myItem: TCustomComboBoxItem;
  qryString: string;
  SenderFullName: string;
  Reporter : TMailMerge;
begin
  myItem := TCustomComboBoxItem(cmbSenders.items.Objects
    [cmbSenders.ItemIndex]);
  qryString := 'SELECT * FROM senders WHERE id=' + myItem.Value.ToString;

  // Execute the sql statement and store the details of the sender.
  with SQLGeneralPurpose do
  begin
    Close;
    SQL.Text := qryString;
    Open
  end;
  if Not(SQLGeneralPurpose.Eof) then
  begin
    // Add the address to the FROM box
    Reporter := TMailMerge.Create;
    try
       Reporter.Name:=SQLGeneralPurpose.FieldByName('name').AsString;
       Reporter.FirstName:=SQLGeneralPurpose.FieldByName('FirstName').AsString;
       Reporter.Address:=SQLGeneralPurpose.FieldByName('Address').AsString;
       Reporter.ZipCode:=SQLGeneralPurpose.FieldByName('ZipCode').AsString;
       Reporter.Phone:=SQLGeneralPurpose.FieldByName('Phone').AsString;
       Reporter.EMail:=SQLGeneralPurpose.FieldByName('Email').AsString;
       Reporter.SocialSecurity:=SQLGeneralPurpose.FieldByName('socialsecurity').AsString;
       MEMOFrom.Lines.Add(Reporter.Name);
    finally
      Reporter.FreeInstance;
    end;

    //MemoFROM.Lines.Add(FDQryListOfSenders.FieldByName('address').AsString);
    //MemoFROM.Lines.Add(FDQryListOfSenders.FieldByName('zipcode').AsString + ' '
    //  + FDQryListOfSenders.FieldByName('city').AsString);
    //MemoFROM.Lines.Add('GSM: ' + FDQryListOfSenders.FieldByName('phone')
    //  .AsString);
    //MemoFROM.Lines.Add('email: ' + FDQryListOfSenders.FieldByName('email')
    //  .AsString);
    //CheckBox_SocialSecurityNr.Caption := FDQryListOfSenders.FieldByName
    //  ('socialsecurity').AsString;
  end;
end;

procedure TmainForm.DS_SendersDataChange(Sender: TObject; Field: TField);
begin

end;

procedure TmainForm.FormCreate(Sender: TObject);
{ Initial settings }
var
  dbPath: string;
begin
  dbPath := 'C:\Development\Lazarus\LetterWizard\Twister.db';
  SQLConnection.DatabaseName:=dbPath;
  SQLConnection.Connected:= true;
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
end;

procedure TmainForm.groupSalutationClick(Sender: TObject);
begin

end;

procedure TmainForm.SQLConnectionAfterConnect(Sender: TObject);
begin

end;

end.
