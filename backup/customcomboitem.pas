{
  CustomComboBoxItem
  Version : 1.0
  Created : 31/01/2023
  DVG
  Description : Represents an item in a combobox with a key/value
}
unit CustomComboItem;

interface

type
  TCustomComboBoxItem = class(TObject)
  private
    FValue: integer;
    FDisplayText: string;
  public
    property Value: integer read FValue write FValue;
    property DisplayText: string read FDisplayText write FDisplayText;
    constructor Create(AValue: integer; ADisplayText: string);
    destructor Destroy; override;
  end;

implementation

constructor TCustomComboBoxItem.Create(AValue: integer; ADisplayText: String);
begin
  FValue := AValue;
  FDisplayText := ADisplayText;
end;

destructor TCustomComboBoxItem.Destroy;
begin

end;

end.
