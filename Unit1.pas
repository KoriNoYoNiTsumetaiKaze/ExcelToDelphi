unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComObj, Buttons;

type
  TForm1 = class(TForm)
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    BitBtn1: TBitBtn;
    GroupBox1: TGroupBox;
    ListBox1: TListBox;
    GroupBox2: TGroupBox;
    ListBox2: TListBox;
    Label1: TLabel;
    Edit1: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var XL,Sheet:Variant;
    XLSFile:String;
    i,row:integer;
    FData:OLEVariant;
    sum:Real;
begin
  if OpenDialog1.Execute then XLSFile:=OpenDialog1.FileName
  else Exit;
  XL:= CreateOleObject('Excel.Application');
  XL.Visible:= false;
  XL.Workbooks.Open(XLSFile);
  Sheet:=XL.ActiveWorkbook.ActiveSheet;
  FData:=Sheet.UsedRange.Value;
  row:=Sheet.UsedRange.Rows.Count;
  sum:=0;
  for i:=1 to row do
    begin
      try
        ListBox1.Items.Add(FData[i,1]);
        ListBox2.Items.Add(FData[i,2]);
        sum:=sum+StrToFloat(FData[i,2]);
      except
        on E : Exception do
          ShowMessage(E.ClassName+' Ошибка добавления строки Excel '+E.Message);
      end;
    end;
  Edit1.Text:=FloatToStr(sum);
  XL.Quit;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
Form1.Close;
end;

end.
