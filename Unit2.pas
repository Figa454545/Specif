unit Unit2;

interface

uses
  Windows, SysUtils, Variants, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, Word2000, OleServer,  WordXP, ComCtrls;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    ProgramIcon: TImage;
    ProductName: TLabel;
    Version: TLabel;
    Copyright: TLabel;
    Comments: TLabel;
    OKButton: TButton;
    CheckBox1: TCheckBox;
    Memo1: TMemo;
    procedure OKButtonClick(Sender: TObject);
    procedure ProgramIconClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutBox: TAboutBox;
  Word_file1: OleVariant;

implementation

uses Unit1;

{$R *.dfm}

procedure TAboutBox.OKButtonClick(Sender: TObject);
begin
Close;
end;

procedure TAboutBox.ProgramIconClick(Sender: TObject);
begin
  Form1.WordApplication1.Connect;
  Word_file1:=ExtractFilePath(Application.ExeName)+'/Шаблон для спецификации.doc';// Form1.OpenDialog1.FileName;
  Form1.WordApplication1.Documents.Open(Word_File1,EmptyParam,EmptyParam,EmptyParam,
                                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                                            EmptyParam,EmptyParam,EmptyParam);
  Form1.WordApplication1.Options.CheckSpellingAsYouType:=false;
  Form1.WordApplication1.Options.CheckGrammarAsYouType:=false;
  Form1.WordDocument1.ConnectTo(Form1.WordApplication1.ActiveDocument);
  Form1.WordApplication1.Visible:=true;
  Form1.WordApplication1.Activate;
  Form1.WordDocument1.Disconnect;
  Form1.WordApplication1.Disconnect;

end;

procedure TAboutBox.FormCreate(Sender: TObject);
begin
  Version.Caption:=vers;
end;

procedure TAboutBox.CheckBox1Click(Sender: TObject);
begin
 if CheckBox1.Checked then AboutBox.Height:=510 else AboutBox.Height:=280;
end;

end.

