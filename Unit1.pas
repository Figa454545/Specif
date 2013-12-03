unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, DBTables, Grids, DBGrids, Word2000, OleServer,
  WordXP, ComCtrls, Menus;

type
  TForm1 = class(TForm)
    Table1: TTable;
    DataSource1: TDataSource;
    Table1Index: TAutoIncField;
    Table1Format: TStringField;
    Table1Name: TStringField;
    Table1Prim: TStringField;
    OpenDialog1: TOpenDialog;
    Table1Oboz: TStringField;
    Table9: TTable;
    DataSource9: TDataSource;
    Table9Name: TStringField;
    Table1Razd: TStringField;
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    Table9Index: TSmallintField;
    Table2: TTable;
    DataSource2: TDataSource;
    Table3: TTable;
    DataSource3: TDataSource;
    Table4: TTable;
    DataSource4: TDataSource;
    DBGrid3: TDBGrid;
    Label2: TLabel;
    Table3Index: TAutoIncField;
    Table3I: TSmallintField;
    Table3Group: TStringField;
    Table4Index: TAutoIncField;
    Table4Razd: TStringField;
    Table4Format: TStringField;
    Table4Oboz: TStringField;
    Table4Name: TStringField;
    Table4Prim: TStringField;
    DBGrid4: TDBGrid;
    DBGrid5: TDBGrid;
    Table2Index: TSmallintField;
    Table4Oboz_ispol: TStringField;
    WordApplication1: TWordApplication;
    WordDocument1: TWordDocument;
    Table1Position: TStringField;
    Table1Kol_vo: TStringField;
    Table4Position: TStringField;
    Table4Kol_vo: TStringField;
    ProgressBar1: TProgressBar;
    Table3Name: TStringField;
    Table2Name: TStringField;
    Label1: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    Word1: TMenuItem;
    Label5: TLabel;
    procedure Table2AfterScroll(DataSet: TDataSet);
    procedure FormCreate(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure Word1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure Table9AfterScroll(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  spec:textfile;
  cat,s,s1,oboz_,posit,im,im__,im_,Prim,datt,format,oboz,kol_vo:string;
  i,int,k,count_ispol,count_group_ispol:integer;

  Word_file,name123: OleVariant;
  W: Variant;

const Str_len = 24;
      vers = '1.06';
implementation

uses Unit2, ComObj;

{$R *.dfm}
procedure Decod(s,cat:string);
begin
  k:=1;
  i:=1;

  while s[k]<>';' do inc(k);  //Доходим по строке до следующей точки с запятой
  format:=copy(s,i,k-i);          //format = формат сборочной единицы
  inc(k);

  while s[k]<>';' do inc(k);  //Доходим по строке до следующей точки с запятой
  inc(k);
  i:=k;
  while s[i]<>';' do inc(i);  //Доходим по строке до следующей точки с запятой
  posit:=copy(s,k,i-k);          //posit = позиция сборочной единицы
  inc(i);
  k:=i;

  while s[k]<>';' do inc(k);    //Доходим по строке до следующей точки с запятой
  oboz:=copy(s,i,k-i);
  if pos(' ',oboz)<>0 then delete(oboz,pos(' ',oboz),1);   //oboz = обозначение сборочной единицы

  inc(k);
  i:=k;

  while s[i]<>';' do inc(i);    //Доходим по строке до следующей точки с запятой
  im:=copy(s,k,i-k);          //im = название сборочной единицы
  inc(i);
  k:=i;

  while s[k]<>';' do inc(k);          //Доходим по строке до следующей точки с запятой
  kol_vo:=copy(s,i,k-i);          //kol_vo = колличество сборочных единиц
  inc(k);
  i:=k;

  while s[i]<>';' do inc(i);    //Доходим по строке до следующей точки с запятой
  prim:=copy(s,k,i-k);          //prim = примечание сборочной единицы

  Form1.Table1.Insert;
  Form1.Table1Razd.value:=Cat;
  Form1.Table1Format.value:=format;
  Form1.Table1Position.value:=posit;
  Form1.Table1Oboz.value:=oboz;
  Form1.Table1Name.value:=im;
  Form1.Table1Kol_vo.value:=kol_vo;
  Form1.Table1Prim.value:=prim;
  Form1.Table1.Post;

  if Form1.Table9Name.value<>Cat then begin
    Form1.Table9.Insert;
    Form1.Table9Name.value:=Cat;
    if (Cat='Сборочные единицы') then Form1.Table9Index.value:=1;
    if (Cat='Детали') then Form1.Table9Index.value:=2;
    if (Cat='Стандартные изделия') then Form1.Table9Index.value:=3;
    if (Cat='Комплекты') then Form1.Table9Index.value:=6;
    if (Cat='Монтажные части') then Form1.Table9Index.value:=7;
    if (Cat='Инструменты') then Form1.Table9Index.value:=8;
    if (Cat='Материалы') then Form1.Table9Index.value:=5;
    if (Cat='Прочие изделия') then Form1.Table9Index.value:=4;
    if (Cat='Документация') then Form1.Table9Index.value:=0;
    Form1.Table9.Post;
  end;
end;

procedure Decod_peremen(s,cat:string);
begin
  k:=1;
  i:=1;

  while s[k]<>';' do inc(k);  //Доходим по строке до следующей точки с запятой
  format:=copy(s,i,k-i);          //format = формат сборочной единицы
  inc(k);

  while s[k]<>';' do inc(k);  //Доходим по строке до следующей точки с запятой
  inc(k);
  i:=k;
  while s[i]<>';' do inc(i);  //Доходим по строке до следующей точки с запятой
  posit:=copy(s,k,i-k);          //posit = позиция сборочной единицы
  inc(i);
  k:=i;

  while s[k]<>';' do inc(k);    //Доходим по строке до следующей точки с запятой
  oboz:=copy(s,i,k-i);
  if pos(' ',oboz)<>0 then delete(oboz,pos(' ',oboz),1);   //oboz = обозначение сборочной единицы

  inc(k);
  i:=k;

  while s[i]<>';' do inc(i);    //Доходим по строке до следующей точки с запятой
  im:=copy(s,k,i-k);          //im = название сборочной единицы
  inc(i);
  k:=i;

  while s[k]<>';' do inc(k);          //Доходим по строке до следующей точки с запятой
  kol_vo:=copy(s,i,k-i);          //kol_vo = колличество сборочных единиц
  inc(k);
  i:=k;

  while s[i]<>';' do inc(i);    //Доходим по строке до следующей точки с запятой
  prim:=copy(s,k,i-k);          //prim = примечание сборочной единицы

  Form1.Table4.Insert;
  Form1.Table4Razd.value:=Cat;
  Form1.Table4Format.value:=format;
  Form1.Table4Position.value:=posit;
  Form1.Table4Oboz.value:=oboz;
  Form1.Table4Name.value:=im;
  Form1.Table4Kol_vo.value:=kol_vo;
  Form1.Table4Prim.value:=prim;
  Form1.Table4Oboz_ispol.value:=Form1.Table2Name.Value;
  Form1.Table4.Post;

  if Form1.Table3Name.value<>Cat then begin
    Form1.Table3.Insert;
    Form1.Table3Name.value:=Cat;
    Form1.Table3Group.value:=Form1.Table2Name.value;
    count_group_ispol:=count_group_ispol+1;
    Form1.Table3I.value:=count_group_ispol;
    Form1.Table3.Post;
  end;


end;

procedure TForm1.Table2AfterScroll(DataSet: TDataSet);
begin
Table4.Filter:='Oboz_ispol='+#39+Table2Name.Value+#39;
Table4.Filtered:=true;
end;

procedure Dobav_strok(i:integer);
var k:integer;
begin
   if (i=30) or ((i>35)and((i-30) mod 31 = 0)) then begin
     Form1.WordDocument1.Tables.Item(1).Rows.Add(EmptyParam);
     Form1.WordDocument1.Tables.Item(1).Cell(i+1,5).Range.Font.Bold:=0;
     Form1.WordDocument1.Tables.Item(1).Cell(i+1,5).Range.Font.Underline:=0;
     Form1.WordDocument1.Tables.Item(1).Cell(i+1,5).Select;
     Form1.WordApplication1.Selection.ParagraphFormat.Alignment:=wdAlignParagraphLeft;
     Form1.WordApplication1.Selection.Collapse(EmptyParam);
     for k:=2 to 31 do Form1.WordDocument1.Tables.Item(1).Rows.Add(EmptyParam);
   end;
end;

procedure Prover_perenos(var i:integer);

begin
    if (i=28) or ((i>35)and((i-30) mod 31 = 29)) then begin   // Перенос раздела на новую страницу
      i:=i+1;
      i:=i+1;
      Dobav_strok(i);
    end;
    if (i=29) or ((i>35)and ((i-30) mod 31 = 30)) then begin
      i:=i+1;
      Dobav_strok(i);
    end;
end;

procedure Perenos(s:string);
var   s1:string;
      pos_space, sum_space: integer;

label m1;
begin
      if pos(' ',s)<>0 then begin                        // в строке есть несколько слов, находим по пробелам
        if pos('ГОСТ ',s)<>0 then
        delete(s,pos('ГОСТ ',s)+4,1);                        // в строке есть ГОСТ, уберем пробел, чтобы не перенесли номер на след. строку
        while length(s)>Str_len do begin
          pos_space:=pos(' ',s);                           // позиция первого пробела
          sum_space:=pos_space;
          if pos_space=0 then goto M1;
          s1:=copy(s,1,pos_space);                         // s1 - первое слово
          s:=copy(s,pos_space+1,length(s)-pos_space);      // убираем из s первое слово
          while pos_space<=Str_len+1 do begin                     // пока позиция пробела меньше 31
            pos_space:=pos(' ',s);                       // позиция первого пробела
            sum_space:=sum_space+pos_space;
            if (sum_space>Str_len) or (pos_space=0) then break;
            s1:=s1+copy(s,1,pos_space);
            s:=copy(s,pos(' ',s)+1,length(s)-pos(' ',s));
          end;
          if pos('ГОСТ',s1)<>0 then insert(' ',s1,pos('ГОСТ',s1)+4);                        // в строке есть ГОСТ, вернем пробел
          Form1.WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=s1;
          i:=i+1; Dobav_strok(i);
        end;
          if pos('ГОСТ',s)<>0 then insert(' ',s,pos('ГОСТ',s)+4);                        // в строке есть ГОСТ, вернем пробел
          Form1.WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=s;
      end else begin                           // в строке только одно слово, переносим без учета слогов
m1:       while length(s)>Str_len do begin
          s1:=copy(s,1,Str_len);
          s:=copy(s,Str_len+1,length(s)-26);
          Form1.WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=s1;
          i:=i+1;  Dobav_strok(i);
        end;
        Form1.WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=s;
      end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
AssignFile(spec,'Z:\Конструкторский отдел\Пупышев Дмитрий\Dim_Spec\version.txt');
Reset(spec);
Readln(spec,s);
if s<>vers then Application.MessageBox('Обновите версию программы!','Обновление',MB_OK);

Table1.DatabaseName:=ExtractFilePath(Application.ExeName)+'Data';
Table2.DatabaseName:=ExtractFilePath(Application.ExeName)+'Data';
Table3.DatabaseName:=ExtractFilePath(Application.ExeName)+'Data';
Table4.DatabaseName:=ExtractFilePath(Application.ExeName)+'Data';
Table9.DatabaseName:=ExtractFilePath(Application.ExeName)+'Data';

Table1.Active:=true;
Table2.Active:=true;
Table3.Active:=true;
Table4.Active:=true;
Table9.Active:=true;
end;

procedure File_read;
begin
if Form1.OpenDialog1.Execute then
begin
  AssignFile(spec,Form1.OpenDialog1.FileName);
  Reset(spec);
  count_group_ispol:=0;
  count_ispol:=0;
with Form1 do begin
  Table1.Active:=False;
  Table2.Active:=False;
  Table3.Active:=False;
  Table4.Active:=False;
  Table9.Active:=False;

  Table1.EmptyTable;
  Table2.EmptyTable;
  Table3.EmptyTable;
  Table4.EmptyTable;
  Table9.EmptyTable;

  Table1.Active:=True;
  Table2.Active:=True;
  Table3.Active:=True;
  Table4.Active:=True;
  Table9.Active:=True;

  s1:=Form1.OpenDialog1.FileName;
  im_:=Form1.OpenDialog1.FileName;
  delete(im_,Length(im_)-3,4);

  while pos('\',s1)<>0 do delete(s1,1,pos('\',s1));

  oboz_:=copy(s1,1,pos('_',s1)-1);
  delete(s1,1,pos('_',s1));
  im__:=copy(s1,1,pos('.',s1)-1);

  Readln(spec);     //Формат;Зона;Поз.;Обозначение;Наименование;Кол.;Примечание
  Readln(spec);     //;;;;;;;

Repeat
  while (pos('Документация',s)=0)and(pos('Сборочные',s)=0)and(pos('Детали',s)=0)and
        (pos('Стандартные',s)=0)and(pos('Монтажные',s)=0)and(pos('Инструменты',s)=0)and
        (pos('Материалы',s)=0)and(pos('Проч',s)=0)and(not Eof(spec))and(pos('Перемен',s)=0)and
        (pos('Комплекты',s)=0) do
    Readln(spec,s); //Читаем файл пока не встретим один из разделов
  if Eof(spec) then exit;
  if pos('Сборочные',s)<>0 then cat:='Сборочные единицы';   // Найден раздел Сборочные единицы
  if pos('Детали',s)<>0 then cat:='Детали';                 // Найден раздел детали
  if pos('Стандартные',s)<>0 then cat:='Стандартные изделия';
  if pos('Монтажные',s)<>0 then cat:='Монтажные части';
  if pos('Инструменты',s)<>0 then cat:='Инструменты';
  if pos('Материалы',s)<>0 then cat:='Материалы';
  if pos('Проч',s)<>0 then cat:='Прочие изделия';
  if pos('Документация',s)<>0 then cat:='Документация';
  if pos('Комплекты',s)<>0 then
  cat:='Комплекты';
  if pos('Перемен',s)<>0 then begin
    Form1.Table2.Insert;
    Form1.Table2Name.value:='Переменные данные для исполнений';
    Form1.Table2Index.value:=0;
    Form1.Table2.Post;
    break;
  end;

    Readln(spec,s);                                           //  ;;;;;;;
    Readln(spec,s);                                           //А4;;1;860.005СП1(М3);Корпус клапана;1;;
    while pos(';;;;;;;',s)=0 do begin
       if (pos('Инструменты',s)<>0) or (pos('Монтажные',s)<>0) then begin
                  Form1.Table9.Insert;
                  Form1.Table9Name.value:=Cat;
                  Form1.Table9Index.value:=6;
                  Form1.Table9.Post;
          break;
       end;
       Decod(s,cat);
       Readln(spec,s);  //Чтение след строки
       if (pos('* -',s)<>0)or(pos('Примечани',s)<>0)or EOF(spec) then exit;
    end;
Until EOF(spec);

Repeat
  while (pos('СБ',s)=0)and(pos('СП',s)=0)and(pos('Документация',s)=0)and(pos('Сборочные',s)=0)and(pos('Детали',s)=0)and
        (pos('Стандартные',s)=0)and(pos('Монтажные',s)=0)and(pos('Инструменты',s)=0)and(pos('Материалы',s)=0)and
        (pos('Комплекты',s)=0)and(pos('Проч',s)=0)and(not Eof(spec)) do
    Readln(spec,s); //Читаем файл пока не встретим один из разделов
  if Eof(spec) then begin
    Table1.First;
    Table2.First;
    Table3.First;
    Table4.First;
    Table9.First;
    exit;
  end;
  if (pos('СП',s)<>0)or(pos('СБ',s)<>0) then begin
    count_ispol:=count_ispol+1;
    Form1.Table2.Insert;
    Form1.Table2Name.value:=copy(s,5,Length(s)-7);
    Form1.Table2Index.value:=count_ispol;
    Form1.Table2.Post;
    Readln(spec,s);
  end else begin
  if pos('Сборочные',s)<>0 then cat:='Сборочные единицы';   // Найден раздел Сборочные единицы
  if pos('Детали',s)<>0 then cat:='Детали';                 // Найден раздел детали
  if pos('Стандартные',s)<>0 then cat:='Стандартные изделия';
  if pos('Комплекты',s)<>0 then
  cat:='Комплекты';
  if pos('Монтажные',s)<>0 then cat:='Монтажные части';
  if pos('Инструменты',s)<>0 then cat:='Инструменты';
  if pos('Материалы',s)<>0 then cat:='Материалы';
  if pos('Проч',s)<>0 then cat:='Прочие изделия';
  if pos('Документация',s)<>0 then cat:='Документация';


    Readln(spec,s);                                           //  ;;;;;;;
    Readln(spec,s);
    while pos(';;;;;;;',s)<>0 do Readln(spec,s);              //Читаем пустые строки пока не встретим элемент спецификации
    while pos(';;;;;;;',s)=0 do begin                        //А4;;1;860.005СП1(М3);Корпус клапана;1;;
       if (pos('Инструменты',s)<>0) or (pos('Монтажные',s)<>0) then begin
              Form1.Table3.Insert;
              Form1.Table3Name.value:=Cat;
              Form1.Table3Group.value:=Form1.Table2Name.value;
              count_group_ispol:=count_group_ispol+1;
              Form1.Table3I.value:=count_group_ispol;
              Form1.Table3.Post;
          break;
       end;
       Decod_peremen(s,cat);
       Readln(spec,s);  //Чтение след строки
       if (pos('* -',s)<>0)or(pos('Примечани',s)<>0)or EOF(spec) then exit;
    end;
  end;
Until ((pos('Примечани',s)<>0)or(pos('* -',s)<>0)or EOF(spec));
end;
end;
end;

procedure TForm1.N3Click(Sender: TObject);
begin
  File_read;

  Table1.First;
  Table2.First;
  Table3.First;
  Table4.First;
  Table9.First;

end;

procedure TForm1.Word1Click(Sender: TObject);
begin
  ProgressBar1.Visible:=true;
  Form1.Enabled:=False;
  ProgressBar1.Position:=0;

  i:=2;
  Table9.First;                                   //Подсчет записей
  While not Table9.Eof do begin
    i:=i+2;
    Table1.First;
    While not Table1.Eof do begin
      i:=i+1;
      Table1.Next;
    end;
    Table9.Next;
  end;
  Table2.First;
  While not Table2.Eof do begin
    i:=i+3;
    Table3.First;
    While not Table3.Eof do begin
      i:=i+3;
      Table4.First;
      While not Table4.Eof do begin
        i:=i+1;
        Table4.Next;
      end;
      Table3.Next;
    end;
    Table2.Next;
  end;
  ProgressBar1.Max:=i;

  WordApplication1.Connect;

  Word_file:=ExtractFilePath(Application.ExeName)+'Шаблон для спецификации.doc';// Form1.OpenDialog1.FileName;
  WordApplication1.Documents.Open(Word_File,EmptyParam,EmptyParam,EmptyParam,
                                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                                            EmptyParam,EmptyParam,EmptyParam,EmptyParam,
                                            EmptyParam,EmptyParam,EmptyParam);
  WordApplication1.Options.CheckSpellingAsYouType:=false;
  WordApplication1.Options.CheckGrammarAsYouType:=false;
  WordDocument1.ConnectTo(WordApplication1.ActiveDocument);

  i:=2;
  Table9.First;
  While not Table9.Eof do begin
    Prover_perenos(i);
    WordDocument1.Tables.Item(1).Cell(i,5).Range.Font.Bold:=1;
    WordDocument1.Tables.Item(1).Cell(i,5).Range.Font.Underline:=1;
    WordDocument1.Tables.Item(1).Cell(i,5).Select;
    WordApplication1.Selection.ParagraphFormat.Alignment:=wdAlignParagraphCenter;
    WordApplication1.Selection.Collapse(EmptyParam);

    ProgressBar1.Position:=i;

    if Length(Table9Name.Value)>Str_len then Perenos(Table9Name.Value)
    else begin
    WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=Table9Name.Value;
    i:=i+1; Dobav_strok(i);
    end;
    Table1.First;
    if Table1.Eof then i:=i-1;
    While not Table1.Eof do begin
      i:=i+1; Dobav_strok(i);
      ProgressBar1.Position:=i;
      if Length(Table1Name.Value)>Str_len then begin
        Prover_perenos(i);
        WordDocument1.Tables.Item(1).Cell(i,1).Range.Text:=Table1Format.Value;
        WordDocument1.Tables.Item(1).Cell(i,3).Range.Text:=Table1Position.Value;
        WordDocument1.Tables.Item(1).Cell(i,4).Range.Text:=Table1Oboz.Value;
        Perenos(Table1Name.Value);
      end else begin
        WordDocument1.Tables.Item(1).Cell(i,1).Range.Text:=Table1Format.Value;
        WordDocument1.Tables.Item(1).Cell(i,3).Range.Text:=Table1Position.Value;
        WordDocument1.Tables.Item(1).Cell(i,4).Range.Text:=Table1Oboz.Value;
        WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=Table1Name.Value;
      end;
      WordDocument1.Tables.Item(1).Cell(i,6).Range.Text:=Table1Kol_vo.Value;
      WordDocument1.Tables.Item(1).Cell(i,7).Range.Text:=Table1Prim.Value;
      Table1.Next;
    end;
    i:=i+1; Dobav_strok(i);
    i:=i+1; Dobav_strok(i);
    Table9.Next;
  end;
  Table2.First;
  While not Table2.Eof do begin
    Prover_perenos(i);
    WordDocument1.Tables.Item(1).Cell(i,5).Range.Font.Bold:=1;
    WordDocument1.Tables.Item(1).Cell(i,5).Range.Font.Underline:=1;
    WordDocument1.Tables.Item(1).Cell(i,5).Select;
    WordApplication1.Selection.ParagraphFormat.Alignment:=wdAlignParagraphCenter;
    WordApplication1.Selection.Collapse(EmptyParam);
    ProgressBar1.Position:=i;
    if Table2Name.Value='Переменные данные для исполнений' then begin
      WordDocument1.Tables.Item(1).Cell(i,4).Range.Font.Bold:=1;
      WordDocument1.Tables.Item(1).Cell(i,4).Range.Font.Underline:=1;
      WordDocument1.Tables.Item(1).Cell(i,4).Select;
      WordApplication1.Selection.ParagraphFormat.Alignment:=wdAlignParagraphRight;
      WordDocument1.Tables.Item(1).Cell(i,4).RightPadding:=0;
      WordApplication1.Selection.Collapse(EmptyParam);
      WordDocument1.Tables.Item(1).Cell(i,4).Range.Text:='Переменные данн';

      WordDocument1.Tables.Item(1).Cell(i,5).Select;
      WordApplication1.Selection.ParagraphFormat.Alignment:=wdAlignParagraphLeft;
      WordDocument1.Tables.Item(1).Cell(i,5).LeftPadding:=0;
      WordApplication1.Selection.Collapse(EmptyParam);
      WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:='ые для исполнений';


    end else WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=Table2Name.Value;
    i:=i+1; Dobav_strok(i);
    i:=i+1; Dobav_strok(i);
    Table3.First;
    While not Table3.Eof do begin
      WordDocument1.Tables.Item(1).Cell(i,5).Range.Font.Bold:=1;
      WordDocument1.Tables.Item(1).Cell(i,5).Range.Font.Underline:=1;
      WordDocument1.Tables.Item(1).Cell(i,5).Select;
      WordApplication1.Selection.ParagraphFormat.Alignment:=wdAlignParagraphCenter;
      WordApplication1.Selection.Collapse(EmptyParam);
      ProgressBar1.Position:=i;
      WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=Table3Name.Value;
      i:=i+1; Dobav_strok(i);
      i:=i+1; Dobav_strok(i);
      Table4.First;
      if Table4.Eof then i:=i-1;
      While not Table4.Eof do begin
        Prover_perenos(i);

      if Length(Table4Name.Value)>Str_len then begin
        Prover_perenos(i);
        WordDocument1.Tables.Item(1).Cell(i,1).Range.Text:=Table4Format.Value;
        WordDocument1.Tables.Item(1).Cell(i,3).Range.Text:=Table4Position.Value;
        WordDocument1.Tables.Item(1).Cell(i,4).Range.Text:=Table4Oboz.Value;
        Perenos(Table4Name.Value);
      end else begin
        WordDocument1.Tables.Item(1).Cell(i,1).Range.Text:=Table4Format.Value;
        WordDocument1.Tables.Item(1).Cell(i,3).Range.Text:=Table4Position.Value;
        WordDocument1.Tables.Item(1).Cell(i,4).Range.Text:=Table4Oboz.Value;
        WordDocument1.Tables.Item(1).Cell(i,5).Range.Text:=Table4Name.Value;
      end;

        WordDocument1.Tables.Item(1).Cell(i,6).Range.Text:=Table4Kol_vo.Value;
        WordDocument1.Tables.Item(1).Cell(i,7).Range.Text:=Table4Prim.Value;
        i:=i+1; Dobav_strok(i);
        Table4.Next;
      end;
      i:=i+1;  Dobav_strok(i);
      Table3.Next;
    end;
    Table2.Next;
  end;
  WordApplication1.Visible:=true;
  W:=WordApplication1.Dialogs.Item(84);
  W.Name:=im_+'.doc';
  W.Show;
  WordDocument1.Disconnect;
  WordApplication1.Disconnect;
  ProgressBar1.Visible:=False;
  Form1.Enabled:=True;

end;

procedure TForm1.N2Click(Sender: TObject);
begin
AssignFile(spec,'Z:\Конструкторский отдел\Пупышев Дмитрий\Dim_Spec\version.txt');
Reset(spec);
Readln(spec,s);
if s<>vers then Application.MessageBox('Обновите версию программы!','Обновление',MB_OK);
AboutBox.ShowModal;
end;

procedure TForm1.Table9AfterScroll(DataSet: TDataSet);
begin
Table1.Filter:='Razd='+#39+Table9Name.Value+#39;
Table1.Filtered:=true;
end;

end.
