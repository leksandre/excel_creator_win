unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExcelXP, OleCtrls, comObj, Grids, DBGrids, DB, ADODB; // , Excel2000

type
  TForm1 = class(TForm)
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    DBGrid1: TDBGrid;
    procedure Button1Click(Sender: TObject);
    procedure report(title:string; source:TDBGrid);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses Math;

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
 begin
 //
 end;
            
procedure TForm1.report(title:string; source:TDBGrid);
 var
    ExcelAppl: OleVariant;
    WorkSheetr : Variant;
    WorkBkr : _WorkBook;
    atitle,aFileName:string;
    i,j,y,c:integer;
begin
if (not (source.DataSource.DataSet.Active)) then Exit;

 if title='' then atitle:=InputBox('введите заголовок','по умолчанию','');
 if atitle<>'' then title:=atitle;

    If (OpenDialog1.Execute<>false) and (OpenDialog1.FileName<>'') then
  begin
  try
    ExcelAppl:=CreateOleObject('Excel.Application');
    ExcelAppl.Visible:=False;
  except
    on E: Exception do
      raise Exception.Create('Ошибка создания объекта Excel: ' + E.Message);
  end;
  try
  aFileName:=OpenDialog1.FileName;
    ExcelAppl.WorkBooks.Open(aFileName);
    ExcelAppl.ActiveWorkBook.WorkSheets.Add;
    WorkSheetr :=  ExcelAppl.ActiveWorkBook.WorkSheets[1];// as _WorkSheet
    ExcelAppl.ActiveWorkBook.WorkSheets[1].name:=title+inttostr(ExcelAppl.ActiveWorkBook.WorkSheets.count);
    c:=source.FieldCount;
    source.DataSource.DataSet.First;
    y:=0;
for i:=0 to c-1 do
    if source.Columns[i].Visible then
    begin
    WorkSheetr.cells[3,y+2]:=source.Fields[i].DisplayName;
    WorkSheetr.Columns[y+2].ColumnWidth:=source.Fields[i].DisplayWidth;  //  Canvas.TextWidth(WorkSheetr.cells[3,i+2])
    inc(y);
    end;
    

    while not source.DataSource.DataSet.Eof do
    begin  inc (j);    y:=0;
for i:=0 to c-1 do
    if source.Columns[i].Visible then  begin
    WorkSheetr.cells[j+3,y+2]:=vartostr(source.Fields[i].CurValue);
    inc(y);
    end;

    source.DataSource.DataSet.Next;       
    end;               

    WorkSheetr.cells[1,2]:=title;

    ExcelAppl.ActiveWorkBook.Save;
  finally
    ExcelAppl.Quit;
    ExcelAppl:=Unassigned;  ShowMessage('Succes!!');
  end;
  end;
  
end;
end.
