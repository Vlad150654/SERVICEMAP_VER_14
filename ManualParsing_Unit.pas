unit ManualParsing_Unit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.Grids,
  Common, Vcl.ExtCtrls;

type
  TManual = class(TForm)
    sg: TStringGrid;
    BitBtn1: TBitBtn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn1Click(Sender: TObject);
    procedure sgDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure sgDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    filename : String;
  end;

var
  Manual: TManual;

implementation

{$R *.dfm}

// mrNone, mrOk, mrCancel, mrAbort, mrRetry, mrIgnore, mrYes, mrNo
procedure TManual.BitBtn1Click(Sender: TObject);
var
  i, j, n, k, first : Integer;
  cn : String;
  sl : TStringList;
begin
  for i := 1 to sg.RowCount-1 do
    sg.Cells[sg.ColCount-1, i] := '';
  sl := TStringList.Create;
  first := 0;
  i := 1;
  try
    while i < sg.RowCount do
    begin
      cn := sg.Cells[0, i];
      j := i + 1;
      while sg.Cells[0, j] = cn do
        Inc(j);//считаем количествр КН
      k := 0;
      for n := i to j-1 do
        if sg.Cells[sg.ColCount-2, n] = '1' then
          Inc(k)
        else
          Dec(k);
      n := (j - i) mod 2;//проверяем на четность
      if ((n = 0) and (k <> 0) or (n > 0) and (k <> 1))  then
      begin
        sl.Add('Ошибка в условиях выбора при кадастровом номере ' + cn);
        for n := i to j-1 do
          sg.Cells[sg.ColCount-1, n] := '???';
        if first = 0 then
          first := i;
      end;
      i := j;
    end;
    if sl.Count > 0 then
    begin
      sg.Row := first;
      ModalResult := MessageDlg('Обнаружены ошибки:' + #10 + sl.Text, mtCustom,
          [mbRetry, mbIgnore, mbCancel], 0);
      if ModalResult = mrRetry then
        ModalResult := mrNone;
      exit;
    end;
  finally
    sl.Free;
  end;

  sl := TStringList.Create;
  for i := 0 to sg.RowCount-1 do
  begin
    cn := sg.Cells[0, i];
    for j := 1 to sg.ColCount-2 do
      cn := cn + ';' + sg.Cells[j, i];
    sl.Add(cn);
  end;
  sl.SaveToFile(filename);//Ручной разбор.csv
  sl.Free;

  ModalResult := mrOk;
end;

procedure TManual.FormActivate(Sender: TObject);
begin
  BitBtn1Click(nil);
end;

procedure TManual.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TManual.sgDblClick(Sender: TObject);
begin
  if sg.Cells[sg.ColCount-2, sg.Row] = '1' then
    sg.Cells[sg.ColCount-2, sg.Row] := '0'
  else
    sg.Cells[sg.ColCount-2, sg.Row] := '1';
  sg.Repaint;
  sg.Refresh;
end;

procedure TManual.sgDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
  State: TGridDrawState);
begin
  if sg.Cells[sg.ColCount-2, ARow] = '1' then
    sg.Canvas.Brush.Color := clGradientInactiveCaption
  else
    sg.Canvas.Brush.Color := clWindow;
  sg.Canvas.FillRect(Rect);
  sg.Canvas.TextOut(Rect.Left, Rect.Top, sg.Cells[ACol, ARow]);
end;

procedure TManual.FormShow(Sender: TObject);
begin
  sg.ColWidths[0] := 160;
  sg.ColWidths[5] := 120;
  sg.ColWidths[8] := 120;
  sg.ColWidths[9] := 120;
  sg.ColWidths[10] := 50;
  sg.ColWidths[11] := 50;
end;

end.
