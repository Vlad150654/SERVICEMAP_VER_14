program CalcSplit;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  Winapi.Windows,
  System.Classes,
  System.SysUtils,
  ComObj,
  Variants,
  Dialogs,
  Winapi.Messages,
  cfCommon,
  MapBasic_INT in '..\MapBasic_INT.pas',
  GeoTransform in '..\GeoTransform.pas',
  Common in '..\Common.pas';

var
  row, max : Integer;
  wnd, mwnd : HWND;
  tab : String;
  scheme : Integer;
  aCopyData : TCopyDataStruct;
  CDS : TCopyDataStruct;

  fields, cn : String;
  sl : TStringList;
  lst : TStringList;

function GetConsoleWindow: HWND; stdcall; external kernel32;
{---------------------определение пересечения----------------------------------}
procedure CalcPlacement;
var
  area, overlap, proc : Double;
  s : String;
  i, j, k : Integer;
begin
  miDo('ob = %s.obj', [tab]);//таблица участки
  area := miEvalDbl('CartesianArea(ob, "sq m")');//площаль участка
  if area = 0 then
    exit;
  //objectA  Partly Within objectB	границы объекта A частично лежат внутри границ объекта B;
  //выбираем части лежащие в границах участка
  miDo('select * from zouit where obj Partly Within ob into temp noSelect');//выбираем из обременений или почв
  k := miEvalInt('TableInfo(temp, %d)', [TAB_INFO_NROWS]);

  for i := 1 to k do//проходим по кусочкам вхождений
  begin
    miDo('Fetch rec %d from temp', [i]);
    if cn <> '-' then
      s := miEval('%s.IDN', [tab]) + #9 + miEval('%s.%s', [tab, cn]) + #9//IDN, кадастровый номер из таблицы
    else
      s := '';
    if scheme = 1 then//ZOUIT
      for j := 0 to sl.Count-1 do
        s := s + miEval('%s.%s', [tab, sl[j]]) + #9//IDN, кадастровый номер из выборки
    else//SOIL
      for j := 0 to sl.Count-1 do
        s := s + miEval('temp.%s', [sl[j]]) + #9;

    try
      miDo('ob2 = Overlap(temp.object, ob)');//вырезаем из tab ЗОУИТ
      overlap := miEvalDbl('CartesianArea(ob2, "sq m")');
    except
      overlap := area;
    end;

    proc := overlap*100/area;
    if proc > 100 then
      proc := 100;
    s := s + FloatToStrF(area, ffFixed, 12, 4) + #9;

    if overlap = area then//участок полностью лежит в зоне
      lst.Add(s + FloatToStrF(area, ffFixed, 12, 4) + #9 + '100')//добавляем всю площадь
    else
      if overlap = 0 then//нет наложений, ничего не делаем
    else
      lst.Add(s + FloatToStrF(overlap, ffFixed, 12, 4) + #9 + FloatToStrF(proc, ffFixed, 12, 4));//добавляем площадь перекрытия

  end;
  miDo('Close table temp');
end;

//Почвы: "D:\Tokyo\SERVICEMAP\test\Почвы_Intersect_CF_4\1\mini.tab" "D:\Tokyo\SERVICEMAP\test\Почвы_Intersect_CF_4\1\soil_map_M2_5-1.0 с семантикой.TAB" "soil0_txt,Номер_подзоны" "КН" 1511698 2
//ZOUIT: "D:\Tokyo\SERVICEMAP\test\ZOUIT_Intersect_CF_4\1\mini.TAB" "D:\Tokyo\SERVICEMAP\test\ZOUIT_Intersect_CF_4\1\Дополнительные_затраты.TAB" "IDN,Кн" "-" 1511698 1
begin
  if ParamStr(1) = '' then
    ShowMessage('Это файл запускается из ценовых факторов')
  else
  try
    mwnd := GetConsoleWindow;
    SetWindowText(mwnd,ParamStr(2));
    mi := CreateOleObject('MapInfo.Application');
    AppWnd := Mi.Eval('SystemInfo(9)');
    Set_Priority(AppWnd, HIGH_PRIORITY_CLASS);
    miDo('Set ProgressBars Off');
    miDo('Dim ob as Object');
    miDo('Dim ob2 as Object');
    sl := TStringList.Create;
    lst := TStringList.Create;

    fields := ParamStr(3);
    sl.CommaText := fields;
    cn := ParamStr(4);
    miDo('Open table "%s"', [ParamStr(1)]);//часть таблицы участков
    tab := miEval('PathToTableName$("%s")', [ParamStr(1)]);
    miDo('Set CoordSys table %s', [tab]);
    Val(ParamStr(5), wnd, row);
    Val(ParamStr(6), scheme, row);//scheme=1 ZOUIT; scheme=2 SOIL
//showmessage(ParamStr(1)+#10+ParamStr(2)+#10+ParamStr(3)+#10+ParamStr(4)+#10+ParamStr(5)+#10+ParamStr(6));

    miDo('Open table "%s" as zouit', [ParamStr(2)]);//таблица обременений
    max := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
    writeln(format('осталось %d', [max]));
    for row := 1 to max do
    begin
      miDo('Fetch rec %d from %s', [row, tab]);
      CalcPlacement;
      writeln(format('осталось %d', [max-row]));
    end;

    miDo('Close table zouit');
    lst.SaveToFile(ExtractFilePath(ParamStr(1)) + 'aaa.txt');

  finally
    sl.Free;
    lst.Free;
//    miDo('Drop table %s', [tab]);
    miDo('UnDim ob');
    miDo('UnDim ob2');
    miDo('Close All');
    miDo('End MapInfo');
    mi := Unassigned;

    CDS.dwData := 0;
    CDS.cbData := (length(ParamStr(2)) + 1) * sizeof(ParamStr(2));
    CDS.lpData := PChar(ParamStr(2));
    SendMessage(wnd, WM_COPYDATA, wnd, Integer(@CDS));
  end;
end.
