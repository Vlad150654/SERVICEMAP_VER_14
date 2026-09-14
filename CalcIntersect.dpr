program CalcIntersect;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  Winapi.Windows,
  System.Classes,
  System.SysUtils,
  ComObj,
  Variants,
  Dialogs,
  cfCommon,
  Winapi.Messages,
  MapBasic_INT in '..\MapBasic_INT.pas',
  GeoTransform in '..\GeoTransform.pas',
  Common in '..\Common.pas';

var
  row, rowcount : Integer;
  tab : String;
  cad : String;
  idn : String;
  kontur : Integer;
  wnd, mwnd : HWND;
  mspn, mspk : TPoint3D;//начальна€ и конечна€ точки в местной системе по MapInfo
  pd : Array of TPartitionData;
  area : Array of Double;
  AnyData : Array of TCostFactory;
  procent : Double;
  cn : TStringList;
  CDS : TCopyDataStruct;
  st16 : Boolean;

function GetConsoleWindow: HWND; stdcall; external kernel32;

procedure LoadAnyData(xlsname : String);
var
  ExcelApp, Workbook, WorkSheet : OleVariant;
  i, j, k{, n, link} : Integer;
  s, ftype : String;
begin
  SetLength(AnyData, 0);
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := False;
  try
    ExcelApp.Workbooks.Open(xlsname);
    WorkBook := ExcelApp.Workbooks.item[1];
    WorkSheet := WorkBook.WorkSheets.item[1];

    i := 2;
    Repeat
      if (String(WorkSheet.Cells[i, 4]) = '') or (String(WorkSheet.Cells[i, 6]) = '') then
        break;
      SetLength(AnyData, i - 1);
      Val(String(WorkSheet.Cells[i, 1]), AnyData[i-2].id, k);
      AnyData[i - 2].table := String(WorkSheet.Cells[i, 2]);
      AnyData[i - 2].field := String(WorkSheet.Cells[i, 3]);
      AnyData[i - 2].Alias := String(WorkSheet.Cells[i, 4]);
//      Val(String(WorkSheet.Cells[i, 5]), AnyData[i - 2].CalcType, k);
      AnyData[i - 2].fType := String(WorkSheet.Cells[i, 5]);
      Val(String(WorkSheet.Cells[i, 6]), AnyData[i - 2].Link, k);
      Val(String(WorkSheet.Cells[i, 8]), AnyData[i - 2].gko, k);
      Inc(i);
    Until false;

  finally
    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;
    WorkSheet := Unassigned;
    WorkBook := Unassigned;
    ExcelApp := Unassigned;
  end;
end;
{---------------------определение пересечени€ с зоной--------------------------}
procedure Intersect;
var
  i, j, k, start, stop : Integer;
  area, overlap : Double;
{---------------------определение зоны по вхождению части----------------------}
  procedure CalcZone(tab, field : String);
  begin//границы объекта B частично лежат внутри границ объекта A;
    miDo('Select * from %s Where obj Contains Part ob into TempMO NoSelect', [tab]);
    if miEvalInt('TableInfo(TempMO, %d)', [TAB_INFO_NROWS]) > 0 then
      if (i < High(AnyData)) and (AnyData[i].Table = AnyData[i+1].Table) then
        start := Length(pd);//сдесь будет перва€ запись

    While miEval('EOT(TempMO)') <> 'T' do
    begin
      miDo('ob2 = TempMO.obj');
      if miEval('Str$(ob2)') = 'Region' then
      begin
        area := miEvalDbl('CartesianArea(ob, "sq m")');//площадь участка
        try
          miDo('ob2 = Overlap(ob, ob2)');
          overlap := miEvalDbl('CartesianArea(ob2, "sq m")');//площадь перекрыти€
          k := Length(pd);
          SetLength(pd, k+1);
          pd[k].idn := idn;
          pd[k].id := i;
          pd[k].Table := tab;
          pd[k].Value := miEval('TempMO.%s', [field]);
          pd[k].area := area;
          pd[k].overlap := overlap;
        except
          on e: Exception do showmessage(e.Message);
        end;
      end;
      miDo('Fetch Next From TempMO');
    end;
  end;
{---------------------определение наличи€ газоснабжени€-------------------------}
//  procedure CalcGaz(tab, field : String);
//  var
//    ar1, ar2, overlap : Double;
//  begin
//    miDo('Select * from %s Where ob Contains obj into TempMO NoSelect', [tab]);//участки газа в нашем участке
//    While miEval('EOT(TempMO)') <> 'T' do
//    begin
//      ar1 := miEvalDbl('CartesianArea(ob, "sq m")');//площадь нового участка (от ƒимы)
//      miDo('ob2 = TempMO.obj');
//      ar2 := miEvalDbl('CartesianArea(ob2, "sq m")');//площадь старого участка (_Ќаличие_газа)
//      if (ar1 >= ar2) or (Abs(ar2 - ar1) < 0.000001) then     //row=91453
//      begin
//        miDo('ob2 = Overlap(ob, ob2)');
//        overlap := miEvalDbl('CartesianArea(ob2, "sq m")');
//        if Round(overlap/ar2*100) >= 95 then
//        begin
//          k := Length(pd);
//          SetLength(pd, k+1);
//          pd[k].idn := idn;
//          pd[k].id := i;
//          pd[k].Table := tab;
//          pd[k].Value := miEval('TempMO.%s', [field]);
//          pd[k].area := ar1;
//          pd[k].overlap := overlap;
//        end;
//      end;
//      miDo('Fetch Next From TempMO');
//    end;
//  end;

begin
  i := 0;
  while i <= High(AnyData) do
  begin
//    if AnyData[i].CalcType = 0 then
//    begin
      start := -1;
//      if AnyData[i].Link = 1 then
      if (AnyData[i].Link = 1) and (st16 or (AnyData[i].gko = 1)) then
        CalcZone(AnyData[i].Table, AnyData[i].Field);
//      else
//      if AnyData[i].Link = 2  then CalcGaz(AnyData[i].Table, AnyData[i].Field);
//    end;

    if start > -1 then
    begin
      stop := High(pd);
      miDo('Fetch First From TempMO');
      while (i < High(AnyData)) and (AnyData[i].Table = AnyData[i+1].Table) do
      begin
        for j := start to stop do
        begin
          k := Length(pd);
          SetLength(pd, k+1);
          pd[k] := pd[j];
          pd[k].id := i + 1;
          pd[k].value := miEval('TempMO.%s', [AnyData[i+1].Field]);
        end;
        Inc(i);
      end;
    end;
    SafeCloseTable('TempMO');
    Inc(i);
  end;
end;
{-------------------------------обновление данных в таблице-------------------}
procedure UpdateTable(var ar : Array of String);
var
  i : Integer;
  data : String;
begin
  data := '';
  for i := Low(ar) to High(ar) do
  begin
    begin
      if kontur > 1 then
      begin
        if Pos('Char', AnyData[i].fType) > 0 then
        begin
          if Pos('"', ar[i]) > 0 then
            ar[i] := StringReplace(ar[i], '"', '""', [rfReplaceAll]);
          ar[i] := '"' + ar[i] + '"'
        end else if ar[i] = '' then
          ar[i] := '0';
        data := format('%s%s=%s,', [data, AnyData[i].Alias, ar[i]]);
      end else
      if AnyData[i].Link <> 1 then//это вычисл€тс€ при создании таблицы
      begin
        if Pos('Char', AnyData[i].fType) > 0 then
        begin
          if Pos('"', ar[i]) > 0 then
            ar[i] := StringReplace(ar[i], '"', '""', [rfReplaceAll]);
          ar[i] := '"' + ar[i] + '"'
        end else if ar[i] = '' then
          ar[i] := '0';
        data := format('%s%s=%s,', [data, AnyData[i].Alias, ar[i]]);
      end;
    end;
  end;
  SetLength(data, Length(data)-1);
  if data <> '' then
    miDo('Update %s Set %s Where rowid=%d', [tab, data, row]);
end;
{--------------------------------вычисление средних значений-------------------}
procedure CalcMedia;
var
  i, j, n, k, id : Integer;
  area_s, x, max : Double;
  ar : Array of String;
  sl : TStringList;
begin
  SetLength(ar, Length(AnyData));
  ar[0] := idn;
  ar[1] := cad;
  area_s := 0;
  for i := 0 to High(area) do
    area_s := area_s + area[i];
  if area_s <= 0  then//дл€ точечных объектов
    area_s := 0.001;
  sl := TStringList.Create;
  sl.Sorted := true;
  sl.Duplicates := dupIgnore;
  try
    for n := 0 to High(AnyData) do//отбираем найденные ценовые факторы
    begin
      sl.Clear;
      for j := 0 to High(pd) do
        if pd[j].id = n then//индекс части участка совпадает с ценовым индексом
          sl.AddObject(pd[j].value, Pointer(n+1));
      max := 0;
      id := -1;
      for i := 0 to sl.Count-1 do//провер€ем по всем найденным значени€м                                                               1
      begin
        k := Integer(sl.Objects[i])-1;//индекс массива данных
        x := 0;
        for j := 0 to High(pd) do
         if pd[j].Value = sl[i] then//провер€ем площади вхождени€ по каждой зоне
         begin
            x := x + pd[j].overlap;//накапливаем значени€
            if max < x then//ищем максимальное
            begin
              max := pd[j].overlap;
              id := j;
            end;
         end;
        if x/area_s*100 > procent then
          ar[k] := sl[i];
//if x = 0 then
//Inc(iii);
        if (id > -1) and (Pos('Ќѕ_ћќ', AnyData[n].Table) = 1) then//ищем максимальное
          ar[k] := pd[id].Value;//площадь наибольшего вхождени€
      end;
    end;
    UpdateTable(ar);
//    for j := 0 to High(pd) do
//      Writeln(f, format('%s%s%s%s%d%s%s%s%s%s%.2f%s%.2f',
//        [pd[j].idn,#9,cad,#9,pd[j].id,#9,pd[j].table,#9,pd[j].value,#9,pd[j].area,#9,pd[j].overlap]));
  finally
    SetLength(ar, 0);
    SetLength(area, 0);
    SetLength(pd, 0);
    sl.Free;
  end;
end;

procedure Work(tablename, xlsname, resfolder, proc  : String);
var
  i, k: Integer;
  path, part, cf : String;
begin
  tab := miEval('PathToTableName$("%s")', [tablename]);
  Val(proc, procent, i);
  Val(ParamStr(5), wnd, i);
  st16 := ParamStr(6) = '1';

  part := ExtractFilePath(tablename);
  part := Copy(part, Length(part)-1, 1);
  cf := ExtractFilePath(xlsname);
  cf := Copy(cf, Length(cf)-4, 4);
  path := ExtractFilePath(tablename) + AddSlesh(cf);

  miDo('Open table "%s"', [tablename]);
  miDo('Set CoordSys table %s', [tab]);
  rowcount := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);//размер полной таблицы
  LoadAnyData(xlsname);

  cn.Clear;
  FindRecursive(AddSlesh(path), '*.tab', False, cn);
  for i := 0 to cn.Count-1 do
    miDo('Open table "%s"', [cn[i]]);

  row := 1;
  miDo('Fetch Rec %d From %s', [row, tab]);
  kontur := miEvalInt('%s.„исло_участков', [tab]);
  try
    While row < rowcount do
    begin
      SetLength(area, kontur);
      for k := 1 to kontur do
      begin
        miDo('Fetch Rec %d From %s', [row, tab]);
        cad := miEval('%s. адастровый_номер', [tab]);
        writeln(format('%s  осталось %d', [cad, (rowcount - row)]));
        idn := miEval('%s.idn', [tab]);
        miDo('ob = %s.obj', [tab]);
        area[k-1] := miEvalDbl('CartesianArea(ob, "sq m")');
        mspn.x := miEvalDbl('CentroidX(ob)');
        mspn.y := miEvalDbl('CentroidY(ob)');
        Intersect;
        if k = kontur then
          CalcMedia;
        Inc(row);
      end;
      SetLength(area, 0);

      if row < rowcount then
      begin
        miDo('Fetch Rec %d From %s', [row, tab]);
        kontur := miEvalInt('%s.„исло_участков', [tab]);
      end;
    end;

  finally
    SetLength(AnyData, 0);
//    miDo('select * from %s where центр_ћќ<>"" into Temp', [tab]);
    miDo('select * from %s where число_участков=0 into Temp', [tab]);
    miDo('Commit table Temp as "%smulti_contour%s.tab"', [resfolder, part]);
  end;
end;

//"D:\Tokyo\SERVICEMAP\test\пересечени€_CF_4\1\Intersect.tab" "D:\Tokyo\SERVICEMAP\Win32\Debug\CF_4\CostFactory.xlsx" "D:\Tokyo\SERVICEMAP\test\пересечени€_CF_4\1\CF_4\" 50 1111
//двинин   "D:\Tokyo\SERVICEMAP\test\пересечени€_CF_4\двинин\1\Intersect.tab" "D:\Tokyo\SERVICEMAP\VER_7\Win32\Debug\CF_4\CostFactory.xlsx" "D:\Tokyo\SERVICEMAP\test\пересечени€_CF_4\двинин\1\" 50 1111
begin
  try
//ShowMessage(ParamStr(1)+#10+ParamStr(2)+#10+ParamStr(3)+#10+ParamStr(4));
    if ParamStr(1) = '' then
      ShowMessage('Ёто файл запускаетс€ из ценовых факторов')
    else
    try
      mwnd := GetConsoleWindow;
      SetWindowText(mwnd,ParamStr(2));

      mi := CreateOleObject('MapInfo.Application');
      AppWnd := Mi.Eval('SystemInfo(9)');
      Set_Priority(AppWnd, HIGH_PRIORITY_CLASS);
      miDo('Set ProgressBars Off');
      cn := TStringList.Create;

      miDo('Set ProgressBars Off');
      miDo('Dim ob as Object');
      miDo('Dim ob1 as Object');
      miDo('Dim ob2 as Object');
      miDo('Dim ztb as Object');
      miDo('Dim igs as Object');
      miDo('Dim first as Object');
      miDo('Dim eburg as Object');

      Work(ParamStr(1),ParamStr(2),ParamStr(3), ParamStr(4));

    finally
      cn.Free;
      miDo('UnDim ob');
      miDo('UnDim ob1');
      miDo('UnDim ob2');
      miDo('UnDim ztb');
      miDo('UnDim igs');
      miDo('UnDim first');
      miDo('UnDim eburg');
      miDo('Close All');
      miDo('End MapInfo');
      mi := Unassigned;

      tab := 'RESULT';
      row := Pos('Intersect', ParamStr(1));
      if row > 0 then
        tab := 'RESULT' + Copy(ParamStr(1), row-2, 1);
      CDS.dwData := 0;
      CDS.cbData := (length(tab) + 1) * sizeof(tab);
      CDS.lpData := PChar(tab);
      SendMessage(wnd, WM_COPYDATA, wnd, Integer(@CDS));
    end;
  except
    on E: Exception do  ShowMessage(E.ClassName + ' : ' + E.Message);
  end;
end.

