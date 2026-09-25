unit Letters_Unit;

interface

uses Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls,
  ComObj,
  System.RegularExpressions,
  System.RegularExpressionsCore,
  System.IniFiles,
  StrUtils,
  MapBasic_INT, Common, Data.DB, Data.Win.ADODB, Vcl.Buttons;

const
  strConnect = 'Provider=SQLOLEDB.1;Password=s;Persist Security Info=True;'+
    'User ID=zisadmin;Initial Catalog=%s;Data Source=TESTSQL;'+
    'Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;'+
    'Workstation ID=G001;Use Encryption for Data=False;Tag with column collation when possible=False';

type
  PLettersObj = ^TLettersObj;
  TLettersObj = Record
    tip : String[35];
    foundationdate  : String[12];
    count : Integer;
    filename : String[250];
  end;

  TLettersForm = class(TForm)
    btCalc: TButton;
    OpenDialog1: TOpenDialog;
    stBar: TStatusBar;
    leNumberInputinfo: TLabeledEdit;
    leDateInputinfo: TLabeledEdit;
    btPacket: TButton;
    btToBase: TButton;
    ADOConnection1: TADOConnection;
    Query: TADOQuery;
    CreateCSV: TBitBtn;
    LabeledEdit1: TLabeledEdit;
    procedure btCalcClick(Sender: TObject);
    procedure btPacketClick(Sender: TObject);
    procedure btToBaseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure CreateCSVClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    BaseName : String;
//    ds : TADODataset;
    ListObjects : TStringList;
    sl : TStringList;
    ZipFile : String;
    MaxID : LongInt;
    Reg: TRegEx;
    ArchiveCount : Integer;
    path, packet, cat, NumberInputinfo, DateInputinfo, temp : String;
    procedure Info(Msg: String);
    procedure UnPack(sourcepath, destpath : String);
    function RunCommand(var cmd : String) : Boolean;
    procedure UpdatePrimXLS(var cmd : String); overload;
    procedure UpdatePrim(var cmd : String); overload;
    procedure UpdatePrim(var sl : TStringList; var cmd : String); overload;
    procedure ReplacePrim(var cmd : String);
  end;

var
  LettersForm: TLettersForm;

implementation

{$R *.dfm}

function SortInfo(Item1, Item2: PLettersObj): Integer;
begin
  if Item1.tip < Item2.tip then
    Result := -1
  else if Item1.tip > Item2.tip then
    Result := 1
  else
  if Item1.foundationdate < Item2.foundationdate then
    Result := -1
  else if Item1.foundationdate > Item2.foundationdate then
    Result := 1
//  else
//    if Item1.count < Item2.count then
//      Result := -1
//    else if Item1.count > Item2.count then
//      Result := 1
  else
  if Item1.filename < Item2.filename then
    Result := -1
  else if Item1.filename > Item2.filename then
    Result := 1
  else
    Result := 0;
end;
{------------------------------------------------------------------------------}
procedure TLettersForm.Info(Msg: String);
begin
  stBar.Panels[0].Text := Msg;
  stBar.Repaint;
  Application.ProcessMessages;
end;
{------------------------------------------------------------------------------}
procedure TLettersForm.UnPack(sourcepath, destpath : String);
var
  i : Integer;
  s : String;
  sl : TStringList;
begin
  sl := TStringList.Create;
  try
    FindRecursive(sourcepath, '*.7z', false, sl);//ишем 7z
    FindRecursive(sourcepath, '*.zip', false, sl);//добавл¤ем zip
    ArchiveCount := ArchiveCount + sl.Count;
    for i := 0 to sl.Count-1 do//извлекаем все архивы
    begin
      info('...' + Copy(sl[i], Length(sl[i])-55, Length(sl[i])));
      s := ChangeFileExt(ExtractFileName(sl[i]), '.mif');
      if not FileExists(destpath + s) then
        ExCreateProcess(format('"%s" x "%s" -y -o"%s"',
          [ZipFile, sl[i], destpath]), '', SW_HIDE, true);
      Application.ProcessMessages;
    end;
//    sl.Clear;//если архир распакован в этот же каталог повторяем
//    FindRecursive(sourcepath, '*.zip', false, sl);//добавл¤ем zip
//    for i := 0 to sl.Count-1 do//извлекаем все архивы
//    begin
//      info('...' + Copy(sl[i], Length(sl[i])-45,  Length(sl[i])));
//      s := ChangeFileExt(ExtractFileName(sl[i]), '.mif');
//      if not FileExists(destpath + s) then
//        ExCreateProcess(format('"%s" x "%s" -y -o"%s"',
//          [ZipFile, sl[i], destpath]), '', SW_HIDE, true);
//      Application.ProcessMessages;
//    end;
    sl.Clear;
    FindFolders(destpath, sl, false);
    for i := 0 to sl.Count-1 do//ищем архивные файлы
      UnPack(AddSlesh(sl[i]), AddSlesh(sl[i]));
  finally
    sl.Free;
  end;
end;
{------------------------------------------------------------------------------}
procedure TLettersForm.btPacketClick(Sender: TObject);
begin
{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\Outbox_Статья_16_2024\2024_04_04\Письма\';{$ENDIF}  //1124,1125...
//{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\Ошибка кн\Первичная обработка писем';{$ENDIF}  //2525
//{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\перечни от 31.08.2026\';{$ENDIF}  //2314,2315...

//  if path = '' then
//    path := 'W:\ЦГКО\IT\Outbox\Статья 16\2024';
  if not GetFolder(path, Handle) then exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  path := DellSlesh(path);//удалим последний слеш

  packet := ExtractFileDir(path);
  packet := StringReplace(packet, '\Письма', '', []);
  packet := ExtractFileName(packet);//дата письма
  packet := AddSlesh(GetAppFolder + 'Письма\' + packet);
  ForceDirectories(packet);//каталог для сохранения расчетов

  NumberInputinfo := ExtractFileName(path);
  leNumberInputinfo.Text := NumberInputinfo;
end;

procedure TLettersForm.btCalcClick(Sender: TObject);
var
  i, j, n, m, ver, matchcount : Integer;
  sl,  csv, comp : TStringList;
  s, s2, ss, cn, xml, mifmid, catzem, objecttype, foundationdate, newxml : String;
  log : TStringList;
  mc, fd, fdcont : TMatchCollection;
  res : TList;
  pInfo, predinfo : PLettersObj;
  ExcelApp : OleVariant;
  sa : StringArray;
begin
//{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\перечни от 31.08.2026\';
//  leDateInputinfo.Text := '31.08.2026';
//  leNumberInputinfo.Text := '2314';
//{$ENDIF}

  if leDateInputinfo.Text = '' then//21.09.2026
    Exc('Дата письма не определена!');
  if leNumberInputinfo.Text = '' then//2314
    Exc('Номер письма не определен!');

  DateInputinfo := leDateInputinfo.Text;
  NumberInputinfo := leNumberInputinfo.Text;
  cat := packet + NumberInputinfo;

  if IsExcelFileOpen(AddSlesh(cat) + 'Группировка.csv') then
    Exc('Файл Excel "Группировка" сейчас открыт и занят другом процессом!');

  ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [cat])), SW_HIDE, true, 0);//удаление каталога со всем содержимым и подкаталогами
  ForceDirectories(AddSlesh(cat));

  xml := cat + '\xml';
  ForceDirectories(AddSlesh(xml));
  mifmid := cat + '\mifmid';
  ForceDirectories(AddSlesh(mifmid));

  sl := TStringList.Create;
  csv := TStringList.Create;
  comp := TStringList.Create;
  res := TList.Create;
  log := TStringList.Create;
  try
    ArchiveCount := 0;
    temp := GetSpecialPath(35) + '\letters';//C:\ProgramData
    ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [temp])), SW_HIDE, true, 0);//удаление каталога со всем содержимым и подкаталогами
    ForceDirectories(AddSlesh(temp));
    ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s\*.zip" "%s"',
     [path, temp])), SW_HIDE, true, 0);
    UnPack(AddSlesh(temp), AddSlesh(temp));

    if ArchiveCount = 0 then
      Exc('Нет архивных данных');

    info('Копирование xml...');
    sl.Clear;
    FindRecursive(AddSlesh(temp), '*.xml', true, sl);
    if sl.Count = 0 then
      Exc('Нет файлов XML');

    for i := 0 to sl.Count-1 do
    begin
      j := 0;
      newxml := AddSlesh(xml) + ExtractFileName(sl[i]);
      while FileExists(newxml) do
      begin
        Inc(j);
        newxml := format('%sкопия %d %s', [AddSlesh(xml), j, ExtractFileName(sl[i])]);
      end;
      CopyFile(PWideChar(sl[i]), PWideChar(newxml), false);
    end;

    info('Копирование mif...');
    sl.Clear;
    FindRecursive(AddSlesh(temp), '*.mif', true, sl);
    for i := 0 to sl.Count-1 do
      CopyFile(PWideChar(sl[i]), PWideChar(AddSlesh(mifmid) + ExtractFileName(sl[i])), false);

    info('Копирование mid...');
    sl.Clear;
    FindRecursive(AddSlesh(temp), '*.mid', true, sl);
    for i := 0 to sl.Count-1 do
      CopyFile(PWideChar(sl[i]), PWideChar(AddSlesh(mifmid) + ExtractFileName(sl[i])), false);

    ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [temp])), SW_HIDE, false, 0);//удаление каталога со всем содержимым и подкаталогами
    sl.Clear;
    FindRecursive(AddSlesh(xml), '*.xml', true, sl);
    for i := 0 to sl.Count-1 do//обработка XML
    begin
      Info('...' + Copy(sl[i], Length(sl[i])-55, Length(sl[i])));
      comp.LoadFromFile(sl[i]);
      s := '';
      for j := 0 to comp.Count-1 do
        s := s + Trim(comp[j]);//все в одну строку
      n := Pos('<Objects>', s);
      n := PosEx('<', s, n+12);//добавляем длину больше длины <Objects> и сразу находим начало catzem
      m := PosEx(' ', s, n+1);
      catzem := Copy(s, n+1, m-n-1);
      //добавим перевод строк для правильной работы Matches
      s2 := format('</%s>', [catzem]);
      s := StringReplace(s, s2, s2+#10, [rfReplaceAll]);
      mc := reg.Matches(s, '<ObjectType>\d\d\d\d\d\d\d\d\d\d\d\d</ObjectType>', [roIgnoreCase, roMultiLine]);
      objecttype:= Copy(mc.Item[0].Value, 13, 12);

      //запрос без кода области 66
      s2 := format('<%s CadastralNumber="\d\d:\d\d:\d{1,7}:\d{1,}".+</%s>',
        [catzem, catzem]);

      //проверяем версию
      mc := reg.Matches(s, '<ListForRating Version="04', [roIgnoreCase, roMultiLine]);
      if mc.count = 1 then
      begin
        ver := 4;
        mc := reg.Matches(sl[0], '\d\d\.\d\d\.\d\d\d\d', [roIgnoreCase, roMultiLine]);
        if mc.count > 0 then//вытаскиваем дату основания включения в перечень из названия файла
          foundationdate := mc.Item[0].Value;//Дата основания = DATEFORM
        mc := reg.Matches(s, s2, [roIgnoreCase, roMultiLine]);
        //запрашиваем подтверждение даты основания включения в переченъ
        if not InputQuery(format('%s (%d объектов)', [ExtractFileName(sl[i]), mc.Count]),
          'Введите дату основания включения в перечень', foundationdate) then
            exit;
      end else
      begin
        ver := 5;//она же 6
        fd := reg.Matches(s, 'FoundationDate="\d\d\d\d-\d\d-\d\d', [roIgnoreCase, roMultiLine]);
        if fd.Count = 0 then
        begin
          ver := 4;
          mc := reg.Matches(sl[0], '\d\d\.\d\d\.\d\d\d\d', [roIgnoreCase, roMultiLine]);
          if mc.count > 0 then//вытаскиваем дату основания включения в перечень из названия файла
            foundationdate := mc.Item[0].Value;//Дата основания = DATEFORM
          mc := reg.Matches(s, s2, [roIgnoreCase, roMultiLine]);
          //запрашиваем подтверждение даты основания включения в переченъ
          if not InputQuery(format('%s (%d объектов)', [ExtractFileName(sl[i]), mc.Count]),
            'Введите дату основания включения в перечень', foundationdate) then
              exit;
        end else
          mc := reg.Matches(s, s2, [roIgnoreCase, roMultiLine]);
      end;

      csv.Clear;//заполняем файл CSV
      csv.Add('CadastralNumber;ObjectType;DATEFORM;NUMBERINPUTINFO;DATEINPUTINFO');
      matchcount := mc.count;
      ss := ChangeFileExt('...' + Copy(sl[i], Length(sl[i])-40, Length(sl[i])), '.csv');
      for j := 0 to mc.count-1 do
      begin
        s2 := mc.Item[j].Value;
        n := Pos('"', s2);
        m := PosEx('"', s2, n + 1);
        cn := Copy(s2, n+1, m-n-1);
        Info(ss + ' ' + cn);

        //проверка КН
        if Copy(cn, 1, 2) <> '66' then
          log.Add('Плохой кадастровый номер '+cn+' в файле'+#10+sl[i]);

        if ver = 5 then//она же 6
        begin
          s2 := Copy(fd.Item[j].Value, 17, 10);
          fdcont := reg.Matches(s, 'FoundationDate="' + s2, [roIgnoreCase, roMultiLine]);
          sa := SplitStr(s2, '-');
          foundationdate := sa[2] + '.' + sa[1] + '.' + sa[0];
          SetLength(sa, 0);
        end;

        csv.Add(format('%s;%s;%s;%s;%s',
          [cn, ObjectType, foundationdate, NumberInputinfo, DateInputinfo]));

        pInfo := AllocMem(SizeOf(pInfo^));
        res.add(pInfo);
        pInfo.tip := catzem;
        pInfo.foundationdate := foundationdate;
        if ver = 5 then//она же 6
          pInfo.count := fdcont.Count
        else
          pInfo.count := mc.count;
        pInfo.filename := ExtractFileName(sl[i]);
      end;

      if matchcount = 0 then
      begin
        DeleteFile(sl[i]);
        log.Add('Удален файл '+ sl[i]);
      end else
        csv.SaveToFile(ChangeFileExt(sl[i], '.csv'));
    end;

    if res.Count > 0 then
    begin
      res.Sort(@SortInfo);
      comp.Clear;
      comp.Add('Тип объекта;Дата основания;количество объектов;файл');
      predinfo := res[0];
      comp.Add(format('%s;%s;%d;%s',
        [predinfo.tip, predinfo.foundationdate, predinfo.count,predinfo.filename]));
      for i := 1 to res.Count-1 do
      begin
        pInfo := res[i];
        if (pInfo.tip <> predinfo.tip) or (pInfo.foundationdate <> predinfo.foundationdate) or
          (pInfo.count <> predinfo.count) or (pInfo.filename <> predinfo.filename) then
         begin
          comp.Add(format('%s;%s;%d;%s', [pInfo.tip, pInfo.foundationdate, pInfo.count, pinfo.filename]));
          predinfo := res[i];
         end;
      end;
      comp.SaveToFile(AddSlesh(cat) + 'Группировка.csv');

      ExcelApp := CreateOleObject('Excel.Application');
      ExcelApp.Workbooks.Open(AddSlesh(cat) + 'Группировка.csv');
      ExcelApp.Visible := true;
      ExcelApp := Unassigned;
      Info('Готово.');
      Application.ProcessMessages;

    end else
      Note('Нет объектов для группировки');

  finally
    Info('Готово.');
    log.SaveToFile(AddSlesh(xml) + 'logError.log');
    if log.Count > 0 then
      ExShellExecute('Notepad.exe', AddSlesh(xml) + 'logError.log', SW_NORMAL, false, 0)
    else
      DeleteFile(AddSlesh(xml) + 'logError.log');
    log.Free;
//    leNumberInputinfo.Text := '';
    for i := 0 to res.Count-1 do
      Dispose(res.Items[i]);
    res.Free;
    sl.Free;
    csv.Free;
    comp.Free;
  end;
  Application.ProcessMessages;
end;

procedure TLettersForm.CreateCSVClick(Sender: TObject);
var
  i, j, n, m, ver, matchcount : Integer;
  sl,  csv, comp : TStringList;
  s, s2, ss, cn, xml, mifmid, catzem, objecttype, foundationdate, newxml : String;
  log : TStringList;
  mc, fd, fdcont : TMatchCollection;
  res : TList;
  pInfo, predinfo : PLettersObj;
  ExcelApp : OleVariant;
  sa : StringArray;
begin
  if leDateInputinfo.Text = '' then
    Exc('Дата письма не определена!');//31.08.2026
  if leNumberInputinfo.Text = '' then
    Exc('Номер письма не определен!');//2314
  if LabeledEdit1.Text = '' then
    Exc('foundationdate письма не определен!');//2026-08-29

 // {$IFDEF DEBUG}path := 'W:\ЦГКО\IT\___Программы\Дежурная карта_VER_14\Письма\14.03.2025\508\xml';{$ENDIF}  //2525
  if not GetFolder(path, Handle) then exit;//D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\Письма\перечни от 31.08.2026\2317\xml\
  xml := path;

  DateInputinfo := leDateInputinfo.Text;
  NumberInputinfo := leNumberInputinfo.Text;
  sl := TStringList.Create;
  csv := TStringList.Create;
  comp := TStringList.Create;
  res := TList.Create;
  log := TStringList.Create;
//  cat := ExtractFilePath(DellSlesh(path));
  try
    FindRecursive(AddSlesh(xml), '*.xml', true, sl);
    for i := 0 to sl.Count-1 do//обработка XML
    begin
      Info('...' + Copy(sl[i], Length(sl[i])-55, Length(sl[i])));
      comp.LoadFromFile(sl[i]);
      s := '';
      for j := 0 to comp.Count-1 do
        s := s + Trim(comp[j]);//все в одну строку
      n := Pos('<Objects>', s);
      n := PosEx('<', s, n+12);//добавляем длину больше длины <Objects> и сразу находим начало catzem
      m := PosEx(' ', s, n+1);
      catzem := Copy(s, n+1, m-n-1);
      //добавим перевод строк для правильной работы Matches
      s2 := format('</%s>', [catzem]);
      s := StringReplace(s, s2, s2+#10, [rfReplaceAll]);
      mc := reg.Matches(s, '<ObjectType>\d\d\d\d\d\d\d\d\d\d\d\d</ObjectType>', [roIgnoreCase, roMultiLine]);
      objecttype:= Copy(mc.Item[0].Value, 13, 12);

      //запрос без кода области 66
      s2 := format('<%s CadastralNumber="\d\d:\d\d:\d{1,7}:\d{1,}".+</%s>',
        [catzem, catzem]);

      ver := 4;//вообще то версия 5 но "foundationdate" нет в файле (В ВЕРСИИ 6 ЕСТЬ!!!)
      foundationdate := LabeledEdit1.Text;//2026-08-29
      mc := reg.Matches(s, s2, [roIgnoreCase, roMultiLine]);

      csv.Clear;//заполняем файл CSV
      csv.Add('CadastralNumber;ObjectType;DATEFORM;NUMBERINPUTINFO;DATEINPUTINFO');
      matchcount := mc.count;
      ss := ChangeFileExt('...' + Copy(sl[i], Length(sl[i])-40, Length(sl[i])), '.csv');
      for j := 0 to mc.count-1 do
      begin
        s2 := mc.Item[j].Value;
        n := Pos('"', s2);
        m := PosEx('"', s2, n + 1);
        cn := Copy(s2, n+1, m-n-1);
        Info(ss + ' ' + cn);

        //проверка КН
        if Copy(cn, 1, 2) <> '66' then
          log.Add('Плохой кадастровый номер '+cn+' в файле'+#10+sl[i]);
        csv.Add(format('%s;%s;%s;%s;%s',
          [cn, ObjectType, foundationdate, NumberInputinfo, DateInputinfo]));

        pInfo := AllocMem(SizeOf(pInfo^));
        res.add(pInfo);
        pInfo.tip := catzem;
        pInfo.foundationdate := foundationdate;
        if ver = 5 then//она же 6
          pInfo.count := fdcont.Count
        else
          pInfo.count := mc.count;
        pInfo.filename := ExtractFileName(sl[i]);
      end;

      if matchcount = 0 then
      begin
        DeleteFile(sl[i]);
        log.Add('Удален файл '+ sl[i]);
      end else
        csv.SaveToFile(ChangeFileExt(sl[i], '.csv'));
    end;

//    if res.Count > 0 then
//    begin
//      res.Sort(@SortInfo);
//      comp.Clear;
//      comp.Add('Тип объекта;Дата основания;количество объектов;файл');
//      predinfo := res[0];
//      comp.Add(format('%s;%s;%d;%s',
//        [predinfo.tip, predinfo.foundationdate, predinfo.count,predinfo.filename]));
//      for i := 1 to res.Count-1 do
//      begin
//        pInfo := res[i];
//        if (pInfo.tip <> predinfo.tip) or (pInfo.foundationdate <> predinfo.foundationdate) or
//          (pInfo.count <> predinfo.count) or (pInfo.filename <> predinfo.filename) then
//         begin
//          comp.Add(format('%s;%s;%d;%s', [pInfo.tip, pInfo.foundationdate, pInfo.count, pinfo.filename]));
//          predinfo := res[i];
//         end;
//      end;
//      comp.SaveToFile(AddSlesh(cat) + 'Группировка.csv');
//
//      ExcelApp := CreateOleObject('Excel.Application');
//      ExcelApp.Workbooks.Open(AddSlesh(cat) + 'Группировка.csv');
//      ExcelApp.Visible := true;
//      ExcelApp := Unassigned;
//      Info('Готово.');
//      Application.ProcessMessages;
//    end else
//      Note('Нет объектов для группировки');

  finally
    Info('Готово.');
    log.SaveToFile(AddSlesh(xml) + 'logError.log');
    if log.Count > 0 then
      ExShellExecute('Notepad.exe', AddSlesh(xml) + 'logError.log', SW_NORMAL, false, 0)
    else
      DeleteFile(AddSlesh(xml) + 'logError.log');
    log.Free;
    csv.Free;
    sl.Free;
    comp.Free;
    res.Free;
  end;
end;

function TLettersForm.RunCommand(var cmd : String) : Boolean;
begin
  Result := true;
  try
    ADOConnection1.BeginTrans;
    ADOConnection1.Execute(cmd);
//ADOConnection1.RollbackTrans;
    ADOConnection1.CommitTrans;
  except
    on e : Exception do
    begin
      ADOConnection1.RollbackTrans;
      Result := false;
      Note(e.Message);
    end;
  end;
end;

(*
foundationdate = DATEFORM Дата основания пакета
DATEINPUTINFO дата входящего письма
NUMBERINPUTINFO номер входяшего письма*)

procedure TLettersForm.btToBaseClick(Sender: TObject);
var
  ExcelApp, Workbook, WorkSheet : OleVariant;
  csv, sl : TStringList;
  sa : StringArray;
  s, cmd, PRIM, DATALOAD, NumberInputinfo, DateInputinfo : String;
  i, j, k : longint;

  procedure Command_Values;
  begin
    //описание полей и значений для первой записи
    cmd := format('Insert into zisadmin.ZAPROS16ST ('+
    'CADASTRALNUMBER,'+//varchar(30)
    'ObjectType,'+//int
    'TYPE_OBJ,'+//varchar(100)
    'DATEFORM,'+//date  Дата основания пакета = foundationdate
    'NUMBERINPUTINFO,'+//varchar(255) номер входяшего письма
    'DATEINPUTINFO,'+//date дата входящего письма
    'PRIM,'+//varchar(255)
    'QUESTION_ID) '+//int = MaxID      //будет незаполнено в excel
    'values(''%s'', %s, ''%s'', ''%s'', ''%s'', ''%s'', ''В ОЖИДАНИИ ОБРАБОТКИ'', %d)',
    [sa[0], sa[1], ListObjects.Values[sa[1]], sa[2], sa[3], sa[4], MaxID]);
  end;

begin
//D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\Письма\2024_04_04\1124\Группировка.csv
//D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\Письма\2024_04_04\1236
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\Письма\перечни от 31.08.2026\2314\'; {$ENDIF}

  try
    ExcelApp := GetActiveOleObject('Excel.Application');
  except
    on e: Exception do s := e.Message + #10;
  end;
  if VarIsEmpty(ExcelApp) then
    Exc(s + 'Нет открытого объекта Excel');

  csv := TStringList.Create;
  sl := TStringList.Create;
  try
    k := 0;
    for i := 1 to ExcelApp.Workbooks.count do
    begin
      Workbook := ExcelApp.Workbooks.item[i];
      if (CText(Workbook.name, 'Группировка.csv')) and
        FindSheetByName(Workbook, WorkSheet, 'Группировка') then
        k := i;
    end;
    if k = 0 then
      Exc('Лист "Группировка" в объекте Workbooks не найден');

    Workbook := ExcelApp.Workbooks.item[k];
    path := ExtractFilePath(Workbook.FullName);
    s := ExtractFileName(ExtractFileDir(Workbook.FullName)) + '_';

    FindRecursive(path + 'xml\', '*.csv', false, csv);
    if csv.Count = 0 then
      Exc('Нет файлов CSV в каталоге' + #10 + path + 'xml\');

    if IsExcelFileOpen(path + s + 'ZAPROS16ST.xlsx') then
      Exc('Файл Excel "'+s+'ZAPROS16ST" сейчас открыт и занят другом процессом!');

    ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s*.csv" "%sAllCSV.csv"',//объединяем все CSV
      [path + 'xml\', path])), SW_HIDE, true, 0);
    csv.LoadFromFile(path + 'AllCSV.csv');
    for i := csv.Count-1 downto 1 do//удаляем заголовки
      if Pos('CadastralNumber', csv[i]) > 0 then
        csv.Delete(i);
    csv.SaveToFile(path + 'AllCSV.csv');

    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;
    ExcelApp.DisplayAlerts := false;//отключаем сообщения для перезаписи
    ExcelApp.Workbooks.Open(AddSlesh(path) + 'AllCSV.csv');
    Workbook := ExcelApp.Workbooks.item[1];
    WorkSheet := ExcelApp.ActiveWorkbook.ActiveSheet;
    WorkBook.SaveAs(Filename:=path + s + 'ZAPROS16ST.xlsx',FileFormat:=xlOpenXMLWorkbook,CreateBackup:=False);

    WorkSheet.Columns.Item['C'].Insert;
    WorkSheet.Cells[1, 3] := 'TYPE_OBJ';
    WorkSheet.Cells[1, 7] := 'PRIM';
    WorkSheet.Cells[1, 8] := 'QUESTION_ID';
    WorkSheet.range[format('G2:G%d', [csv.Count])].Select;
    ExcelApp.Selection.Value := 'В ОЖИДАНИИ ОБРАБОТКИ';

    WorkBook.Save;

    DeleteFile(AddSlesh(path) + 'AllCSV.csv');

//    //находим MaxID
//    ds.Close;
//    ds.CommandText := 'select top 1 * from zisadmin.ZAPROS16ST ORDER BY QUESTION_ID DESC';
//    ds.Open;
//    MaxID := ds.FieldByName('QUESTION_ID').AsInteger;
//
//    //запрос на общее количество MaxID
//    ds.Close;
//    ds.CommandText :=  format('select * from zisadmin.ZAPROS16ST'+
//      ' Where question_id = %d', [MaxID]);
//    ds.Open;
//    k := ds.RecordCount;
//
//    //запрос на добавление записей к MaxID (все "В ОЖИДАНИИ ОБРАБОТКИ")
//    ds.Close;
//    ds.CommandText :=  format('select * from zisadmin.ZAPROS16ST'+
//      ' Where (question_id=%d) AND (PRIM = ''В ОЖИДАНИИ ОБРАБОТКИ'')', [MaxID]);
//    ds.Open;
//
//    if k = ds.RecordCount then
//    begin
//      if MessageDlg(format('Пакет %d в ожидании обработки %s Добавить записи в пакет?',
//        [MaxID, #10]), mtWarning, mbOKCancel, 0) <> mrOK then
//          exit;
//    end else
//
//    //запрос на формирование нового пакета
//    begin
//      ds.Close;
//      ds.CommandText :=  format('select * from zisadmin.ZAPROS16ST'+
//        ' Where (question_id=%d) AND ((NOT answer_id is Null) OR'+
//        ' ((answer_id is Null) AND (PRIM = ''В ОБРАБОТКЕ'')))', [MaxID]);
//      ds.Open;
//      if k = ds.RecordCount then
//      begin
//        if MessageDlg(format('Пакет %d завершен %s Сформировать новый пакет?',[MaxID, #10]),
//          mtWarning, mbOKCancel, 0) <> mrOK then
//            exit;
//        Inc(MaxID);
//      end else
//        Exc('Нет записей удовлетворяющих условиям запроса');
//    end;
//
/////=========================== заполнение ZAPROS16ST ===========================
//    for i := 0 to csv.Count-1 do
//    begin
//      info(ExtractFileName(csv[i]));
//      sl.LoadFromFile(csv[i]);
//      sa := SplitStr(sl[1], ';');//заголовок пропускаем
//      Command_Values;
//      SetLength(sa, 0);
//      //описание значений для остальных записей
//      for j := 2 to sl.Count-1 do//продолжаем со второй строки
//      begin
//        sa := SplitStr(sl[j], ';');//CadastralNumber,ObjectType,DATEFORM,NUMBERINPUTINFO,DATEINPUTINFO
//        NumberInputinfo := sa[3];//для вставки в LIST_PACKAGE
//        DateInputinfo := sa[4];//для вставки в LIST_PACKAGE
//        k := j + 1;
//        if k mod 900 = 0 then//максимум 1000 строк в запросе
//        begin
//          if not RunCommand(cmd) then
//            exit;
//          Command_Values;
//        end else
//          cmd := format('%s,(''%s'', %s, ''%s'', ''%s'', ''%s'', ''%s'', ''В ОЖИДАНИИ ОБРАБОТКИ'', %d)',
//          [cmd, sa[0], sa[1], ListObjects.Values[sa[1]], sa[2], sa[3], sa[4], MaxID]);
//        SetLength(sa, 0);
//      end;
//      if not RunCommand(cmd) then
//        exit;
//    end;
/////=========================== заполнение LIST_PACKAGE =========================
//    ds.Close;
//    ds.CommandText := 'select Top 1 * from zisadmin.LIST_PACKAGE ORDER BY ID DESC';
//    ds.Open;
//    k := ds.FieldByName('ID').AsInteger;
//
//    Query.sql.Clear;
//    if k < MaxID then
//    begin//новая запись
//      Query.sql.Add(format('Insert into zisadmin.LIST_PACKAGE ('+
//      'ID,'+
//      'DATALOAD,'+//date
//      'TYPE_PACKAGE,'+//varchar(100)
//      'PRIM)'+//varchar(255)
//      ' values(%d, ''%s'', %s, ''%s'')',
//      [MaxID, DateInputinfo, '16', NumberInputinfo]));
//      Query.ExecSQL;
//    end else
//    if k = MaxID then
//    begin//обновлени записи
//      PRIM := ds.FieldByName('PRIM').AsString;
//
//      if Length(PRIM) = 0 then
//        PRIM := NumberInputinfo
//      else
//      begin
//        if PRIM[Length(PRIM)] = ',' then
//          SetLength(PRIM, Length(PRIM) - 1);
//        PRIM := PRIM + ',' + NumberInputinfo;
//      end;
//
//      Query.sql.Add(format('update zisadmin.LIST_PACKAGE SET PRIM=''%s'''+
//        ' Where id=%d', [PRIM, MaxID]));
//      Query.ExecSQL;
//    end else
//      Note('Таблица LIST_PACKAGE не была изменена!');

    info('Готово');
  finally
    csv.Free;
    sl.Free;
    ExcelApp := Unassigned;
    Workbook := Unassigned;
    WorkSheet := Unassigned;
  end;
end;

procedure TLettersForm.UpdatePrimXLS(var cmd : String);
var
  Query: TADOQuery;
//'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\ZAPROS16ST.xlsx;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=0;READONLY=0;';
begin//kn, DATEINmPUTINFO, NUMBERINPUTINFO
  Query := TADOQuery.Create(nil);
  try
    Query.ConnectionString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\ZAPROS16ST.xlsx;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=0;READONLY=0;';
    Query.sql.Add(format(cmd,  [MaxID]));
    Query.ExecSQL;
  finally
    Query.Free;
  end;
end;

procedure TLettersForm.UpdatePrim(var cmd : String);
begin//kn, DATEINmPUTINFO, NUMBERINPUTINFO
  Query.sql.Clear;
  Query.sql.Add(format(cmd,  [MaxID]));
  Query.ExecSQL;
end;

procedure TLettersForm.UpdatePrim(var sl : TStringList; var cmd : String);
var
  i, k : Integer;
  s : String;
begin
  Query.sql.Clear;
  Query.sql.Add('CREATE TABLE #temcadnum (cn VARCHAR(30))');

  s := format('Insert into #temcadnum (cn) values(''%s'')', [sl[0]]);
  for i := 1 to sl.Count-1 do
  begin
    k := i + 1;
    if k mod 900 = 0 then
    begin
      Query.sql.Add(s);
      s := format('Insert into #temcadnum (cn) values(''%s'')', [sl[i]])
    end else
      s := format('%s,(''%s'')', [s, sl[i]]);
  end;
  Query.sql.Add(s);
  cmd := cmd + ' and (CADASTRALNUMBER=any(Select cn from #temcadnum))';
  Query.sql.Add(format(cmd,  [MaxID]));
  Query.ExecSQL;
end;

procedure TLettersForm.ReplacePrim(var cmd : String);
var
  id : LongInt;
begin
//  ds.Close;
//  ds.CommandText := format(cmd, [MaxID]);
//  ds.Open;
//  if ds.IsEmpty then Exit;
//
//  if not ds.FindFirst then exit;
//  ds.Next;
//  while not ds.Eof do
//  begin
//    id := ds.FieldByName('ID').AsInteger;
//    Query.sql.Clear;
//    Query.sql.Add(format('update zisadmin.ZAPROS16ST SET PRIM=''ДУБЛЬ'', METKA=''%s'','+
//      'question_id=null Where id=%d', [IntToStr(MaxID), id]));
//    Query.ExecSQL;
//    ds.Next;
//  end;
end;

procedure TLettersForm.FormCreate(Sender: TObject);
var
  k : Integer;
  ini : TMemIniFile;
begin
  sl := TStringList.Create;;
  ListObjects := TStringList.Create;
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
   ZipFile := ini.ReadString('sevenzip', 'path', 'C:\Program Files\7-Zip\7z.exe');
   ini.ReadSectionValues('ListObjects', ListObjects);
   BaseName := ini.ReadString('BaseName', 'name', 'btizu_21st_2022');
  finally
    ini.Free;
  end;

//  ds := TADODataSet.Create(nil);
//  ds.CacheSize := 100;
//  ds.CursorLocation := clUseServer;
//  ds.Connection := ADOConnection1;
//
//  ADOConnection1.Close;    //tr!!!
//  ADOConnection1.ConnectionString := format(strConnect, [BaseName]);
//  ADOConnection1.Connected := true;
//
//  ds.Close;
//  ds.CommandText := 'select top 1 * from zisadmin.ZAPROS16ST ORDER BY QUESTION_ID DESC';
//  ds.Open;
//  MaxID := ds.FieldByName('QUESTION_ID').AsInteger;
end;

procedure TLettersForm.FormDestroy(Sender: TObject);
begin
//  ds.Close;
//  ds.Free;
  sl.Free;
  ListObjects.Free;
  ADOConnection1.Close;
end;

end.
