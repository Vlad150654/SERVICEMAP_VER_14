unit ServiceMap_Unit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  System.IniFiles,
  System.Win.ComObj,
  Data.DB, Data.Win.ADODB, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Menus,
  System.RegularExpressions,
  System.Net.URLClient,
  ActiveX,
  System.Contnrs,
  System.Net.HttpClient,
  System.StrUtils,
  System.Zip,
  cfCommon,
  ManualParsing_Unit,
  Letters_Unit,
  Common, MapBasic_INT,
  CostFactors_Unit,
  Path_Unit,
  GeoTransform,
  PrepareFD_Unit,
OXmlUtils, OXmlPDOM, OxmlXPath,
  PsAPI,
//  ShellAPI,
  dbTables,
  csCommon, csGDALOGR, csGDALDataSet, {PythonEngine,} Vcl.PythonGUIInputOutput;

type
  TServiceMapForm = class(TForm)
    Panel4: TPanel;
    stBar: TStatusBar;
    pg : TProgressBar;
    MainMenu1: TMainMenu;
    mnCalcCF: TMenuItem;
    mnGraphics: TMenuItem;
    mnUpdateBase: TMenuItem;
    mnCreateTables: TMenuItem;

    cbProgramMB: TComboBox;
    mnExamen: TMenuItem;
    CatZip: TLabeledEdit;
    mnCompareXML: TMenuItem;
    cbCategories: TComboBox;
    Label1: TLabel;
    mnUnion: TMenuItem;
    mnPrepareLetters: TMenuItem;
    mnLetters: TMenuItem;
    cbMapBasic: TCheckBox;
    cbRunUpdate: TCheckBox;
    cbPacked: TCheckBox;
    ogr1: TMenuItem;
    ADOConnection1: TADOConnection;
    mnAnalogs: TMenuItem;
    mnDKSEMANTIC: TMenuItem;
    CorrectXML: TMenuItem;
    mnFD: TMenuItem;
    mnCalc: TMenuItem;
    mnCheck: TMenuItem;
    mnSelectUserFields: TMenuItem;
    mnFullCorrectFD: TMenuItem;
    mnCorrect: TMenuItem;
    mnCalcGroup: TMenuItem;
    mnImportXML: TMenuItem;
    mnSt16: TMenuItem;
    mnGKO: TMenuItem;
    OpenDialog1: TOpenDialog;
    mnCorrectFDPlus: TMenuItem;
    Telda2НСПД: TMenuItem;
    N2: TMenuItem;
    mnAny: TMenuItem;
    mnTransactions: TMenuItem;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure mnUpdateBaseClick(Sender: TObject);
    procedure mnCreateTablesClick(Sender: TObject);
    procedure mnExamenClick(Sender: TObject);
    procedure mnCompareXMLClick(Sender: TObject);
    procedure mnUnionClick(Sender: TObject);
    procedure mnPrepareLettersClick(Sender: TObject);
    procedure ogr1Click(Sender: TObject);
    procedure mnDKSEMANTICClick(Sender: TObject);
    procedure CorrectXMLClick(Sender: TObject);
    procedure mnCalcClick(Sender: TObject);
    procedure mnCheckClick(Sender: TObject);
    procedure mnSelectUserFieldsClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure mnCalcGroupClick(Sender: TObject);
    procedure mnSt16Click(Sender: TObject);
    procedure mnGKOClick(Sender: TObject);
    procedure mnImportXMLClick(Sender: TObject);
    procedure OpenDialog1Close(Sender: TObject);
    procedure mnFullCorrectFDClick(Sender: TObject);
    procedure mnCorrectFDPlusClick(Sender: TObject);
    procedure Telda2НСПДClick(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure mnTransactionsClick(Sender: TObject);
  private
    { Private declarations }
    radius, decimal : Integer;
    procent : Double;
    path, source, xls, xml, date, tabdbf, fieldsdbf : String;
    cn : TStringList;
    TempPath : String;//C:\Users\TroitskikhVV\AppData\Local\Temp\MIPRO
    ResFolder : String;
    procedure Info(msg : String);
    function OnCreateDecompressStream(const InStream: TStream;
      const ZipFile: TZipFile; const Item: TZipHeader;
      IsEncrypted: Boolean): TStream;
    procedure CorrectXLS(xlsname : String);
    function CorrectFDproc : Boolean;
    procedure Contains(var id, row, kodnp, idkp : Integer;
      var fields : TStringList; var WorkSheet, DataParcels : OleVariant;
      tab, tname, fname : String);
    procedure Intersects(var id, row, radius : Integer;
      var fields : TStringList; var WorkSheet, DataParcels : OleVariant;
      tab, tname, fname : String);
    procedure CalcOne(var ExcelApp : OleVariant);
    public
    { Public declarations }
  end;

const
  strConnectBTI = 'Provider=SQLOLEDB.1;Password=s;Persist Security Info=True;'+
    'User ID=zisadmin;Initial Catalog=%s;Data Source=TESTSQL';

var
  ServiceMapForm: TServiceMapForm;

implementation

{$R *.dfm}

function TServiceMapForm.OnCreateDecompressStream(const InStream: TStream;
  const ZipFile: TZipFile; const Item: TZipHeader;
  IsEncrypted: Boolean): TStream;
begin
  try
    if IsEncrypted then
      ShowMessage('Зашифровано')
    else
      Result := InStream;
  except
    on E: Exception do
      Result := InStream;
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.OpenDialog1Close(Sender: TObject);
begin
  Self.SetFocus;
end;
{------------------------------------------------------------------------------}
procedure ExportToCSVRaw(const AFileName: string);
var
  F: TextFile;
  Row, RowCount: Integer;
  Line: string;
  ID, Name, Qty, Price, DateStr: string;
begin
  RowCount := 1000000; // Проверяем на 1 миллионе строк!

  // 1. Связываем файловую переменную с именем файла
  AssignFile(F, AFileName);
  try
    // 2. Открываем файл для перезаписи
    Rewrite(F);
    // 3. Устанавливаем увеличенный буфер для дисковых операций (стандартный 128 байт — слишком мал).
    // Буфер в 64 КБ (65536 байт) резко снижает количество обращений к диску и ускоряет запись в разы.
    SetTextBuf(F, AllocMem(65536)^, 65536);

    // Для Unicode версий Delphi, чтобы принудительно задать UTF-8
    Write(F, #$EF#$BB#$BF); // Запись UTF-8 BOM

    // 4. Специфика Excel: Явно задаем разделитель столбцов
    WriteLn(F, 'sep=;');

    // 5. Записываем заголовки столбцов
    WriteLn(F, 'ID;Наименование;Количество;Цена;Дата');

    // 6. Цикл генерации и мгновенной записи данных
    for Row := 1 to RowCount do
    begin
      // Подготавливаем данные (переводим в строки заранее)
      ID      := IntToStr(Row);

      // Текст обязательно оборачиваем в двойные кавычки на случай,
      // если внутри названия попадется точка с запятой или пробел
      Name    := '"Товар № ' + IntToStr(Row) + '"';

      Qty     := IntToStr(Row * 2);
      Price   := FloatToStr(Row * 1.25);
      DateStr := DateToStr(Date);

      // Собираем строку через разделитель ';'
      Line := ID + ';' + Name + ';' + Qty + ';' + Price + ';' + DateStr;

      // Сразу сбрасываем строку в буфер файла
      WriteLn(F, Line);
    end;

  finally
    // 7. Закрываем файл (при этом буфер принудительно сбрасывается на диск)
    CloseFile(F);
  end;
  ShellExecute(0, 'open', PChar(AFileName), nil, nil, SW_SHOWNORMAL);
end;

procedure CheckMem;
type
  TVariantArray = array[0..MaxInt div SizeOf(Variant) - 1] of Variant;
  PVariantArray = ^TVariantArray;

var
  ExcelApp, Workbook, WorkSheet : OleVariant;
  d: TDateTime;
  vData: OLEVariant;
  pData: PVariantArray;
  Row, Col: Integer;
  Value: Variant;
  MaxRow, LinearIndex : Integer;
  s : String;
begin
//ExportToCSVRaw('D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\Telda2НСПД\Примеры файлов из Телда\temp.csv');
//exit;
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;
//  ExcelApp.Workbooks.Open('D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\Telda2НСПД\Примеры файлов из Телда\31560-0-5127 (8).xlsx');
//    ExcelApp.Workbooks.Open('D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\Telda2НСПД\Примеры файлов из Телда\31560-0-1.xlsx');
  ExcelApp.Workbooks.Open('D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\Telda2НСПД\Примеры файлов из Телда\Книга1.xlsx');

  WorkSheet := ExcelApp.ActiveWorkbook.ActiveSheet;
  vData := WorkSheet.UsedRange.Value;

  d:=Now;
    for Row := VarArrayLowBound(vData, 1) to VarArrayHighBound(vData, 1) do
      for Col := VarArrayLowBound(vData, 2) to VarArrayHighBound(vData, 2) do
        Value := vData[Row, Col];
//   note('Время чтения всего листа1: '+FormatDateTime('hh:mm:ss:z', Now()-d));//0.407

  d:=Now;
    for Row := VarArrayLowBound(vData, 1) to VarArrayHighBound(vData, 1) do
      for Col := VarArrayLowBound(vData, 2) to VarArrayHighBound(vData, 2) do
        Value := VarArrayGet(vData, [Row, Col]);//это быстрее? НЕТ!
//   note('Время чтения всего листа2: '+FormatDateTime('hh:mm:ss:z', Now()-d));//0.440

  MaxRow := VarArrayHighBound(vData, 1);
  d:=Now;
  pData := VarArrayLock(vData);// Быстро читаем значение напрямую из памяти
  try
    for Row := VarArrayLowBound(vData, 1) to VarArrayHighBound(vData, 1) do//строки
      for Col := VarArrayLowBound(vData, 2) to VarArrayHighBound(vData, 2) do//столбцы
      begin
        // Вычисляем плоский (линейный) индекс элемента в памяти Excel-массива
         LinearIndex := (Col - 1) * MaxRow + Row - 1;
         Value := pData^[LinearIndex];
         s := Value;
    end;
  finally
    VarArrayUnlock(vData);// Обязательно разблокируем!
  end;
// note('Время чтения всего листа3: '+FormatDateTime('hh:mm:ss:z', Now()-d));//0.190

  ExcelApp.Workbooks.Close;
  ExcelApp.Quit;
  Workbook := Unassigned;
  WorkSheet := Unassigned;

  d:=Now;
  VarClear(vData);
//  vData := Unassigned;
//  note('Время удаления данных: '+FormatDateTime('hh:mm:ss:z', Now()-d));//0.810
end;

function GetProcessMemoryUsage: UInt64;
var
  MemCounters: TProcessMemoryCounters;
begin
  MemCounters.cb := SizeOf(TProcessMemoryCounters);
  if GetProcessMemoryInfo(GetCurrentProcess, @MemCounters, MemCounters.cb) then
    Result := MemCounters.WorkingSetSize div 1024
  else
    RaiseLastOSError;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.Telda2НСПДClick(Sender: TObject);
var
  sourcepath, s, proc, kod_model, tab : String;
  i, j, k, n, m : Integer;
  sl, fd, csv, log, cof, nspd : TStringList;
  sa  : StringArray;
  mc : TMatchCollection;
  ExcelApp, Workbook, WorkSheet : OleVariant;
begin
//см. Telda2НСПДClick.txt
//  note(inttostr(GetProcessMemoryUsage));
//  CheckMem;
//  note(inttostr(GetProcessMemoryUsage));
//EXIT;

  if not FileExists(GetAppFolder + 'TopnLab\Справочник соответствия модели и ЦОФ.xlsx') then
    Exc('Не найден файл "Справочник соответствия модели и ЦОФ.xlsx"');
  sourcepath := GetAppFolder;
  {$IFDEF DEBUG}sourcepath := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\Telda2НСПД\Примеры файлов из Телда\';{$ENDIF}
  if not GetFolder(sourcepath, Handle) then
    exit;

  sl := TStringList.Create;
  log := TStringList.Create;
  fd := TStringList.Create;
  csv := TStringList.Create;
  cof := TStringList.Create;
  nspd := TStringList.Create;
  FindRecursive(sourcepath, '*.xlsx', false, sl);
  for i := sl.Count-1 downto 0 do
  begin
    s := ExtractFileName(sl[i]);
    if s[1] = '~' then
      sl.Delete(i);
  end;
  try
    if sl.Count = 0 then
      Exc('Нет файлоы Excel в указанном каталоге!');

    Info('Чтение справочника...');
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.Application.EnableEvents := false;
    ExcelApp.DisplayAlerts := false;
    try
      SafeCloseTable(miEval('PathToTableName$("%s")', [GetAppFolder + 'TopnLab\Справочник соответствия модели и ЦОФ.tab']));
      tab := CreateTableFromExcel(GetAppFolder + 'TopnLab\Справочник соответствия модели и ЦОФ.xlsx', 'Справочник', '');
      miDo('Select Код_Модели, count(*) from %s group by Код_Модели into groupTable', [tab]);

      s := 'кадастровый номер объекта;значение площади;';//формируем заголовок
      cof.Add('кадастровый номер объекта');
      cof.Add ('значение площади');
      for i := 1 to miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]) do
      begin
        miDo('Fetch Rec %d From %s', [i, tab]);
        cof.Add(miEval('%s.Наименование_ЦОФ_Телда', [tab]));//полный список полей
        s := s + miEval('%s.Наименование_ЦОФ_НСПД', [tab]) + ';';
      end;
      cof.Add('номер процедуры');
      s := s + 'номер процедуры';
      csv.Add(s);//готовим 1 общий файл

      Info('Сборка выгрузки...');
      pg.Max := sl.Count;
      for i := 0 to sl.Count-1 do
      begin
        proc := ExtractFileName(sl[i]);
        j := Pos('-', proc);
        if j > 0 then
          proc := Copy(proc, 1, j-1);

        ExcelApp.Workbooks.Open(sl[i]);//открываем и сохраняем в CSV
        Workbook := ExcelApp.Workbooks.item[1];
        s := ChangeFileExt(sl[i], '.csv');
        WorkBook.SaveAs(Filename:=s, FileFormat:=xlCSV, CreateBackup:=False);
        ExcelApp.Workbooks.Close;
        fd.LoadFromFile(s);
        DeleteFile(s);
        sa := SplitStr(fd[4], ';');//Из данных выгрузки берем строку со списком полей

        //проверка соответстия полей справочника и выгрузки
        for j := Low(sa) to High(sa) do
        begin
          mc := reg.Matches(sa[j], '\d\d\.\d', []);//ищем сигнатуру кода модели
          if mc.Count > 0 then
          begin
            kod_model := Copy(mc.item[0].value, 1, 2);
            //выберем в справочнике все записи по коду модели
            miDo('Select * from %s Where Код_Модели = %s into tempTab', [tab, kod_model]);
            for n := 1 to miEvalInt('TableInfo(tempTab, %d)', [TAB_INFO_NROWS]) do
            begin
              miDo('Fetch Rec %d From tempTab', [n]);
              s := miEval('tempTab.Наименование_ЦОФ_Телда');
              m := -1;//проверяем каждый код выбранной модели на наличие в выгрузке
              for k := Low(sa) to High(sa) do
                if Pos(s, sa[k]) = 1 then//проверяем по вхождению (строки в MI могут быть обрезаны!)
                  Inc(m);
              if m = -1 then//не все поля найдены
                log.Add('файл ' + sl[i] + ' отсутствует ' + s);
            end;
            break;
          end;
        end;

        nspd.Clear;
        for j := 0 to cof.Count-1 do
        begin
          m := -1;
          for n := Low(sa) to High(sa) do
            if cof[j] = sa[n] then
              m := n;
          if m >=0 then
            nspd.Add(intToStr(m))
          else
            nspd.Add('');
        end;

        for j := 5 to fd.Count-1 do
        begin
          SetLength(sa, cof.Count);
          s := '';
          sa := SplitStr(fd[j], ';');
          for n := 0 to nspd.Count-1 do
          begin
            if nspd[n] = '' then
              s := s + ';'
            else
              s := s + sa[StrToInt(nspd[n])] + ';'
          end;
          csv.Add(s + proc);
          SetLength(sa, 0);
        end;

        pg.Position := i+1;
      end;
      pg.Max := 0;

      SetLength(sa, 0);

      csv.SaveToFile(sourcepath + 'Combine.csv');

      s := sourcepath + 'Error.log';
      DeleteFile(s);
      if log.Count > 0 then
      begin
        log.Insert(0, 'Обнаружены недостающие поля');
        log.SaveToFile(s);
        ExShellExecute('Notepad.exe', s, SW_NORMAL, false, 0);
        exit;
      end;

    finally
      ExcelApp.Workbooks.Close;
      ExcelApp.Quit;
      Workbook := Unassigned;
      WorkSheet := Unassigned;
      if TableExists(tab) then
        miDo('Drop table %s', [tab]);
      Info('Готово');
    end;

  finally
    sl.Free;
    log.Free;
    fd.Free;
    csv.Free;
    cof.Free;
    nspd.Free;
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.CorrectXLS(xlsname : String);
var
  zip : TZipFile;
  temppath, s, s2 : String;
  sl, xml : TStringList;
  i, j : Integer;
  mc : TMatchCollection;
begin
  temppath := ExtractFilePath(miEval('TempFileName$("")')) + 'temp\';
  ForceDirectories(temppath);
  sl := TStringList.Create;
  xml := TStringList.Create;
  zip := TZipFile.Create;
  zip.UTF8Support := true;
  zip.OnCreateDecompressStream := OnCreateDecompressStream;
  zip.Open(xlsname, zmReadWrite);
  try
    zip.ExtractAll(temppath);
    FindRecursive(temppath, '*.xml', true, sl);
    for i := 0 to sl.Count-1 do
    begin
      if Pos('workbook.xml', sl[i]) > 0 then
      begin
        xml.LoadFromFile(sl[i]);
        s := xml.Text;
        s := StringReplace(s, '</definedName>', '</definedName>'+#10, [rfReplaceAll]);
        s2 := '<definedName.+_FilterDatabase.+</definedName>';
        mc := reg.Matches(s, s2, [roIgnoreCase, roMultiLine]);
        if mc.count > 1 then
        begin
          for j := mc.count-1 downto 1 do
            s := StringReplace(s, mc.Item[j].Value, '', [rfReplaceAll]);
          xml.Text := s;
          xml.SaveToFile(sl[i]);
          zip.Close;
          s := ExtractFilePath(miEval('TempFileName$("")')) + ExtractFileName(xlsname);
          zip.ZipDirectoryContents(s, temppath, zcDeflate);
          zip.Close;
//          ExShellExecute('cmd', format('/C xCopy /b /Y "%s" "%s"',
          ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',
            [s, xlsname]), SW_HIDE, true, 0);
          exit;
        end;
      end;
    end;
    s := format('/c rd /S /Q  "%s"', [temppath]);
    ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
  finally
    sl.Free;
    xml.Free;
    zip.Free;
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnImportXMLClick(Sender: TObject);
var
  s, ss, analogs : String;
  area, d : Double;
  i, j, k, n, m, p, row, kn, lot_area, cad_area : Integer;
  xml : TStringList;
  mc, sub, mc1 : TMatchCollection;
  xy : TPoint3D;
  ExcelApp, Workbook, WorkSheet, range, DataParcels, data : OleVariant;

begin//TServiceMapForm.mnImportXMLClick
  if not FileExists(GetAppFolder + 'TopnLab\TopnLab.xlsx') then
    Exc('Не найден файл TopnLab.xlsx');
  if not FileExists(GetAppFolder+'Base\Дежурная_карта.tab') then
    Exc(GetAppFolder+'Base\Дежурная_карта.tab не найдена');
  miDo('Open table "%s"', [GetAppFolder+'Base\Дежурная_карта.tab']);
  miDo('Set CoordSys table Дежурная_карта');

{$IFDEF DEBUG} ServiceMapForm.OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\Новый источник данных'; {$ENDIF}
  OpenDialog1.FileName := '*.xml';//feed (1).XML
  if not OpenDialog1.Execute then
    exit;
  xml := TStringList.Create;
  xml.LoadFromFile(OpenDialog1.FileName, TEncoding.UTF8);

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;
  Info('Чтение формата выгрузки...');
  try
    ExcelApp.Workbooks.Open(GetAppFolder + 'TopnLab\TopnLab.xlsx');
    Workbook := ExcelApp.Workbooks.item[1];
    WorkSheet := Workbook.WorkSheets.item[1];
    row := WorkSheet.UsedRange.Rows.Count;
    range := WorkSheet.range[format('A2:C%d', [row])];
    DataParcels := range.value;
    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;

    for j := VarArrayLowBound(DataParcels,1)+1 to VarArrayHighBound(DataParcels,1) do
    begin//создаем запросы по таблицам
      ss := DataParcels[j, 1];
      if (ss <> '') and (ss[1] = '#') then//дополнительные поля берем из таблицы Mahinfo
      begin
        ss := Copy(ss, 2, Length(ss));
        if not FileExists(GetAppFolder+'Analogs\'+ss+'.TAB') then
          Exc(GetAppFolder+'Analogs\'+ss+'.TAB не найдена');
        miDo('Open table "%s.tab"', [GetAppFolder+'Analogs\'+ss]);
      end;
    end;

    Workbook := ExcelApp.WorkBooks.Add;//Создаем Книгу (Workbook)
    WorkSheet := WorkBook.WorkSheets.Item[1];
    WorkSheet := ExcelApp.ActiveSheet;
    kn := 0;//колонка с кадастровым номером
    lot_area := 0;//колонка с площадью по лотам
    cad_area := 0;//колонка с кадастровой площадью

    m := 1;//создаем заголовок (m - порядковый номер колонки без учета объединений)
    for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
    begin//идем по строкам
      if DataParcels[i, 1] = DataParcels[i, 2] then
      begin
        k := i+1;//подсчет ячеек для объединения
        while (k <= VarArrayHighBound(DataParcels,1)) and (DataParcels[k, 1] = '') do
          Inc(k);
        s := ExcelNum2Str(WorkSheet, m);
        ss := ExcelNum2Str(WorkSheet, k-2 - (i-m));

        range := WorkSheet.range[format('%s1:%s1', [s, ss])];
        range.Merge;//объединение ячеек
        range.HorizontalAlignment := xlCenter;//центрируем по середине
        WorkSheet.Cells[1, m] := (DataParcels[i, 3]);

        if CText(DataParcels[i, 1], '<lot-area>') then
          lot_area := -1;//помечаем как потенциально возможную для работы
      end else
      begin
        if CText(DataParcels[i, 2], '<value>') and (lot_area = -1) then
          lot_area := m;//нашли колонку с площадью по лотам

        if CText(DataParcels[i, 1], '$cad_area') then
          cad_area := m;//нашли колонку с площадью MI

        WorkSheet.Cells[2, m] := (DataParcels[i, 3]);
        if CText(DataParcels[i, 1], '<cadastral-number>') then
          kn := m;
        Inc(m);
      end;
    end;

    if lot_area = 0 then
      Exc('Не найдена колонка с площадью по лотам!');
    if cad_area = 0 then
      Exc('Не найдена колонка с кадастровой площадью!');

    data := VarArrayCreate([1, 1, 1, m-1], varOleStr);//одна строка
    s := '';
    for i := 0 to xml.Count-1 do
      s := s + Trim(xml[i]);//все в одну строку
    s := StringReplace(s, '</offer>', '</offer>' + lf, [rfReplaceAll]);
    mc := reg.Matches(s, '<offer internal-id="\d{7,}">.+</offer>', [roMultiLine]);//<offer internal-id="102080425111">
    if mc.Count = 0 then
      Exc('Нет offer internal');

    Info('парсим XML');
    pg.Max := mc.Count;
    for i := 0 to mc.Count-1 do
    begin
      pg.Position := i;
      s := mc.item[i].Value;
      for j := VarArrayLowBound(DataParcels,1)+1 to VarArrayHighBound(DataParcels,1) do
        if DataParcels[j, 1] = DataParcels[j, 2] then
          s := StringReplace(s, DataParcels[j, 1], lf + DataParcels[j, 1], [rfReplaceAll]);

      for k := VarArrayLowBound(data,2) to VarArrayHighBound(data,2) do
        data[1, k] := '';
      k := PosEx('"', s, 21);
      data[1, 1] := Copy(s, 21, k-21);//ID
      xy.x := 0;

      m := 0;//число подузлов
      for j := VarArrayLowBound(DataParcels,1)+1 to VarArrayHighBound(DataParcels,1) do
        if DataParcels[j, 1] <> '' then//идем по колонке тегов
      begin//создаем запросы
        ss := DataParcels[j, 1];
        if (ss[1] = '#') or (ss[1] = '$') then//дополнительные поля
          Continue;
        if ss = DataParcels[j, 2] then//выбираем вложенные узлы
          Inc(m);
        System.Insert('/', ss, 2);
        k := Length(ss);
        ss := DataParcels[j,1] + '.+?' + ss;//? ленивый режим
        sub := reg.Matches(s, ss, [roMultiLine]);
        if sub.Count > 0 then
        begin
          if DataParcels[j, 1] = DataParcels[j, 2] then//выбираем вложенные узлы
          begin
            for n := j+1 to VarArrayHighBound(DataParcels,1) do
              if DataParcels[n, 1] = '' then//выбираем аттрибуты узла используя пустые значениея
              begin
                ss := DataParcels[n, 2];
                System.Insert('/', ss, 2);
                k := Length(ss);
                ss := DataParcels[n, 2] + '.+?' + ss;
                mc1 := reg.Matches(sub.item[0].Value, ss, [roMultiLine]);
                if mc1.Count > 0 then
                begin
                  data[1, n-m] := Copy(mc1.item[0].Value, k, Length(mc1.item[0].Value)-k-k+1);//данные из XML

                  if (n-m = lot_area+1) and CText(data[1, lot_area+1], 'сотка') then
                  begin
                    Val(data[1, cad_area], area, p);//площадь взятая из объекта MI
                    Val(data[1, lot_area], d, p);//площадь по лотам
                    if (Abs(d-area)*100/d > 10) and (kn > 0) then
                      data[1, kn] := '';//это КН удаляем
                  end;

                  if CText(DataParcels[n, 2], '<latitude>') then
                    Val(data[1, n-m], xy.y, p);
                  if CText(DataParcels[n, 2], '<longitude>') then
                    Val(data[1, n-m], xy.x, p);
                end;
              end else
                break;
          end else//выбираем данные при узле
          begin
            data[1, j-m] := Copy(sub.item[0].Value, k, Length(sub.item[0].Value)-k-k+1);
            if CText(DataParcels[j, 1], '<description>') then
            begin
              ss := data[1, j-m];
              ss := StringReplace(ss, '<br />', '', [rfReplaceAll]);
              ss := StringReplace(ss, format('<![CDATA[Арт. %s ', [data[1, 1]]), '', [rfReplaceAll]);
              ss := StringReplace(ss, ']]>', '', [rfReplaceAll]);
              data[1, j-m] := ss;

              if kn <> 0 then//есть колонка с КН
              begin//в описании ищем кадастровый номер
                mc1 := reg.Matches(ss, '\d\d:\d\d:\d{1,7}:\d{1,}', [roMultiLine]);
                if mc1.Count > 0 then//первый попавшийся КН
                begin
                  data[1, kn] := mc1.Item[0].Value;//заносим КН из описания
                  ss := ExcelNum2Str(WorkSheet, kn);
                  range := WorkSheet.range[format('%s%d:%s%d', [ss, i+3, ss, i+3])];
                  range.Interior.Color := 65535;
                end;
              end;
            end;
          end;
        end else
        if CText(DataParcels[j,1], '<lot-area>') then//пллощадь по лоту не найдена
          data[1, cad_area] := '';//площадь взятая из объекта MI


        if CText(DataParcels[j, 1], '<cadastral-number>') and (xy.x <> 0) then
        begin
          Wgs2Msk_(xy.y, xy.x, 2);
          miDo('Create Point Into Variable ob	(%.9f, %.9f)', [xy.x, xy.y]);
          if (kn = 0) or (data[1, kn] = '') then//нет КН в описании
          begin
            miDo('Select * from Дежурная_карта where obj Contains ob Into TempTab NoSelect');//центроид объекта B лежит в границах объекта A
            if miEvalInt('TableInfo(TempTab, %d)', [TAB_INFO_NROWS]) > 0 then
            begin
              data[1, j-m] := miEval('TempTab.КН');
              data[1, cad_area] := format('%.f', [miEvalDbl('CartesianArea(TempTab.obj, "sq m")')/100]);//площадь в сотках
            end;
          end;
        end;
      end;

      if xy.x <> 0  then//точечный объект наверняка есть
      j := VarArrayLowBound(DataParcels,1)+1;
      While j <= VarArrayHighBound(DataParcels,1) do
      begin//создаем запросы по таблицам MI
        ss := DataParcels[j, 1];
        if (ss <> '') and (ss[1] = '#') then
        begin
          ss := Copy(ss, 2, Length(ss));
          miDo('Select * from %s where obj Contains ob Into TempTab NoSelect', [ss]);//центроид объекта B лежит в границах объекта A
          if miEvalInt('TableInfo(TempTab, %d)', [TAB_INFO_NROWS]) > 0 then
          begin
            if DataParcels[j, 2] <> '' then//есть дополнительные поля
            begin
              data[1, j-m] := miEval('TempTab.%s', [DataParcels[j, 2]]);//берем значение этого поля
              while j < VarArrayHighBound(DataParcels,1) do//продолжаем брать из других полей
              begin
                if (j < VarArrayHighBound(DataParcels,1)) and (DataParcels[j, 1] = DataParcels[j+1, 1]) then
                begin
                  Inc(j);
                  data[1, j-m] := miEval('TempTab.%s', [DataParcels[j, 2]]);
                end else
                  Break;
              end;
            end else//для таблиц без полей (Буфер_от500, water_polygon)
              data[1, j-m] := 1;//есть вхождение
          end else
            if DataParcels[j, 2] = '' then//для таблиц без полей (Буфер_от500, water_polygon)
              data[1, j-m] := 0;//нет вхождения
        end;
        Inc(j);
      end;

      ss := ExcelNum2Str(WorkSheet, VarArrayHighBound(data,2));
      range := WorkSheet.range[format('A%d:%s%d', [i+3, ss, i+3])];
      range.value := data;
    end;

    pg.Max := 0;
    SafeCloseTable('TempTab');
    ExcelApp.Visible := true;

  finally
    if (lot_area = 0) or (cad_area = 0) then
    begin
      ExcelApp.Workbooks.Close;
      ExcelApp.Quit;
    end;
    Workbook := Unassigned;
    WorkSheet := Unassigned;
//    data := Unassigned;//tr!!!
    VarClear(data);
    VarClear(DataParcels);
    xml.Free;
    Info('Готово');
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.Info(msg : String);
begin
  stBar.Panels[1].Text := msg;
  stBar.Repaint;
  Application.ProcessMessages;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.FormActivate(Sender: TObject);
var
  ini: TMemIniFile;
  s : String;
begin
  if not FileExists(CatZip.Text) then
  begin
    if InputQuery('7z.exe не найден!', 'Введите путь к файлу', s) then
    begin
      CatZip.Text := s;
      ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
      try//обновим путь к файлу 7z.exe
        ini.WriteString('sevenzip', 'path', CatZip.Text);
      finally
        ini.UpdateFile;
        ini.Free;
      end;
    end;
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Mi.MBApplications.Item(mbName).Do('CloseServiceMap');
  Action := caFree;
end;
{------------------------------------------------------------------------------}
//C:\Users\TroitskikhVV\AppData\Local\Programs\Python\Python313-32
procedure TServiceMapForm.FormCreate(Sender: TObject);
var
  ver : String;
  sl : TStringList;
  i : Integer;
  ini: TMemIniFile;
begin
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
    CatZip.Text := ini.ReadString('sevenzip', 'path', 'C:\Program Files\7-Zip\7z.exe');
    cbCategories.Items.CommaText := ini.ReadString('Categories', 'names', '');
    if cbCategories.Items.Count > 0 then
      cbCategories.ItemIndex := 0;
  finally
    ini.Free;
  end;

  sl := TStringList.Create;
  try
    FindRecursive(GetAppFolder, '*.mbx', false, sl);
    for i := sl.Count-1 downto 0 do
    begin
      sl[i] := ExtractFileName(sl[i]);
      if CText(mbName, sl[i]) then
        sl.Delete(i);
    end;
    cbProgramMB.Items.Assign(sl);
    if cbProgramMB.Items.Count > 0 then
      cbProgramMB.ItemIndex := 0;
  finally
    sl.Free;
  end;

  if GetVersionInfo(GetModuleName(HInstance), ver) then
   Caption := Caption + ' ' + ver;

  cn := TStringList.Create;
  cn.Duplicates := dupIgnore;
  miDo('Dim ob as Object');
  miDo('Dim ob1 as Object');
  miDo('Dim ob2 as Object');
  miDo('Dim ztb as Object');
  miDo('Dim igs as Object');
  miDo('Dim eburg as Object');
  TempPath := ExtractFilePath(miEval('TempFileName$("")'));
  ResFolder := GetAppFolder + 'RESULT\';
  Set_Priority(AppWnd, HIGH_PRIORITY_CLASS);
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.FormDestroy(Sender: TObject);
begin
  cn.Free;
  miDo('UnDim ob');
  miDo('UnDim ob1');
  miDo('UnDim ob2');
  miDo('UnDim ztb');
  miDo('UnDim igs');
  miDo('UnDim eburg');
  mi := Unassigned;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 27 then
    BreakCommand := true;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnUnionClick(Sender: TObject);
var
  sl, examen_sl, replace_sl : TStringList;
  i, j, k : Integer;
  s, s2 : String;
begin
{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP\Win32\Debug\Письма';{$ENDIF}    //2024_04_01
{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\';{$ENDIF}//для расчета от шмелева 2026_09_10

//  if path = '' then
//    path := 'W:\ЦГКО\IT\Outbox\Статья 16\2024\';
  if not GetFolder(path, Handle) then exit;

  sl := TStringList.Create;
  try
    FindFolders(path, sl, false);
    for i := 0 to sl.Count-1 do
      if CText(ExtractFileName(sl[i]), 'All') then
        Exc('В выбранном каталоге имеется папка "All"' + #10 +
          'в которую заносится результаты объединения.' + #10 + 'Удалите эту папку.');
  finally
    sl.Free;
  end;

  sl := TStringList.Create;
  sl.Duplicates := dupAccept;
  examen_sl := TStringList.Create;
  examen_sl.Duplicates := dupError;
  examen_sl.Sorted := True;
  replace_sl := TStringList.Create;
  try
    FindRecursive(path, '*.xml', true, sl);
    Info(' Проверка дубликатов файлов...');
    for i := 0 to sl.Count-1 do
    try
      s := ChangeFileExt(ExtractFileName(sl[i]), '');//имя без расширения
      examen_sl.Add(s);
    except
      k := 2;
      while examen_sl.Find(format('%s_%d', [s, k]), j) do
        Inc(k);
      s := format('%s_%d', [s, k]);
      examen_sl.Add(s);
      replace_sl.Add(sl[i] + '=' + ExtractFilePath(sl[i]) + s + '.xml');
      s := ChangeFileExt(sl[i], '.csv');
      if not FileExists(s) then
        Exc('Не найден файл ' + s +#10 + 'Обработка прервана');
    end;
    //изменения вносятся в каталоги полученные в результате обработки писем
    for i := 0 to replace_sl.Count-1 do
    begin
      s := replace_sl.Names[i];
      s2 := replace_sl.ValueFromIndex[i];
      RenameFile(s, s2);
      RenameFile(ChangeFileExt(s, '.csv'), ChangeFileExt(s2, '.csv'));
    end;
    DeleteFile(path + 'Изменение имен файлов.txt');
    if replace_sl.Count > 0 then
      replace_sl.SaveToFile(path + 'Изменение имен файлов.txt');

    sl.Clear;
    FindFolders(path, sl, false);
    if not ForceDirectories(path + 'All') then
      Exc('Не удалось создать каталог "All"');
    Info(' Копирование файлов');
    pg.Max := sl.Count;
    for i := 0 to sl.Count-1 do
    begin
      pg.Position := i + 1;
//      ExShellExecute('cmd', PChar(format('/C xCopy /b /Y "%s\*" "%s"',
      ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s\*" "%s"',
        [AddSlesh(sl[i])+'xml', path + 'All'])), SW_HIDE, true, 0);
    end;
    pg.Max := 0;
  finally
    sl.Free;
    examen_sl.Free;
    replace_sl.Free;
    Info(' Готово');
  end;
end;
{-------------------------------------------------------------------------------}
function gmd(date : String) : String;
begin
  Result := Copy(date, 1, 4) + Copy(date, 5, 2) + Copy(date, 7, 2);
end;
{-------------------------------------------------------------------------------}
procedure TServiceMapForm.mnUpdateBaseClick(Sender: TObject);
var
  log : TStringList;
  tab, cn, s_z, s_b, tz, tb, dmg : String;
  arz, arb : Double;
  i, kz, kb, row, rowcount, maxid, maxrow : Integer;
begin
  if not FileExists(GetAppFolder+'Base\Дежурная_карта.tab') then
    Exc(GetAppFolder+'Base\Дежурная_карта.tab не найдена');

  dmg := StringReplace(CurrentData, '_', '', [rfReplaceAll]);
  BreakCommand := false;
  if not TableExists('AllTables') then
  begin
    OpenDialog1.FileName := '*.tab';
    if not OpenDialog1.Execute then
      exit;

    if not ExamenOGR (stBar, OpenDialog1.FileName, TempPath) then
      Exc('Найдена плохая графика.');

    OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
    miDo('Open table "%s"', [OpenDialog1.FileName]);
    tab := miEval('PathToTableName$("%s")', [OpenDialog1.FileName]);
  end else
    tab := 'AllTables';

  miDo('Create Index On %s (КН)', [tab]);
  log := TStringList.Create;
  log.Add(format('информация%sкадастровый номер%sдата базы%sдата пакета%s'+
    'частей в базе%sчастей в пакете%sплощадь в базе%sплощадь в пакете%sdS',
    [#9, #9, #9, #9, #9, #9, #9, #9]));
  try
    Info(' Выбор объектов..');
    miDo('select * from %s Order By КН into work NoSelect', [tab]);//выбираем объекты
    rowcount := miEvalInt('TableInfo(work, %d)', [TAB_INFO_NROWS]);

    miDo('Open table "%s"', [GetAppFolder+'Base\Дежурная_карта.tab']);
    miDo('Create Index On Дежурная_карта (КН)');

    Info(' Поиск максимального номера...');
    miDo('Select Max(idn) from Дежурная_карта into Temp');
    maxid := miEvalInt('Temp.Col1');
    maxrow := miEvalInt('TableInfo(Дежурная_карта, %d)', [TAB_INFO_NROWS]);

    Info(' Создание ZIP файла...');
    ExCreateProcess(format('"%s" a -tzip -mx5 "%sДежурная_карта_%s.zip" "%sДежурная_карта.*"',
      [CatZip.Text, GetAppFolder+'Base\', CurrentData, GetAppFolder+'Base\']), '', SW_HIDE, true);

    miDo('Set coordsys table Дежурная_карта');
    Info(' Обновление Дежурной карты...');
    pg.Max := rowcount;
    row := 1;
    While row <= rowcount do
    begin
      Application.ProcessMessages;
      if BreakCommand then
        exit;
      stBar.Panels[0].Text := IntToStr(row);
      miDo('Fetch Rec %d From work', [row]);//выборка из таблицы AllTables или из выбранной

      cn := MiEval('work.КН');
      miDo('select * from work where КН="%s" into selwork NoSelect', [cn]);
      miDo('select * from Дежурная_карта where КН="%s" into selbase NoSelect', [cn]);

      kz := miEvalInt('TableInfo(selwork, %d)', [TAB_INFO_NROWS]);//число выборок в выгрузке
      kb := miEvalInt('TableInfo(selbase, %d)', [TAB_INFO_NROWS]);//число выборок в базе
      arz := 0; arb := 0;
      try
        tz := miEval('Str$(selwork.obj)');//тип объекта в выгрузке
        if tz <> 'Point' then//объединяем в 1 многоконтурный объект
          miDo('Create Object As Union From selwork Into Variable ob')
        else
          miDo('ob = selwork.obj');
        s_z := miEval('selwork.Дата_пакета');//дата-строка в выгрузке
        if s_z = '' then
        begin
          s_z := dmg;
          miDo('Update selwork Set Дата_пакета = "%s"', [s_z]);
        end;
        arz := miEvalDbl('CartesianArea(ob, "sq m")');
      except
        log.Add(format('Ошибка в выгрузке%s%s',[#9, cn]));
      end;

      if kb > 0 then//если в базе есть такие кадастровые номера
      begin//замена
        try
          tb := miEval('Str$(selbase.obj)');//тип объекта в базе
          if tb <> 'Point' then//объединяем в 1 многоконтурный объект
            miDo('Create Object As Union From selbase Into Variable ob2')
          else
            miDo('ob2 = selbase.obj');
          s_b := miEval('selbase.Дата_пакета');//Дата пакета в базе
          if s_b = '' then
            s_b := '20200101';
          arb := miEvalDbl('CartesianArea(ob2, "sq m")');
        except
          log.Add(format('Ошибка в Дежурной карте%s%s',[#9, cn]));
        end;

        if CText(tz,'Region') and CText(tb, 'Point') then//сравниваем типы объетов
        begin//точечный объект в дежурке заменяем площадным без проверки обновления
          miDo('Insert Into дежурная_карта Select * From selwork');//пакет-загрузка
          log.Add(format('замена%s%s%s%s%s%s%s%d%s%d%s%.2f%s%.2f%s%.2f',
            [#9, cn, #9, gmd(s_b), #9, gmd(s_z), #9, kb, #9, kz, #9, arb, #9, arz, #9, arb-arz]));
          miDo('Delete from selbase');//объект в базе удаляем всегда
        end else
        if s_z > s_b then//сравниваем по дате, более новую выгрузку заносим в базу
        begin
          miDo('Insert Into дежурная_карта Select * From selwork');//пакет-загрузка
          log.Add(format('изменение%s%s%s%s%s%s%s%d%s%d%s%.2f%s%.2f%s%.2f',
            [#9, cn, #9, gmd(s_b), #9, gmd(s_z), #9, kb, #9, kz, #9, arb, #9, arz, #9, arb-arz]));
          miDo('Delete from selbase');//объект в базе удаляем всегда
        end;
      end else
      begin//этого номера в базе нет, добавляем все выбранные части КН
        miDo('Insert Into дежурная_карта Select * From selwork');
        log.Add(format('добавление%s%s%s%s%s%s%s%s%s%d%s%s%s%.2f',
          [#9, cn, #9, '', #9, gmd(s_z), #9, '', #9, kz, #9, '', #9, arz]));
      end;
      Inc(row, kz);
      pg.Position := row;
    end;
    SafeCloseTable('work');
    SafeCloseTable('selbase');

  finally
    log.SaveToFile(GetAppFolder+'Base\logUpdate.log');
    if log.Count > 1 then
      ExShellExecute('Notepad.exe', GetAppFolder+'Base\logUpdate.log', SW_NORMAL, false, 0)
    else
      DeleteFile(GetAppFolder+'Base\logUpdate.log');
    log.Free;
    pg.Max := 0;
    stBar.Panels[0].Text := '';
    if not BreakCommand then
    begin
      miDo('Select * from Дежурная_карта Where rowid>%d into temp NoSelect', [maxrow]);
      miDo('Update temp Set IDN = rowid+%d DropIndex Auto', [maxid]);
      Info(' Сохранение Дежурной карты...');
      miDo('Commit table Дежурная_карта');
      if cbPacked.Checked then
      begin
        Info(' Упаковка Дежурной карты...');
        miDo('Pack Table Дежурная_карта Graphic Data');
      end;
    end;
    Info(' Готово');
  end;
end;
{-----------------------------------------------------------------------------}
procedure TServiceMapForm.N2Click(Sender: TObject);
var
  xls, s : String;
  i, j, row, col : Integer;
  sl : TStringList;
  ExcelApp, Workbook, Workbook_, WorkSheet, WorkSheet_, data : OleVariant;
begin
//  path := 'W:\ЦГКО\IT\Outbox\Статья 16\2026\2026_06_18\SOKS_a_2026-06-15_2\SOKC_Result\';
  path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\Сервер\';//tr!!!

  xls := GetAppFolder + 'ODS\Новая форма.ods';
  if not FileExists(xls) then
    Exc('Не найден файл' + lf + xls);
  if not GetFolder(path, Handle) then
    Exit;

  Info('Подготовка файлов...');
  sl := TStringList.Create;
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.DisplayAlerts := false;
  ExcelApp.Application.EnableEvents := false;

  try
    FindRecursive(path, '*.ods', true, sl);
    for i := 0 to sl.Count-1 do//проверка открытых файлов ODS
    begin
      s := ExtractFilePath(sl[i]) + 'Новая форма ' + ExtractFileName(sl[i]);
      if FileExists(s) then
        if not DeleteFile(s) then
          Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.',
            [s, lf]));
    end;

    Info('Создание новых форм');
    sl.Clear;
    FindRecursive(path, '*.ods', true, sl);
    pg.Max := sl.Count;
    for i := 0 to sl.Count-1 do//перезапись новых ODS
    begin
      pg.Position := i+1;
      ExcelApp.Workbooks.Open(xls);
      Workbook := ExcelApp.Workbooks.item[1];
      WorkSheet := Workbook.WorkSheets.item[1];

      WorkSheet.Cells[7, 7] := 0;
      WorkSheet.Cells[8, 7] := 0;

      for j := 1 to 7 do
        WorkSheet.Cells[11, j] := '';
      for j := 1 to 7 do
        WorkSheet.Cells[14, j] := '';

      s := ExtractFilePath(sl[i]) + 'Новая форма ' + ExtractFileName(sl[i]);
      WorkBook.SaveAs(Filename:=s, CreateBackup:=False);
      ExcelApp.Workbooks.Open(sl[i]);
      Workbook_ := ExcelApp.Workbooks.item[2];//старый ODS
      if FindSheetByName(WorkBook_, WorkSheet_, 'I') then
      begin
        WorkSheet_ := Workbook_.WorkSheets.item[1];
        s := ExtractFileName(sl[i]);
        j := Pos('_', s);
        WorkSheet.name := Copy(s, 1, j-1);
        WorkSheet.Cells[2, 2] := WorkSheet_.Cells[2, 2];
        WorkSheet.Cells[2, 7] := WorkSheet_.Cells[2, 8];
        WorkSheet.Cells[5, 7] := WorkSheet_.Cells[5, 8];
        WorkSheet.Cells[6, 7] := WorkSheet_.Cells[6, 8];
      end;

      if FindSheetByName(WorkBook_, WorkSheet_, 'IV') then
      begin
        WorkSheet.Cells[15, 6] := WorkSheet_.Cells[1, 3];
        WorkSheet.Cells[17, 6] := WorkSheet_.Cells[4, 3];
      end;

      if FindSheetByName(WorkBook_, WorkSheet_, 'III', row, col) then
      begin
//        data := VarArrayCreate([1, row-2, 1, 1], varVariant);// 1 колонка

        s := '15:' + IntToStr(15+row-4);//вставляем пустые строки перед 15 строкой
        WorkSheet.Rows[s].Insert(Shift:=xlShiftDown);

//        data := WorkSheet_.range[format('A3:A%d', [row])].value;
        WorkSheet.range[format('A14:A%d', [row-3+14])].value := WorkSheet_.range[format('A3:A%d', [row])].value;
        WorkSheet.Range[format('A14:B%d', [14+row-3])].Merge(True);

//        data := WorkSheet_.range[format('B3:B%d', [row])].value;
        WorkSheet.range[format('C14:C%d', [row-3+14])].value := WorkSheet_.range[format('B3:B%d', [row])].value;

//        data := WorkSheet_.range[format('C3:C%d', [row])].value;
        WorkSheet.range[format('D14:D%d', [row-3+14])].value := WorkSheet_.range[format('C3:C%d', [row])].value;
        WorkSheet.Range[format('D14:E%d', [14+row-3])].Merge(True);

//        data := WorkSheet_.range[format('D3:D%d', [row])].value;
        WorkSheet.range[format('F14:F%d', [row-3+14])].value := WorkSheet_.range[format('D3:D%d', [row])].value;
        WorkSheet.Range[format('F14:G%d', [14+row-3])].Merge(True);

        WorkSheet.Cells[8, 7] := row-2;
//        VarClear(data);
      end;

      if FindSheetByName(WorkBook_, WorkSheet_, 'II', row, col) then
      begin
//        data := VarArrayCreate([1, row-2, 1, 1], varVariant);// 1 колонка

        s := '12:' + IntToStr(12+row-4);//вставляем пустые строки перед 12 строкой
        WorkSheet.Rows[s].Insert(Shift:=xlShiftDown);

//        data := WorkSheet_.range[format('A3:A%d', [row])].value;
        WorkSheet.range[format('A11:A%d', [row-3+11])].value := WorkSheet_.range[format('A3:A%d', [row])].value;
       //Объединить строки построчно (например, A2:B2, A3:B3 и т.д.) Передаем параметр Across:=True
        WorkSheet.Range[format('A11:B%d', [11+row-3])].Merge(True);

//        data := WorkSheet_.range[format('B3:B%d', [row])].value;
        WorkSheet.range[format('C11:C%d', [row-3+11])].value := WorkSheet_.range[format('B3:B%d', [row])].value;

//        data := WorkSheet_.range[format('C3:C%d', [row])].value;
        WorkSheet.range[format('D11:D%d', [row-3+11])].value := WorkSheet_.range[format('C3:C%d', [row])].value;

//        data := WorkSheet_.range[format('D3:D%d', [row])].value;
        WorkSheet.range[format('E11:E%d', [row-3+11])].value := WorkSheet_.range[format('D3:D%d', [row])].value;
        WorkSheet.Range[format('E11:F%d', [11+row-3])].Merge(True);

//        data := WorkSheet_.range[format('E3:E%d', [row])].value;
        WorkSheet.range[format('G11:G%d', [row-3+11])].value := WorkSheet_.range[format('E3:E%d', [row])].value;
        WorkSheet.Cells[7, 7] := row-2;
//        VarClear(data);
      end;
      WorkBook.Save;
      ExcelApp.Workbooks.Close;
      ExcelApp.Quit;
    end;
    pg.Max := 0;

    xls := GetAppFolder + 'ODS\' + AddSlesh(CurrentData);
    ForceDirectories(xls);
    for i := 0 to sl.Count-1 do//переименование ODS
    begin
      s := ExtractFilePath(sl[i]) + 'Новая форма ' + ExtractFileName(sl[i]);
      ExShellExecute('cmd', PChar(format('/c MOVE /Y "%s" "%s"',
        [s, xls + ExtractFileName(sl[i])])), SW_HIDE, true, 0);//перенести файл с переименованием
    end;

  finally
    sl.Free;
    ExcelApp := Unassigned;
    Workbook := Unassigned;
    WorkSheet := Unassigned;
    Workbook_ := Unassigned;
    WorkSheet_ := Unassigned;
    Info('Готово');
  end;
end;
{-----------------------------------------------------------------------------}
procedure TServiceMapForm.ogr1Click(Sender: TObject);
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\temp\двинин\БРАК'; {$ENDIF}
  OpenDialog1.FileName := '*.tab';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  ExamenOGR(stBar, OpenDialog1.FileName, TempPath);
//  InsEnvVar('PATH', 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\GDAL\');
//  SetDllDirectory('D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\GDAL\');
//  csGDALInit('D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\GDAL\');
end;
{------------------------------------------------------------------------------}
{В каталоге "W:\ЦГКО\IT\___Программы\Дежурная карта_VER_12\Письма" выбираем папку, например "2024_05_24"
Ищем все подкаталоги в выбанном, за искдюченим каталога "All" (23 подкаталога)
В каждом из них входим в каталог xml и ищем все файлы csv
Собираем всю информацию в один список из csv: CadastralNumber,ObjectType,DATEFORM,NUMBERINPUTINFO,DATEINPUTINFO (819 записей)
В выбранных каталогах поочередно ищем файлы "mif" (310 файлов)
Из этих файлов вычленяем кадастровый номер и ищем строку с КН в списке CSV
В случае находжения создаем запись типа
  TMif = Record
    fullname : WideString;		полное имя файла
    shortname : ShortString;	короткое имя файла (для сортировки)
    dat : Integer;				DATEFORM в формате г*10000 + м*100 + д  (для сортировки)
    dateform : ShortString;		DATEFORM в формате д.м.г
  end;
Полученный список сортируем по короткому имени и на увеличение даты пакета для перезаписи, таким образом КН номера с более поздней датой перезапишут более ранние
В каталоге программы создаем подкаталог с той же датой "Графика\2024_05_24"
Далее идем по отсортированному списку и
1 загружаем MIF файл по полному имени
2 меняем систему координат
3 Добавляем колонки Дата_проверки и Дата_пакета
4 Сохраняем в новом каталоге "Графика\2024_05_24\Mid_Mid"
5 Загружаем MID файл
6 Добавляем dateform и пустое значение для даты проверки
7 Сохраняем в новом каталоге

Затем идет Создание таблиц, их слияние  и т.д.}
procedure TServiceMapForm.mnCreateTablesClick(Sender: TObject);
type
  PMif = ^TMif;
  TMif = Record
    fullname : WideString;
    shortname : ShortString;
    dat : Integer;
    dateform : ShortString;
  end;

var
  i, j, n, k, kz, kb : Integer;
  mif : PMif;
  mifmid : TList;
  sl, fold, log, csv, tmp : TStringList;
  s, x, y,  dir, tab, tabs, work, cn, s_z, s_b : String;
  sa : StringArray;

  function SortMif(Item1, Item2: PMif): Integer;
  begin
    if Item1.shortname < Item2.shortname then
      Result := -1
    else if Item1.shortname > Item2.shortname then
      Result := 1
    else
    if Item1.dat < Item2.dat then
      Result := -1
    else if Item1.dat > Item2.dat then
      Result := 1
    else
      Result := 0;
  end;

begin//mnCreateTablesClick
{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\TEST\';{$ENDIF}
{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\Письма';{$ENDIF}    //2024_04_27
//{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\TEST\Outbox_Статья_16_2024';{$ENDIF}//2024_04_04


  if not FileExists(GetAppFolder+'Base\Дежурная_карта.tab') then
    Exc(GetAppFolder+'Base\Дежурная_карта.tab не найдена');

//  if path = '' then//рабочий пример W:\ЦГКО\IT\___Программы\OLD\Первичная обработка писем_ver2\Письма
//    path := ResFolder;
  if not GetFolder(path, Handle) then exit;

  log := TStringList.Create;
  fold := TStringList.Create;
  sl := TStringList.Create;
  mifmid := TList.Create;
  csv := TStringList.Create;
  tmp := TStringList.Create;
  work := AddSlesh(AddSlesh(GetCurrentDir) + 'Графика\' + ExtractFileName(DellSlesh(path)));
  ForceDirectories(work + 'Графика_вся\');

  try
    FindFolders(path, fold, false);
    if fold.Count = 0 then
      Exc('Нет подкаталогов в каталоге ' + path);
    for i := fold.Count-1 downto 0 do
      if CText(fold[i], path + 'All') then
        fold.Delete(i);

    info(' Поиск и сортировка Mif');
    pg.Max := fold.Count;//23
    for i := 0 to fold.Count-1 do
    begin
      pg.Position := i+1;

      csv.Clear;
      sl.Clear;
      FindRecursive(AddSlesh(fold[i])+'xml\', '*.csv', true, sl);
      for j := 0 to sl.Count-1 do//8,...
      begin//собираем всю информацию из csv
        tmp.LoadFromFile(sl[j]);
        for n := 1 to tmp.Count-1 do//заголовок пропускаем
          csv.Add(tmp[n]);//819
      end;

      sl.Clear;
      FindRecursive(AddSlesh(fold[i]), '*.mif', true, sl);
      for j := 0 to sl.Count-1 do
      begin
        s := ExtractFileName(sl[j]);//вычленяем КН из имени файла
        s := StringReplace(Copy(s, 1, Length(s)-4), '_', ':', [rfReplaceAll]) + ';';
        k := 0;
        for n := 0 to csv.Count-1 do//ищем строку с КН в списке CSV
          if Pos(s, csv[n]) = 1 then
          begin
            k := 1;
            mif := AllocMem(SizeOf(mif^));
            mif.fullname := sl[j];
            mif.shortname := ExtractFileName(sl[j]);
            mifmid.Add(mif);
            sa := SplitStr(csv[n], ';');
            mif.dateform := sa[2];
            sa := SplitStr(mif.dateform, '.');
            mif.dat := StrToInt(sa[2])*10000 + StrToInt(sa[1]) * 100 + StrToInt(sa[0]);
            SetLength(sa, 0);
            break;
          end;
        if k = 0 then
          log.Add('Кадастровый номер ' + Copy(s, 1, Length(s)-1) + ' файле ' + sl[j] + ' не найден в файлах CSV');
      end;
    end;
    pg.Max := 0;
    if mifmid.Count = 0 then
      Exc('Нет графики!');

    mifmid.Sort(@SortMif);//сортировка на увеличение даты пакета для перезаписи

    dir := AddSlesh(work) + 'Mid_Mid\';
    ForceDirectories(dir);
    tabs := dir + '\tabs\';
    ForceDirectories(tabs);

    info(' Копирование MifMid');
    pg.Max := mifmid.Count;
    for i := 0 to mifmid.Count-1 do
    begin
      pg.Position := i+1;
      mif := mifmid[i];
      sl.LoadFromFile(mif.fullname);
      if Pos ('CoordSys NonEarth Units "m" Bounds (2', sl[3]) > 0 then//зона 2
        sl[3] := cs21 //изменение системы координат
      else
        sl[3] := cs11;
      j := Pos('Columns ', sl[4]);
      Val(Copy(sl[4], j+8, Length(sl[4])), n, j);
      sl[4] := 'Columns ' + IntToStr(n + 2);
      for j := 0 to sl.Count-1 do
      begin
        if CText(sl[j], 'Data') then
        begin//вставка полей перед данными
          sl.Insert(j, '  Дата_проверки Date');
          sl.Insert(j, '  Дата_пакета Date');
          sl.SaveToFile(dir + mif.shortname);
          break;
        end;
      end;
      //добавляем дату в MID
      s := ChangeFileExt(mif.fullname, '.mid');
      sl.LoadFromFile(s);
      for j := 0 to sl.Count-1 do
        sl[j] := sl[j] + format(',"%s",""', [mif.dateform]);
      s := ChangeFileExt(mif.shortname, '.mid');
      sl.SaveToFile(dir + s);//новый каталог
    end;

    miDo('Set ProgressBars Off');
    Info(' Создание таблиц');
    sl.Clear;
    FindRecursive(dir, '*.mif', false, sl);
    miDo('Close All');
    pg.Max := sl.Count;
    for i := 0 to sl.Count-1 do
    begin
      pg.Position := i+1;
      try
        tab := ChangeFileExt(ExtractFileName(sl[i]), '.tab');
        miDo('Import "%s" Type "MIF" Into "%s" Overwrite', [sl[i], tabs+tab]);
      except
        on e: Exception do log.Add(sl[i]+#9+e.Message);
      end;
      miDo('Close All');
    end;

    Info(' Слияние таблиц');
    sl.Clear;
    FindRecursive(tabs, '*.tab', false, sl);
    pg.Position := 0;
    for i := 0 to sl.Count-1 do
    try
      pg.Position := i+1;
      tab := miEval('PathToTableName$("%s")', [sl[i]]);
      miDo('Open Table "%s"', [sl[i]]);

      if i = 0 then
      begin
        miDo('select NOTE,Дата_пакета,Дата_проверки from %s into AllTables', [tab]);
        miDo('Commit Table AllTables As "%sAllTables.TAB" TYPE NATIVE Charset "WindowsCyrillic"', [work]);
        miDo('Close All');
        miDo('Open table "%sAllTables.tab"', [work]);
      end else
      begin
        miDo('Insert into AllTables Select NOTE,Дата_пакета,Дата_проверки from %s', [tab]);
        miDo('Close Table %s', [tab]);
      end;
    except
      on e: Exception do log.Add(sl[i]+#9+e.Message);
    end;
    miDo('Commit table AllTables');
    miDo('Alter Table "AllTables" (ADD IDN Integer,Черезполосники Char(10),'+
      'Самопересечение Char(10),За_Границей Char(10),В_работу Char(10),'+
      'Примечание Char(250) RENAME NOTE КН '+
      'MODIFY КН Char(50) order IDN,КН,Дата_пакета,Дата_проверки,Черезполосники,'+
      'Самопересечение,За_Границей,В_работу,Примечание)');
    pg.Max := 0;
    miDo('Set ProgressBars On');

//    ExShellExecute('cmd', PChar(format('/C xCopy /b /Y "%sAllTables.*" "%s"',
    ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%sAllTables.*" "%s"',
      [work, work+'Графика_вся\'])), SW_HIDE, true, 0);
    miDo('Close table AllTables');
    miDo('Open table "%sГрафика_вся\AllTables.tab"', [work]);

  finally
    log.SaveToFile(work + 'logCreate.log');
    if log.Count > 0 then
      ExShellExecute('Notepad.exe', work + 'logCreate.log', SW_NORMAL, false, 0)
    else
      DeleteFile(work + 'logCreate.log');
    info(' Готово');
    for i := 0 to mifmid.Count-1 do
      Dispose(mifmid.Items[i]);
    mifmid.Free;
    log.Free;
    fold.Free;
    sl.Free;
    csv.Free;
    tmp.Free;
  end;

  if TableExists('AllTables') and cbMapBasic.Checked and FileExists(GetAppFolder + cbProgramMB.Text) then
  begin
    miDo('Open table "%s"', [GetAppFolder+'Base\Дежурная_карта.tab']);
    if not ExamenOGR(stBar, Work + 'AllTables.tab', TempPath) then
    begin
      miDo('Select * from AllTables, ErrorTable where AllTables.КН=ErrorTable.КН into TempTab NoSelect');
      miDo('Update TempTab Set В_работу = "Нет" DropIndex Auto');
      miDo('Commit table AllTables');
      miDo('Close table ErrorTable');
    end;
    //чисто MapBasic
    miDo('Map from AllTables');
    miDo('Set Map Window FrontWindow() Layer 1 Editable On');
    if cbProgramMB.ItemIndex = -1 then
      Exc('Программа Mapbasic не определена');
    Info(' Работает ' + cbProgramMB.Text + '...');
    miDo('Run Application "%s"',[GetAppFolder + cbProgramMB.Text]);
    Application.ProcessMessages;
    s := StringReplace(cbProgramMB.Text, '.mbx', '', [rfIgnoreCase]);
    mi.MBApplications.Item(s).Eval(format('Work:%s', [GetAppFolder]));
    s := mi.MBApplications.Item(s).Eval(format('Path:%s', [path]));
    Info(s);
    miDo('Terminate Application "%s"',[GetAppFolder + cbProgramMB.Text]);

    if FileExists(format('%sГрафика_вся\брак\AllTables.tab', [work])) then
    begin
      SafeCloseTable('AllTables');//переоткрываем AllTables с другими полями из брака
      miDo('Open table "%sГрафика_вся\брак\AllTables.tab"', [work]);
      SafeCloseTable('Дежурная_карта_брак');
      if not FileExists(GetAppFolder+'Base\Дежурная_карта_брак.tab') then
        miDo('Commit table AllTables As "%sДежурная_карта_брак.tab"', [GetAppFolder+'Base\']);
      miDo('Open table "%s"', [GetAppFolder+'Base\Дежурная_карта_брак.tab']);

      Info(' Обновление брака...');
      miDo('select * from AllTables Order By КН into brak NoSelect');//выбираем все объекты
      pg.Max := miEvalInt('TableInfo(AllTables, %d)', [TAB_INFO_NROWS]);
      i := 1;
      While i <= pg.Max do
      begin
        miDo('Fetch Rec %d From brak', [i]);
        cn := MiEval('brak.КН');
        miDo('select * from brak where КН="%s" into selbrak NoSelect', [cn]);
        miDo('select * from Дежурная_карта_брак where КН="%s" into selbase NoSelect', [cn]);
        kz := miEvalInt('TableInfo(selbrak, %d)', [TAB_INFO_NROWS]);//число выборок в выгрузке
        kb := miEvalInt('TableInfo(selbase, %d)', [TAB_INFO_NROWS]);//число выборок в базе
        s_z := miEval('selbrak.Дата_пакета');//дата-строка в выгрузке

        if kb > 0 then//если в базе есть такие кадастровые номера
        begin//замена
          s_b := miEval('selbase.Дата_пакета');//Дата пакета в базе
          if s_z > s_b then//сравниваем по дате, более новую выгрузку заносим в базу
          begin
            miDo('Insert Into дежурная_карта_брак Select * From selbrak');//пакет-загрузка
            miDo('Delete from selbase');//объект в базе удаляем всегда
          end;
        end else//этого номера в базе нет, добавляем все выбранные части КН
          miDo('Insert Into дежурная_карта_брак Select * From selbrak');
        Inc(i, kz);
        pg.Position := i;
      end;
      pg.Max := 0;
      SafeCloseTable('brak');
      SafeCloseTable('selbase');

      miDo('Select * from AllTables, Дежурная_карта where'+
        ' AllTables.КН=Дежурная_карта.КН Into TempTab NoSelect');
      miDo('Update TempTab Set В_работу = "OK" DropIndex Auto');
      miDo('Select * from AllTables Where В_работу = "OK" into TempTab NoSelect');
      miDo('Delete from TempTab');
      miDo('Commit table AllTables');

      Info(' Поиск графики в Дежурной карте...');
      miDo('Select * from Дежурная_карта_брак, Дежурная_карта where'+
        ' Дежурная_карта_брак.КН=Дежурная_карта.КН Into TempTab NoSelect');
      miDo('Update TempTab Set В_работу = "OK" DropIndex Auto');
      miDo('Select * from Дежурная_карта_брак Where В_работу = "OK" into TempTab NoSelect');
      miDo('Delete from TempTab');
      miDo('Commit table Дежурная_карта_брак');
      miDo('Pack Table Дежурная_карта_брак Graphic Data');
      miDo('Browse * From Дежурная_карта_брак');

      dir := GetAppFolder + 'Base\Новый брак\';
      ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [dir])), SW_HIDE, true, 0);//удаление каталога со всем содержимым и подкаталогами
      ForceDirectories(dir);
      miDo('Select * from Дежурная_карта_брак Where Дата_проверки="%s" into TempTab noSelect',
        [miEval('AllTables.Дата_проверки')]);
      miDo('Commit table TempTab As "%sAllTables.tab"', [dir]);
      miDo('Close table TempTab');
      if FileExists(format('%sГрафика_вся\брак\Самопересечение.TAB', [work])) then
//        ExShellExecute('cmd', PChar(format('/C xCopy /b /Y "%sГрафика_вся\брак\Самопересечение.*" "%s"',
        ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%sГрафика_вся\брак\Самопересечение.*" "%s"',
         [work, dir])), SW_HIDE, false, 0);
      SafeCloseTable('AllTables');
    end;
    info(' Готово');
  end;

  if cbRunUpdate.Checked then
  begin
    if not TableExists('AllTables') then
      miDo('Open table "%sГрафика_вся\AllTables.tab"', [work]);
    mnUpdateBaseClick(nil);
  end;
end;

procedure TServiceMapForm.mnExamenClick(Sender: TObject);
var
  sl : TStringList;
  i : Integer;
  t : Cardinal;
begin
{$IFDEF DEBUG} path := 'D:\temp\'; {$ENDIF}
  if not GetFolder(path, Handle) then exit;
  path := '\\?\' + path;
  t := GetTickCount;

  ExCreateProcess(format('"%sNormalization.exe" "%s" "%s"',
    [GetAppFolder, path, CatZip.Text]), '', SW_NORMAL, true);

  sl := TStringList.Create;
  try
    path := Copy(path, 1, Length(path)-1) + '_' +  CurrentData + '\';//отсекаем слеш
    FindFolders(path, sl, false);
    for i := 0 to sl.Count-1 do//проходим по всем каталогам
      if not CText(sl[i], path + 'MID_MIF') then//этот каталог не удаляем
        ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [sl[i]])), SW_HIDE, true, 0);//удаление каталога
  finally
    sl.Free;
     Note(format('процес завершен.%sВремя обработки %d минут',
        [#10, Round((gettickCount - t) / 60000)]));
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnPrepareLettersClick(Sender: TObject);
begin
//  if IsExcelFileOpen('d:\temp\МС.xlsx') then
//    ShowMessage('Файл Excel сейчас открыт и занят другом процессом!')
//  else
//    ShowMessage('Файл свободен, его можно открывать/редактировать.');
//    exit;
  LettersForm.ShowModal;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnCompareXMLClick(Sender: TObject);
type
  Pxls = ^Txls;
  Txls = Record//csv : CadastralNumber	ObjectType	DATEFORM(FoundationDate)	NUMBERINPUTINFO	DATEINPUTINFO
    del : Integer;
    ind : Integer;
    trunkline : Byte;//признак обрезанных данных
    CadastralNumber : ShortString;
    FoundationDate : ShortString;//DATEFORM
    NUMBERINPUTINFO : ShortString;
    DATEINPUTINFO : ShortString;
    full : WideString;
  end;

  PTransactions = ^TTransactions;
  TTransactions = Record//csv : CadastralNumber	ObjectType	DATEFORM(FoundationDate)	NUMBERINPUTINFO	DATEINPUTINFO
    id : Integer;
    registration_date : tdatetime;
    price : Double;//признак обрезанных данных
    cadastral_number : ShortString;
  end;

var
  i, j, k, n, m, c_file, c_area, c_add, kk : Integer;
  s, s2, s3, catzem, csvfile, title, bottomxml, work, kn, ft : String;
  xml, sl, stl, csv, log, ListObjects, FieldsTypes : TStringList;
  xls, xlscopy : TList;
  xl, xl2 : Pxls;
  sa, sb : StringArray;
  ini: TMemIniFile;
  mc : TMatchCollection;

  function ReversData(sa : String) : String;
  begin
    sb := SplitStr(sa, '.');
    if Length(sb) = 3 then
      Result := sb[2] + '-' + sb[1] + '-' + sb[0]//переворачивам дату
    else
      Result := sa;
    SetLength(sb, 0);
  end;

  function SortXls(Item1, Item2: Pxls): Integer;
  begin
    if Item1.CadastralNumber < Item2.CadastralNumber then
      Result := -1
    else if Item1.CadastralNumber > Item2.CadastralNumber then
      Result := 1
    else
    if Item1.FoundationDate < Item2.FoundationDate then
      Result := -1
    else if Item1.FoundationDate > Item2.FoundationDate then
      Result := 1
    else
    if Item1.DATEINPUTINFO < Item2.DATEINPUTINFO then
      Result := -1
    else if Item1.DATEINPUTINFO > Item2.DATEINPUTINFO then
      Result := 1
    else
    if Item1.NUMBERINPUTINFO < Item2.NUMBERINPUTINFO then
      Result := -1
    else if Item1.NUMBERINPUTINFO > Item2.NUMBERINPUTINFO then
      Result := 1
    else
      Result := 0;
  end;

  function SortXls_2(Item1, Item2: Pxls): Integer;
  begin
    if Item1.NUMBERINPUTINFO < Item2.NUMBERINPUTINFO then
      Result := -1
    else if Item1.NUMBERINPUTINFO > Item2.NUMBERINPUTINFO then
      Result := 1
    else
    if Item1.DATEINPUTINFO < Item2.DATEINPUTINFO then
      Result := -1
    else if Item1.DATEINPUTINFO > Item2.DATEINPUTINFO then
      Result := 1
    else
      Result := 0;
  end;

  function SortTransactions(Item1, Item2: PTransactions): Integer;
  begin
    if Item1.registration_date < Item2.registration_date then
      Result := -1
    else if Item1.registration_date > Item2.registration_date then
      Result := 1
    else
      if Item1.price < Item2.price then
        Result := -1
      else if Item1.price > Item2.price then
        Result := 1
      else
        Result := 0;
  end;

  procedure Clear_xls(var list : TList);
  var
    i : Integer;
  begin
    for i := list.Count-2 downto 0 do
    begin
      if Pxls(list[i]).del = 1 then
      begin
        Dispose(list.Items[i]);//удаляем выбранные уникальные значения
        list.Delete(i);
      end;
    end;
  end;

  procedure ChackDuplicates(var list : TList);
  var
    i : Integer;
  begin
    i := 0;
    While i < list.Count-1 do
    begin
      xl := list[i];
      xl2 := list[i+1];
      While xl.CadastralNumber = xl2.CadastralNumber do//двигаемся по дублям
      begin
        if (xl.CadastralNumber = xl2.CadastralNumber) and
           (xl.FoundationDate = xl2.FoundationDate) then
          begin
            xl2.del := 1;//помечаем для удаления
            log.Add('Удален дубликат'+xl.CadastralNumber+#9+
              xl2.FoundationDate+#9+xl2.NUMBERINPUTINFO+#9+xl2.DATEINPUTINFO);
          end;
        Inc(i);
        xl := list[i];
        xl2 := list[i+1];
      end;
      Inc(i);
    end;
    Clear_xls(list);
  end;

  //в качестве параметра дается xls или xlscopy
  function PreparePacket(var list : TList) : Boolean;
  var
    i, k : Integer;
  begin
    stl.Clear;
    stl.Add(title);
    i := 0;
    While i < list.Count-1 do//заполняем массив данными из списка уникальнх КН
    begin
      xl := list[i];
      xl2 := list[i+1];
      if xl.CadastralNumber <> xl2.CadastralNumber then
      begin
        xl.del := 1;//помечаем для удаления
        stl.Add(xl.full);//заносим в выборку
        Inc(i);
      end else
      While xl.CadastralNumber = Pxls(list[i]).CadastralNumber do//двигаемся по дублям
        Inc(i);
    end;
    Clear_xls(list);

    i := 0;
    While i < list.Count-1 do//дополняем массив данными из списка дублей
    begin
      xl := list[i];
      xl2 := list[i+1];
      if xl.CadastralNumber = xl2.CadastralNumber then//КН дублируются
      begin
        if xl.FoundationDate < xl2.FoundationDate then
        begin
          xl.del := 1;//помечаем для удаления
          stl.Add(xl.full);//заносим в выборку
          Inc(i);
        end;
        While xl.CadastralNumber = Pxls(list[i]).CadastralNumber do//двигаемся по дублям
          Inc(i);
      end;
    end;

    Clear_xls(list);
    Result := stl.count > 1;
  end;

  procedure CreateExcel;
  var
    ExcelApp, Workbook, WorkSheet, range : OleVariant;
    scol, colname : String;
    col, row, i, j, n, m, k, nn : integer;

    procedure ActivateSheet(n : Integer);
    begin
      WorkBook.WorkSheets.Item[n].Activate;
      WorkSheet := WorkBook.WorkSheets.Item[n];
      col := WorkSheet.UsedRange.Columns.Count;//берем размеры на новом листе
      row := WorkSheet.UsedRange.Rows.Count;
      WorkSheet.Columns[1].ColumnWidth := 20;
      WorkSheet.Columns[2].ColumnWidth := 12;
      WorkSheet.Columns[3].ColumnWidth := 12;
      WorkSheet.Columns[5].ColumnWidth := 12;
      WorkSheet.Columns[6].ColumnWidth := 12;
      WorkSheet.Columns[7].ColumnWidth := 12;
      WorkSheet.Columns[8].ColumnWidth := 20;

      ExcelApp.ActiveWindow.SplitColumn := 0;
      ExcelApp.ActiveWindow.SplitRow := 1;
//      ExcelApp.ActiveWindow.FreezePanes := True;
  end;

    procedure List2Xls(sheetname : String);
    var
      i, n, r, part, c : Integer;
      sl_ : TStringList;
      s, colname : String;

    procedure SetNumberFormat;//устанавливаем числовой формат для площадей
    begin
      if c_area > 0 then
      begin
        colname := ExcelNum2Str(WorkSheet, c_area);
        range := WorkSheet.range[format('%s:%s', [colname, colname])];
        range.NumberFormat := '0';
      end;
    end;

    begin//List2Xls
      Info('Excel ' + sheetname);
      WorkBook.WorkSheets.Item[k].Name := sheetname;

      sa := SplitStr(title, '~');//строка-заголовок
      c := High(sa);
      s := ExcelNum2Str(WorkSheet, c+1);//последняя колонка
      SetLength(sa, 0);
      sl_ := TStringList.Create;
      part := 1;
      n := 1;
      try
        r := 0;
        pg.Max := stl.Count;

        While r < stl.Count do//заполняем массив данными из списка значений
        begin
          if r mod 100 = 0 then
            pg.Position := r;
          sl_.Add(stl[r]);
          sa := SplitStr(stl[r], '~');
          if (c_area > 0) and ((sa[c_area-1] = '') or (sa[c_area-1] = '0')) then
            Log.Add('Нет площади ' + sa[0] + #9 + sa[c_area-1]);
          SetLength(sa, 0);
          Inc(r);

          if r mod 100000 = 0 then//ограничиваем размер XML
          begin
            sl_.SaveToFile(temppath + 'tempsheet.txt');
            ExcelImportText(ExcelApp, temppath+'tempsheet.txt', ft, '~', '1251');
            SetNumberFormat;
            sl_.Clear;
            sl_.Add(title);//не проверено!!!
            n := r;
            WorkBook.WorkSheets.Item[k].Name := format('Пакет_%dчасть_%d', [k, part]);
            Inc(k);
            Inc(part);
            WorkBook.WorkSheets.Add(After:=ExcelApp.ActiveSheet, Count:=1);
            ActivateSheet(k);
          end;
        end;

        if sl_.Count > 0 then
        begin
          sl_.SaveToFile(temppath + 'tempsheet.txt');
          ExcelImportText(ExcelApp, temppath+'tempsheet.txt', ft, '~', '1251');
          SetNumberFormat;
          if part > 1 then
            WorkBook.WorkSheets.Item[k].Name := format('Пакет_%dчасть_%d', [k, part]);
        end;
      finally
        sl_.Free;
      end;
      pg.Max := 0;
      Inc(k);
    end;

    function CreateTitleXML : String;
    var
      g : TGUID;
      codeobject, guid : String;
    begin
      CoCreateGuid(g);
      guid := GUIDToString(g);
      guid := Copy(guid, 2, Length(guid)-2);
      if CText(catzem,	'Buildings') then codeobject := '002001002000'
      else
       if CText(catzem,	'Flats') then codeobject := '002001003000'
      else
       if CText(catzem,	'Constructions') then codeobject := '002001004000'
      else
        if CText(catzem,	'Uncompleteds') then codeobject := '002001005000'
      else
        if CText(catzem,	'CarParkingSpaces') then codeobject := '002001009000'
      else
        if CText(catzem,	'Parcels') then codeobject := '002001001000'
      else
       codeobject := '000000000000';

      Result := '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+
      '<ListForRating Version="06">'+
      format('  <ListInfo DateForm="%s" GUID="%s">', [CurrentData, guid])+
      '    <ListType>4</ListType>'+
      '    <Region>66</Region>'+
      '    <ObjectsType>'+
      format('      <ObjectType>%s</ObjectType>', [codeobject])+
      '    </ObjectsType>'+
      '  </ListInfo>'+
      '  <Objects>'+
      format('<%s>', [ListObjects[kk]]) + lf;
    end;

    function CreateXML(packet : String) : Integer;
    const
     KTPN_UTF8: array[0..2] of Byte = ($E2, $80, $A9);
    var
      i, j : Integer;

      procedure CorrectXML;
      var
        sl : TStringList;
        s, s2, catzem : String;
        i, n, m : Integer;
      begin
        if cbCategories.ItemIndex > 0 then
          exit;
        sl := TStringList.Create;
        try
          sl.LoadFromFile(work + packet);
          s := '';
          for i := 0 to sl.Count-1 do
            s := s + Trim(sl[i]);//все в одну строку

          n := Pos('<Objects>', s);
          n := PosEx('<', s, n+12);//добавляем длину больше длины <Objects> и сразу находим начало catzem
          m := PosEx(' ', s, n+1);
          catzem := Copy(s, n+1, m-n-1);
          //добавим перевод строк для правильной работы Matches
          s2 := format('</%s>', [catzem]);
          s := StringReplace(s, s2, s2+#10, [rfReplaceAll]);

          Info(packet + ' Корректировка SubParcels');
          mc := reg.Matches(s, '<SubParcels>.+</SubParcels>', [roIgnoreCase, roMultiLine]);
          n := mc.Count;
          pg.Max := n;
          for i := 0 to n-1 do
          begin
            pg.Position := i + 1;
            s := StringReplace(s, mc.Item[i].Value, '', [rfReplaceAll]);
          end;

          Info(packet + ' Корректировка ZonesAndTerritories');
          mc := reg.Matches(s, '<ZonesAndTerritories>.+</ZonesAndTerritories>', [roIgnoreCase, roMultiLine]);
          m := mc.Count;
          pg.Max := m;
          for i := 0 to m-1 do
          begin
            pg.Position := i + 1;
            s := StringReplace(s, mc.Item[i].Value, '', [rfReplaceAll]);
          end;

          sl.Text := s;
        finally
          packet := ChangeFileExt(packet, 'испр.xml');
          sl.SaveToFile(work + packet);
          sl.Free;
          log.Add(format('Пакет %s исправлено: SubParcels %d ZonesAndTerritories %d', [packet, n, m]));
        end;
      end;

    begin
      Result := 0;
      s3 := TEncoding.UTF8.GetString(KTPN_UTF8);
      Info('Создание файла ' + packet);
      pg.Max := row;
      xml.Text := CreateTitleXML;
      for i := 2 to row do
      try
        if i mod 100 = 0 then
          pg.Position := i;
        s := WorkSheet.Cells[i, col];//исходный файл
        sl.LoadFromFile(s, TEncoding.UTF8);
        s := '';
        for j := 0 to sl.Count-1 do
          s := s + Trim(sl[j]);

        s := StringReplace(s, s3, '', [rfReplaceAll]);
        //добавим перевод строк для правильной работы Matches
        s2 := format('</%s>', [catzem]);
        s := StringReplace(s, s2, s2 + lf, [rfReplaceAll]);
        kn := WorkSheet.Cells[i, 1];
        s2 := format('<%s CadastralNumber="%s".+</%s>', [catzem, kn, catzem]);
        mc := reg.Matches(s, s2, [roIgnoreCase, roMultiLine]);
        if mc.count > 0 then//выдергиваем из файла нужную категорию земель
        begin
          xml.Add(mc.Item[0].Value + lf);
          xl := AllocMem(SizeOf(xl^));
          xl.CadastralNumber := kn;
          xl.NUMBERINPUTINFO := String(WorkSheet.Cells[i, 4]);
          xl.DATEINPUTINFO := String(WorkSheet.Cells[i, 5]);
          xls.Add(xl);
        end else
          log.Add('файл '+String(WorkSheet.Cells[i, col])+#10+'  КН '+kn+' Ошибка парсинга');
      except
        on e : Exception do
        begin
          log.Add(format('%s в строке %d', [e.Message, i]));
          exit;
        end;
      end;
      xml.Add(bottomxml);
      xml.SaveToFile(work + packet, TEncoding.UTF8);
      CorrectXML;
      pg.Max := 0;
      ExCreateProcess(format('"%s" a -tzip "%s" "%s"',
        [CatZip.Text, work + ChangeFileExt(packet, '.zip'), work + packet]), '', SW_HIDE, false);
      Result := 1;
    end;

    procedure CreateTelDaXLS(packet : String);
    var
      ExcelApp_, Workbook_, WorkSheet_ : OleVariant;
      DataParcels : OleVariant;
      r : Integer;
    begin
      Info('Создание файла ' + packet);
      r := WorkSheet.UsedRange.Rows.Count;
      range := WorkSheet.range[format('A1:%s%d', [ExcelNum2Str(WorkSheet, 5), r])];
      DataParcels := range.value;

      ExcelApp_ := CreateOleObject('Excel.Application');
      ExcelApp_.DisplayAlerts := false;
      ExcelApp_.Application.EnableEvents := false;
      try
//        ExcelApp_.CutCopyMode := False;//Отменяет режим выреза или копирования и удаляет перемещаемую границу
        Workbook_ := ExcelApp_.WorkBooks.Add;//Создаем Книгу (Workbook)
        WorkSheet_ := WorkBook_.WorkSheets.Item[1];

        range := WorkSheet_.range[format('B3:%s%d', [ExcelNum2Str(WorkSheet, 5+1), r+2])];
        range.value := DataParcels;
        WorkSheet_.Columns['C:D'].Select;
        ExcelApp_.Selection.Delete(Shift:=xlToLeft);//удаление колонок
        WorkSheet_.Columns['C:C'].Select;
        ExcelApp_.Selection.Cut;//вырезка колонок
        WorkSheet_.Columns['E:E'].Select;
        ExcelApp_.Selection.Insert(Shift:=xlToRight);//вставка колонок
        WorkSheet_.Columns[2].ColumnWidth := 20;
        WorkSheet_.Columns[3].ColumnWidth := 20;

        WorkSheet_.Cells[1, 1].Value := '&T0'; // заполнение первого столбца (3 строки)
        WorkSheet_.Cells[2, 1].Value := '&H1';
        WorkSheet_.Cells[3, 1].Value := '&H2';

        WorkSheet_.Cells[2, 2].Value := 'cadastralNumber';
        WorkSheet_.Cells[2, 3].Value := 'additional_62441';
        WorkSheet_.Cells[2, 4].Value := 'additional_62439';

        WorkSheet_.Cells[3, 2].Value := 'кадастровый номер';
        WorkSheet_.Cells[3, 3].Value := 'Дата входящего письма';
        WorkSheet_.Cells[3, 4].Value := 'Номер входящего письма';

        DeleteFile(work + packet + '.xlsx');
        WorkBook_.SaveAs(Filename:=work + packet, FileFormat:= xlOpenXMLWorkbook, CreateBackup:=False);
        ExcelApp_.Workbooks.Close;
        ExcelApp_.Quit;
      finally
        VarClear(DataParcels);
        ExcelApp_ := Unassigned;
        Workbook_ := Unassigned;
        WorkSheet_ := Unassigned;
      end;
    end;

  begin//CreateExcel
    Info('Подготовка пакетов...');
    ExcelApp := CreateOleObject('Excel.Application');
    try
      ExcelApp.Application.EnableEvents := false;
      ExcelApp.Displayalerts := false;
      ExcelApp.CutCopyMode := False;
      Workbook := ExcelApp.WorkBooks.Add;//Создаем Книгу (Workbook)
      ActivateSheet(1);

      k := 1;
      if PreparePacket(xlscopy) then
        List2Xls('Пакет_' + IntToStr(k));
      Repeat
        if not PreparePacket(xlscopy) then break;
        WorkBook.WorkSheets.Add(After:=ExcelApp.ActiveSheet, Count:=1);
        ActivateSheet(k);
        List2Xls('Пакет_' + IntToStr(k));
      Until false;

      nn := 0;//число сформированных пакетов
      n := Workbook.Sheets.Count;
      for i := 1 to n do
      begin
        ActivateSheet(i);
        CreateTelDaXLS(format('%s_Пакет_%d', [ListObjects[kk], i]));
        //очищаем список
        for j := 0 to xls.Count-1 do
          Dispose(xls.Items[j]);
        xls.Clear;

        nn := nn + CreateXML(format('%s_Пакет_%d.xml', [ListObjects[kk], i]));

        //заполняем список
        xls.Sort(@SortXls_2);
        xl := AllocMem(SizeOf(xl^));
        xls.Add(xl);//вставим пустую строку для проверки последнего значения

        xl := xls[0];
        for j := 1 to xls.Count-1 do
        begin
          xl2 := xls[j];
          if (xl.NUMBERINPUTINFO <> xl2.NUMBERINPUTINFO) or
            (xl.DATEINPUTINFO <> xl2.DATEINPUTINFO) then//последняя запись пустая сработает
          begin
//            s2 := format('update zisadmin.ZAPROS16ST set PRIM=''В ОБРАБОТКЕ'''+
//              ' Where (NUMBERINPUTINFO=''%s'') and (DATEINPUTINFO=''%s'')',
//              [xl.NUMBERINPUTINFO, xl.DATEINPUTINFO]) + ' and (question_id=%d)';
////            LettersForm.UpdatePrim(cad, s2);
//            LettersForm.UpdatePrim(s2);
//            LettersForm.UpdatePrimXLS(s2);
            xl := xl2;
          end else
        end;

      end;
      pg.Max := 0;
      log.Add(format('Сформировано пакетов: %d.', [nn]));
    finally
      ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
      ExcelApp.Visible := true;
      WorkBook.SaveAs(Filename := path + 'RESULT\' + ListObjects[kk],
        FileFormat:= xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False);

      ExcelApp := Unassigned;
      Workbook := Unassigned;
      WorkSheet := Unassigned;
      range := Unassigned;
    end;
  end;

begin//mnCompareXMLClick
{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\письма';{$ENDIF}
{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\письма\Письма2026_03_12';{$ENDIF}//выбрать All

//  if path = '' then
//    path := 'W:\ЦГКО\IT\Outbox\Статья 16\2024\';
  if not GetFolder(path, Handle) then exit;//выбираем каталог All

  ListObjects := TStringList.Create;//категория земель
  FieldsTypes := TStringList.Create;
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
    if CText(cbCategories.Text, 'oks') then
      ListObjects.CommaText := ini.ReadString('oks', 'names', '')
    else
      ListObjects.Add(cbCategories.Text);
    for kk := 0 to ListObjects.Count-1 do
//    for kk := 1 to 1 do//tr!!!
    begin
      cn.CommaText := ini.ReadString(ListObjects[kk], 'nodes', '');
      if cn.Count = 0 then
      begin
        ListObjects.Free;
        Exc('Файл ServiceMap.ini не содержит информацию об узлах XML файла для категории ' + ListObjects[kk]);
      end;
      catzem := path + 'RESULT\' + ListObjects[kk] + '.xlsm';
      if FileExists(catzem) then
        if not DeleteFile(catzem) then
        begin
          ListObjects.Free;
          Exc(format('Файл "%s.xlsm" открыт в Excel.%s Закройте файл и повторите попытку.', [catzem, #10]));
        end;
    end;
    FieldsTypes.CommaText := ini.ReadString('FieldsTypes', 'nodes', '');
    if FieldsTypes.Count = 0 then
      Exc('Ошибка чтения файла ServiceMap.ini в секции FieldsTypes');
    FieldsTypes.Sort;
  finally
    ini.Free;
  end;

  log := TStringList.Create;
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
    for kk := 0 to ListObjects.Count-1 do
//    for kk := 1 to 1 do//tr!!!
    begin
      bottomxml := format('%s%s</%s>%s%s%s</Objects>%s%s</ListForRating>',
        [#9, #9, ListObjects[kk], #13, #10, #9, #13, #10]);
      work := path + 'RESULT\' + AddSlesh(ListObjects[kk]);
      ForceDirectories(work);
      stBar.Panels[0].Text := ListObjects[kk];
      log.Add(ListObjects[kk]);
      cn.CommaText := ini.ReadString(ListObjects[kk], 'nodes', '');
      c_area := 0;//колонка с площадью
      c_add := 0;//число колонок для сравнения значений
      catzem := Copy(ListObjects[kk], 1, Length(ListObjects[kk])-1);
////////////////////подготавливаем список узлов и аттрибутов////////////////////
      SetLength(neednode, cn.Count);
      n := 0;
      for i := 0 to cn.Count-1 do//формируем список узлов и аттрибутов для колонок EXCEL
      begin
        sa := SplitStr(cn[i], ':');//разбиваем на название узла и аттрибуты
        neednode[i].node := sa[0];//добавляем название узла
        neednode[i].pos := n;//позиция узла
        Inc(n);

        if neednode[i].node[1] = '+' then
        begin
          neednode[i].node := Copy(neednode[i].node, 2, Length(neednode[i].node));
          Inc(c_add);
          SetLength(neednode[i].check, Length(neednode[i].check)+1);
          neednode[i].check[Length(neednode[i].check)-1] := neednode[i].pos+1;
          if CText(neednode[i].node, 'Area') then
            c_area := n;//колонка с площадью
        end;

        if Length(sa) > 1 then
        begin
          Inc(n, Length(sa)-2);
          SetLength(neednode[i].attr, Length(sa)-1);
          for j := 1 to High(sa) do//заносим список аттрибутов
          begin
            if sa[j][1] = '+' then
            begin
              sa[j] := Copy(sa[j], 2, Length(sa[j]));
              Inc(c_add);
              SetLength(neednode[i].check, Length(neednode[i].check)+1);
              neednode[i].check[Length(neednode[i].check)-1] := neednode[i].pos + j;
            end;
            neednode[i].attr[j-1] := sa[j];
          end;
        end;

        SetLength(sa, 0);
      end;

      title := '';//строка-заголовок для ручного разбора
      for i := Low(neednode) to High(neednode) do
      begin
        if Length(neednode[i].attr) > 0 then
          for j := Low(neednode[i].attr) to High(neednode[i].attr) do
            title := title + neednode[i].attr[j] + '~'
        else
          title := title + neednode[i].node + '~';
      end;
      title := title + 'XML name';//заголовок для XLS файла
      //типы полей (стлбцоы) в XLS
      ft := '';
      sa := SplitStr(title, '~');//разбиваем заголовок
      for i := Low(sa) to High(sa) do
      begin
        s := FieldsTypes.Values[sa[i]];
        if s = '' then
          s := '1';
        ft := ft + s + ',';
      end;
      SetLength(sa, 0);
      SetLength(ft, Length(ft)-1);

      xml := TStringList.Create;
      sl := TStringList.Create;
      stl := TStringList.Create;
      csv := TStringList.Create;
      xlscopy := TList.Create;
      xls := TList.Create;

      try
        FindRecursive(AddSlesh(path), '*.xml', false, xml);
        if xml.Count = 0 then
          Exc('Нет файлов XML в каталоге ' + path);
        //читаем XML и создаем список для сортировки
        for i := 0 to xml.Count-1 do
        begin
          info(ExtractFileName(xml[i]));
          //правка машиномест
          if CText(ListObjects[kk], 'CarParkingSpaces') then
          begin
            sl.LoadFromFile(xml[i]);
            for j := 0 to sl.Count-1 do
              sl[j] := StringReplace(sl[j], 'CarParkinSpace', 'CarParkingSpace', [rfReplaceAll]);
            sl.SaveToFile(xml[i]);
          end;
          sl.Clear;
          ReadOXmlPDOM(xml[i], ListObjects[kk], log, sl);

          for j := 0 to sl.Count-1 do
          begin
            mc := reg.Matches(sl[j], '\d\d:\d\d:\d{1,7}:\d{1,}', [roIgnoreCase, roMultiLine]);
            if mc.Count = 0 then
              log.Add(format('Ошибка кадастрового номера объекта № %d файле %s',
                [j+1, xml[i]]))
            else
            begin
              csvfile := ChangeFileExt(xml[i], '.csv');
              if FileExists(csvfile) then
              begin//открываем файл CSV и достаем из него данные
                csv.LoadFromFile(csvfile);
                for n := 0 to csv.Count-1 do
                  if Pos(mc[0].Value + ';', csv[n]) = 1 then//нашли вхождение КН
                  try
                    xl := AllocMem(SizeOf(xl^));
                    sa := SplitStr(csv[n], ';');//CadastralNumber,ObjectType,DATEFORM,NUMBERINPUTINFO,DATEINPUTINFO
                    xl.CadastralNumber := sa[0];
                    xl.FoundationDate := ReversData(sa[2]);
                    xl.NUMBERINPUTINFO := sa[3];
                    xl.DATEINPUTINFO := ReversData(sa[4]);
                    s := sa[3];
                    s2 := ReversData(sa[4]);
                    SetLength(sa, 0);

                    sa := SplitStr(sl[j], '~');
                    sa[2] := xl.FoundationDate;
                    sa[3] := s;
                    sa[4] := s2;
                    xl.full := sa[0];
                    for k := Low(sa)+1 to High(sa) do
                      xl.full := xl.full + '~' + sa[k];
                    xl.full := xl.full + '~' + xml[i];
                    xl.trunkline := Integer(sl.Objects[j]);
                    xls.Add(xl);
                    break;
                  finally
                    SetLength(sa, 0);
                  end;
              end;
            end;
          end;
        end;

        if xls.Count = 0 then
          log.Add('Нет данных') else
        begin
          xls.Sort(@SortXls);
          xl := AllocMem(SizeOf(xl^));
          xls.Add(xl);//вставим пустую строку
          ChackDuplicates(xls);

          for j := 0 to xls.Count-1 do//создаем копию данных
          begin
            xl := xls[j];
            xl.ind := j;
            xl2 := AllocMem(SizeOf(xl2^));
            xl2^ := xl^;
            xl.trunkline := 0;//обнуляем чтобы не использовать информацию
            xlscopy.Add(xl2);
          end;

          k := 1;//распределяем по пакетам чтобы выявить данные для ручного разбора
          while PreparePacket(xls) do//xls при этом очищается, остается xlscopy
            Inc(k);
  //////////////////////создаем копию дублей в StringGrid/////////////////////////
          if xls.count > 1 then
          begin
            Manual := TManual.Create(nil);
            try//настраиваем форму ручного разбора
              Manual.Caption := 'Ручной разбор ' + ListObjects[kk];
              sa := SplitStr(title, '~');//строка-заголовок
              Manual.sg.RowCount := xls.count;
              Manual.sg.ColCount := Length(sa)+2;
              for i := Low(sa) to High(sa) do
                Manual.sg.Cells[i, 0] := sa[i];
              Manual.sg.Cells[High(sa)+1, 0] := 'Взать';
              Manual.sg.Cells[High(sa)+2, 0] := 'Сообщения';
              Manual.sg.ColWidths[0] := 120;
              SetLength(sa, 0);
              for j := 0 to xls.Count-2 do//последняя строка пустая
              begin
                sa := SplitStr(Pxls(xls[j]).full, '~');//строка-заголовок
                for i := Low(sa) to High(sa) do
                  Manual.sg.Cells[i, j+1] := sa[i];
                Manual.sg.Cells[High(sa)+1, j+1] := '0';
                SetLength(sa, 0);
              end;
              Manual.filename := path + 'RESULT\' + ListObjects[kk] + '_Ручной разбор.csv';

              //приблизительно заполняем колонку выборки
              j := 1;
              While j < Manual.sg.RowCount do
              begin
                kn := Manual.sg.Cells[0, j];
                While kn = Manual.sg.Cells[0, j] do
                begin
                  Inc(j);
                  if (Manual.sg.Cells[2, j] <> Manual.sg.Cells[2, j-1]) then
                    Manual.sg.Cells[Manual.sg.ColCount-2, j-1] := '1';
                end;
                Manual.sg.Cells[Manual.sg.ColCount-2, j-1] := '1'
              end;

              k := Manual.ShowModal;
              if k = mrCancel then
              begin
                log.Clear;
                Exit;
              end;

              //удаляем отмеченные КН начиная с конца
              n := log.Count;
              for j := Manual.sg.RowCount-1 downto 1 do
                if Manual.sg.Cells[Manual.sg.ColCount-2, j] = '1' then
                begin
                  xl := xls[j-1];//номер строки StringGrid соответствует индексу xls-1
                  log.Add('Удален дубликат'+xl.CadastralNumber+#9+
                    xl2.FoundationDate+#9+xl2.NUMBERINPUTINFO+#9+xl2.DATEINPUTINFO);//FoundationDate=DATEFORM
                  k := xl.ind;
                  Dispose(xlscopy.Items[k]);
                  xlscopy.Delete(k);
                end;

            finally
              Manual.Free;
            end;
          end;
          CreateExcel;
        end;

      finally
        for i := Low(neednode) to High(neednode) do
          SetLength(neednode[i].attr, 0);
        SetLength(neednode, 0);
        for i := 0 to xlscopy.Count-1 do
          Dispose(xlscopy.Items[i]);
        for i := 0 to xls.Count-1 do
          Dispose(xls.Items[i]);
        xls.Free;
        xlscopy.Free;
        sl.Free;
        stl.Free;
        xml.Free;
        csv.Free;
      end;
      stBar.Panels[0].Text := '';
    end;

    //обработка дублей
    for i := 0 to log.Count-1 do
    begin
      j := Pos('Удален дубликат', log[i]);
      if j = 1 then
      begin
        s := Trim(Copy(log[i], 16, Length(log[i])));//DATEFORM(FoundationDate)
        sa := SplitStr(s, #9);//CadastralNumber,FoundationDate(DATEFORM),NUMBERINPUTINFO,DATEINPUTINFO
//        s2 := format('Select * from zisadmin.ZAPROS16ST'+
//        ' Where (CadastralNumber=''%s'') and (DATEFORM=''%s'')',
//        [sa[0], sa[1]]) + ' and (question_id=%d) Order By ID';
//        LettersForm.ReplacePrim(s2);
        Setlength(sa, 0);
      end;
    end;

  finally
    log.SaveToFile(path + 'RESULT\' + cbCategories.Text + '_log.log');
    if log.Count > 0 then
      ExShellExecute('Notepad.exe', path + 'RESULT\' + cbCategories.Text + '_log.log', SW_NORMAL, false, 0);
    log.Free;
    ini.Free;
    ListObjects.Free;
    FieldsTypes.Free;
    info('Готово.');
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnSt16Click(Sender: TObject);
var
  i : Integer;
begin
  Form2.TempPath := TempPath;
  Form2.Caption := 'Расчет ценовых факторов ' + (Sender as TmenuItem).Caption;

  for i := 0 to Form2.mnCalc.Count-1 do
    Form2.mnCalc.Items[i].Visible := False;
  Form2.mnObjectsFromText.Visible := True;
  Form2.mnObjectsFromXls.Visible := True;

  Form2.mnTable.Visible := False;
  Form2.mnBuffers.Visible := False;

  Form2.st16 := true;
  Form2.worktable := '';
  Form2.ResFolder := '';
  Form2.ftab := '';
  Form2.Show;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnTransactionsClick(Sender: TObject);
const
  XLStitle = 'registration_type	registration_date	document_date	price	quantity	'+
    'cadastral_number	part_numerator	part_denomenator	part_right_price	share_description';
var
  i, j, n, k, m, quantity : Integer;
  s, s2, ss, xmsname, cn: String;
  sa : StringArray;
  sl, xml, xls : TStringList;
  mc, mc2 : TMatchCollection;
  ExcelApp, Workbook, WorkSheet : OleVariant;
begin
{$IFDEF DEBUG}path := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\сделки\';{$ENDIF}
//  if path = '' then
//    path := 'W:\ЦГКО\IT\Сотрудники отдела IT\Лемеш\Реестр_сделок\';
  if not GetFolder(path, Handle) then exit;
  sl := TStringList.Create;
  xml := TStringList.Create;
  xls := TStringList.Create;
  xls.Add(XLStitle);
  try
    FindRecursive(path, '*.xml', false, sl);
    if sl.Count = 0 then
      Exc('Нет файлов со сделками!');
    for i := 0 to sl.Count-1 do
    begin
//      if Pos('_sale.xml', sl[i]) > 0 then
//        xmsname := temppath + 'Sale.txt';
//      if Pos('_lease.xml', sl[i]) > 0 then
//        xmsname := temppath + 'Lease.txt';
//      if Pos('_mortgage.xml', sl[i]) > 0 then
//        xmsname := temppath + 'Mortgage.txt';
//      if Pos('_shared.xml', sl[i]) > 0 then
//        xmsname := temppath + 'Shared.txt';
      if Pos('_sale.xml', sl[i]) > 0 then
      begin
        Info('Парсим ' + ExtractFileName(sl[i]));
        xmsname := temppath + 'Sale.txt';
        xml.LoadFromFile(sl[i]);
        s := '';
        for j := 0 to xml.Count-1 do
          s := s + Trim(xml[j]);//все в одну строку
        s := StringReplace(s, '</sale_contract>', '</sale_contract>'+lf, [rfReplaceAll]);
        sa := SplitStr(XLStitle, #9);
        mc := reg.Matches(s, '<sale_contract>.+</sale_contract>', [roIgnoreCase, roMultiLine]);//блоки

        for j := 0 to mc.Count-1 do
        begin
          s2 := '';
          quantity := 0;

          for n := Low(sa) to High(sa) do
          begin
            mc2 := reg.Matches(mc.Item[j].Value, '<'+sa[n]+'>.+?</'+sa[n]+'>', []);//секции (? ленивый поиск)
            if mc2.Count > 0 then
            begin//выбираем чистое значение
              ss := StringReplace(mc2.Item[0].Value, '<'+sa[n]+'>', '', []);
              ss := StringReplace(ss, '</'+sa[n]+'>', '', []);
              s2 := s2 + ss + #9;
            end else//поиск в узлах
            begin
              ss := sa[n]+'="';
              k := Pos(ss, mc.Item[j].Value);
              if k > 0 then
              begin
                m := PosEx('"', mc.Item[j].Value, k+Length(ss));//закрывающая кавычка
                if m > 0 then
                 s2 := s2 + Copy(mc.Item[j].Value, k+Length(ss), m-k-Length(ss)) + #9;

                 if sa[n] = 'quantity' then
                   quantity := StrToInt(Copy(mc.Item[j].Value, k+Length(ss), m-k-Length(ss)));
                 if sa[n] = 'cadastral_number' then//кадастровый номер для замены
                   cn := Copy(mc.Item[j].Value, k+Length(ss), m-k-Length(ss));

              end else
                s2 := s2 + #9;
            end;
          end;
          ss := Copy(s2, 1, Length(s2)-1);
          xls.Add(ss);

          if quantity > 1 then
          begin
            mc2 := reg.Matches(mc.Item[j].Value, '\d\d:\d\d:\d{1,7}:\d{1,}', []);//кадастровые номера
            for n := 1 to mc2.Count-1 do
            begin
              s2 := StringReplace(ss, cn, mc2.Item[n].Value, []);
              xls.Add(s2);
            end;
          end;
        end;
        if xls.Count > 1 then
        begin
          Info('Создаем таблицу Excel ');
          xls.SaveToFile(temppath + 'Sale.txt');
          ExcelApp := CreateOleObject('Excel.Application');
          try
            ExcelApp.Application.EnableEvents := false;
            ExcelApp.Displayalerts := false;
            ExcelApp.WorkBooks.OpenText(Filename:='"' + xmsname + '"',
                  Origin:=1251, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote,
                  ConsecutiveDelimiter:=False, Tab:=True, Other:=False, TrailingMinusNumbers:=True);
            WorkSheet := ExcelApp.ActiveSheet;
            for k := 1 to WorkSheet.UsedRange.Columns.Count do
              WorkSheet.Columns[k].ColumnWidth := 15;
            WorkSheet.Columns[6].ColumnWidth := 18;
            ExcelApp.ActiveWindow.SplitColumn := 0;
            ExcelApp.ActiveWindow.SplitRow := 1;
            ExcelApp.ActiveWindow.FreezePanes := True;

          finally
            ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
            ExcelApp.Visible := true;
//            WorkBook.SaveAs(Filename := path + ,
//              FileFormat:= xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False);
            WorkSheet := Unassigned;
            Workbook := Unassigned;
            ExcelApp := Unassigned;
          end;
          Info('Готово');
        end else
          Note('Ошибка парсинга');

        SetLength(sa, 0);
      end;
    end;
  finally
    sl.Free;
    xml.Free;
    xls.Free;
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnGKOClick(Sender: TObject);
var
  i : Integer;
begin
  Form2.TempPath := TempPath;
  Form2.Caption := 'Расчет ценовых факторов ' + (Sender as TmenuItem).Caption;

  for i := 0 to Form2.mnCalc.Count-1 do
    Form2.mnCalc.Items[i].Visible := True;
  Form2.mnObjectsFromText.Visible := False;
  Form2.mnObjectsFromXls.Visible := False;

  Form2.mnTable.Visible := True;
  Form2.mnBuffers.Visible := True;

  Form2.st16 := false;
  Form2.worktable := '';
  Form2.ResFolder := '';
  Form2.ftab := '';
  Form2.Show;
end;
{--------------------корректировка узла <Location>-----------------------------}
procedure TServiceMapForm.mnCorrectFDPlusClick(Sender: TObject);
begin
  mnFullCorrectFDClick(Sender);
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnFullCorrectFDClick(Sender: TObject);
var
  i, j, k, n : Integer;
  dest, sub, res, s, ss : String;
  cat, sl, names : TStringList;
begin
  PrepareForm := TPrepareForm.Create(nil);
  PrepareForm.TempPath := TempPath;
  PrepareForm.Location := Sender As TmenuItem = mnCorrectFDPlus;
//  //ZIP
//{$IFDEF DEBUG}PrepareForm.edSource.text := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\ОКС_ИТОГ\';{$ENDIF}
//{$IFDEF DEBUG}PrepareForm.SourcPath := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\ОКС_ИТОГ';{$ENDIF}
//  //пакеты
//{$IFDEF DEBUG}PrepareForm.edPackets.text := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\SOKS_a_2026-02-06_3\';{$ENDIF}
//{$IFDEF DEBUG}PrepareForm.PacketPath := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\SOKS_a_2026-02-06_3\';{$ENDIF}
  //ZIP
{$IFDEF DEBUG}PrepareForm.edSource.text := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\2\ЗУ_ИТОГ\';{$ENDIF}
{$IFDEF DEBUG}PrepareForm.SourcPath :=     'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\2\ЗУ_ИТОГ';{$ENDIF}
  //пакеты
{$IFDEF DEBUG}PrepareForm.edPackets.text := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\2\SOKS_Zu_a_2026-02-19_1\';{$ENDIF}
{$IFDEF DEBUG}PrepareForm.PacketPath :=     'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\2\SOKS_Zu_a_2026-02-19_1';{$ENDIF}

  Info('Подготовка файлов FD..');
  if PrepareForm.ShowModal = mrOk then
  begin
    path := PrepareForm.edPackets.Text;//Каталог с пакетами
    if CorrectFDproc then
    begin
      s := format('/c del /F /Q  "%s"', [TempPath + '*.*']);//удалим и пересоздадим
      ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//чистим каталог
      res := path + 'SOKC_Result\';
      cat := TStringList.Create;//список номеров писем
      FindFolders(res, cat, false);
      Info('Создание НСПД и приложений....');
      pg.Max := cat.Count;
      for i := 0 to cat.Count-1 do
      begin
        pg.Position := i+1;
        sub := res + AddSlesh(ExtractFileName(cat[i]));

        s := format('/c rd /S /Q  "%s"', [sub + 'НСПД\']);//удалим и пересоздадим
        ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
        ForceDirectories(sub + 'НСПД\');

        cn.Clear;
        FindRecursive(sub, '*.zip', true, cn);//копируем приложения к актам
        for j := 0 to cn.Count-1 do
        begin
          if Pos('Приложение к Акту', cn[j]) > 0 then
            ExShellExecute('cmd', format('/C Copy /b  /Y "%s" "%s"',
              [cn[j], sub + 'НСПД\']), SW_HIDE, false, 0);
        end;

        cn.Clear;
        FindRecursive(sub, '*.ods', true, cn);//копируем акты
        for j := 0 to cn.Count-1 do
          ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',
            [cn[j], sub + 'НСПД\']), SW_HIDE, false, 0);

        names := TStringList.Create;//список коротких имен FD со ссылкой на полные
        ss := '';
        for j := 0 to cn.Count-1 do
        begin
          s := Copy(ExtractFileName(cn[j]), 1, 19);
          s[8] := '/';
          s[13] := '/';
          ss := ss + s + ',';
        end;
        names.Add(Copy(ss, 1, Length(ss)-1));
        names.SaveToFile(sub + 'НСПД\' + ExtractFileName(DellSlesh(sub)) + '.txt');
        names.Free;

        cn.Clear;
        FindRecursive(sub, '*.zip', true, cn);//ищем все архивы
        for j := 0 to cn.Count-1 do//раскладываем по каталогам для проверки уникальности
        begin
          s := TempPath + AddSlesh(IntTostr(j));
          ForceDirectories(s);
          ExCreateProcess(format('"%s" x "%s" -y -o"%s"', [CatZip.Text, cn[j], s]), '', SW_HIDE, true);
        end;

        //проверка на уникальность и изменение названия FD для дублей
        sl := TStringList.Create;
        names := TStringList.Create;//список коротких имен FD со ссылкой на полные
        try
          FindRecursive(TempPath, 'FD_*.xml', true, sl);//ищем все архивы
          for j := 0 to sl.Count-1 do//создаем список коротких имен для сортировки
            names.AddObject(ExtractFileName(sl[j]), Pointer(j));
          names.Sort;
          j := 1;
          While j < names.Count do//ищем и исправляем одинаковые названия
          begin
            k := 2;
            While (j < names.Count) and (names[j-1] = names[j]) do
            begin
              s := Copy(names[j] , 1, Length(names[j] )-4) + format('(%d).xml', [k]);
              n := Integer(names.Objects[j]);
              RenameFile(sl[n], ExtractFilePath(sl[n]) + s);
              Inc(k);
              Inc(j);
            end;
            Inc(j);
          end;
        finally
          sl.Free;
          names.Free;
        end;

        //проверка на уникальность и изменение названия COST для дублей
        sl := TStringList.Create;
        names := TStringList.Create;//список коротких имен COST со сылкой на полные
        try
          FindRecursive(TempPath, 'COST_*.xml', true, sl);//ищем все архивы
          for j := 0 to sl.Count-1 do//создаем список коротких имен для сортировки
            names.AddObject(ExtractFileName(sl[j]), Pointer(j));
          names.Sort;
          j := 1;
          While j < names.Count do//ищем и исправляем одинаковые названия
          begin
            k := 2;
            While (j < names.Count) and (names[j-1] = names[j]) do
            begin
              s := Copy(names[j] , 1, Length(names[j] )-4) + format('(%d).xml', [k]);
              n := Integer(names.Objects[j]);
              RenameFile(sl[n], ExtractFilePath(sl[n]) + s);
              Inc(k);
              Inc(j);
            end;
            Inc(j);
          end;
        finally
          sl.Free;
          names.Free;
        end;

//        s := format('/c rd /S /Q  "%s"', [TempPath + 'FD\']);//удалим и пересоздадим
//        ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
//        ForceDirectories(TempPath + 'FD\');
//
//        s := format('/c rd /S /Q  "%s"', [TempPath + 'COST\']);//удалим и пересоздадим
//        ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
//        ForceDirectories(TempPath + 'COST\');
//
//        s := format('/c rd /S /Q  "%s"', [TempPath + 'ALL\']);//удалим и пересоздадим
//        ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
//        ForceDirectories(TempPath + 'ALL\');
//
//        for j := 0 to cn.Count-1 do//чистим каталоги
//        begin
//          ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',//копируем в каталог все FD
//            [TempPath + AddSlesh(IntTostr(j)) + 'FD*.xml', TempPath + 'FD\']), SW_HIDE, true, 0);
//
//          ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',//копируем в каталог все COST
//            [TempPath + AddSlesh(IntTostr(j)) + 'COST*.xml', TempPath + 'COST\']), SW_HIDE, true, 0);
//
//          s := format('/c rd /S /Q  "%s"', [TempPath + AddSlesh(IntTostr(j))]);
//          ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
//        end;
//
//        ExCreateProcess(format('%s a -tzip -mx5 -r0 %s %s',//создание архива FD
//          [CatZip.Text, TempPath + 'ALL\FD.zip', TempPath + '\FD_*.xml']), '', SW_HIDE, true);
//
//        ExCreateProcess(format('%s a -tzip -mx5 -r0 %s %s',//создание архива COST
//          [CatZip.Text, TempPath + 'ALL\COST.zip', TempPath + '\COST_*.xml']), '', SW_HIDE, true);
//
//        ExCreateProcess(format('%s a -tzip -mx5 -r0 %s %s',//создание архива Приложения
//          [CatZip.Text, TempPath + 'ALL\Приложения.zip', TempPath + 'ALL\*.zip']), '', SW_HIDE, true);
//
//        ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',//забираем готовые Приложения
//          [TempPath + 'ALL\Приложения.zip', sub + 'НСПД\']), SW_HIDE, false, 0);
//
//        s := format('/c rd /S /Q  "%s"', [TempPath + 'FD\']);
//        ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
//        s := format('/c rd /S /Q  "%s"', [TempPath + 'COST\']);
//        ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
//        s := format('/c rd /S /Q  "%s"', [TempPath + 'ALL\']);
//        ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
      end;//конец списка номеров писем

      pg.Max := 0;
      cat.Free;
    end;
  end;
  Info('Готово');
  PrepareForm.Free;
end;
{------------------------------------------------------------------------------}
function TServiceMapForm.CorrectFDproc : Boolean;
var
  sl, fd, xml, cn, kn, packet, cat, log : TStringList;
  i, j, k, n, m, n_proc, l : Integer;
  s, ss, tab, proc, nom, aoks, sheetname, xlsname, res : String;
  zip : TZipFile;
  sa : StringArray;
  mc : TMatchCollection;

  function GetSheetNameFirst : String;
  var
    ADO : TADOConnection;
    output: ttableitems;
  begin
    ADO := TADOConnection.Create(nil);
    ADO.Provider := 'MSDASQL.1';
    ADO.ConnectionString := 'Provider=MSDASQL.1;Persist Security Info=False;'+
    'Extended Properties="DSN=Excel Files;DBQ=' + xlsname + ';DefaultDir='+
    ExtractFilePath(xlsname) + ';DriverId=1046;MaxBufferSize=2048;PageTimeout=5;"';
    ADO.Open;

    output := ADODbTables(ADO, [ttTable, ttView, ttSynonym]);
    //output := ADODbTables(ADOConnection1, [ttSystemTable, ttAccessTable]);
//    for i := Low(output) to High(output) do
//    begin
//      sl.Add(output[i].ItemName + '---' + output[i].ItemType);
//    end;
    Result := output[0].ItemName;//только имя первого листа
    output := nil;
    ADO.Close;
    ADO.Free;
 end;
 //рабочий каталог MapInfo  C:\Users\TroitskikhVV\AppData\Local\Temp\MIPRO
  procedure LoadXML;
  var
    i, j, k : Integer;
  begin
    for i := 0 to sl.Count-1 do
    begin
      xml := TStringList.Create;
      xml.LoadFromFile(sl[i], TEncoding.UTF8);
      fd.AddObject(sl[i], xml);//название xml и ссылка на его содержимое в объекте
      mc := reg.Matches(xml.text, '<CadastralNumber>\d\d:\d\d:\d{1,7}:\d{1,}</CadastralNumber>',
        [roIgnoreCase, roMultiLine]);
      kn := TStringList.Create;//список кадастровых номеров
      for j := 0 to mc.Count-1 do
      begin
        k := Pos('</CadastralNumber>', mc.Item[j].Value);//конец кадастрового номера
        s := Copy(mc.Item[j].Value, 18, k-18);
        kn.Add(s);//накапливаем кадастровые номера в xml в списке
      end;
      kn.Sort;
      cn.AddObject(sl[i], kn);//кадастровые номера
    end;
  end;

begin//CorrectFDproc
  Result := False;
//  path := '\\?\' + path;
  sl := TStringList.Create;
  fd := TStringList.Create;//список списков название xml и ссылка на его содержимое в объект
  cn := TStringList.Create;//список списков кадастровых номеров в xml (симметричен fd)
  packet := TStringList.Create;//список каталогов в каталоге "ПАКЕТЫ"
  cat := TStringList.Create;//список каталогов содержащих Приложение к Акту АОКС
  log := TStringList.Create;
  try
    FindFolders(path + 'Пакеты\', packet, false);
    if packet.Count = 0 then
      Exc('В выбранном каталоге не найдены пакеты.');

    FindRecursive(path, '*.zip', false, sl);
    for i := 0 to sl.Count-1 do//FD.zip формируется а модуле PrepareFD_Unit
      if CText(ExtractFileName(sl[i]), 'FD.zip') then//разархивация файлов XML
        ExCreateProcess(format('"%s" x "%s" -y -o"%s"',
          [CatZip.Text, sl[i], TempPath]), '', SW_HIDE, true);

    sl.Clear;//проверяем количество разархивированных файлов FD
    FindRecursive(TempPath, '*.xml', false, sl);
    if sl.Count = 0 then
      Exc('Нет FD.zip архива или XML файлов в архиве!');

    ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [path + 'SOKC_Result'])), SW_HIDE, true, 0);//удаление каталога со всем содержимым и подкаталогами
    ForceDirectories(path + 'SOKC_Result');//создаем каталог для сохранения расчетов
    ExShellExecute('cmd', format('/C xCopy /y /s /e "%s" "%s"',//копируем исходные данные из пакетов
      [path + 'Пакеты', path + 'SOKC_Result']), SW_HIDE, true, 0);

    res := path + 'SOKC_Result\';
    packet.Clear;
    FindFolders(res, packet, false);//те же пакеты в результирующем файле

    sl.Clear;
    FindRecursive(res, '*.xlsx', true, sl);
    for k := sl.Count-1 downto 0 do
      if Pos('Проверочный', ExtractFileName(sl[k])) = 1 then
        DeleteFile(sl[k]) else//удалим проверочные файлы, они не нужны
      if Pos('АОКС', ExtractFileName(sl[k])) = 1 then
        DeleteFile(sl[k]);//удалим АОКС файлы, они не нужны

    sl.Clear;
    FindRecursive(res, '*.ods', true, sl);
    for k := sl.Count-1 downto 0 do
      if Pos('Приложение к Акту', ExtractFileName(sl[k])) = 1 then
        DeleteFile(sl[k]);//удалим лишние файлы ODS, они не нужны

    miDo('Close All');
    pg.Max := packet.Count;
    for i := 0 to packet.Count-1 do
    begin
      pg.Position := i+1;
      cat.Clear;
      FindFolders(AddSlesh(packet[i]), cat, false);
      for j := 0 to cat.Count-1 do//
      begin
        Info(format('пакет %s каталог %s', [ExtractFileName(packet[i]), ExtractFileName(cat[j])]));
        sl.Clear;
        FindRecursive(AddSlesh(cat[j]), '*.xlsx', false, sl);
        if sl.count = 0 then
          Continue;
        xlsname := sl[0];//должен быть единственный файл - это наше приложение в акту
        sheetname := GetSheetNameFirst;//берем название первого листа в книге

        tab := CreateTableFromExcel(sl[0], sheetname, '');//создаем таблицу Mapinfo из Приложение к Акту АОКС
        miDo('Select * from %s where Кадастровая_стоимость__руб<>"" group by N_PROCEDURE into Proc', [tab]);
        n_proc := miEvalInt('TableInfo(Proc, %d)', [TAB_INFO_NROWS]);
        if n_proc > 0 then//число уникальных номеров процедур
        begin
          for n_proc := 1 to miEvalInt('TableInfo(Proc, %d)', [TAB_INFO_NROWS]) do//идем по группам процедур
          begin
            for k := 0 to cn.Count-1 do
              cn.Objects[k].Free;
            cn.Clear;
            for k := 0 to fd.Count-1 do
              fd.Objects[k].Free;
            fd.Clear;

            miDo('Fetch rec %d From Proc', [n_proc]);
            proc := miEval('Proc.N_PROCEDURE');
            log.Add(format('пакет %s каталог %s процедура %s',
              [ExtractFileName(packet[i]), ExtractFileName(cat[j]), proc]));
            Info(log[log.Count-1]);
            miDo('Select * from %s where Кадастровая_стоимость__руб<>"" and '+
              'N_PROCEDURE="%s" into TempTab', [tab, proc]);
            sl.Clear;
            FindRecursive(TempPath, '*.xml', false, sl);
            for k := sl.Count-1 downto 0 do
              if Pos(proc, sl[k]) = 0 then
                sl.Delete(k);//удалим лишние файлы
            LoadXML;

            if sl.count > 0 then
            begin
              for n := 1 to miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]) do
              begin
                miDo('Fetch rec %d From %s', [n, tab]);
                nom := miEval('%s.col2', [tab]);//кадастровый номер
                s := miEval('%s.N_PROCEDURE', [tab]);
                for k := 0 to cn.Count-1 do
                if s = proc then
                begin
                  kn := TStringList(cn.Objects[k]);
                  if kn.Find(nom, m) then
                    kn.Delete(m);
                end;
              end;

              Info('Корректировка XML');
              for l := 0 to fd.Count-1 do
              begin
                s := ExtractFileName(fd[l]);
                xml := TStringList(fd.Objects[l]);//список FD...
                kn := TStringList(cn.Objects[l]);

                for k := 0 to kn.Count-1 do//оставшиеся КН удаляем
                begin
                  log.Add(kn[k] + ' удален');
                  ss := format('<CadastralNumber>%s</CadastralNumber>', [kn[k]]);
                  for n := xml.Count-1 downto 0 do
                    if Pos(ss, xml[n]) > 0 then
                    begin
                      xml.Delete(n);
                      break;
                    end;
                end;

                mc := reg.Matches(xml.text, '<CadastralNumber>\d\d:\d\d:\d{1,7}:\d{1,}</CadastralNumber>',
                  [roIgnoreCase, roMultiLine]);
                if mc.Count = 0 then
                  log.Add('файл ' + s + ' удален')
                else
                  xml.SaveToFile(AddSlesh(cat[j]) + s);
              end;
            end else
              log.Add(format('Для процедуры "%s" файлы "FD_???.XML" не найдены!!!', [proc]));
          end;
          DeleteFile(xlsname);

          sl.Clear;
          FindRecursive(AddSlesh(cat[j]), '*.ods', false, sl);
          if sl.Count > 0 then
          begin
            nom := ExtractFileName(sl[0]);//АОКС-66-2024-002005_от_23.09.2024_на_2780.ods
            aoks := Copy(nom, 6, 14);//66-2024-002005
            nom := Copy(nom, 24, 10);//23.09.2024
            sa := SplitStr(nom, '.');
            nom := sa[2] + '-' + sa[1] + '-' + sa[0];
            SetLength(sa, 0);
            sa := SplitStr(aoks, '-');
            aoks := sa[0] + '/' + sa[1] + '/' + sa[2];
            SetLength(sa, 0);
          end else
            log.Add('Нет файла с расширением "ODS !!!');

          sl.Clear;
          FindRecursive(AddSlesh(cat[j]), '*.xml', false, sl);
          if sl.Count > 0 then//пустые Приложение к Акту АОКС пропускаем
          begin
            zip := TZipFile.Create;
            zip.UTF8Support := true;
            ss := Copy(ExtractFileName(xlsname), 1, 37) + '.zip';
            s := AddSlesh(cat[j]) + ss;
            zip.Open(s, zmWrite);
            for k := 0 to sl.Count-1 do
            begin
              if (Length(nom) = 10) and (Pos('FD', ExtractFileName(sl[k])) = 1) then
              begin
                xml := TStringList.Create;
                try
                  xml.LoadFromFile(sl[k], TEncoding.UTF8);
                  for l := 0 to xml.Count-1 do
                  begin
                    mc := reg.Matches(xml[l], '<Date_Doc>\d\d\d\d-\d\d-\d\d</Date_Doc>', [roIgnoreCase, roMultiLine]);
                    if mc.Count > 0 then
                      xml[l] := StringReplace(xml[l], mc.Item[0].Value,
                        format('%s', ['<Date_Doc>' + nom + '</Date_Doc>']), []);

                    mc := reg.Matches(xml[l], '<Report_Details Date="\d\d\d\d-\d\d-\d\d"', [roIgnoreCase, roMultiLine]);
                    if mc.Count > 0 then
                      xml[l] := StringReplace(xml[l], mc.Item[0].Value,
                        format('%s', ['<Report_Details Date="' + nom + '"']), []);

                    mc := reg.Matches(xml[l], '<N_Doc>АОКС-\d\d/\d\d\d\d/\d\d\d\d\d\d</N_Doc>', [roIgnoreCase, roMultiLine]);
                    if mc.Count > 0 then
                      xml[l] := StringReplace(xml[l], mc.Item[0].Value,
                        format('%s', ['<N_Doc>АОКС-' + aoks + '</N_Doc>']), []);

                    mc := reg.Matches(xml[l], 'Number="АОКС-\d\d/\d\d\d\d/\d\d\d\d\d\d">', [roIgnoreCase, roMultiLine]);
                    if mc.Count > 0 then
                      xml[l] := StringReplace(xml[l], mc.Item[0].Value,
                        format('%s', ['Number="АОКС-' + aoks + '">']), []);
                  end;
                  xml.SaveToFile(sl[k], TEncoding.UTF8);
                finally
                  FreeNil(xml);
                end;
              end;
              zip.Add(sl[k], ExtractFileName(sl[k]), zcDeflate);
              DeleteFile(sl[k]);
            end;
            zip.Close;
            zip.Free;
          end else
            log.Add(format('пакет %s каталог %s: НЕТ ДАННЫХ для архивации!!!',
              [ExtractFileName(packet[i]), ExtractFileName(cat[j])]));
        end else
          log.Add(format('пакет %s каталог %s: НЕТ ДАННЫХ "Кадастровая_стоимость__руб"',
            [ExtractFileName(packet[i]), ExtractFileName(cat[j])]));
        miDo('Drop table %s', [tab]);
      end;

    end;
    sl.Clear;
    FindRecursive(res, '*.xlsx', true, sl);
    for i := 0 to sl.Count-1 do
      DeleteFile(sl[i]);

    pg.Max := 0;
    Info('Готово');
    pg.max := 0;
    Result := True;
  finally
    s := path + 'xmlLog.log';
    log.SaveToFile(s);
    if Result then
      ExShellExecute('Notepad.exe', s, SW_NORMAL, false, 0)
    else
      DeleteFile(s);
    log.Free;
    packet.Free;
    cat.Free;
    sl.Free;
    for i := 0 to cn.Count-1 do
      cn.Objects[i].Free;
    cn.Free;
    for i := 0 to fd.Count-1 do
      fd.Objects[i].Free;
    fd.Free;
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.Contains(var id,  row, kodnp, idkp : Integer;
  var fields : TStringList; var WorkSheet, DataParcels : OleVariant;
  tab, tname, fname : String);
var
  data, range : OleVariant;
  t, s, ss : String;
  i, j, k, n, col : Integer;
  sa : StringArray;
begin
  col := id + 1;
  t := miEval('PathToTableName$("%s")', [tname]);
  Info('Проверка попадания в ' + t);
  miDo('Open table "%s"', [tname]);
  SafeCloseTable('TempTab');
  if CText(fname, 'КодНП_МО') then
    kodnp := 2;
  if CText(fname, 'ID_КП_Сады') then
    idkp := 3;

  n := id+1;
  while n < fields.Count do
  begin//ищем вхождение названия таблицы в следующих полях
    if not AnsiContainsText(fields.ValueFromIndex[n], t) then
      break;
    Inc(n);
  end;

  data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, n-id], varVariant);
  pg.Max := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
  for i := 1 to pg.Max do
  begin
    if i mod 10 = 0 then pg.Position := i;
    miDo('Fetch Rec %d From %s', [i, tab]);
    miDo('ob = %s.obj', [tab]);
    miDo('Select * from %s Where %s.Obj Contains ob Into TempTab NoSelect', [t, t]);
    k := miEvalInt('TableInfo(TempTab, %d)', [TAB_INFO_NROWS]);
    if fname = '' then
    begin//названия поля нет, учитываем только количество
      if k > 0 then
        data[i, 1] := 1
      else
        data[i, 1] := 0;
    end else
    begin//по названию поля выбираем значения
      if k > 0 then
      begin
        for j := 1 to n-id do//распределяем данные по числу полей
        begin
          s := Copy(fields.ValueFromIndex[id+j-1], 10, 100);//выбираем таблицу и поле после Contains
          sa := SplitStr(s, ';');//=fname
          data[i, j] := miEval('TempTab.%s', [sa[1]]);
          SetLength(sa, 0);
        end;
        if CText(fname, 'КодНП_МО') then
          DataParcels[i, kodnp] := data[i, 1];//для создание буферных зон
        if CText(fname, 'ID_КП_Сады') then
          DataParcels[i, idkp] := data[i, 1];//для создание буферных зон
        if CText(fname, 'НаименованиеНП_МО') then
          data[i, 1] := StringReplace(String(data[i, 1]), 'ё', 'е', [rfReplaceAll, rfIgnoreCase]);
      end else
        for j := 1 to n-id do
          data[i, j] := '';
    end;
  end;
  SafeCloseTable(t);
  s := ExcelNum2Str(WorkSheet, col);
  ss := ExcelNum2Str(WorkSheet, col+n-id-1);
  range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
  range.value := data;
  pg.Max := 0;
  VarClear(data);
  range := Unassigned;
  id := n-1;
end;

procedure TServiceMapForm.Intersects(var id, row, radius : Integer;
    var fields : TStringList; var WorkSheet, DataParcels : OleVariant;
    tab, tname, fname : String);
var
  data, range : OleVariant;
  t, s, ss, f : String;
  i, j, k, n, col : Integer;
  sa : StringArray;
begin
  col := id + 1;
  t := miEval('PathToTableName$("%s")', [tname]);
  Info('Проверка пересечения с ' + t);
  miDo('Open table "%s"', [tname]);
  SafeCloseTable('TempTab');

  n := id+1;
  while n < fields.Count-2 do
  begin//ищем вхождение названия таблицы в следующих полях
    if not AnsiContainsText(fields.ValueFromIndex[n], t) then
      break;
    Inc(n);
  end;

  data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, n-id], varVariant);
  pg.Max := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
  for i := 1 to pg.Max do
  begin
    if i mod 10 = 0 then pg.Position := i;
    miDo('Fetch Rec %d From %s', [i, tab]);
    miDo('ob = %s.obj', [tab]);
    miDo('ob=Buffer(ob, 12, %d, "m")', [radius]);
    miDo('Select * from %s Where %s.Obj Intersects ob Into TempTab NoSelect', [t, t]);

    for j := 1 to n-id do//распределяем данные по числу полей
    begin
      ss := ';';
      for k := 1 to miEvalInt('TableInfo(TempTab, %d)', [TAB_INFO_NROWS]) do
      begin
        miDo('Fetch rec %d From TempTab', [k]);
        s := Copy(fields.ValueFromIndex[id+j-1], 12, 100);//выбираем таблицу и поле после Intersects
        sa := SplitStr(s, ';');//=fname
        s := miEval('TempTab.%s', [sa[1]]);
        SetLength(sa, 0);
        if Pos(s, ss) = 0 then
          ss := ss + s + ';';
      end;
      if ss = ';' then
        ss := '';
      data[i, j] := ss;
    end;

  end;
  SafeCloseTable(t);
  s := ExcelNum2Str(WorkSheet, col);
  ss := ExcelNum2Str(WorkSheet, col+n-id-1);
  range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
  range.value := data;
  pg.Max := 0;
  VarClear(data);
  range := Unassigned;
  id := n-1;
end;

procedure TServiceMapForm.CalcOne(var ExcelApp : OleVariant);
var
  Workbook, WorkSheet, range, DataParcels, data : OleVariant;
  row, col, col2, i, j, k, n, m, bufsize, bufdist, idkp, kodnp : Integer;
  colarea, colcena, colcomment, colprocent, collat, collon, colpdf, coladr, coladr2, colpred : Integer;//нужны для работы
  sl, sl2, log, fields, dbf : TStringList;
  tab, txt, id, s, ss, analogs, respathPDF : String;
  s3 : String[200];
  semantik : Variant;
  sa : StringArray;
  xy : TPoint3D;
  ini: TMemIniFile;
  mc : TMatchCollection;

  function FindColValue(var ind : Integer; value : String) : Boolean;
  var
    i : Integer;
  begin
    Result := false;
    for i := 0 to fields.Count-1 do
      if CText(fields.ValueFromIndex[i], value) then
      begin
        ind := i+1;//учет индекса
        Result := true;
        exit;
      end;
  end;
  //поиск КН найденных в Дежурной карте
  procedure ADO;
  var
    ds: TADODataset;
    i, j, k, n : Integer;
  begin
    miDo('Open table "%s"', [GetAppFolder+'Base\Дежурная_карта.tab']);
    ADOConnection1.ConnectionString := format(strConnectBTI, ['btiall']);
    ADOConnection1.Connected := true;
    ds := TADODataset.Create(nil);
    ds.CacheSize := 100;
    ds.CursorLocation := clUseServer;
    ds.Connection := ADOConnection1;
    try
      col2 := WorkSheet.UsedRange.Columns.Count + 1;//начало новых данных
      sl.Clear;
      ds.CommandText := 'Select Top 1 * from geo.dk_semantic';
      ds.Open;
      n := ds.FieldDefs.Count + 1;
      for i := 0 to ds.FieldDefs.Count-1 do
      begin
        WorkSheet.Cells[1, col2+i] := ds.FieldDefs[i].Name;
        if CText('SEGMENTOFUSAGE', ds.FieldDefs[i].Name) then
          WorkSheet.Columns[col2+i].NumberFormat := '@';
      end;
      WorkSheet.Cells[1, col2+ds.FieldDefs.Count] := 'KN_Count';

      Info('Поиск объектов в Дежурной карте...');
      SafeCloseTable('TempTab');
      pg.Max := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
      for i := 1 to pg.Max do
      begin
        if i mod 10 = 0 then pg.Position := i;
        miDo('Fetch Rec %d From %s', [i, tab]);
        miDo('ob =%s.obj', [tab]);
        miDo('Select Дежурная_карта.КН from Дежурная_карта '+
          'Where Дежурная_карта.Obj Contains ob Into TempTab NoSelect');
        j := miEvalInt('TableInfo(TempTab, %d)', [TAB_INFO_NROWS]);
        if j > 0 then
        begin
          ds.Close;
          ss := #39 + miEval('TempTab.КН') + #39;
          for k := 2 to j do
          begin
            miDo('Fetch Rec %d From TempTab', [k]);
            ss := ss + ',' + #39 + miEval('TempTab.КН') + #39;
          end;
          ds.CommandText := format('select * from geo.dk_semantic where '+
            'CADASTRALNUMBER in (%s) order by DATECREATED Desc', [ss]);
          ds.Open;
          s := '';
          for k := 0 to ds.FieldDefs.Count-1 do
            if Pos('DATECREATED', ds.FieldDefs[k].Name) = 1 then
              s := s + Copy(ds.FieldByName(ds.FieldDefs[k].Name).AsString, 1, 10) + '~'
            else
              s := s + ds.FieldByName(ds.FieldDefs[k].Name).AsString + '~';
          s := s + format('%d', [j]);
          sl.Add(s);
        end else
          sl.Add('~~~~~~~');
      end;
      MiDo('Close table Дежурная_карта');
    finally
      pg.Max := 0;
      ds.Close;
      ds.Free;
      ADOConnection1.Close;
    end;

    data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, n], varVariant);
    for i := 0 to sl.Count-1 do
    begin
      sa := SplitStr(sl[i], '~');
      for j := Low(sa) to High(sa) do
        data[i+1, j+1] := sa[j];
      Setlength(sa, 0);
    end;
    s := ExcelNum2Str(WorkSheet, col2);
    ss := ExcelNum2Str(WorkSheet, col2+n-1);
    range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
    range.value := data;
    VarClear(data);
    WorkBook.Save;
  end;
  //поиск КН найденных в комментариях (максимум 2) в базе
  procedure Dop_Sem;
  var
    ds: TADODataset;
    i, j, k, n : Integer;
  begin
    miDo('Open table "%s"', [GetAppFolder+'Base\Дежурная_карта.tab']);
    ADOConnection1.ConnectionString := format(strConnectBTI, ['btiall']);
    ADOConnection1.Connected := true;
    ds := TADODataset.Create(nil);
    ds.CacheSize := 100;
    ds.CursorLocation := clUseServer;
    ds.Connection := ADOConnection1;
    Info('Поиск в Дежурной карте по комментариям...');
    try
      for k := 1 to 2 do
      begin
        n := 0;//проверяем наличие данных по 2 колонкам
        for i := VarArrayLowBound(semantik,2) to VarArrayHighBound(semantik,2) do
          if semantik[k, i] <> '' then
            Inc(n);
        if n = 0 then
          break;

        col2 := WorkSheet.UsedRange.Columns.Count + 1;//начало новых данных
        ds.Close;
        ds.CommandText := 'Select Top 1 * from geo.dk_semantic';
        ds.Open;
        n := ds.FieldDefs.Count + 1;
        for i := 0 to ds.FieldDefs.Count-1 do
        begin
          WorkSheet.Cells[1, col2+i] := format('%s_%d', [ds.FieldDefs[i].Name, k]);
          if CText('SEGMENTOFUSAGE', ds.FieldDefs[i].Name) then
            WorkSheet.Columns[col2+i].NumberFormat := '@';
        end;

        data := VarArrayCreate([1, 1, 1, ds.FieldDefs.Count], varVariant);
        for i := VarArrayLowBound(semantik,2) to VarArrayHighBound(semantik,2) do
        begin
          ds.Close;
          s := semantik[k, i];
          ds.CommandText := format('select * from geo.dk_semantic where '+
            'CADASTRALNUMBER = ''%s'' order by DATECREATED Desc', [s]);
          ds.Open;
          if ds.RecordCount > 0 then
          begin
            for j := 0 to ds.FieldDefs.Count-1 do
              if Pos('DATECREATED', ds.FieldDefs[k].Name) = 1 then
                data[1, j+1] := Copy(ds.FieldByName(ds.FieldDefs[j].Name).AsString, 1, 10)
              else
                data[1, j+1] := ds.FieldByName(ds.FieldDefs[j].Name).AsString;
            s := ExcelNum2Str(WorkSheet, col2);
            ss := ExcelNum2Str(WorkSheet, col2+ds.FieldDefs.Count-1);
            n := semantik[3, i];
            range := WorkSheet.range[format('%s%d:%s%d', [s, n, ss, n])];
            range.value := data;
          end;
        end;
        VarClear(data);
      end;
    finally
      ds.Close;
      ds.Free;
      ADOConnection1.Close;
    end;
    WorkBook.Save;
  end;

  function ModifyExcel : Boolean;
  var
    i, j, k, n : Integer;
  begin
    Info('Выбор колонок');
    Result := false;
    col := WorkSheet.UsedRange.Columns.Count;//размеры на исходном листе
    row := WorkSheet.UsedRange.Rows.Count;
    DataParcels := VarArrayCreate([1, row, 1, fields.Count], varVariant);
    data := VarArrayCreate([1, row, 1, 1], varVariant);
    k := 1;//число колонок в результирующей таблице
    pg.Max := fields.Count;
    for i := 0 to fields.Count-1 do
    begin
      pg.Position := i+1;
      if fields.ValueFromIndex[i] = '' then
      begin
        for j := 1 to col do
          if CText(fields[i], String(WorkSheet.Cells[1, j])) then
          begin
            s := ExcelNum2Str(WorkSheet, j);
            range := WorkSheet.range[format('%s1:%s%d', [s, s, row])];
            data := range.Value;
            if CText(String(data[1, 1]), 'Площадь участка') then
              colarea := k;
            if CText(String(data[1, 1]), 'Цена') then
              colcena := k;
            if CText(String(data[1, 1]), 'Комментарий') then
              colcomment := k;
            if CText(String(data[1, 1]), 'mapLat') then
              collat := k;
            if CText(String(data[1, 1]), 'mapLng') then
              collon := k;
            if CText(String(data[1, 1]), 'pdf') then
              colpdf := k;
            if CText(String(data[1, 1]), 'Адрес') then
              coladr := k;
            if CText(String(data[1, 1]), 'Адрес из парсинга') then
              coladr2 := k;
            if CText(String(data[1, 1]), 'Название предложения') then
              colpred := k;
            for n := 1 to row do
              DataParcels[n, k] := data[n, 1];
            break;
          end;
        Inc(k);
      end else
        break;
    end;
    pg.Position := pg.Max-3;
    //вставляем новый лист перед активным, старый удаляем,название переносим
    pg.Position := pg.Max-2;
    WorkBook.WorkSheets.Add(Before:=ExcelApp.ActiveSheet, Count:=1);
    WorkBook.Worksheets[2].Delete;
    WorkBook.WorkSheets.Item[1].Activate;
    WorkSheet := ExcelApp.ActiveSheet;
    //на новый лист на месте старого вставляем данные
    pg.Position := pg.Max-1;
    s := ExcelNum2Str(WorkSheet, k-1);
    range := WorkSheet.range[format('A1:%s%d', [s, row])];
    range.Value := DataParcels;
    range.RowHeight := 15;
    WorkBook.Save;
    VarClear(data);
    VarClear(DataParcels);
    Result := true;
    pg.Max := 0;
  end;

begin//TServiceMapForm.CalcOne
  if not FileExists(GetAppFolder+'Base\Дежурная_карта.tab') then
    Exc(GetAppFolder+'Base\Дежурная_карта.tab не найдена');
  analogs := GetAppFolder + 'analogs\';
  if not DirectoryExists(analogs) then
    Exc('Не найденн каталог "Analogs" с таблицами для расчеты анвлогов.');

  Info('Подготовка данных...');
  sl := TStringList.Create;
  sl2 := TStringList.Create;
  log := TStringList.Create;
  fields := TStringList.Create;//набор полей исходного файла
  dbf := TStringList.Create;//набор полей для файла DBF (Штань)
  bufsize := 0; bufdist := 0;
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
    ini.ReadSectionValues('Interface', fields);
    for i := 0 to fields.Count-1 do
      if CText(fields.Names[i], 'Buffer_Analogs') then
        Val(fields.ValueFromIndex[i], bufsize, k)
      else
      if CText(fields.Names[i], 'Distance_Analogs') then
        Val(fields.ValueFromIndex[i], bufdist, k);
    if bufsize = 0 then
       bufsize := 250;
    if bufdist = 0 then
       bufdist := 10;
    fields.Clear;

    ini.ReadSectionValues('Analysis', fields);
    for i := fields.Count-1 downto 0 do
      if Pos('//', fields[i]) = 1 then
        fields.Delete((i));
    ini.ReadSectionValues('FIELDS_GRAF', dbf);
  finally
    ini.Free;
  end;

  try
    ExcelApp.Workbooks.Open(xls);
    Workbook := ExcelApp.Workbooks.item[1];
    WorkBook.WorkSheets.Item[1].Activate;
    WorkSheet := ExcelApp.ActiveSheet;

    if tabdbf = '' then//нет готового файла от Штаня делаем все вычисления сами
    begin
      Info('Подготовка данных...');
      if not ModifyExcel then
      begin
        ExcelApp.Workbooks.Close;
        ExcelApp.Quit;
        exit;
      end;
      log.Clear;

      col := WorkSheet.UsedRange.Columns.Count;//берем размеры на созданном листе
      row := WorkSheet.UsedRange.Rows.Count;
      k := 0;
      for i := 0 to fields.Count-1 do
      begin//вставляем дополнительные солонки
        if fields.ValueFromIndex[i] <> '' then
        begin
          Inc(k);
          WorkSheet.Cells[1, col+k] := fields.Names[i];
        end;
      end;
//-------------заполняем колонку Площадь округленная----------------------------
      if FindColValue(n, '1') and (colarea <> 0) then
      begin
        s := WorkSheet.Cells[1, n];//запоминаем номер колонки
        WorkSheet.Columns[colarea].Select;//Q
        ExcelApp.Selection.Copy;
        WorkSheet.Columns[n].Select;
        ExcelApp.Selection.PasteSpecial(Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks :=False, Transpose:=False);
        WorkSheet.Columns[n].NumberFormat := '0';
        WorkSheet.Cells[1, n] := s;//восстанавлтвае название колонки
        range := WorkSheet.range[format('A2:%s%d', [ExcelNum2Str(WorkSheet, n), row])];
      end;
//-------------отдельный запуск программы CompareStrings.exe--------------------
      if FindColValue(n, '07') and FindColValue(m, '007') and FindColValue(k, '0007') then
      begin
        colprocent := n;//AQ
        ExShellExecute('cmd', format('/C xCopy /y /s /e "%s" "%s"',
          [source, TempPath]), SW_HIDE, true, 0);
        source := TempPath + ExtractFileName(source);//копируем во временный каталог
        ExCreateProcess(format('"%sCompareStrings.exe" "%s" "%.5f" "%d"',
          [GetAppFolder, source, procent, decimal]), '', SW_MINIMIZE, false);
      end else
        colprocent := 0;
//------------------------------------------------------------------------------
      Info('Импортируем таблицу...');
      DataParcels := range.value;
      pg.Max := VarArrayHighBound(DataParcels,1);
//-------------отдельный запуск программы CopyFiles.exe-------------------------
      sl.Clear;
//      respathPDF := 'W:\ЦГКО\IT\Outbox\Аналоги\PDF\' + AddSlesh(date);//tr!!!
      respathPDF := 'D:\ЦГКО\IT\Outbox\Аналоги\PDF\' + AddSlesh(date);//tr!!!
      ForceDirectories(respathPDF);
      for j := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
        if not VarIsEmpty(DataParcels[j, colpdf]) then
          sl.Add(DataParcels[j, 1] + '~' + DataParcels[j, colpdf]);//ID + URL
      sl.SaveToFile(TempPath + 'files.txt');
      ExCreateProcess(format('"%sCopyFiles.exe" "%s" "%s" "%s" "%s"',
        [GetAppFolder, TempPath + 'files.txt', xml,
        AddSlesh(path), respathPDF, Handle]), '', SW_MINIMIZE, false);
//------------------------------------------------------------------------------
      Info('Проверка данных');
      for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
        for j := VarArrayLowBound(DataParcels,2) to VarArrayHighBound(DataParcels,2) do
      begin
        pg.Position := i;
        if TVarData(DataParcels[i, j]).VType = varOleStr then
        begin
          s := DataParcels[i, j];
          s := StringReplace(StringReplace(s, #10, ' ', [rfReplaceAll]), #13, ' ', [rfReplaceAll]);
          s := StringReplace(StringReplace(s, 'ё', 'е', [rfReplaceAll]), 'Ё', 'Е', [rfReplaceAll]);
          s := StringReplace(s, #9, ' ', [rfReplaceAll]);
          s := StringReplace(s, '"', '', [rfReplaceAll]);
          DataParcels[i, j] := s;
        end;
      end;
      range.value := DataParcels;
//---------------------создание объединенных координат--------------------------
      if FindColValue(n, '4') and (collat <> 0) and (collon <> 0) then
      begin
        s := ExcelNum2Str(WorkSheet, n);//создание объединенных координат  Z
        range := WorkSheet.range[format('%s2:%s%d', [s, s, row])];
        data := range.value;
        for i := VarArrayLowBound(data,1) to VarArrayHighBound(data,1) do
        begin
          pg.Position := i;
          data[i, 1] := String(DataParcels[i, collat]) + ' ' + String(DataParcels[i, collon]);
          DataParcels[i, col] := data[i, 1];//для Заполнение группировки площадей
        end;
        range.value := data;
        WorkBook.Save;
        VarClear(data);
        pg.Max := 0;
      end;
//------------------------------------------------------------------------------
      txt := ChangeFileExt(xls, '.txt');//сохраняем в текст
      WorkBook.SaveAs(Filename:=txt, FileFormat:=xlText, CreateBackup:=False);
      ExcelApp.Workbooks.Close;
      //восстанавливаем таблицу
      ExcelApp.Workbooks.Open(xls);
      Workbook := ExcelApp.Workbooks.item[1];
      WorkBook.WorkSheets.Item[1].Activate;
      WorkSheet := ExcelApp.ActiveSheet;
//----------------- создание иаблицы MapInfo -----------------------------------
      Info('Создание таблицы...');
      tab := ChangeFileExt(xls, '.tab');//полное имя
      SafeCloseTable(miEval('PathToTableName$("%s")', [tab]));
      s := miEval('TempFileName$("")');
      ss := miEval('PathToTableName$("%s")', [s]);
      miDo('Register Table "%s" TYPE ASCII Delimiter 9 Titles Charset "WindowsCyrillic" Into "%s"', [txt, s]);
      miDo('Open Table "%s" Hide', [s]);
      miDo('Commit Table %s As "%s"', [ss, tab]);
      miDo('Close Table %s', [ss]);
      miDo('Open Table "%s"', [tab]);
      tab := miEval('PathToTableName$("%s")', [tab]);//короткое имя
      miDo('Create Map For %s CoordSys Earth Projection 8, 1001, "m", 60.05, '+
        '0, 1, 1500000, -5911057.63',  [tab]);

      miDo('Alter table %s (Add Теги1 Integer, t1 Integer, t2 Integer)', [tab]);
      miDo('Update %s Set Теги1 = RowId', [tab]);
//      miDo('Update %s Set Комментарий = Left$(Комментарий, 200)', [tab]);
      miDo('commit table %s', [tab]);
      DeleteFile(txt);
//------------------------------------------------------------------------------
      if FindColValue(n, '4') and (collat <> 0) and (collon <> 0) then
      begin
        Info('Создание графических объектов');
        pg.Max := VarArrayHighBound(DataParcels,1);
        data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 2], varVariant);
        miDo('Set CoordSys Table %s', [tab]);
        for i := 1 to pg.Max do
        begin
          pg.Position := i;
          miDo('Fetch Rec %d From %s', [i, tab]);
          s := miEval('%s.COL%d', [tab, n]);//объединенные координаты   X
          j := Pos(' ', s);
          ss := Copy(s, j+1, Length(s));
          s := Copy(s, 1, j-1);

          Val(ss, xy.x, m);
          Val(s, xy.y, m);
          Wgs2Msk_(xy.y, xy.x, 2);
          miDo('Create Point Into Variable ob	(%.9f, %.9f) Symbol MakeSymbol(34,%d,%d)', [xy.x, xy.y, RGB(0, 0, 255), 10]);
          miDo('Update %s Set obj=ob Where rowId=%d', [tab, i]);

          data[i, 1] := xy.x;
          data[i, 2] := xy.y;
        end;
        miDo('Commit Table %s', [tab]);

        FindColValue(n, '2'); FindColValue(m, '3');
        s := ExcelNum2Str(WorkSheet, n);//x_41
        ss := ExcelNum2Str(WorkSheet, m);//y_41
        range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
        range.value := data;
        VarClear(data);
      end;
//------------------------------------------------------------------------------
      idkp := 0;//для создание буферных зон
      kodnp := 0;//для создание буферных зон
      i := 0;
      While i < fields.Count-1 do
      begin
        if Pos('Contains', fields.ValueFromIndex[i]) = 1 then
        try
          s := Copy(fields.ValueFromIndex[i], 10, 100);
          sa := SplitStr(s, ';');
          if FileExists(analogs + sa[0] + '.tab') then//полное название таблицы
            Contains(i, row, kodnp, idkp, fields, WorkSheet, DataParcels,
              tab, analogs + sa[0] + '.tab', sa[1])
          else
            log.Add(format('Файл %s не найден', [analogs + sa[0] + '.tab']));
        finally
          SetLength(sa, 0);
        end;
        Inc(i);
      end;
//------------------------расчет по дорогам-------------------------------------
      i := 0;
      While i < fields.Count-1 do
      begin
        if Pos('Intersects', fields.ValueFromIndex[i]) = 1 then
        try
          s := Copy(fields.ValueFromIndex[i], 12, 100);
          sa := SplitStr(s, ';');
          if FileExists(analogs + 'Дороги_WGS84\' + sa[0] + '.tab') then//полное название таблицы
            Intersects(i, row, radius,  fields, WorkSheet, DataParcels,
              tab, analogs + 'Дороги_WGS84\' + sa[0] + '.tab', sa[1])
          else
            log.Add(format('Файл %s не найден', [analogs + 'Дороги_WGS84\' + sa[0] + '.tab']));
        finally
          SetLength(sa, 0);
        end;
        Inc(i);
      end;
//------------------------------------------------------------------------------
      if FindColValue(n, '01') and FindColValue(k, '001') and FindColValue(m, '4') then
      if k = n+1 then
      begin
        Info('Группировка по координатам');
        s := TempPath + 'group.txt';
        miDo('Select COL%d,Count(*) from %s group by COL%d into TempTab NoSelect', [m, tab, m]);//объединенные координаты
        miDo('Select * From TempTab where Count > 1 Into group  NoSelect');
        miDo('Export group Into "%s" Type "ASCII" Overwrite Delimiter "=" CharSet "WindowsCyrillic"', [s]);
        miDo('Close table TempTab');

        sl.LoadFromFile(s);
        for i := 0 to sl.Count-1 do
          sl[i] := StringReplace(sl[i], '"', '', [rfReplaceAll]);

        data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 2], varVariant);
        for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
        begin
          pg.Position := i;
          for j := 0 to sl.Count-1 do
          begin
            if sl.Names[j] = DataParcels[i, col] then
            begin
              data[i, 1] := j+1;
              data[i, 2] := sl.ValueFromIndex[j];
            end;
          end;
        end;

        s := ExcelNum2Str(WorkSheet, n);//id_Gdouble
        ss := ExcelNum2Str(WorkSheet, k);//Count_Gdouble
        range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
        range.value := data;
        VarClear(data);
      end;
//------------------------------------------------------------------------------
      if FindColValue(n, '05') and FindColValue(k, '005') then
      if k = n+1 then
      begin
        Info('Группировка по расстояниям');
        SafeCloseTable('TempBuffer');
        miDo('Create Table TempBuffer (ID Integer, Area_Dist Float) File "%sTempBuffer.tab"', [TempPath]);
        miDo('Create Map For TempBuffer CoordSys table %s', [tab]);
        miDo('Set coordsys table TempBuffer');
        miDo('Fetch Rec 1 From %s', [tab]);
        miDo('ob = Buffer(%s.obj, 20, %d, "m")', [tab, bufdist]);
        xy.z := miEvalDbl('CartesianArea(ob, "sq m")')+0.1;//площадь для сравнения буферов
        miDo('Create Object As Buffer from %s Into table TempBuffer '+
          'Width %d Units "m" Type Spherical Resolution 12', [tab, bufdist]);
        miDo('select * from TempBuffer');
        miDo('Objects Disaggregate Into Table TempBuffer');
        miDo('delete from TempBuffer where rowid = 1');
        miDo('Select * from TempBuffer into selBuffer Where CartesianArea(obj, "sq m")>%.f', [xy.z]);
        miDo('Update selBuffer Set ID = rowid, Area_Dist = CartesianArea(obj, "sq m")');
        data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 2], varVariant);
        pg.max := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
        for i := 1 to pg.max do//идем по основной  таблице
        begin
          pg.Position := i;
          miDo('Fetch rec %d From %s', [i, tab]);
          mido('ob = %s.obj', [tab]);//выбираем буфер в котором находится объект
          miDo('select * from selBuffer where ob Within selBuffer.obj into temptab NoSelect');
          data[i, 1] := miEval('temptab.col1');
          data[i, 2] := miEval('temptab.col2');
        end;
        pg.max := 0;

        s := ExcelNum2Str(WorkSheet, n);//id_Dist
        ss := ExcelNum2Str(WorkSheet, k);//Area_Dist
        range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
        range.value := data;
        VarClear(data);
        miDo('Drop table TempBuffer');
      end;
//------------------------Группировка по площадям-------------------------------
      if FindColValue(n, '02') and FindColValue(k, '002') and FindColValue(m, '0002') then
      if (k = n+1) and (k = m-1) then
      begin
        Info('Группировка по площадям');
        s := TempPath + 'group.txt';
        miDo('Select COL%d,Count(*) from %s group by COL%d into TempTab NoSelect', [colarea, tab, colarea]);
        miDo('Select * From TempTab where Count > 1 Into group  NoSelect');
        miDo('Export group Into "%s" Type "ASCII" Overwrite CharSet "WindowsCyrillic"',[s]);
        miDo('Close table TempTab');

        pg.Max := VarArrayHighBound(DataParcels, 1);
        sl.LoadFromFile(s);
        for i := 0 to sl.Count-1 do
          sl.Objects[i] := Pointer(i+1);//номер группы

        data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 3], varVariant);
        for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
        begin
          pg.Position := i;
          for j := 0 to sl.Count-1 do
          begin
            sa := SplitStr(sl[j], #9);
            if DataParcels[i, colarea]=sa[0] then
            begin
              data[i, 1] := sa[0];
              data[i, 2] := sa[1];
              data[i, 3] := Integer(sl.Objects[j]);
            end;
            SetLength(sa, 0);
          end;
        end;

        s := ExcelNum2Str(WorkSheet, n);//Площадь   AC
        ss := ExcelNum2Str(WorkSheet, m);//Площадь_ID  AE
        range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
        range.value := data;
        pg.Max := 0;
        VarClear(data);
      end;
//-----------------------------УПКС=--------------------------------------------
      if FindColValue(n, '06') and FindColValue(k, '006') and FindColValue(m, '0006') then
      if (k = n+1) and (k = m-1) then
      begin
        Info('Группировка по УПКС');
        s := TempPath + 'group.txt';
        miDo('Select COL%d, COL%d from %s where COL%d>0 into TempTab NoSelect',
          [colcena, colarea, tab, colarea]);//исключи площади = 0
        miDo('Update TempTab Set COL1 = Round(COL1/Col2, 1)');//округляем до целого

        miDo('Select Цена,Count(*) from TempTab group by Цена into group NoSelect');
        miDo('Select * From group where Count > 1 Into TempTab NoSelect');
        miDo('Export TempTab Into "%s" Type "ASCII" Overwrite CharSet "WindowsCyrillic"',[s]);
        sl.LoadFromFile(s);

        sl2.Clear;
        for i := 0 to sl.Count-1 do
        begin
          sl.Objects[i] := Pointer(i+1);//индекс в списке sl
          j := Pos(#9, sl[i]);
          sl2.AddObject(Copy(sl[i], 1, j-1), Pointer(i));//указатель на список sl
        end;
        sl2.Sort;

        pg.Max := VarArrayHighBound(DataParcels, 1);
        data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 3], varVariant);
        for i := 1 to miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]) do
        begin
          pg.Position := i;
          miDo('Fetch Rec %d From %s', [i, tab]);
          s :=  miEval('%s.COL%d', [tab, colcena]);
          if sl2.Find(s, j) then
          begin
            j := Integer(sl2.Objects[j]);
            sa := SplitStr(sl[j], #9);
            data[i, 1] := sa[0];
            data[i, 2] := sa[1];
            data[i, 3] := Integer(sl.Objects[j]);
            SetLength(sa, 0);
          end;

        end;
        s := ExcelNum2Str(WorkSheet, n);//AN
        ss := ExcelNum2Str(WorkSheet, m);//AP
        range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
        range.value := data;
        pg.Max := 0;
        VarClear(data);
      end;
//------------------------------------------------------------------------------
      if FindColValue(n, '03') and FindColValue(k, '003')  and FindColValue(i, '0003') then
      if (k = n+1) and (k = i-1) and (idkp <> 0) and (kodnp <> 0) then
      begin
        Info('Создание буферных зон...');
        SafeCloseTable('TempBuffer');
        miDo('Create Table TempBuffer (ID Integer, Count Integer, Area Float) File "%sTempBuffer.tab"', [TempPath]);
        miDo('Create Map For TempBuffer CoordSys Earth Projection 1, 104');
        for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
          if (Pos(':000:0:0:0', String(DataParcels[i, kodnp])) = 3) and (String(DataParcels[i, idkp]) = '') then
          begin
            miDo('Fetch Rec %d From %s', [i, tab]);
            miDo('ob = %s.obj', [tab]);
            miDo('Insert Into TempBuffer (Obj) Values (ob)');
          end;
        if miEvalInt('TableInfo(TempBuffer, %d)', [TAB_INFO_NROWS]) > 0 then
        begin
          miDo('Create Object As Buffer from TempBuffer Into table TempBuffer Width %d '+
            'Units "m" Type Spherical Resolution 12', [bufsize]);
          i := miEvalInt('TableInfo(TempBuffer, %d)', [TAB_INFO_NROWS]);
          miDo('select * from TempBuffer into selBuffer Where Rowid < %d', [i]);
          miDo('delete from selBuffer');
          miDo('select * from TempBuffer');
          miDo('Objects Disaggregate Into Table TempBuffer');
          miDo('Commit Table TempBuffer Interactive');
          miDo('Pack Table TempBuffer Graphic Data');
          miDo('Set CoordSys table %s', [tab]);
          miDo('Update TempBuffer Set ID = rowid, Area = Area(obj, "sq m") DropIndex Auto');
          Info('Группировка по буферам');
          pg.max := miEvalInt('TableInfo(TempBuffer, %d)', [TAB_INFO_NROWS]);
          for i := 1 to pg.max do//идем по буферам
          begin
            pg.Position := i;
            miDo('Fetch rec %d From TempBuffer', [i]);
            mido('ob = TempBuffer.obj');//запрос для каждого буфера
            miDo('select * from %s where ob Contains %s.obj into temptab noSelect', [tab, tab]);
            s := miEval('TableInfo(temptab, %d)', [TAB_INFO_NROWS]);//число объектов попавших в буфер
            miDo('update TempBuffer set count = %s where rowid=%d', [s, i]);
          end;

          data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 3], varVariant);
          pg.max := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
          for i := 1 to pg.max do//идем по основной  таблице
          begin
            pg.Position := i;
            miDo('Fetch rec %d From %s', [i, tab]);
            mido('ob = %s.obj', [tab]);//выбираем буфер в котором находится объект
            miDo('select * from TempBuffer where ob Within TempBuffer.obj into temptab NoSelect');
            data[i, 1] := miEval('temptab.col3');
            data[i, 2] := miEval('temptab.col2');
            data[i, 3] := miEval('temptab.col1');
          end;
          pg.max := 0;
          s := ExcelNum2Str(WorkSheet, n);//Буфер_площадь
          ss := ExcelNum2Str(WorkSheet, k);//Буфер_ID
          range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
          range.value := data;
          VarClear(data);
        end;
        //miDo('commit table TempBuffer');
        miDo('Drop table TempBuffer');
      end;
//------------------------------------------------------------------------------
      if FindColValue(n, '04') and FindColValue(k, '004') and FindColValue(m, '0004') then
      if (k = n+1) and (k = m-1) then
      begin
        Info('Группировка по комментариям');
        s := TempPath + 'group.txt';
        miDo('Select Комментарий,count(*) from %s group by Комментарий into TempTab NoSelect', [tab]);
        miDo('Select * From TempTab where Count > 1 Into group  NoSelect');
        miDo('Export group Into "%s" Type "ASCII" Overwrite CharSet "WindowsCyrillic"',[s]);
        miDo('Close table TempTab');

        pg.Max := VarArrayHighBound(DataParcels, 1);
        sl.LoadFromFile(s);
        for i := 0 to sl.Count-1 do
        begin
          sl[i] := StringReplace(sl[i], '"', '', [rfReplaceAll]);
          sl.Objects[i] := Pointer(i+1);//номер группы
        end;

        data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 3], varVariant);
        for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
        begin
          pg.Position := i;
          for j := 0 to sl.Count-1 do
          begin
            sa := SplitStr(sl[j], #9);
            if Pos(sa[0], String(DataParcels[i, colcomment])) = 1 then
            begin
              data[i, 1] := sa[0];
              data[i, 2] := sa[1];
              data[i, 3] := Integer(sl.Objects[j]);
            end;
            SetLength(sa, 0);
          end;
        end;
        s := ExcelNum2Str(WorkSheet, n);//Коментарий_площадь
        ss := ExcelNum2Str(WorkSheet, m);//Коментарий_ID
        range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
        range.value := data;
        sl.Clear;
        pg.Max := 0;
        VarClear(data);
      end;
//------------------------------------------------------------------------------
      if FindColValue(n, '011') then
      begin
        for j := 0 to fields.Count-1 do
          if AnsiContainsText(fields.ValueFromIndex[j], 'НаименованиеНП_МО') then
          begin
            col2 := j+1;// колонка naim_np
            Info('Проверяем Соответствие НП');
            pg.Max := VarArrayHighBound(DataParcels,1);
            data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 1], varOleStr);
            for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
            begin
              pg.Position := i;
              s := WorkSheet.Cells[i+1, col2];//для WorkSheet заголовок пропускаем
              if ContainsText(DataParcels[i, coladr], s) or ContainsText(DataParcels[i, coladr2], s) or
                ContainsText(DataParcels[i, colpred], s) or ContainsText(DataParcels[i, colcomment], s) then
                  data[i, 1] := 'Да'
              else
                data[i, 1] := 'Нет';
            end;
            s := ExcelNum2Str(WorkSheet, n);
            range := WorkSheet.range[format('%s2:%s%d', [s, s, row])];//BH Correspondence_NP
            range.value := data;
            VarClear(data);
            break;
          end;
      end;
//--------обработка файла полученного после программы CopyFiles.exe-------------
      if FindColValue(col2, '012') then
      begin
        Info('Заполнение Скан объявления...');
        data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 1], varOleStr);
        s := AddSlesh(respathPDF) + 'files.txt';
        if FileExists(s) then
        begin
          sl.LoadFromFile(s);
          DeleteFile(s);
          for j := 0 to sl.Count-1 do
            data[j+1, 1] := sl[j];
          s := ExcelNum2Str(WorkSheet, col2);
          range := WorkSheet.range[format('%s2:%s%d', [s, s, row])];
          range.value := data;
        end;
        VarClear(data);
      end;
//--------обработка файла полученного после программы CompareStrings.exe--------
      if FileExists(ExtractFilePath(source) + 'text.txt') and (colprocent <> 0)then
      begin
        Info('Заполнение сравнения строк...');
        sl2.Clear;
        sl.LoadFromFile(ExtractFilePath(source) + 'text.txt');
        data := VarArrayCreate([1, VarArrayHighBound(DataParcels,1), 1, 3], varOleStr);
        for i := 0 to sl.Count-1 do
        begin
          j := Pos(';', sl[i]);
          sl2.AddObject(Copy(sl[i], 1, j-1), Pointer(i));
        end;
        sl2.Sort;
        for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
        begin
          s := DataParcels[i, 1];
          if sl2.Find(s, j) then
          begin
            n := Integer(sl2.Objects[j]);
            sa := SplitStr(sl[n], ';');
            data[i, 1] := sa[1];
            data[i, 2] := sa[2];
            data[i, 3] := sa[3];
            SetLength(sa, 0);
          end;
        end;
        s := ExcelNum2Str(WorkSheet, colprocent);//AQ
        ss := ExcelNum2Str(WorkSheet, colprocent+2);//AS
        range := WorkSheet.range[format('%s2:%s%d', [s, ss, row])];
        range.value := data;
        VarClear(data);
      end;
//------------------------------------------------------------------------------
      ADO;
//-------------кадастровые номера из семантики----------------------------------
      semantik := VarArrayCreate([1, 3, 0, 0], varVariant);
      j := 0;
      for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
      begin//ищем начальные кадастровые номера в комментариях
        s := StringReplace(DataParcels[i, colcomment], ' ', '', [rfReplaceAll]);
        for n := High(s) downto Low(s) do
          if ord(s[n]) > 11030 then
            System.Delete(s, n, 1);
        mc := reg.Matches(s, '\d\d:\d\d:\d{1,7}:\d{1,}', [roIgnoreCase, roMultiLine]);
        if mc.Count > 0 then    //1=14 15
        begin
          Inc(j);
          VarArrayRedim(semantik, j);
          semantik[3, j-1] := i+1;//с учетом заголовка
          semantik[1, j-1] := mc[0].Value;
          if mc.Count > 1 then
            semantik[2, j-1] := mc[1].Value;
        end;
      end;
      if VarArrayHighBound(semantik, 2) > 0 then
        Dop_Sem;
//------------------------------------------------------------------------------
      Info('Создание файла DBF...');
      sl2.Clear;
      for i := 1 to WorkSheet.UsedRange.Columns.Count do
        sl2.Add(AnsiUpperCase(WorkSheet.Cells[1, i]));

      for i := dbf.Count-1 downto 0 do
      begin
        k := -1;
        for j := 0 to sl2.Count-1 do
          if Pos(AnsiUpperCase(dbf.Names[i]), sl2[j]) > 0 then
            k := j;
        if k = -1 then
        begin
          log.Add(dbf[i] + ' Нет в списке полей');
          dbf.Delete(i);
        end;
      end;
      date := dbf.Names[0];
      for i := 1 to dbf.Count-1 do
        date := date + ',' + dbf.ValueFromIndex[i];//.Names[i];
      txt := ChangeFileExt(xls, '.txt');//сохраняем в текст
      WorkBook.SaveAs(Filename:=txt, FileFormat:=xlText, CreateBackup:=False);
      ExcelApp.Workbooks.Close;
      //восстанавливаем таблицу
      ExcelApp.Workbooks.Open(xls);
      Workbook := ExcelApp.Workbooks.item[1];
      WorkBook.WorkSheets.Item[1].Activate;
      WorkSheet := ExcelApp.ActiveSheet;

      SafeCloseTable(tab);
      RenameFile(ChangeFileExt(xls, '.id'), ChangeFileExt(xls, '_.id'));
      RenameFile(ChangeFileExt(xls, '.map'), ChangeFileExt(xls, '_.map'));

      s := miEval('TempFileName$("")');
      ss := miEval('PathToTableName$("%s")', [s]);
      miDo('Register Table "%s" TYPE ASCII Delimiter 9 Titles Charset "WindowsCyrillic" Into "%s"', [txt, s]);
      miDo('Open Table "%s" Hide', [s]);
      miDo('Commit Table %s As "%s"', [ss, ChangeFileExt(xls, '.tab')]);
      miDo('Close Table %s', [ss]);

      RenameFile(ChangeFileExt(xls, '_.id'), ChangeFileExt(xls, '.id'));
      RenameFile(ChangeFileExt(xls, '_.map'), ChangeFileExt(xls, '.map'));
      miDo('Open Table "%s"', [ChangeFileExt(xls, '.tab')]);
      DeleteFile(txt);

      ss := format('Alter Table "%s" (rename ', [tab]);
      for i := 0 to dbf.Count-1 do
        ss := ss + dbf.Names[i] + ' ' + dbf.ValueFromIndex[i] + ',';
      miDo(Copy(ss, 1, Length(ss)-1) + ')');

      s := ExtractFilePath(xls) + 'Z' + ExtractFileName(DellSlesh(respathPDF)) + 'graf.dbf';
      miDo('Select %s From %s Into TempTab NoSelect', [date, tab]);
      miDo('Export TempTab Into "%s" Type "DBF" CharSet "WindowsCyrillic"', [s]);
//------------------------------------------------------------------------------
    end else//есть готовый файл от Штаня
    begin
      Info('Копирование исключений...');
      col := WorkSheet.UsedRange.Columns.Count;//размеры на листе
      row := WorkSheet.UsedRange.Rows.Count;
      miDo(fieldsdbf);//выбор из исправленной таблицы (от Штаня) сформированный в Path_Unit
      col2 := 1;
      for i := 1 to Length(fieldsdbf) do//подсчитываем количество колонок через разделитель(запятую)
        if fieldsdbf[i] = ',' then Inc(col2);
      miDo('Export tempTab Into "%stempTab.txt" Type "ASCII" Overwrite CharSet "WindowsCyrillic" Titles',[TempPath]);
      sl.LoadFromFile(TempPath + 'tempTab.txt');
      for i := 0 to sl.Count-1 do
        sl[i] := StringReplace(sl[i], '"', '', [rfReplaceAll]);

      data := VarArrayCreate([1, row, 1, col2], varVariant);//массив для заполнеия данными с заголовком
      range := WorkSheet.range[format('A2:A%d', [row])];
      DataParcels := range.value;//только номера

      sa := SplitStr(sl[0], #9);//зпголовок полей из секции [FIELDS_DBF]
      for j := Low(sa) to High(sa) do
        data[1, j+1] := StringReplace(sa[j], '"', '', [rfReplaceAll]);
      SetLength(sa, 0);
      pg.Max := VarArrayHighBound(DataParcels,1);
      for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
      begin
        pg.Position := i;
        for j := 1 to sl.Count-1 do
        begin
          sa := SplitStr(sl[j], #9);
          if DataParcels[i, 1] = sa[0] then
          begin
            for k := Low(sa) to High(sa) do
              data[i+1, k+1] := sa[k];
            break
          end;
          SetLength(sa, 0);
        end;
      end;
      //добавляем в конец анализ от Штаня
      s := ExcelNum2Str(WorkSheet, col+1);
      ss := ExcelNum2Str(WorkSheet, col+col2);
      range := WorkSheet.range[format('%s1:%s%d', [s, ss, row])];
      range.value := data;
      VarClear(DataParcels);
      VarClear(data);
      s := ChangeFileExt(xls, '_2.xlsx');//новое имя результата
      WorkBook.SaveAs(Filename:=s, FileFormat:= 51, CreateBackup:=False);
    end;

    pg.Max := 0;
    SafeCloseTable('TempTab');
    Info('Готово');
    WorkSheet.Rows[1].Select;
    ExcelApp.Selection.AutoFilter;
    ExcelApp.ActiveWindow.SplitColumn := 0;
    ExcelApp.ActiveWindow.SplitRow := 1;
    ExcelApp.ActiveWindow.FreezePanes := True;
    ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
    WorkBook.Save;
  finally
    txt := ChangeFileExt(xls, '.log');
    log.SaveToFile(txt);
    if log.Count > 0 then
      ExShellExecute('Notepad.exe', txt, SW_NORMAL, false, 0)
    else
      DeleteFile(txt);
    Workbook := Unassigned;
    WorkSheet := Unassigned;
    range := Unassigned;
    VarClear(DataParcels);
    VarClear(data);
    sl.Free;
    sl2.Free;
    fields.Free;
    log.Free;
    dbf.Free;
  end;
end;

procedure TServiceMapForm.mnCalcClick(Sender: TObject);
var
  analogs : String;
  ExcelApp : OleVariant;
begin//TServiceMapForm.mnCalcClick
  if not FileExists(GetAppFolder+'Base\Дежурная_карта.tab') then
    Exc(GetAppFolder+'Base\Дежурная_карта.tab не найдена');
  analogs := GetAppFolder + 'analogs\';
  if not DirectoryExists(analogs) then
    Exc('Не найденн каталог "Analogs" с таблицами для расчеты анвлогов.');

  Path_Form := TPath_Form.Create(nil);
  try
    if Path_Form.ShowModal <> mrOk then
      exit;
    xls := Path_Form.xls;
    tabdbf := Path_Form.tabdbf;
    if not Path_Form.cbDBF.Checked then
    begin
      tabdbf := '';
      CorrectXLS(xls);
    end;
    xml := Path_Form.Edit3.Text;
    path := Path_Form.path;
    date := Path_Form.Edit5.Text;
    fieldsdbf := Path_Form.fieldsdbf;
    radius := Path_Form.radius;
    procent := Path_Form.procent;
    decimal := Path_Form.decimal;
    source := Path_Form.source;
  finally
    FreeNil(Path_Form);
  end;

  ExcelApp := CreateOleObject('Excel.Application');
  CalcOne(ExcelApp);
  ExcelApp.Visible := true;
  ExcelApp := unassigned;
end;

procedure TServiceMapForm.mnCalcGroupClick(Sender: TObject);
var
  ExcelApp, WorkBook, WorkSheet : OleVariant;
  sl, files, dbf : TStringList;
  s, ss, tab, txt : String;
  i , j, k, r, col, row : Integer;
  ini : TMemIniFile;
begin
  if not DirectoryExists(GetAppFolder + 'analogs\') then
    Exc('Не найден каталог "Analogs" с таблицами для расчеты анвлогов.');

  sl := TStringList.Create;
  dbf := TStringList.Create;
  files := TStringList.Create;

  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
    ini.ReadSectionValues('FIELDS_GRAF', dbf);
  finally
    ini.Free;
  end;

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;

  try
    OpenDialog1.FileName := '*.txt';//D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\Аналоги_test4\Аналоги_группа.txt
    if not OpenDialog1.Execute then
      exit;
    OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
    sl.LoadFromFile(OpenDialog1.FileName);//, TEncoding.ASCII
    for i := sl.Count-1 downto 0 do
      if Pos('//', sl[i]) = 1 then
        sl.Delete(i);
    while sl[sl.Count-1] = '' do
        sl.Delete(sl.Count-1);
    if (sl.Count = 0) or (sl.Count mod 5 > 0) then
      Exc('Количество записей в файле должно быть кратно 5!');

    Path_Form := TPath_Form.Create(nil);
    try
      i := 0;
      Repeat
        Path_Form.Edit1.Text := sl[i];
        Path_Form.Edit2.Text := sl[i+1];
        Path_Form.Edit3.Text := sl[i+2];
        Path_Form.Edit4.Text := sl[i+3];
        Path_Form.Edit5.Text := sl[i+4];
        path := Path_Form.Edit4.Text;
        Path_Form.btOkClick(nil);
        if Path_Form.ModalResult <> mrOk then
          Exit;
        Inc(i, 5);
      Until i >= sl.Count;
    finally
      Path_Form.Free;
    end;

    xls := GetAppFolder + 'Аналоги\' + AddSlesh(CurrentData) + 'Общий_Результат.xlsx';
    if FileExists(xls) then
      if not DeleteFile(xls) then
        Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [xls, #10]));

    s := ExtractFilePath(xls) + 'Общий_Результат_graf.dbf';
    if FileExists(s) then
      if not DeleteFile(s) then
        Exc(format('Файл "%s" открыт.%s Закройте файл и повторите попытку.', [s, #10]));

    Path_Form := TPath_Form.Create(nil);
    try
      i := 0;
      Repeat
        Path_Form.Edit1.Text := sl[i];
        Path_Form.Edit2.Text := sl[i+1];
        Path_Form.Edit3.Text := sl[i+2];
        Path_Form.Edit4.Text := sl[i+3];
        Path_Form.Edit5.Text := sl[i+4];
        Path_Form.btOkClick(nil);

        xls := Path_Form.xls;
        tabdbf := Path_Form.tabdbf;
        if not Path_Form.cbDBF.Checked then
        begin
          tabdbf := '';
          CorrectXLS(xls);
        end;
        xml := Path_Form.Edit3.Text;
        path := Path_Form.Edit4.Text;
        date := Path_Form.Edit5.Text;
        fieldsdbf := Path_Form.fieldsdbf;
        radius := Path_Form.radius;
        procent := Path_Form.procent;
        decimal := Path_Form.decimal;
        source := Path_Form.source;

        CalcOne(ExcelApp);
        ExcelApp.Workbooks.Close;
        ExcelApp.Quit;

        files.Add(xls);
        Inc(i, 5);
//        PostMessage(FindWindow('XLMAIN', nil), WM_QUIT, 0, 0);
      Until i >= sl.Count;
    finally
      Path_Form.Free;
    end;

    for i := 0 to files.Count-1 do
    begin
      xls := files[i];
      if FileExists(xls) then
      begin
        if i = 0 then
        begin
          s := ExtractFilePath(xls) + 'Общий_Результат.xlsx';
          ExShellExecute('cmd', PChar(format('/C Copy /Y "%s" "%s"', [xls, s])), SW_HIDE, true, 0);
          xls := s;
          ExcelApp.Workbooks.Open(xls);
          WorkSheet := ExcelApp.Workbooks.item[1].WorkSheets.Item[1];
          col := WorkSheet.UsedRange.Columns.Count;//берем размеры на новом листе
          r := WorkSheet.UsedRange.Rows.Count + 1;
        end else
        try
          ExcelApp.Workbooks.Open(xls);
          WorkSheet := ExcelApp.Workbooks.item[2].WorkSheets.Item[1];
          row := WorkSheet.UsedRange.Rows.Count;
          s := ExcelNum2Str(WorkSheet, col);
          WorkSheet.range[format('A2:%s%d', [s, row])].Select;
          ExcelApp.Selection.Copy;

          ExcelApp.Workbooks.item[1].Activate;//активируем первую книгу (первый рабочий лист становится активным)
          WorkSheet := ExcelApp.Workbooks.item[1].ActiveSheet;
          WorkSheet.Activate;//обязательно активировать

          WorkSheet.range[format('A%d:%s%d', [r, s, r + row])].Select;
          ExcelApp.Selection.PasteSpecial(Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks :=False, Transpose:=False);
          r := r + row + 1;
        finally
          ExcelApp.Workbooks.item[2].Close;
        end;
      end;
    end;
    WorkSheet.Rows[1].Select;
    ExcelApp.Selection.AutoFilter;
    ExcelApp.ActiveWindow.SplitColumn := 0;
    ExcelApp.ActiveWindow.SplitRow := 1;
    ExcelApp.ActiveWindow.FreezePanes := True;
    ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
    ExcelApp.Workbooks.item[1].Save;
//------------------------------------------------------------------------------
    Info('Создание файла DBF...');
    Workbook := ExcelApp.Workbooks.item[1];
    xls := ExtractFilePath(files[0]) + 'Общий_Результат.xlsx';
    sl.Clear;
    for i := 1 to WorkSheet.UsedRange.Columns.Count do
      sl.Add(AnsiUpperCase(WorkSheet.Cells[1, i]));

    for i := dbf.Count-1 downto 0 do
    begin
      k := -1;
      for j := 0 to sl.Count-1 do
        if Pos(AnsiUpperCase(dbf.Names[i]), sl[j]) > 0 then
          k := j;
      if k = -1 then
        dbf.Delete(i);
    end;

    date := dbf.Names[0];
    for i := 1 to dbf.Count-1 do
      date := date + ',' + dbf.ValueFromIndex[i];//.Names[i];
    txt := ChangeFileExt(xls, '.txt');//сохраняем в текст
    WorkBook.SaveAs(Filename:=txt, FileFormat:=xlText, CreateBackup:=False);
    ExcelApp.Workbooks.Close;
    //восстанавливаем таблицу
    ExcelApp.Workbooks.Open(xls);
    Workbook := ExcelApp.Workbooks.item[1];
    WorkSheet := ExcelApp.Workbooks.item[1].ActiveSheet;
    WorkSheet.Activate;//обязательно активировать

    s := miEval('TempFileName$("")');
    ss := miEval('PathToTableName$("%s")', [s]);
    miDo('Register Table "%s" TYPE ASCII Delimiter 9 Titles Charset "WindowsCyrillic" Into "%s"', [txt, s]);
    miDo('Open Table "%s" Hide', [s]);
    tab := ChangeFileExt(xls, '.tab');
    miDo('Commit Table %s As "%s"', [ss, tab]);
    miDo('Close Table %s', [ss]);
    DeleteFile(txt);
    miDo('Open table "%s"', [tab]);
    tab := miEval('PathToTableName$("%s")', [tab]);//короткое имя

    ss := format('Alter Table "%s" (rename ', [tab]);
    for i := 0 to dbf.Count-1 do
      ss := ss + dbf.Names[i] + ' ' + dbf.ValueFromIndex[i] + ',';
    miDo(Copy(ss, 1, Length(ss)-1) + ')');

    s := ExtractFilePath(xls) + 'Общий_Результат_graf.dbf';
    miDo('Select %s From %s Into TempTab NoSelect', [date, tab]);
    miDo('Export TempTab Into "%s" Type "DBF" CharSet "WindowsCyrillic"', [s]);
//------------------------------------------------------------------------------

    ExcelApp.Visible := true;
    Info('Готово');

  finally
    ExcelApp := Unassigned;
    WorkSheet := Unassigned;
    sl.Free;
    dbf.Free;
    files.Free;
  end;
end;

//расчет ценовых факторов по введенным координатам для проверки
procedure TServiceMapForm.mnCheckClick(Sender: TObject);
var
  ExcelApp, Workbook, WorkSheet, DataParcels : OleVariant;
  s, ss, xls, txt, tab, analogs : String;
  i, j, n, m, kodnp, idkp, col, row, radius : Integer;
  xy : TPoint3D;
  sa : StringArray;
  sl, log, fields : TStringList;
  ini: TMemIniFile;
begin
  analogs := GetAppFolder + 'analogs\';
  if not DirectoryExists(analogs) then
    Exc('Не найден каталог "Analogs" с таблицами для расчеты анвлогов.');

  tab := TempPath + 'Проверка.tab';
  xls := ChangeFileExt(tab, '.xlsx');//имя файда excel
  if FileExists(xls) then
    if not DeleteFile(xls) then
      Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [xls, #10]));
  if not InputQuery('Ввод координат', 'x/y', s) then
    exit;//Exc('Выбор прерван');

  fields := TStringList.Create;//набор полей исходного файла
  radius := 0;
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
    ini.ReadSectionValues('Interface', fields);
    for i := 0 to fields.Count-1 do
      if CText(fields.Names[i], 'Radius_Roads') then
        Val(fields.ValueFromIndex[i], radius, n);
    if radius = 0 then
       radius := 30;
    fields.Clear;

    ini.ReadSectionValues('Analysis', fields);
    for i := fields.Count-1 downto 0 do
      if Pos('//', fields[i]) = 1 then
        fields.Delete((i));
    for i := fields.Count-1 downto 0 do
    begin
      if (Pos('=Contains', fields[i]) = 0) and (Pos('=Intersects', fields[i]) = 0) then
        fields.Delete((i));;
    end;
  finally
    ini.Free;
  end;

  Info('Создание таблицы...');
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;
  Workbook := ExcelApp.WorkBooks.Add;//  Создаем Книгу (Workbook)
  WorkSheet := WorkBook.WorkSheets.item[1];//активный 1 лист
  try
    SafeCloseTable('Проверка');
    miDo('Create table Проверка (coord_wqs Char(50)) File "%s"', [tab]);
    miDo('Create Map For Проверка CoordSys Earth Projection 8, 1001, "m", 60.05, '+
      '0, 1, 1500000, -5911057.63');
    miDo('Set CoordSys Table Проверка');

    row := 1;
    col := fields.Count+1;
    DataParcels := VarArrayCreate([1, row, 1, col], varVariant);
    WorkSheet.Cells[2, 1] := s;
    j := Pos(' ', s);
    ss := Copy(s, j+1, Length(s));
    s := Copy(s, 1, j-1);
    Val(ss, xy.x, m);
    Val(s, xy.y, m);
    Wgs2Msk_(xy.y, xy.x, 2);
    miDo('Create Point Into Variable ob	(%.9f, %.9f) Symbol MakeSymbol(34,%d,%d)', [xy.x, xy.y, RGB(0, 0, 255), 10]);
    miDo('Insert into Проверка (obj) Values (ob)');
    idkp := 0;//для создание буферных зон
    kodnp := 0;//для создание буферных зон
    i := 0;
    While i < fields.Count-1 do
    begin
      if Pos('Contains', fields.ValueFromIndex[i]) = 1 then
      try
        s := Copy(fields.ValueFromIndex[i], 10, 100);
        sa := SplitStr(s, ';');
        if FileExists(analogs + sa[0] + '.tab') then//полное название таблицы
          Contains(i, row, kodnp, idkp, fields, WorkSheet, DataParcels,
            'Проверка', analogs + sa[0] + '.tab', sa[1])
        else
          log.Add(format('Файл %s не найден', [analogs + sa[0] + '.tab']));
      finally
        SetLength(sa, 0);
      end;
      Inc(i);
    end;
//------------------------расчет по дорогам-------------------------------------
      i := 0;
      While i < fields.Count-1 do
      begin
        if Pos('Intersects', fields.ValueFromIndex[i]) = 1 then
        try
          s := Copy(fields.ValueFromIndex[i], 12, 100);
          sa := SplitStr(s, ';');
          if FileExists(analogs + 'Дороги_WGS84\' + sa[0] + '.tab') then//полное название таблицы
            Intersects(i, row, radius, fields, WorkSheet, DataParcels,
              'Проверка', analogs + 'Дороги_WGS84\' + sa[0] + '.tab', sa[1])
          else
            log.Add(format('Файл %s не найден', [analogs + 'Дороги_WGS84\' + sa[0] + '.tab']));
        finally
          SetLength(sa, 0);
        end;
        Inc(i);
      end;
//------------------------------------------------------------------------------
    WorkSheet.Cells[1, 1] := 'coord_wqs';
    for i := 0 to fields.Count-1 do
      WorkSheet.Cells[1, i+2] := fields.names[i];
    WorkBook.SaveAs(Filename:=xls, FileFormat:= 51, CreateBackup:=False);
    ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
    ExcelApp.Visible := true;
    Info('Готово');
  finally
    miDo('Drop table Проверка');
    fields.Free;
    ExcelApp := Unassigned;
    Workbook := Unassigned;
    WorkSheet := Unassigned;
    VarClear(DataParcels);
  end;
end;

//ziis системный пароль для управления данными в базе
{Для обработки аналогов мы используем таблицу geo_dk_semantic  в БД btiall
Ее необходимо регулярно обновлять сведениями из ПО Телды
Из Телды мы можем только выгружать данные по всем ЗУ, т.к. не знаем какие объекты изменились.
Выгрузка получается в виде архива. Внутри однородные xlsx файлы

Соответствия между полями таблицы и столбцами xlsx прописаны в файле. Возможно в будущем мы расширим количество забираемых в таблицу столбцов. Подумай над необходимостью системы настройки соответствий. 
Необходимо достать данные из xlsx файлов. Сравнить с данными в таблице geo.DK_SEMANTIC.
Добавить в geo.DK_SEMANTIC новые записи и обновить данные по записям у которых есть различия по любому из полей (за исключением ключевого CADASTRALNUMBER) . 
Текущий архив в W:\ЦГКО\IT\Outbox\Текущие задачи\2024\Сбор Аналогов\DK_SEMANTIC}
procedure TServiceMapForm.mnDKSEMANTICClick(Sender: TObject);
var
  ExcelApp, Workbook, WorkSheet, range, data : OleVariant;
  sl, fields, log : TStringList;
  ini : TMemIniFile;
  s, ss, s2, cmd, f, basename : String;
  index : Array of Integer;
  ds : TADODataset;
  sa : StringArray;
  i, j, k, n, col, row, newrow, upprow : Integer;

  function RunCommand(var cmd : String) : Boolean;
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
        log.Add('Операция отменена' + #10 + e.Message);
      end;
    end;
  end;

  Function str_(data : OleVariant) : String;
  var
    s : String;
  begin
    s := data;
    Result := StringReplace(StringReplace(s, #10, ' ', [rfReplaceAll]), #13, ' ', [rfReplaceAll]);
    Result := StringReplace(s, #9, ' ', [rfReplaceAll]);
    Result := StringReplace(s, #39, #34, [rfReplaceAll]);
  end;

begin//mnDKSEMANTICClick
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\Аналоги_test\DK_SEMANTIC'; {$ENDIF}
  OpenDialog1.FileName := '*.zip';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;
  sl := TStringList.Create;
  fields := TStringList.Create;
  log := TStringList.Create;
  try
    ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
    try
      basename := ini.ReadString('DataBase', 'BaseName', 'btiAll');
      ss := ini.ReadString('sevenzip', 'path', 'C:\Program Files\7-Zip\7z.exe');
      ini.ReadSectionValues('DK_SEMANTIC', fields);
    finally
      ini.Free;
    end;

    ExCreateProcess(format('"%s" x "%s" -y -o"%s"', [ss, OpenDialog1.FileName, TempPath]), '', SW_HIDE, true);
    FindRecursive(AddSlesh(TempPath), '*.xlsx', false, sl);

    newrow := 0;
    upprow := 0;
    ADOConnection1.ConnectionString := format(strConnectBTI, [basename]);
    ADOConnection1.Connected := true;
    ds := TADODataset.Create(nil);
    ds.CacheSize := 100;
    ds.CursorLocation := clUseServer;
    ds.Connection := ADOConnection1;
    f := fields.Names[0];
    for i := 1 to fields.Count-1 do
      f := f + ',' + fields.Names[i];
    SetLength(index, fields.Count);//номера колонок, соответствующие полям ds

    for k := 0 to sl.Count-1 do
    begin
      Info('Открытие ' + ExtractFileName(sl[k]));
      ExcelApp.Workbooks.Open(sl[k]);
      Workbook := ExcelApp.Workbooks.item[1];
      WorkBook.WorkSheets.Item[1].Activate;
      WorkSheet := ExcelApp.ActiveSheet;
      col := WorkSheet.UsedRange.Columns.Count;
      row := WorkSheet.UsedRange.Rows.Count;

      for i := 0 to fields.Count-1 do
      begin
        s := fields.ValueFromIndex[i];//найдем название колонки в Excel
        for j := 1 to col do//заполняем номера колонок
          if CText(String(WorkSheet.Cells[5, j]), s) then
            index[i] := j;
      end;

      data := VarArrayCreate([1, 1, 1, col], varVariant);
      s2 := ExcelNum2Str(WorkSheet, col);//по строкам
      pg.Max := row;
      Info('Обработка ' + ExtractFileName(sl[k]));
      for i := 6 to row do
      try
        if i mod 100 = 0 then
          pg.Position := i;
        range := WorkSheet.range[format('A%d:%s%d', [i, s2, i])];
        data := range.value;
        ds.Close;
        ds.CommandText := format('Select %s from geo.dk_semantic where CADASTRALNUMBER=''%s''',
          [f, String(data[1, index[0]])]);
        ds.Open;
        if ds.IsEmpty then//новая запись
        begin
          cmd := format('Insert into geo.DK_SEMANTIC (%s) values(''%s'')', [f, String(data[1, index[0]])]);
          for j := 1 to fields.Count-1 do
            System.Insert(format(',''%s''', [str_(data[1, index[j]])]), cmd, Length(cmd));
          RunCommand(cmd);
          Inc(newrow);
        end else
        begin
          for j := 1 to fields.Count-1 do
          begin
            s := ds.FieldByName(fields.Names[j]).AsString;//строка из базы, DATECREATED: yyyy-mm-dd
            ss := data[1, index[j]];//строка из Excel, DATECREATED: dd.mm.yyyy
            if CText(fields.Names[j], 'DATECREATED') then
            begin
              if s <> '' then
              begin
                sa := SplitStr(Copy(s, 1, 10), '-');//строка всегда не пустая
                s := sa[2] + '.' + sa[1] + '.' + sa[0];
                SetLength(sa, 0);
              end;
              if ss = '' then
              begin
                ss := s;//порядок года, месяца и дня значения не имеет
                log.Add(String(data[1, index[0]]) + ' DATECREATED нет данных');
              end;
            end;
            if not CText(s, ss) then
            begin
              cmd := 'Update geo.DK_SEMANTIC Set ';
              for n := 1 to fields.Count-1 do
                cmd := cmd + format('%s = ''%s'',', [fields.Names[n], str_(data[1, index[n]])]);
              cmd := Copy(cmd, 1, Length(cmd)-1) + format(' Where CADASTRALNUMBER = ''%s''',
                [String(data[1, index[0]])]);
              RunCommand(cmd);
              Inc(upprow);
              Break;
            end;
          end;
        end;
      except
        on e : Exception do
        begin
          Info(String(data[1, index[0]]));
          Note(e.Message);
        end;
      end;
      pg.Max := 0;

      VarClear(data);
      ExcelApp.Workbooks.Close;
      ExcelApp.Quit;
      Application.ProcessMessages;
    end;
  finally
    log.Add(format('Добавлено записей %d%sОбновлено записей %d', [newrow, #10, upprow]));
    s := ChangeFileExt(OpenDialog1.FileName, '.log');
    log.SaveToFile(s);
    ExShellExecute('Notepad.exe', s, SW_NORMAL, false, 0);

    ExcelApp := Unassigned;
    Workbook := Unassigned;
    WorkSheet := Unassigned;
    VarClear(data);
    sl.Free;
    fields.Free;
    log.Free;
    ds.Close;
    ds.Free;
    ADOConnection1.Close;
    SetLength(index, 0);
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.CorrectXMLClick(Sender: TObject);
var
  sl : TStringList;
  s, s2, catzem : String;
  i, n, m : Integer;
  mc : TMatchCollection;
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\Аналоги_test\XML'; {$ENDIF}
  OpenDialog1.FileName := '*.xml';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  sl := TStringList.Create;
  try
    sl.LoadFromFile(OpenDialog1.FileName);
    s := '';
    for i := 0 to sl.Count-1 do
      s := s + Trim(sl[i]);//все в одну строку

    n := Pos('<Objects>', s);
    n := PosEx('<', s, n+12);//добавляем длину больше длины <Objects> и сразу находим начало catzem
    m := PosEx(' ', s, n+1);
    catzem := Copy(s, n+1, m-n-1);
    //добавим перевод строк для правильной работы Matches
    s2 := format('</%s>', [catzem]);
    s := StringReplace(s, s2, s2+#10, [rfReplaceAll]);

    mc := reg.Matches(s, '<SubParcels>.+</SubParcels>', [roIgnoreCase, roMultiLine]);
    n := mc.Count;
    for i := 0 to n-1 do
      s := StringReplace(s, mc.Item[i].Value, '', [rfReplaceAll]);

    mc := reg.Matches(s, '<ZonesAndTerritories>.+</ZonesAndTerritories>', [roIgnoreCase, roMultiLine]);
    m := mc.Count;
    for i := 0 to m-1 do
      s := StringReplace(s, mc.Item[i].Value, '', [rfReplaceAll]);//447278

    sl.Text := s;
  finally
    s := ExtractFilePath(OpenDialog1.FileName);
    s2 := ExtractFileName(OpenDialog1.FileName);
    s2 := ChangeFileExt(s2, 'испр.xml');
    sl.SaveToFile(s + s2);
    sl.Free;
    Note(format('Удалено:%s SubParcels %d%s ZonesAndTerritories %d', [#10, n, #10, m]))
  end;
end;
{------------------------------------------------------------------------------}
procedure TServiceMapForm.mnSelectUserFieldsClick(Sender: TObject);
var
  sl : TStringList;
  ini : TMemIniFile;
  i, j, k, n, col : Integer;
  xls, alert : String;
  ExcelApp, Workbook, WorkSheet : OleVariant;
begin
  OpenDialog1.FileName := '*.xlsx';//D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\TEST\от_Шмелева_Сводный_Результат\Сводный_Результат_2.xlsx
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;
  sl := TStringList.Create;
  try
    ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
    try
      ini.ReadSectionValues('FIELDS_USER', sl);
    finally
      ini.Free;
    end;
    xls := ChangeFileExt(OpenDialog1.FileName, '_Пользователь.xlsx');
    if FileExists(xls) then
      if not DeleteFile(xls) then
        Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [xls, #10]));

    Info('Загрузка данных...');
    ExcelApp.Workbooks.Open(OpenDialog1.FileName);
    Workbook := ExcelApp.Workbooks.item[1];
    WorkBook.WorkSheets.Item[1].Activate;
    WorkSheet := ExcelApp.ActiveSheet;
    col := WorkSheet.UsedRange.Columns.Count;//берем размеры на листе

    Info('Выборка данных...');
    alert := '';
    WorkBook.WorkSheets.Add(Before:=ExcelApp.ActiveSheet, Count:=1);//вставляем новый лист перед активным
    n := 1;
    for i := 0 to sl.Count-1 do
    begin
      WorkBook.WorkSheets.Item[2].Activate;
      WorkSheet := ExcelApp.ActiveSheet;
      pg.Position := i+1;
      k := 0;
      for j := 1 to col do
        if CText(sl[i], String(WorkSheet.Cells[1, j])) then
      begin
        WorkSheet.Columns[j].Select;
        ExcelApp.Selection.Copy;
        WorkBook.WorkSheets.Item[1].Activate;
        WorkSheet := ExcelApp.ActiveSheet;
        WorkSheet.Columns[n].Select;
        ExcelApp.Selection.PasteSpecial(Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks :=False, Transpose:=False);
        Inc(n);
        k := 1;
        break;
      end;
      if k = 0 then
        alert := alert + sl[i] + #10;
    end;

    WorkBook.Worksheets[2].Delete;//старый лист удаляем
    pg.Max := col;
    pg.Max := 0;
    Info('Готово');
    WorkBook.SaveAs(Filename:=xls, FileFormat:=51, CreateBackup:=False);
    ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация - альбомная
    ExcelApp.Visible := true;
    if alert <> '' then
      Note('Не найдены колонки:' + #10 + alert + 'Проверьте секцию [FIELDS_USER]');
  finally
    ExcelApp := Unassigned;
    Workbook := Unassigned;
    WorkSheet := Unassigned;
    sl.Free;
  end;
end;

end.
