unit CostFactors_Unit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, System.Net.HttpClient, Data.DB,
  Data.Win.ADODB, Vcl.Menus, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.ExtCtrls,
  System.RegularExpressions,
  System.RegularExpressionsCore,
  System.IniFiles,
  StrUtils,
  ComObj,
  ActiveX,
  System.Contnrs,
  GeoTransform,
  MapBasic_INT,
  cfCommon,
  System.Net.URLClient,
  System.NetEncoding,
  System.Net.HttpClientComponent,
//  Clipbrd,
  Common, Vcl.Buttons,
  Vcl.CheckLst;

const
  none = 'Значение не определено';

//ConnectionString для ADOConnection1 менять для каждой новой версии CF
  // Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Tokyo\SERVICEMAP\Win32\Debug\Ценовые факторы_3\коммуникации\Коммуникации_свод.accdb;Persist Security Info=False
  strConnect1 =
    'Provider=Microsoft.ACE.OLEDB.12.0;Data Source="%s";Persist Security Info=False';

//  // Provider=SQLOLEDB.1;Password=s;Persist Security Info=True;User ID=zisadmin;Initial Catalog=btizu2022;Data Source=TESTSQL;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=G001;Use Encryption for Data=False;Tag with column collation when possible=False
//  strConnect2 = 'Provider=SQLOLEDB.1;Password=s;Persist Security Info=True;' +
//    'User ID=zisadmin;Initial Catalog=%s;Data Source=TESTSQL;' +
//    'Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;' +
//    'Workstation ID=G001;Use Encryption for Data=False;Tag with column collation when possible=False';

  NUMBER_MESSAGE = WM_USER + 102;//идентификатор сообщенияв диапазоне от WM_USER+100 до Ç7FFFF

type
  PzouitObj = ^TzouitObj;
  TzouitObj = Record
    idn : Integer;
    cn : String[35];
    zone: String[150];
    subzone : Integer;
    soil_ap : String[160];
    area : Double;
    overlap : Double;
    proc : Double;
  end;

  // класс-замена талиц мапинфо,содержащщий координаты центроидов
  TCoord = class(TObject)
  public
    table: String;
    dim: Array of dPoint;
    constructor Create(name: String; dimcount: Integer);
    destructor Destroy; override;
  end;

  TForm2 = class(TForm)
    OpenDialog1: TOpenDialog;
    Panel4: TPanel;
    stBar: TStatusBar;
    pg: TProgressBar;
    MainMenu1: TMainMenu;
    mnCostFactory: TMenuItem;
    Procent: TLabeledEdit;
    mnCalc: TMenuItem;
    mnGraphic: TMenuItem;
    mnBase: TMenuItem;
    ADOConnection1: TADOConnection;
    mnCommunicationItog: TMenuItem;
    mnComplexCommunication: TMenuItem;
    mnZOUIT: TMenuItem;
    mnSoil: TMenuItem;
    mnComplexOther: TMenuItem;
    mnCommunications: TMenuItem;
    N7: TMenuItem;
    mnOthers: TMenuItem;
    N9: TMenuItem;
    mnObjectsFromText: TMenuItem;
    cbFindCN: TCheckBox;
    TimeOut: TLabeledEdit;
    mnObjectsFromXls: TMenuItem;
    Query: TADOQuery;
    cbExamenOGR: TCheckBox;
    mnTable: TMenuItem;
    mnCreateTableFromTab: TMenuItem;
    mnApplaCosFactory: TMenuItem;
    N1: TMenuItem;
    mnBuffers: TMenuItem;
    SpedTest1: TMenuItem;
    mnCF: TMenuItem;
    mnOpenWorkTable: TMenuItem;
    mnTerZone: TMenuItem;
    mnUtils: TMenuItem;
    ZIP_mifmid_tab: TMenuItem;
    mnMi2Excel: TMenuItem;
    Elipce2Rect: TMenuItem;
    XML1: TMenuItem;
    mnQueryRR: TMenuItem;
    Access1: TMenuItem;
    MultiCircuitObject1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure mnMultiContourClick(Sender: TObject);
    procedure mnTimesClick(Sender: TObject);
    procedure mnZonesClick(Sender: TObject);
    procedure mnCode_MOClick(Sender: TObject);
    procedure mnLinkClick(Sender: TObject);
    procedure LoadAnyData(Sender: TObject);
    procedure mnGraphicClick(Sender: TObject);
    procedure mnBaseClick(Sender: TObject);
    procedure mnCommunicationItogClick(Sender: TObject);
    procedure mnComplexCommunicationClick(Sender: TObject);
    procedure mnZOUITClick(Sender: TObject);
    procedure mnSoilClick(Sender: TObject);
    procedure mnComplexOtherClick(Sender: TObject);
    procedure mnItogCFClick(Sender: TObject);
    function CF_from_CN : Boolean;
    procedure mnObjectsFromTextClick(Sender: TObject);
    procedure mnObjectsFromXlsClick(Sender: TObject);
    procedure mnCreateTableFromTabClick(Sender: TObject);
    procedure mnApplaCosFactoryClick(Sender: TObject);
    procedure mnBuffersClick(Sender: TObject);
    procedure SpedTest1Click(Sender: TObject);
    procedure mnCFClick(Sender: TObject);
    procedure mnOpenWorkTableClick(Sender: TObject);
    procedure mnTerZoneClick(Sender: TObject);
    procedure OpenDialog1Close(Sender: TObject);
    procedure ZIP_mifmid_tabClick(Sender: TObject);
    procedure mnMi2ExcelClick(Sender: TObject);
    procedure Elipce2RectClick(Sender: TObject);
    procedure XML1Click(Sender: TObject);
    procedure mnQueryRRClick(Sender: TObject);
    procedure Access1Click(Sender: TObject);
    procedure MultiCircuitObject1Click(Sender: TObject);
  private
    { Private declarations }
//    ds: TADODataset;
    client: THTTPClient;
    ss: TStringStream;
    AnyData: Array of TCostFactory;
    cad, prevcad: String[32]; // кадастровый номер текущий и следующий за ним
    row: Integer;
    multi: Integer;
    pn, pk: TPoint3D;//начальная и конкчная точки в долготе/широте по MapInfo
    cn: TStringList; // кадасстровые номера
    po_cf: String; // версия ПО и СФ
    direction: TObjectList; // список TCoord для расчета расстояний
    cpu: Integer;
    LoadFields: TStringList;
    AddFields: TStringList;
    PatternFields: TStringList;

    procedure WMCopyData(var Msg: TWMCopyData); message WM_COPYDATA;
    procedure PrepareCoords;
    function OpenTable(source, title : String): Boolean;
    function CreateWorkTable : Boolean;
    procedure CreatePrimaryObjects;
    procedure CalcZone(zonecf: Integer);
    procedure WeightedAverage;
    procedure AverageWeighted;
    procedure CalcMultiContour;
    procedure CopyTable(table, path: String);
    procedure MoveCostFactory(part: Integer; cf: String);
    function Find_CN_IDN(table: String): String;
    procedure Info(Msg: String);
    function ExamenBufferSize : Boolean;
    procedure CalcMerge;
    procedure CalcIntersect2;
    procedure CalcFactors;
    function CalcCPUCount: Integer;
    function CompactDatabase(SrcName: String): Boolean;
    procedure FillEmptyFields(table: String);
    procedure TimeCalc(tab, tif, field: String);
    function CheckWorkTable(procname : String) : Boolean;
    procedure ServerCertValidate(const Sender: TObject;
      const ARequest: TURLRequest; const Certificate: TCertificate;
      var Accepted: Boolean);
    function QueryRosReestrV2(var id: Integer; var cadastr: String): Integer;
    function CheckSprPath : Boolean;
    function Xls2Txt : Boolean;
    procedure CreateLoaderFile;
    procedure Compare_13_segment;
    procedure RunMapBasic(dir : String);
  public
    { Public declarations }
    st16 : Boolean;
    TempPath : String;
    feat : Integer;
    curdata : String;
    ftab, fcn: String[32];
    worktable: String;// названи базовой таблицы
    ResFolder: String;
  end;

var
  Form2: TForm2;
  cs1 : String = 'CoordSys Earth Projection 8, 1001, "m", 60.05, 0, 1, 1500000, -5911057.6299999999';
  cs2 : String = 'CoordSys Earth Projection 8, 1001, "m", 66.05, 0, 1, 2500000, -5911057.6299999999';
  cs11 : String = 'CoordSys Earth Projection 8, 9999, 3, 23.57, -140.95, -79.8, 0, -0.35, -0.79, -0.22, 0, "m", 60.05, 0, 1, 1500000, -5911057.63';
  cs21 : String = 'CoordSys Earth Projection 8, 9999, 3, 23.57, -140.95, -79.8, 0, -0.35, -0.79, -0.22, 0, "m", 66.05, 0, 1, 2500000, -5911057.63';

implementation

uses ServiceMap_Unit;

{$R *.dfm}

constructor TCoord.Create;
begin
  inherited Create;
  table := name;
  SetLength(dim, dimcount);
end;
{------------------------------------------------------------------------------}
destructor TCoord.Destroy;
begin
  SetLength(dim, 0);
  inherited Destroy;
end;
{------------------------------------------------------------------------------}
procedure TForm2.WMCopyData(var Msg: TWMCopyData);
var
  s: string;
begin
  s := PChar(Msg.CopyDataStruct.lpData);
  if Pos('RESULT', s) > 0 then
    Dec(cpu);
  Msg.Result := 2006; // Отошлем что-нибудь назад
end;
{------------------------------------------------------------------------------}
procedure TForm2.Info(Msg: String);
begin
  stBar.Panels[1].Text := Msg;
  stBar.Repaint;
end;
{ ------------ищем поле с кадастровым номером и проверяем индексное поле------- }
function TForm2.Find_CN_IDN(table: String): String;
var
  i: Integer;
  value, idn: String;
  sl: TStringList;
begin
  Result := '';
  sl := getFieldsList(table);
  miDo('Fetch First From %s', [table]);
  try
    idn := '';
    for i := 0 to sl.Count - 1 do
    begin
      value := miEval('%s.%s', [table, sl[i]]);
      if (Result = '') and reg.IsMatch(value, '\d\d:\d\d:\d\d\d\d\d\d{1,}:\d{1,}') then
        Result := sl[i];
      if CText('idn', sl[i]) then
        idn := sl[i];
    end;

    if Result <> '' then
    begin
      value := '';
      for i := 0 to sl.Count - 1 do
      begin
        miDo('Drop index %s (%s)', [table, sl[i]]); // удалим индексы
        value := value + ',' + sl[i];//список поолей через запятую
      end;
      if idn = '' then
        miDo('Alter Table "%s" (add IDN Integer order %s)', [table, 'IDN'+value]);
    end;
  finally
    sl.Free;
  end;
end;
{------------------------------------------------------------------------------}
procedure TForm2.FormCreate(Sender: TObject);
begin
  LoadFields := TStringList.Create;
  AddFields := TStringList.Create;
  PatternFields := TStringList.Create;
//  ds := TADODataset.Create(nil);
//  ds.CacheSize := 100;
//  ds.CursorLocation := clUseServer;

  direction := TObjectList.Create;
  cn := TStringList.Create;
  client := THTTPClient.Create;
  ss := TStringStream.Create('', TEncoding.ASCII);
  curdata := CurrentData;
end;
{------------------------------------------------------------------------------}
procedure TForm2.FormDestroy(Sender: TObject);
begin
  LoadFields.Free;
  AddFields.Free;
  PatternFields.Free;
  while direction.Count > 0 do
    direction.Delete(0);
  direction.Free;

//  ds.Close;
//  ds.Free;
  cn.Free;
  client.Free;
  ss.Free;
  ADOConnection1.Close;
end;
{------------------------------------------------------------------------------}
procedure TForm2.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 27 then
    BreakCommand := true;
end;
{ -------------заполнение рабочего массова из Excel ---------------------------- }
procedure TForm2.LoadAnyData(Sender: TObject);
var
  ExcelApp, Workbook, WorkSheet: OleVariant;
  i, k : Integer;
  s, xlsname : String;
begin
  if Sender <> nil then
  begin
    GetVersionInfo(GetModuleName(HInstance), s);
    po_cf := format('%s/%s', [Copy(s, 8, 100), (Sender as TMenuItem).caption]);
    stBar.Panels[0].Text := (Sender as TMenuItem).caption;
  end;
  xlsname := GetAppFolder + stBar.Panels[0].Text + '\CostFactory.xlsx';
  SetLength(AnyData, 0);

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;
  try
    ExcelApp.Workbooks.Open(xlsname);
    Workbook := ExcelApp.Workbooks.item[1];
    WorkSheet := Workbook.WorkSheets.item[1];

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

    LoadFields.Clear;
    AddFields.Clear;
    PatternFields.Clear;
    try
      WorkSheet := Workbook.WorkSheets.item[2];
      i := 1;
      if String(WorkSheet.Cells[i, 1]) = '' then
        i := 2;
      Repeat
        s := String(WorkSheet.Cells[i, 1]);
        if s = '' then
          break;
        // выбираем список полей для загрузки
        LoadFields.Add(s + '=' + String(WorkSheet.Cells[i, 2]) + ',' +
          String(WorkSheet.Cells[i, 3]) + ',' + String(WorkSheet.Cells[i, 4]));

        // выбираем дополнительные поля
        s := String(WorkSheet.Cells[i, 5]);
        if s <> '' then
          AddFields.Add(s + '=' + String(WorkSheet.Cells[i, 6]) + ',' +
            String(WorkSheet.Cells[i, 7]));

        // выбираем заполнение дополнительных полей
        s := String(WorkSheet.Cells[i, 8]);
        if s <> '' then
          PatternFields.Add(s + '=' + String(WorkSheet.Cells[i, 9]));

        Inc(i);
      Until false;
    except
      LoadFields.Clear;
    end;

  finally
    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;
    WorkSheet := Unassigned;
    Workbook := Unassigned;
    ExcelApp := Unassigned;
  end;
end;
{------------------------------------------------------------------------------}
procedure TForm2.FormShow(Sender: TObject);
var
  sl, cat : TStringList;
  s : String;
  i, j : Integer;
begin
  try
    mnCostFactory.Clear;
    sl := TStringList.Create;
    cat := TStringList.Create;
    FindFolders(GetAppFolder, cat, false);
    for j := cat.Count - 1 downto 0 do
    begin
      sl.Clear;
      if Pos('CF_', cat[j]) > 0 then
        FindRecursive(AddSlesh(cat[j]), '*.xlsx', false, sl);
      for i := sl.Count - 1 downto 0 do
        if Pos('\costfactory.xlsx', AnsiLowerCase(sl[i])) > 0 then
        begin
          s := ExtractFileName(ExtractFileDir(sl[i]));
          mnCostFactory.Add(NewItem(s, 0, false, true, LoadAnyData, 0, 'mn' + s));
        end;
    end;
    if mnCostFactory.Count > 0 then
      mnCostFactory.Items[0].Click;
  finally
    sl.Free;
    cat.Free;
  end;
end;
{ -----------------------расчет времянных характеристик------------------------- }
procedure TForm2.TimeCalc(tab, tif, field: String);
var
  sl : TStringList;
  path : String;
  i, j : Integer;
begin
  path := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'RASTER\';
  SetCurrentDir(path);
  tif := ChangeFileExt(tif, '');
  if not FileExists(tif + '.tif') then
    Exc('TimeCalc : Нет файла ' + tif);
  sl := TStringList.Create;
  try
    Info(field + ': Вычисление временных характеристик...');
    sl.LoadFromFile('gdallocationinfo.bat');
    sl[0] := format(sl[0], [tif]);
    ExShellExecute('cmd', PChar(format('/c %s', [sl[0]])), SW_HIDE, true, 0);

    sl.LoadFromFile('output.txt');
    Application.ProcessMessages;
    Info(tif + ': Запись временных характеристик...');
    for i := 0 to sl.Count - 1 do
    begin
      Val(sl[i], pn.z, j);
      if pn.z = 0 then
        pn.z := -1; // помечаем как невычисленную
      sl[i] := IntToStr(i+1) + #9 + IntToStr(Round(pn.z));
    end;
//    sl.SaveToFile('outputtable.txt');
//    miDo('Register Table "%soutputtable.txt"  TYPE ASCII Delimiter 9 '+
//      'Charset "WindowsCyrillic" Into "%soutputtable.TAB"', [path, path]);
//    miDo('Open Table "%soutputtable.TAB"', [path]);
//    miDo('Add Column %s (%s) From outputtable Set To _COL2 Where COL1 = COL1', [tab, field]);
//    miDo('Close Table outputtable');
  finally
    sl.Free;
  end;
end;
{----------------------копирование ценовых характеристик-----------------------}
procedure TForm2.MoveCostFactory(part: Integer; cf: String);
var
  path: String;
begin
  path := format('%s%d\%s', [TempPath, part, cf]);
  ForceDirectories(path);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s*.tab" "%s"',
    [GetAppFolder + cf, path])), SW_HIDE, true, 0);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s*.map" "%s"',
    [GetAppFolder + cf, path])), SW_HIDE, true, 0);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s*.dat" "%s"',
    [GetAppFolder + cf, path])), SW_HIDE, true, 0);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s*.id" "%s"',
    [GetAppFolder + cf, path])), SW_HIDE, true, 0);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s*.ind" "%s"',
    [GetAppFolder + cf, path])), SW_HIDE, true, 0);
end;
{------------------------------------------------------------------------------}
procedure TForm2.MultiCircuitObject1Click(Sender: TObject);
var
  i : Integer;
  cn : String;
  x0, y0, x1, y1 : Double;
begin
  miDo('Set ProgressBars Off');
  miDo('Open table "d:\temp\ЗУ_ДК"');
  miDo('Set coordsys table ЗУ_ДК');
  miDo('Select КН,count(*) from ЗУ_ДК group by КН into TempTab noselect');
  miDo('select * from TempTab where count > 1 into MultiCirc noselect');
  miDo('Commit table MultiCirc as "%sMultiCirc.tab"', [TempPath]);
  miDo('Close table "MultiCirc"');
  miDo('Open table "%sMultiCirc.tab"', [TempPath]);
  miDo('Alter Table MultiCirc (Add area Decimal(15,2), area_MBR Decimal(15,2),'+
    ' diagonal Decimal(12,2) Rename _col2 Count)');
  pg.Max := miEvalInt('TableInfo(MultiCirc, %d)', [TAB_INFO_NROWS]);
  for i := 1 to pg.Max do
  begin
    if i mod 100 = 0 then
      pg.Position := i;
    miDo('Fetch Rec %d From MultiCirc', [i]);
    cn := miEval('MultiCirc.кн');
    miDo('Select * from ЗУ_ДК where КН="%s" into TempTab', [cn]);
    miDo('Create Object As Union From TempTab Into Variable ob');
    miDo('ob2 = MBR(ob)');
    x0 := miEvalDbl('ObjectGeography(ob2, %d)', [OBJ_GEO_MINX]);
    y0 := miEvalDbl('ObjectGeography(ob2, %d)', [OBJ_GEO_MINY]);
    x1 := miEvalDbl('ObjectGeography(ob2, %d)', [OBJ_GEO_MAXX]);
    y1 := miEvalDbl('ObjectGeography(ob2, %d)', [OBJ_GEO_MAXY]);
    x0 := Sqrt(Sqr(x1-x0) + Sqr(y1-y0));
    miDo('Update MultiCirc Set area=CartesianArea(ob, "sq m"), area_MBR=CartesianArea(ob2, "sq m"),'+
      ' diagonal=%.2f Where rowid=%d', [x0, i]);
  end;
  miDo('Commit table MultiCirc');
  pg.Max := 0;
  Note('Ok');
  miDo('Set ProgressBars On');
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnTerZoneClick(Sender: TObject);
var
  ExcelApp, Workbook, WorkSheet: OleVariant;
  UserAgentId, i, j : Integer;
  cn, tab : String;
begin
  OpenDialog1.FileName := '*.xlsx';
  if not OpenDialog1.Execute then
    exit;

  miDo('Open Table "%s"', [GetAppFolder + 'Base\Дежурная_карта.tab']);
  cn := ChangeFileExt(OpenDialog1.FileName, '.tab');
  tab := miEval('PathToTableName$("%s")', [cn]);
  SafeCloseTable(tab);
  miDo('Create Table "%s" (id Char(20)) file "%s" '+
    'TYPE NATIVE Charset "WindowsCyrillic"', [tab, cn]);
  miDo('Create Map For %s %s',
    [tab, miEval('TableInfo("Дежурная_карта", %d)',[TAB_INFO_COORDSYS_CLAUSE])]);
  feat := 6;

  Inc(UserAgentId);
  if UserAgentId >= High(UserAgent) then
    UserAgentId := 0;
  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;

  try
    ExcelApp.Workbooks.Open(OpenDialog1.FileName);
    Workbook := ExcelApp.Workbooks.item[1];
    WorkSheet := Workbook.WorkSheets.item[1];

    if WorkSheet.UsedRange.Columns.Count = 2 then
    begin
      i := 1;
      Repeat
        if String(WorkSheet.Cells[i, 1]) = '' then
          break;
        cn := String(WorkSheet.Cells[i, 2]);
        Info('Поиск ' + cn);
        if cn <> '' then
          if QueryRosReestrV2(UserAgentId, cn) = 0 then
          begin
            if Pos('-', cn) = 0 then//например кадастровые номера вставляем не изменяя графику
              miDo('Insert into %s (id, obj) Values("%s", ob)', [tab, String(WorkSheet.Cells[i, 1])])
            else//вставляем по частям
              begin
                for j := 1 to  miEvalInt('ObjectInfo(ob, %d)', [OBJ_INFO_NPOLYGONS]) do
                begin
                  miDo('ob1 = ExtractNodes(ob, %d, 1, 1, 1)', [j]);//вставляем по частям
                  miDo('Insert into %s (id, obj) Values("%s", ob1)', [tab, String(WorkSheet.Cells[i, 1])]);
                end;
              end;
          end;
        miDo('Commit table %s', [tab]);
        Inc(i);
      Until false;
      miDo('Map From %s', [tab]);
    end else
      Note('В файле ' + cn + #10 + 'должно быть 2 колонки!');
  finally
    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;
    WorkSheet := Unassigned;
    Workbook := Unassigned;
    ExcelApp := Unassigned;
  end;
  Info(' Готово');
end;
{ ----------------------копирование таблицы в другой каталог-------------------- }
procedure TForm2.CopyTable(table, path: String);
var
  k: Integer;
begin
  ForceDirectories(path);
  k := Pos('.tab', AnsiLowerCase(table));
  if k > 0 then
    table := Copy(table, 1, k - 1); // название таблицы без расширения
  ExShellExecute('cmd', PChar(format('/c copy /b "%s.tab" "%s"', [table, path])
    ), SW_HIDE, true, 0);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s.map" "%s"', [table, path])
    ), SW_HIDE, true, 0);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s.dat" "%s"', [table, path])
    ), SW_HIDE, true, 0);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s.id" "%s"', [table, path])),
    SW_HIDE, true, 0);
  ExShellExecute('cmd', PChar(format('/c copy /b "%s.ind" "%s"', [table, path])
    ), SW_HIDE, true, 0);
end;
{ -----------перенос таблиц рассояний во внутренний список--------------------- }
procedure TForm2.CreatePrimaryObjects;
var
  i, j : Integer;
  c : TCoord;
begin
  miDo('Set CoordSys table %s', [ftab]);
  Info(' Открытие справочников...');
  cn.Clear;
  FindRecursive(GetAppFolder + AddSlesh(stBar.Panels[0].Text), '*.tab',
    false, cn);
  for i := 0 to cn.Count - 1 do
    miDo('Open table "%s"', [cn[i]]);

  miDo('select * from Расстояние_до_центра_МО where LCase$(Центр_МО)="екатеринбург" into TempTab NoSelect');
  miDo('eburg = TempTab.obj');
  miDo('Close table TempTab');
//далее НЕ АКТУАЛЬНО!!!
//  while direction.Count > 0 do
//    direction.Delete(0);
//
//  for i := Low(AnyData) to High(AnyData) do
//    try
//      if AnyData[i].CalcType = 2 then
//      begin
//        c := TCoord.Create(AnyData[i].table, miEvalInt('TableInfo(%s, %d)',
//          [AnyData[i].table, TAB_INFO_NROWS]));
//        direction.Add(c);
//        j := 0;
//        miDo('Fetch First From %s', [AnyData[i].table]);
//        While miEval('EOT(%s)', [AnyData[i].table]) <> 'T' do
//        begin
//          c.dim[j].x := miEvalDbl('CentroidX(%s.obj)', [AnyData[i].table]);
//          c.dim[j].y := miEvalDbl('CentroidY(%s.obj)', [AnyData[i].table]);
//          miDo('Fetch Next From %s', [AnyData[i].table]);
//          Inc(j);
//        end;
//      end;
//
//      if AnyData[i].Link = 101 then
//        miDo('ztb = Ekb_ring.obj')//зоны транспортного безразличия по объекту
//      else if AnyData[i].Link = 102 then
//      // найдем границу всех ИЖС в объект исключив Каменск и Нижний Тагил
//      begin
//        miDo('Select Зона_вляния_для_ИЖС.ID, Зона_вляния_для_ИЖС.Zona, ' +
//          'Внутризональное_расстояние.Центр_Зоны, Внутризональное_расстояние.Код_НП '
//          + 'from Зона_вляния_для_ИЖС, Внутризональное_расстояние where ' +
//          'Зона_вляния_для_ИЖС.Obj Contains Внутризональное_расстояние.Obj into TempIGS NoSelect');
//        miDo('Create Object As Union from TempIGS into Variable igs');
//        miDo('Close table TempIGS');
//      end;
//
//    except
//      on E: Exception do
//        Exc('CreatePrimaryObjects : Ошибка при создании справочников' + #10 +
//          E.message);
//    end;

//    rowcount := miEvalInt('TableInfo(%s, %d)', [ftab, TAB_INFO_NROWS]);//размер полной таблицы
//    row := 1;
end;
{------------------------------------------------------------------------------}
function EnumWindowsProc(wHandle: HWND; lb: TListBox): BOOL; stdcall;
var
  Title, ClassName: array [0 .. 255] of char;
begin
  GetWindowText(wHandle, Title, 255);
  GetClassName(wHandle, ClassName, 255);
  if IsWindowVisible(wHandle) then
    lb.Items.AddObject(string(Title), Pointer(wHandle));
  Result := true;
end;
{ ------------------пересчет координат из местной СК в WGS-------------------- }
procedure TForm2.PrepareCoords;
var
  i, j, k: Integer;
  path: String;
begin
  path := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'RASTER\';
  Info(' Подготовка координат...');
  miDo('Select Area, Area from %s into input NoSelect', [ftab]);//выбираем 2 колонки float
  miDo('Commit table input As "%sCoord.tab"', [path]);
  miDo('Close table input');
  miDo('Open table "%sCoord.tab"', [path]);
  miDo('Update Coord Set COL1=CentroidX(obj), COL2=CentroidY(obj)');
  miDo('Export "Coord" Into "%sinput.txt" Type "ASCII" Overwrite Delimiter "," '+
    'CharSet "WindowsCyrillic"', [path]);
  miDo('Drop table Coord');
  cn.LoadFromFile(path + 'input.txt');

  for i := 0 to cn.Count - 1 do // пересчитываем мск в WGS
  begin
    k := Pos(',', cn[i]);
    Val(Copy(cn[i], 1, k - 1), pn.x, j);
    Val(Copy(cn[i], k + 1, 100), pn.y, j); // коордитаты математические (MI)
    Msk2Wgs_(pn.y, pn.x, 2);// для точного позиционирование вычисляем по эллипсоиду 2
    cn[i] := format('%.9f %.9f', [pn.x, pn.y]);
  end;
  cn.SaveToFile(path + 'input.txt');//координаты в широте/долготе
//  path := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'RASTER\';
//  Info(' Подготовка координат...');
//  miDo('Select X, Y from %s into input NoSelect', [ftab]);//выбираем колонки координат из таблицы
//  miDo('Export "input" Into "%sinput.txt" Type "ASCII" Overwrite Delimiter "," CharSet "WindowsCyrillic"',
//    [path]);
//  miDo('Close table input');
//  cn.LoadFromFile(path + 'input.txt');
//
//  for i := 0 to cn.Count - 1 do // пересчитываем мск в WGS
//  begin
//    k := Pos(',', cn[i]);
//    Val(Copy(cn[i], 1, k - 1), pn.x, j);
//    Val(Copy(cn[i], k + 1, 100), pn.y, j); // коордитаты математические (MI)
//    Msk2Wgs_(pn.y, pn.x, 2);// для точного позиционирование вычисляем по эллипсоиду 2
//    cn[i] := format('%.9f %.9f', [pn.x, pn.y]);
//  end;
//  cn.SaveToFile(path + 'input.txt');//координаты в широте/долготе
end;
{ ---------------------определение пересечения с зоной-------------------------- }
procedure TForm2.CalcZone(zonecf: Integer);
var
  ar1, ar2, overlap: Double;
begin
  Info(AnyData[zonecf].table + ': определение зоны');// центроид объекта B лежит в границах объекта A;
  miDo('Select %s.%s, %s.%s from %s, %s where %s.Число_участков=1 '+
    'AND %s.Obj Contains %s.Obj into SelTemp noSelect',
    [ftab, AnyData[zonecf].Alias, AnyData[zonecf].table,
    AnyData[zonecf].field, ftab, AnyData[zonecf].table, ftab,
    AnyData[zonecf].table, ftab]);
  if miEvalInt('TableInfo(SelTemp, %d)', [TAB_INFO_NROWS]) > 0 then
    miDo('Update SelTemp Set col1=col2');
//      else
//        While (zonecf < High(AnyData)) and
//          (AnyData[zonecf].table = AnyData[zonecf + 1].table) do
//          Inc(zonecf);
  miDo('close table SelTemp');
end;
{---------таблица создается в каталоге ..Ценовые факторы_X\curdata---------}
//для создания рабочей таблицы выбирается таблица с кадастровым номером
function TForm2.CreateWorkTable : Boolean;
var
  columnlist: String;
  i: Integer;
begin
  Result := false;
  fcn := Find_CN_IDN('Source');
  if fcn = '' then
    Exc('CreateWorkTable : Не найдено поле с кадастровым номером');

  ResFolder := GetAppFolder + AddSlesh(StringReplace(stBar.Panels[0].Text, 'CF',
    'Ценовые факторы', [rfIgnoreCase])) + curdata + '\';
  ForceDirectories(ResFolder);

  try
    miDo('Set CoordSys table Source');
    Info(' Группировка объектов...');//добавим колонку с количеством частей участка
    miDo('Select %s,count(*) from Source group by %s into TempTab noselect',
      [fcn, fcn]);
    try
      miDo('Add Column Source(число_участков integer) from TempTab set to col2 where %s=%s', [fcn, fcn]);
    except
      Info(' число_участков существует');
    end;

    Info(' Сортировка объектов...');
    miDo('select idn,%s,число_участков from Source Order By %s where obj into TempTab noSelect',
      [fcn, fcn]);
    Info(' Создание таблицы...');
    miDo('Commit table TempTab As "%s" TYPE NATIVE Charset "WindowsCyrillic"',
      [ResFolder + ftab + '.tab']);
    miDo('Close table Source'); // TempTab тоже закроется
    worktable := ResFolder + ftab + '.tab';
    miDo('Open table "%s"', [worktable]);

    Info(' Добавление колонок....');
    columnlist := '';
    for i := Low(AnyData) + 3 to High(AnyData) do
      columnlist := columnlist + AnyData[i].Alias + ' ' + AnyData[i].fType + ',';
//    columnlist := columnlist + 'x float, y float';//добавим колонки координат
    columnlist := columnlist + 'Area float, версия_ПО_СФ Char(20)';//добавим колонки
    miDo('Alter table %s (Add %s Rename %s Кадастровый_номер)',
      [ftab, columnlist, fcn]);//добавляем поля из EXCEL
    fcn := 'Кадастровый_номер';

    Info(' Обновление колонок....');
//    miDo('Update %s Set IDN=rowid, Area=CartesianArea(obj, "sq m"),'+
//      'x=CentroidX(obj), y=CentroidY(obj), версия_ПО_СФ="%s"', [ftab, po_cf]);
    miDo('Update %s Set IDN=rowid, Area=CartesianArea(obj, "sq m"),'+
      'версия_ПО_СФ="%s"', [ftab, po_cf]);
    miDo('Commit table %s', [ftab]);

    PrepareCoords;

//    Info(' Удаление колонок....');
//    miDo('Alter Table %s (Drop x ,y)', [ftab]);//удалим колонок координат
    CreatePrimaryObjects;

    Info(ftab + ': Готова к работе...');
    mido('Browse * From %s', [ftab]);
    Result := true;
  except
    on E: Exception do
      Note(E.message);
  end;
end;
{------------------------------------------------------------------------------}
(*
Register Table "D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\от Двинина\КН_700000.txt"  TYPE ASCII Delimiter 9 Charset "WindowsCyrillic" Into "C:\Users\TROITS~1\AppData\Local\Temp\MIPRO\MapInfo.Application_12668\~MAP0001.TAB"
Open Table "C:\Users\TROITS~1\AppData\Local\Temp\MIPRO\MapInfo.Application_12668\~MAP0001.TAB" Hide
Commit Table ~MAP0001 As "D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\от Двинина\КН_700000.TAB"
Close Table ~MAP0001
Open Table "D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\от Двинина\КН_700000.TAB"
Browse * From КН_700000
Alter Table "КН_700000" ( add idn Integer ) Interactive
Select * from Дежурная_карта, КН_700000 where Дежурная_карта.КН = КН_700000._COL1 into Selection
Browse * From Selection
Commit Table Query1 As "D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\от Двинина\Link_КН_700000.TAB" TYPE NATIVE Charset "WindowsCyrillic" Interactive
*)
procedure TForm2.mnCreateTableFromTabClick(Sender: TObject);
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST'; {$ENDIF}//KN17.TAB
  if not CheckSprPath then exit;
  miDo('Close All');
  if not OpenTable('Source', 'Открыть таблицу источник') then//открывается таблица источник
    exit;
  CreateWorkTable;

  SafeClosetable('BaseTableCalc');
  if not st16 then//C:\Users\TROITS~1\AppData\Local\Temp\MIPRO\MapInfo.Application_13704\
    ExShellExecute('cmd', PChar(format('/c copy /b /Y "%s.*" "%s.*"',
      [ChangeFileExt(worktable, ''), TempPath + 'BaseTableCalc'])), SW_HIDE, false, 0);
end;
{------------------------------------------------------------------------------}
procedure TForm2.OpenDialog1Close(Sender: TObject);
begin
  OpenDialog1.Title := 'Открыть';
  Self.SetFocus;
end;
{-------открывается таблица и под нее создается каталог-------}
//если название таблицы Source то открывается таблица источник
//в противном случае таблица открывается из каталога ..Ценовые факторы_4\CurrentDat
function TForm2.OpenTable(source, title : String): Boolean;
begin
  Result := false;
  OpenDialog1.FileName := '*.tab';
  if title <> '' then
    OpenDialog1.Title := title;
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);

  curdata := CurrentData;
  ResFolder := GetAppFolder + AddSlesh(StringReplace(stBar.Panels[0].Text, 'CF',
    'Ценовые факторы', [rfIgnoreCase])) + curdata + '\';//table Source
  if source = '' then//это рабочая таблица
    ResFolder := ExtractFilePath(OpenDialog1.FileName);

  ForceDirectories(ResFolder);
  ADOConnection1.Close;
  ADOConnection1.ConnectionString := format(strConnect1, [ResFolder + 'коммуникации\Коммуникации_свод.accdb']);

  Result := true;
  SafeCloseTable('Source');
  worktable := OpenDialog1.FileName;
  ftab := miEval('PathToTableName$("%s")', [worktable]);

  SafeCloseTable(ftab);
  if source <> '' then
    miDo('Open table "%s" As Source', [OpenDialog1.FileName])
  else
    miDo('Open table "%s"', [OpenDialog1.FileName]);
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnOpenWorkTableClick(Sender: TObject);
var
  fields: TStringList;
  i, j, k, res: Integer;
  s: String;
  ind: Array of Integer;
begin
  OpenDialog1.InitialDir := GetAppFolder +
    AddSlesh(StringReplace(stBar.Panels[0].Text, 'CF', 'Ценовые факторы', [rfIgnoreCase]));
  if not OpenTable('', 'Открыть рабочую таблицу') then
    exit;

  worktable := OpenDialog1.FileName;
  ftab := miEval('PathToTableName$("%s")', [worktable]);

  // проверка наличия полей и возможно исправление
  SetLength(ind, High(AnyData));
  fields := getFieldsList(ftab);//список полей таблицы
  try
    s := '*';
    res := -1;
    for i := Low(AnyData) to High(AnyData) do
    begin
      ind[i] := 1; // отметим очередной фактор
      k := -1;//сравниваем список полей таблицы с данными Excel
      for j := 0 to fields.Count - 1 do
        if CText(AnyData[i].Alias, fields[j]) then
          k := j;
      if k = -1 then // нет фактора
      begin
        if res = -1 then // первое расхождение
          res := MessageBox(Handle, 'Перестроить таблицу?',
            'Несовпадение количества полей табицы и их описания в файле CosFactory.xlsx',
            MB_OK + MB_OKCANCEL + MB_ICONQUESTION);
        if res <> mrOk then
          exit;
        Info('Добавление полей...');
        miDo('Alter Table %s (Add %s %s)', [ftab, AnyData[i].Alias,
          AnyData[i].fType]);
      end
      else
        ind[i] := 0; // снимем отметку с очередного фактора
    end;

    if res <> -1 then // были изменения
    begin
      for i := Low(AnyData) to High(AnyData) do
        if (AnyData[i].Link = -1) and (ind[i] <> 0) then
          TimeCalc(ftab, AnyData[i].table, AnyData[i].Alias);
      miDo('Commit table %s', [ftab]);
      s := '';
      for j := Low(AnyData) to High(AnyData) do
        s := s + AnyData[j].Alias + ',';
      miDo('Alter Table %s (Order %s)', [ftab, Copy(s, 1, Length(s) - 1)]);
    end;
  finally
    fields.Free;
    SetLength(ind, 0);
  end;

  CreatePrimaryObjects;
  Info(ftab + ': Готова к работе...');
end;
{---------------------запросы в РострРеестр по КН------------------------------}
procedure TForm2.mnQueryRRClick(Sender: TObject);
var
  ExcelApp, Workbook, WorkSheet: OleVariant;
  UserAgentId, i, j : Integer;
  cn, tab : String;
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\ЗАГРУЗКИ'; {$ENDIF}
  OpenDialog1.FileName := '*.xlsx';
  if not OpenDialog1.Execute then
    exit;

  miDo('Open Table "%s"', [GetAppFolder + 'Base\Дежурная_карта.tab']);
  cn := ChangeFileExt(OpenDialog1.FileName, '.tab');
  tab := miEval('PathToTableName$("%s")', [cn]);
  SafeCloseTable(tab);
  miDo('Create Table "%s" (КН Char(30)) file "%s" '+
    'TYPE NATIVE Charset "WindowsCyrillic"', [tab, cn]);
  miDo('Create Map For %s %s',
    [tab, miEval('TableInfo("Дежурная_карта", %d)',[TAB_INFO_COORDSYS_CLAUSE])]);
//midO('Open table "%s"', [cn]);//tr!!!

  feat := 5;

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := false;
  ExcelApp.DisplayAlerts := false;

  try
    ExcelApp.Workbooks.Open(OpenDialog1.FileName);
    Workbook := ExcelApp.Workbooks.item[1];
    WorkSheet := Workbook.WorkSheets.item[1];
    pg.Max := WorkSheet.UsedRange.Rows.Count;

    //выбираем данные медленно (из ячеек) чтобы умышленно затянуть время между запросами (не давать Sllep)
    i := 1;
    Repeat
      if i mod 10 = 0 then pg.Position := i;
      Inc(UserAgentId);
      if UserAgentId >= High(UserAgent) then
        UserAgentId := 0;

      if String(WorkSheet.Cells[i, 1]) = '' then
        break;
      cn := String(WorkSheet.Cells[i, 1]);

      Info('Поиск ' + cn);
      if (cn <> '') and (Pos('66:', cn) > 0) then
        if QueryRosReestrV2(UserAgentId, cn) = 0 then
          miDo('Insert into %s (КН, obj) Values("%s", ob)', [tab, String(WorkSheet.Cells[i, 1])]);
      Inc(i);
      if i > 500 then
        miDo('Commit table %s', [tab]);
    Until false;
    pg.Max := 0;
    miDo('Commit table %s', [tab]);
    miDo('Map From %s', [tab]);

  finally
    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;
    WorkSheet := Unassigned;
    Workbook := Unassigned;
    ExcelApp := Unassigned;
  end;
  Info(' Готово');
end;

procedure TForm2.Access1Click(Sender: TObject);
var
  sl : TStringList;
  sa, sb : StringArray;
  s : String;
  i, j, k : Integer;
  F : TextFile;
begin
  sl := TStringList.Create;
  sl.LoadFromFile('d:\temp\Здания-ЗУ.txt');
  AssignFile(F, 'D:\temp\Здания-ЗУ новые.txt');
  Rewrite(F);
  Writeln(F, sl[0]);
  k := 1;
  for i := 1 to sl.Count-1 do
  begin
    sa := SplitStr(sl[i], #9);
    s := Copy(sa[2], 2, Length(sa[2]) - 2);
    sb := SplitStr(s, ';');

    if Length(sb) = 0 then
    begin
      Writeln(F, IntToStr(k) + #9 + sa[1] + #9 + '""');
      Inc(k);
    end else

    for j := 1 to Length(sb) do
    begin
      Writeln(F, IntToStr(k) + #9 + sa[1] + #9 + '"'+sb[j-1]+'"');
      Inc(k);
    end;

    SetLength(sb, 0);
    SetLength(sa, 0);
  end;
  CloseFile(F);
  Note('Ok');
end;
{----------Вычисление средневзвешенных значений для многоконтурных объектов средствыми MI-----}
procedure TForm2.AverageWeighted;
var
  field: String;
  i : Integer;
  sl : TStringList;
begin
  Info(' Вычисление средневзвешенных значений МКО...');
  sl := TStringList.Create;
  for i := Low(AnyData) to High(AnyData) do
    if (st16 or (AnyData[i].gko = 1)) and (AnyData[i].Link = -1) then
      sl.Add(AnyData[i].Alias);
  if sl.Count > 0 then
  begin
    field := 'idn,число_участков, Кадастровый_номер,';
    for i := 0  to sl.Count-1 do
      field := field + sl[i] + ',';
    field := field + 'Area';
    miDo('select %s from %s where Число_участков>1 into temptab100 NoSelect',
      [field, ftab]);
    for i := 0  to sl.Count-1 do
    begin
      miDo('Select idn, Sum(Area), WtAvg(%s, Area) From temptab100 Group '+
        'By Кадастровый_номер Into temptab%d noSelect', [sl[i], i+1]);
//    miDo('Update temptab Set col%d = col%d', [i, i + k]);
    end;
//    miDo('Commit table %s', [ftab]);

///
//  for i := 1 to k do//обновляем колонки
//    miDo('Update temptab Set col%d = col%d', [i, i + k]);
//  miDo('Commit table %s', [ftab]);
///
  end;
  sl.Free;
end;
{----------Вычисление средневзвешенных значений для многоконтурных объектов-----}
procedure TForm2.WeightedAverage;
var
  i, j, k, n: Integer;
  cf, sum: Double;
  minus: Boolean;
  field: String;
  sl: TStringList;
  sa: StringArray;
  Data: Array of Double_;
  t:Cardinal;

  function PrepareFields : Boolean;
  var
    i : Integer;
  begin
    Result := False;
    for i := Low(AnyData) to High(AnyData) do
    begin
      if (st16 or (AnyData[i].gko = 1)) and (AnyData[i].Link = -1) then
      begin
        field := field + AnyData[i].Alias + ',';
        Inc(k);
        Result := True;
      end;
    end;
  end;
{select * from BaseTableCalc where число_участков>1 into seltable
Select idn, Sum(Area), WtAvg(TimeMO, Area) From seltable Group By Кадастровый_номер  Into  Lit_query   TODO}
begin
  t := GetTickCount;

  Info(' Вычисление средневзвешенных значений МКО...');
  field := 'idn,число_участков,';
  k := 2;//idn,число_участков
  if not PrepareFields then Exit;

  field := field + 'Area'; // площадь последняя переменная
  miDo('select %s from %s where Число_участков>1 into temptab NoSelect',
    [field, ftab]);
  miDo('Export temptab Into "%sAnaliz.MIF" Type "MIF" Overwrite', [ResFolder]);
  // создаем MIF MID для многоконтурных объектов по выбранным ЦФ
  cn.LoadFromFile(ResFolder + 'Analiz.MID');
  sa := SplitStr(field, ',');
  sl := TStringList.Create;
  try
    sl.Add('  Version   300');
    sl.Add('  Charset "WindowsCyrillic"');
    sl.Add('  Delimiter ","');
    sl.Add(format('  Columns %d', [k + 1]));
    for i := 0 to High(sa) do
      sl.Add(format('    %s float', [sa[i]]));
    sl.Add('  Data');
    for i := 0 to cn.Count - 1 do
      sl.Add('none');
    sl.SaveToFile(ResFolder + 'Analiz.MIF');
  finally
    SetLength(sa, 0);
    miDo('Close table temptab');
    sl.Free;
  end;

  cn.LoadFromFile(ResFolder + 'Analiz.MID');
  SetLength(Data, cn.Count);
  try
    for i := 0 to cn.Count - 1 do
      SplitDbl(',', cn[i], Data[i]);
    i := 0;
    while i < cn.Count do
    begin
      k := Round(Data[i][1]);//количество частей для текущего многоконтурника
      sum := 0;
      for j := 0 to k - 1 do // идем по строкам
        sum := sum + Data[i + j, High(Data[i])];//накапливаем сумму площадей
      for n := 2 to High(Data[i]) - 1 do
      // идем по колонкам пропустив IDN и кол-во частей
      begin
        cf := 0;
        minus := false;
        for j := 0 to k - 1 do//идем по строкам
        begin
          cf := cf + Data[i + j, High(Data[i])] / sum * Data[i + j, n];
          // ЦФ = Сумма (Si/S*ЦФi)
          if Data[i + j, n] < 0 then // один из факторов = -1
            minus := true;
        end;
        if minus then
          cf := -1;
        Data[i + k - 1, n] := cf; // заменяем цф в последней строке
      end;
      Data[i + k - 1, High(Data[i])] := sum;//заменяем сумму в последней строке

      for n := 2 to High(Data[i]) do // обнуляем вместе с площадями
      begin
        for j := 0 to k - 2 do
          Data[i + j, n] := 0;
      end;
      Inc(i, k);
    end;

    for i := 0 to cn.Count - 1 do
    begin
      field := format('%d,%d', [Round(Data[i, 0]), Round(Data[i, 1])]);
      for n := 2 to High(Data[i]) do
        field := format('%s,%.2f', [field, Data[i, n]]);
      cn[i] := field;
    end;

  finally
    SetLength(Data, 0);
  end;

  cn.SaveToFile(ResFolder + 'Analiz.MID');
  // импортируем средневзвешенные значения
  SafeCloseTAble('Analiz');
  miDo('Import "%sAnaliz.mif" Type "MIF" Into "%sAnaliz.TAB" Overwrite',
    [ResFolder, ResFolder]);

  field := 'Select ';
  k := 1;// выбираем время из основной таблицы учитывая площадь
  for i := Low(AnyData) to High(AnyData) do
  begin
    if (st16 or (AnyData[i].gko = 1)) and (AnyData[i].Link = -1) then
    begin
      field := field + format('%s.%s,', [ftab, AnyData[i].Alias]);
      Inc(k);
    end;
  end;
  field := field + format('%s.Area,', [ftab]); // добавим площади

  //выбираем расстояние и время из средневзвешенных значений
  for i := Low(AnyData) to High(AnyData) do
  begin
    if (st16 or (AnyData[i].gko = 1)) and (AnyData[i].Link = -1) then
      field := field + format('Analiz.%s,', [AnyData[i].Alias]);
  end;
  field := field + 'Analiz.Area';
  field := field +
    format(' from %s,Analiz where %s.idn=Analiz.idn into temptab NoSelect',
    [ftab, ftab]);
  miDo(field);

  for i := 1 to k do//обновляем колонки
    miDo('Update temptab Set col%d = col%d', [i, i + k]);
  miDo('Commit table %s', [ftab]);

  miDo('Drop table Analiz');
  ExShellExecute('cmd.exe', format('/c del /F /Q  "%s"',
    [ResFolder + 'Analiz.*']), SW_HIDE, false, 0);//удаление MIF/MID

//  caption := format('Время1 %d  ',  [Round((GetTickCount - t) / 60000)]);
//  t := GetTickCount;
//AverageWeighted;
//  caption := caption + format('Время2 %d',  [Round((GetTickCount - t) / 60000)]);
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnApplaCosFactoryClick(Sender: TObject);
begin
  LoadAnyData(nil);
end;
{------------------------------------------------------------------------------}
function TForm2.CalcCPUCount: Integer;
var
  si: TSystemInfo;
begin
  GetSystemInfo(si);
  Result := si.dwNumberOfProcessors;
  if Result > 4 then
    Result := 4;
end;
{------==-8 частей 479, 4 части 454/47,1 3 части 2 части 784------------------}
procedure TForm2.CalcMultiContour;
var
  i, j : Integer;

  procedure SlpitBaseTable;
  var
    i, n, k, m, jj: Integer;

    procedure SeekRow;//ищем начало объектов с текущим кадастровым номером
    begin
      miDo('Fetch Rec %d From TempTab', [k]);//ищем во всех многоконтурниках
      cad := miEval('TempTab.Кадастровый_номер');
      Repeat
        miDo('Fetch Prev From TempTab');
        prevcad := miEval('TempTab.Кадастровый_номер');
        if prevcad = cad then
          Dec(k);//смещаемся выше по записям
      Until prevcad <> cad;
    end;

  begin
    for jj := 1 to j do
      MoveCostFactory(jj, AddSlesh(stBar.Panels[0].Text));
    m := multi div j;
    k := m;
    n := 1;
    for jj := 1 to j do
    begin
      SeekRow;
      Info(format(' копирование объектов часть %d', [jj]));
      miDo('Select * from TempTab where RowID >=%d AND RowID < %d into Temp1 noSelect',
        [n, k]);
      miDo('Commit table Temp1 As "%s%d\Intersect.tab" TYPE NATIVE Charset "WindowsCyrillic"',
        [TempPath, jj]);
      miDo('Close table Temp1');
      ExCreateProcess
        (format('"%sCalcIntersect.exe" "%s%d\Intersect.tab" "%sCostFactory.xlsx" "%s" %s %d %d',
        [GetAppFolder, TempPath, jj, GetAppFolder + AddSlesh(stBar.Panels[0]
        .Text), ResFolder, Procent.Text, Handle, Integer(st16)]), '', SW_MINIMIZE, false);
      Inc(cpu);
      n := k;
      k := k + m;
      if jj = j-1 then
        k := k + k;//гарантированно увеличиваем конец последней выборки
    end;
  end;

begin//CalcMultiContour
  for j := 1 to 10 do//гарантированно удалим все предыдущие расчеты
    if FileExists(ResFolder + format('multi_contour%d.tab', [j])) then
    begin
      miDo('Open table "%smulti_contour%d.tab"', [ResFolder, j]);
      miDo('Drop table multi_contour%d', [j]);
    end;

  cpu := 0;
  multi := -1;//проверка необходимости расчета для многоконтурных объектов
  for i := Low(AnyData) to High(AnyData) do
    if (AnyData[i].Link = 1) and (st16 or (AnyData[i].gko = 1)) then
      Inc(multi);
  if multi = -1 then Exit;

  Info(' выбор объектов...');
  miDo('Select * from %s where Число_участков > 1 into TempTab noSelect', [ftab]);
  multi := miEvalInt('TableInfo(TempTab, %d)', [TAB_INFO_NROWS]);//размер таблицы многоконтурников
  if multi = 0 then
    exit;

  j := CalcCPUCount;
  Info(' копирование справочников...');
  if multi > 1000 then // делим на части
    SlpitBaseTable // делим на части
  else
  begin // расчитываем как 1 часть
    MoveCostFactory(2, AddSlesh(stBar.Panels[0].Text));
    Info(' копирование объектов...');
    miDo('Commit table TempTab As "%s2\Intersect.tab" TYPE NATIVE Charset "WindowsCyrillic"',
      [TempPath]);
    miDo('Close table TempTab');
    ExCreateProcess
      (format('"%sCalcIntersect.exe" "%s2\Intersect.tab" "%sCostFactory.xlsx" "%s" %s %d %d',
      [GetAppFolder, TempPath, GetAppFolder + AddSlesh(stBar.Panels[0].Text),
      ResFolder, Procent.Text, Handle, Integer(st16)]), '', SW_MINIMIZE, false);
    cpu := 1;
  end;
  Winapi.Windows.SetFocus(Handle);
end;
{------------------------------------------------------------------------------}
procedure SetHeader(var h: TNetHeaders; const AHeader, AValue: String);
var
  i, Max: Integer;
begin
  i := 0;
  Max := Length(h);
  while i < Max do
  begin
    if CText(h[i].name, AHeader) then
      break;
    Inc(i);
  end;
  if i = Max then
  begin
    SetLength(h, Max + 1);
    h[i].name := AHeader;
  end;
  h[i].value := AValue;
end;
{------------------------------------------------------------------------------}
procedure TForm2.ServerCertValidate(const Sender: TObject;
  const ARequest: TURLRequest; const Certificate: TCertificate;
  var Accepted: Boolean);
begin
  Accepted := true;
end;
{------------------------------------------------------------------------------}
procedure TForm2.SpedTest1Click(Sender: TObject);
var
  i, j, k : Integer;
  tab, s : String;
  sl : TStringList;
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\CF_5'; {$ENDIF}
  OpenDialog1.FileName := '*.tab';
  if not OpenDialog1.Execute then
    exit;

  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  sl := TStringList.Create;
  miDo('Close All');
  for i := 0 to OpenDialog1.Files.Count-1 do
  begin
    miDo('Open table "%s"', [OpenDialog1.Files[i]]);
    tab := miEval('PathToTableName$("%s")', [OpenDialog1.Files[i]]);
    s := ChangeFileExt(OpenDialog1.Files[i], '.mif');
    miDo('Export "%s" Into "%s" Type "MIF" Overwrite', [tab, s]);
    sl.LoadFromFile(s);
    k := 0;
    for j := 0 to 10 do
      if Pos('1500000,', sl[j]) > 0 then
      begin
        if Pos('1500000,', sl[j]) > 0 then
          sl[j] := cs11
        else
          sl[j] := cs21;
        k := 1;
        Break;
      end;
    if k = 1 then
    begin
      sl.SaveToFile(s);
      miDo('Close All');
      miDo('Import "%s" Type "MIF" Into "%s" Overwrite', [s, OpenDialog1.Files[i]]);
    end;
    DeleteFile(s);
    s := ChangeFileExt(OpenDialog1.Files[i], '.mid');
    DeleteFile(s);
    miDo('Close All');
  end;
  sl.Free;
  Note('OK');
end;
{------------------------------------------------------------------------------}
function TForm2.QueryRosReestrV2(var id: Integer; var cadastr: String): Integer;
var
  ss : TStringStream;
  client : THTTPClient;
  h : TNetHeaders;
  resp : IHTTPResponse;
  code, s, sxy, command : String;
  i, j, k, n : Integer;
  x, y : Double;
  mc, xy : TMatchCollection;

begin
  Result := 2;
  ss := TStringStream.Create('', TEncoding.UTF8);
  ss.Clear;
  client := THTTPClient.Create;
  try
    client.SecureProtocols := [THTTPSecureProtocol.TLS12];
    client.OnValidateServerCertificate := ServerCertValidate;
    client.UserAgent := UserAgent[id];
    SetHeader(h, 'authority', 'nspd.gov.ru/map');
    SetHeader(h, 'accept', 'text/html,application/xhtml+xml,application/xml;' +
      'q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9');
    SetHeader(h, 'DNT', '1');
    SetHeader(h, 'Connection', 'Keep-Alive');
    SetHeader(h, 'Cache-Control', 'no-cache');
//    SetHeader(h, 'Referer', 'https://nspd.gov.ru/map /'); // !!!!! ОБЯЗАТЕЛЬНО!
    SetHeader(h, 'Referer', 'https://nspd.gov.ru/map/'); // !!!!! ОБЯЗАТЕЛЬНО!
    // Участки 1, ОКС 5, Кварталы 2, Районы 3, Округа 4, Границы 7, ЗОУИТ 10, Тер. зоны 6, Красные линии 13, Лес 12, СРЗУ 15, ОЭЗ 16, ГОК 9
    // старые Коды ответа HTTP : https://developer.mozilla.org/ru/docs/Web/HTTP/Status
    // новые Коды ответа HTTP : https://api-cloud.ru/rosreestr
    try
//запрос через браузер: https://nspd.gov.ru/api/geoportal/v2/search/geoportal?query=66:41:0303161:193&1
      resp := client.Get(format('https://nspd.gov.ru/api/geoportal/v2/search/geoportal?query=%s&$s verify=False',
        [System.NetEncoding.TNetEncoding.URL.Encode(cadastr), feat]), ss, h);
    except
      on E: Exception do
      begin
        code := E.message;
        exit;
      end;
    end;
    code := ss.DataString;
  finally
    if (resp <> nil) and (resp.StatusCode = 200) then
    begin // проверка координат найденного объекта
      mc := reg.Matches(StringReplace(code, ']],', ']],'+#10, [rfReplaceAll]), '\[\[\d.*\d\]\]');
      if mc.Count > 0 then
      begin
        command := format('Create Region Into Variable ob %d', [mc.Count]);
        for i  := 0 to mc.Count-1 do
        begin
          s := StringReplace(mc.item[i].value, '[', '(', [rfReplaceAll]);
          s := StringReplace(s, ']', ')', [rfReplaceAll]);
          s := Copy(s, 2, Length(s)-2);
          s := StringReplace(s, '),(', ')(', [rfReplaceAll]);
          sxy := StringReplace(s, '(', '', [rfReplaceAll]);
          xy := reg.Matches(StringReplace(sxy, ')', #10, [rfReplaceAll]), '\d.*\d');
          command := command + format(' %d ', [xy.Count]);
          for j := 0 to xy.Count-1 do
          begin
            s := xy.item[j].value;
            k := Pos(',', s);
            Val(Copy(s, 1, k-1), y, n);//координаты в проекции Популярная визуализация
            Val(Copy(s, k+1, Length(s)), x, n);
            EPSG_4326_meter2degress(x, y);//координаты в проекции WGS 84
            Wgs2Msk_(x, y, 2);//координаты в проекции МСК-66 Свердловская область
            command := command + format('(%.9f,%.9f)', [y, x]);
          end;
        end;
        miDo('Set CoordSys table Дежурная_карта');
        miDo(command);
        Result := 0;
      end;// else так дает центроиды кварталов
//      begin
//        i := Pos('"geometry":{"type":"Point","coordinates":[', code);
//        j := PosEx(']', code, i+42);
//        if (i > 0) and (j > i) then
//        begin
//          s := Copy(code, i + 42, j-i-42);
//          k := Pos(',', s);
//          Val(Copy(s, 1, k-1), y, n);//координаты в проекции Популярная визуализация
//          Val(Copy(s, k+1, Length(s)), x, n);
//          EPSG_4326_meter2degress(x, y);//координаты в проекции WGS 84
//          Wgs2Msk_(x, y, 2);//координаты в проекции МСК-66 Свердловская область
//          miDo('Set CoordSys table Дежурная_карта');
//          miDo('Create Point Into Variable ob	(%.9f, %.9f)', [y, x]);
//          Result := 0;
//        end;
//      end;
    end;
    client.Free;
    ss.Free;
    SetLength(h, 0);
  end;
end;
{------------------------------------------------------------------------------}
function TForm2.CF_from_CN : Boolean;
var
  i, k, n : Cardinal;
  x0, y0 : Double;
  UserAgentId, maxid, time_ : Integer;
  tab, cn, cn2, x, y, dp, dir : String;
  mc : TMatchCollection;

  function DataProverki: String;
  var
    y, m, d : Word;
    dt : TDateTime;
  begin
    dt := Now;
    DecodeDate(dt, y, m, d);
    Result := format('%d.%.2d.%.2d', [d, m, y]);
  end;

  procedure CheckCol2;
  begin
    if miEvalInt('TableInfo(tempTable, %d)', [TAB_INFO_NCOLS]) = 3 then
    begin
      cn2 := miEval('tempTable._COL2');//КН ЗУ
      mc := reg.Matches(cn2, '\d\d:\d\d:\d{1,7}:\d{1,}', [roIgnoreCase, roMultiLine]);
      if mc.count > 0 then
      begin//ищем КН из второй колонки в дежурной карте
        miDo('Select * from Дежурная_карта where КН="%s" into Temp noSelect', [cn2]);
        if miEvalInt('TableInfo(Temp, %d)', [TAB_INFO_NROWS]) > 0 then
        begin
          miDo('ob = Centroid(Temp.obj)');
          miDo('Insert into Map_%s (КН, obj) Values("%s", ob)', [curdata, cn]);//КН здания
          miDo('Update tempTable Set id=2 Where RowId=%d', [i]);//отметим запись как имеющую объект
        end;
      end;
    end;
  end;

begin//CF_from_CN
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST'; {$ENDIF}//ЗУ для зданий-313.txt
  Result := false;

  curdata := CurrentData;
  ResFolder := GetAppFolder + AddSlesh(StringReplace(stBar.Panels[0].Text, 'CF',
    'Ценовые факторы', [rfIgnoreCase])) + curdata + '\';//сам назначаю путь
  ForceDirectories(ResFolder);
  time_ := StrToInt(TimeOut.Text);
  dp := DataProverki;//Дата_проверки
  miDo('Close All');
  miDo('Open Table "%s"', [GetAppFolder + 'Base\Дежурная_карта.tab']);
  miDo('Create Index On Дежурная_карта (КН)');

  Info(' Поиск максимального номера...');
  miDo('Select Max(idn) from Дежурная_карта into Temp noSelect');
  maxid := miEvalInt('Temp.Col1');

  worktable := ChangeFileExt(OpenDialog1.FileName, '.tab');;
  ftab := miEval('PathToTableName$("%s")', [worktable]);

  SafeCloseTable(ftab);
  //создаем таблицу из кадастровых номеров
  Info(' Регистрация таблицы...');
  tab := CreateTableFromText(OpenDialog1.FileName, 'CoordSys table Дежурная_карта');//короткое имя
  miDo('Alter Table "%s" (Add id Integer)', [tab]);//добавляем ID для отметки номеров

  SafeCloseTable(format('Map_%s', [curdata]));
  Info(' Выбока по кадастровым номерам...');//ищем все КН в дежурной карте
  miDo('Select Дежурная_карта.IDN, Дежурная_карта.КН, Дата_проверки from '+
    'Дежурная_карта, %s where Дежурная_карта.КН = %s._COL1 into Map_%s NoSelect',
    [tab, tab, curdata]);
  miDo('Commit Table Map_%s As "%sMap_%s.TAB" TYPE NATIVE Charset "WindowsCyrillic"',
    [curdata, TempPath, curdata]);//сохраняем выборку с объектами ИЗ дежурной карты
  miDo('Close table Map_%s', [curdata]);//переоткрываем таблицу выборку
  miDo('Open Table "%sMap_%s"', [TempPath, curdata]);//таблица может оказаться пустой
  miDo('Create Index On Map_%s (КН)', [curdata]);
  Info(' Выбока отсутсвующих номеров...');//
  //связываем исходную таблицу с данными выборки из дежурной карты
  miDo('Select * from %s, Map_%s where %s._COL1 = Map_%s.КН into tempTable noSelect',
    [tab, curdata, tab, curdata]);
  miDo('Update tempTable Set id=1');//отмечаем в curdata имеющиеся в дежурке объекты
  miDo('select * from %s where id=0 into tempTable noSelect', [tab]);//выбираем в curdata отсутствующие объекты
  //если не все КН найдены в дежурке
  if cbFindCN.Checked then
  begin
    Case ServiceMapForm.cbCategories.ItemIndex of
      0 : feat := 1;
      1 : feat := 5;
    end;
    UserAgentId := -1;

    pg.Max := miEvalInt('TableInfo(tempTable, %d)', [TAB_INFO_NROWS]);
    for i := 1 to pg.Max do
    begin
      if i mod 10 = 0 then pg.Position := i;
      miDo('Fetch rec %d from tempTable', [i]);
      cn := miEval('tempTable._COL1');
      Info(' Поиск КН ' + cn);
      Repeat
        Inc(UserAgentId);
        if UserAgentId >= High(UserAgent) then
          UserAgentId := 0;
        k := 2;
        n := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NCOLS]);
        if ((n = 3) and (miEval('%s._COL2', [tab]) = '')) or (n = 2) then
        begin//в нормальном режиме надо проверять КН в Росреестре
          k := QueryRosReestrV2(UserAgentId, cn);
          Sleep(time_);//после запроса отдыхаем
        end;
        case k of
          0:begin//объект вставляем в карту
              miDo('Insert into Map_%s (КН, obj) Values("%s", ob)', [curdata, cn]);
              miDo('Update tempTable Set id=3 Where RowId=%d', [i]);//отметим запись как имеющую объект
            end;
          2:begin
              k := 0;
              CheckCol2;//поиск по земельным участкам
            end;
        end;//k=1 повторить
      Until k = 0;
    end;
  end else

  if miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NCOLS]) = 3 then
  begin
    pg.Max := miEvalInt('TableInfo(tempTable, %d)', [TAB_INFO_NROWS]);
    for i := 1 to pg.Max do
    begin
      if i mod 10 = 0 then pg.Position := i;
      miDo('Fetch rec %d from tempTable', [i]);
      cn := miEval('tempTable._COL1');
      CheckCol2;//поиск по земельным участкам
    end;
  end;
  pg.Max := 0;

  //выбираем в CurrentData объекты по центроидам и вставлеем в дежурную карту
  SafeCloseTable('AllTables');
  miDo('Select Map_%s.IDN, Map_%s.КН from Map_%s, '+
    '%s where Map_%s.КН=%s._COL1 and %s.id=2 into AllTables noSelect',
    [curdata, curdata, curdata, tab, curdata, tab, tab]);
  for i := 1 to miEvalInt('TableInfo(AllTables, %d)', [TAB_INFO_NROWS]) do
  begin
    miDo('Fetch rec %d From AllTables', [i]);
    Inc(maxid);
    cn := miEval('AllTables.КН');
    miDo('ob = AllTables.obj');
    miDo('Insert Into дежурная_карта (Obj,IDN,КН,Дата_пакета,Дата_проверки) '+
      'Values(ob,%d,"%s","01.01.2020","%s")', [maxid, cn, dp]);
  end;
  miDo('Commit Table дежурная_карта');

  //выбираем в CurrentData объекты из Росреестра ( id = 3)
  SafeCloseTable('AllTables');
  miDo('Select Map_%s.IDN, Map_%s.КН from Map_%s, '+
    '%s where Map_%s.КН=%s._COL1 and %s.id=3 into AllTables noSelect',
    [curdata, curdata, curdata, tab, curdata, tab, tab]);
  if miEvalInt('TableInfo(AllTables, %d)', [TAB_INFO_NROWS]) > 0 then
  begin
    dir := GetAppFolder + 'Base\БракЦФ\';
    ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [dir])), SW_HIDE, true, 0);//удаление каталога со всем содержимым и подкаталогами
    ForceDirectories(dir);
    miDo('Commit table AllTables As "%sAllTables.tab"', [dir]);
    SafeCloseTable('AllTables');
    miDo('Open table "%sAllTables.tab"', [dir]);
    miDo('Alter table AllTables(Add Дата_пакета Date,'+
      'Дата_проверки Date,Черезполосники Char(10),Самопересечение Char(10),'+
      'За_Границей Char(10),В_работу Char(10), Примечание Char(250))');
    miDo('Update AllTables Set дата_пакета = "01.01.2020"');
    miDo('Commit table AllTables');
    miDo('Browse * From AllTables');
    Exc('Найденная в Росреестре графика требует проверки.');
  end;

  //выбираем КН без объектов
  SafeCloseTable('TableLessObjects');
  miDo('select * from tempTable where id=0 into TableLessObjects noSelect');
  miDo('Commit Table TableLessObjects As "%sTableLessObjects.TAB" TYPE NATIVE'+
    ' Charset "WindowsCyrillic" Interactive', [ResFolder]);
  miDo('Close table %s', [tab]);

  miDo('Commit Table Map_%s As "%s" TYPE NATIVE Charset "WindowsCyrillic"',
    [curdata, worktable]);
  miDo('Close All');
  miDo('Open table "%s" As Source', [worktable]);

  if miEvalInt('TableInfo(Source, %d)', [TAB_INFO_NROWS]) = 0 then
    Exc('Нет записей в таблице Source.');

  if cbExamenOGR.Checked then
    if not ExamenOGR (stBar, ExtractFilePath(worktable) + 'Source', TempPath) then
      Exc('Найдена плохая графика.');


  if CreateWorkTable then
  begin
    CalcMultiContour;//intersect проверить на очистку памяти при сбое!!!!

    mnTimesClick(nil);
    mnZonesClick(nil);

    WeightedAverage; //вычисление средневзвешенных значений расстояние по прямой или время
    mnLinkClick(nil); // Ожидание многоконтурных объектов
    mnCode_MOClick(nil); // определение кодов МО
    mnItogCFClick(nil); // создание итоговой таблицы

    mnGraphicClick(nil);
    mnBaseClick(nil);
    mnCommunicationItogClick(nil);

    mnZOUITClick(nil);
    mnSoilClick(nil);

    CreateLoaderFile;
    Result := true;
  end;
end;
{------------------------------------------------------------------------------}
function TForm2.CheckWorkTable(procname : String) : Boolean;
begin
  Result := False;
  if worktable = '' then
    Exc(procname + ' : Рабочая таблица не открыта.');

  miDo('Open table "%s"', [worktable]);
  ftab := miEval('PathToTableName$("%s")', [worktable]);
  fcn := Find_CN_IDN(ftab);

  Result := True;
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnCFClick(Sender: TObject);
var
  t : Cardinal;
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\Ценовые факторы_4'; {$ENDIF}//2025_09_09_
  if not CheckWorkTable('mnCFClick') then Exit;

  t := GetTickCount;
  CreatePrimaryObjects;

  CalcMultiContour;
  mnTimesClick(nil);
  mnZonesClick(nil);

  WeightedAverage;//вычисление средневзвешенных значений
  mnLinkClick(nil); // Ожидание многоконтурных объектов
  mnCode_MOClick(nil);

  mnItogCFClick(nil);

  Note(format('процес завершен.%sВремя обработки %d минут',
    [#10, Round((GetTickCount - t) / 60000)]));
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnTimesClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to High(AnyData) do
    if (AnyData[i].Link = -1) and (st16 or (AnyData[i].gko = 1)) then
      TimeCalc(ftab, AnyData[i].table, AnyData[i].Alias);
  miDo('Commit table %s', [ftab]);
  Info(ftab + ': Готова к работе...');
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnZonesClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to High(AnyData) do
    if (AnyData[i].Link = 1) and (st16 or (AnyData[i].gko = 1)) then
      CalcZone(i);
  miDo('Commit table %s', [ftab]);
  Info(ftab + ': Готова к работе...');
end;
{------------------------------------------------------------------------------}
function SortInfo1(Item1, Item2: PzouitObj): Integer;
begin
  if Item1.cn < Item2.cn then
    Result := -1
  else if Item1.cn > Item2.cn then
    Result := 1
  else
    if Item1.idn < Item2.idn then
      Result := -1
    else if Item1.idn > Item2.idn then
      Result := 1
    else
      Result := 0;
end;
{------------------------------------------------------------------------------}
function SortInfo2(Item1, Item2: PzouitObj): Integer;
begin
  if Item1.cn < Item2.cn then
    Result := -1
  else if Item1.cn > Item2.cn then
    Result := 1
  else
  begin
    if Item1.zone < Item2.zone then
      Result := -1
    else if Item1.zone > Item2.zone then
      Result := 1
    else
    if Item1.subzone < Item2.subzone then
      Result := -1
    else if Item1.subzone > Item2.subzone then
      Result := 1
    else
      Result := 0;
  end;
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnLinkClick(Sender: TObject);
var
  i, j, k, n: Integer;
  s, m, path: String;
begin
//  miDo('Select * from %s where Число_участков > 1 into TempTab noSelect', [ftab]);
//  multi := miEvalInt('TableInfo(TempTab, %d noSelect)', [TAB_INFO_NROWS]);//размер таблицы многоконтурников
//  miDo('Close table TempTab');
  if multi <= 0 then
    exit;

  path := ExtractFilePath(worktable);
  Info(' Ожидание расчета многоконтурных объектов..');
  While cpu > 0 do
  begin
    Sleep(1000);
    Application.ProcessMessages;
  end;

  j := CalcCPUCount;
  pg.Max := j;
  for n := 1 to j do
  begin
    m := 'multi_contour' + IntToStr(n);
    if FileExists(path + m + '.tab') then
    begin
      Info(format(' Добавляется таблица multi_contour%d...', [n]));
      miDo('Open table "%s"', [path + m + '.tab']);
      s := '';
      k := 0;
      for i := 0 to High(AnyData) do
        if AnyData[i].Link = 1 then
        begin
          s := s + format('%s.%s,', [ftab, AnyData[i].Alias]);
          Inc(k);
        end;
      for i := 0 to High(AnyData) do
        if AnyData[i].Link = 1 then
          s := s + format('%s.%s,', [m, AnyData[i].Alias]);
      SetLength(s, Length(s) - 1);
      s := format('Select %s from %s, %s where %s.IDN=%s.IDN into Temp noSelect',
        [s, ftab, m, ftab, m]);
      miDo(s);
      if miEvalInt('TableInfo(Temp, %d)', [TAB_INFO_NROWS]) > 0 then
      begin
        s := 'Update Temp Set ';
        for i := 1 to k do
          s := s + format('col%d=col%d,', [i, i + k]);
        SetLength(s, Length(s) - 1);
        miDo(s);
      end;
      miDo('Close table Temp');
    end;
  end;

  miDo('Commit table %s', [ftab]);
  Info(ftab + ': Готова к работе..');

  for j := 1 to 10 do//гарантированно удалим все предыдущие расчеты
    if FileExists(ResFolder + format('multi_contour%d.tab', [j])) then
    begin
      miDo('Open table "%smulti_contour%d.tab"', [ResFolder, j]);
      miDo('Drop table multi_contour%d', [j]);
    end;
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnMi2ExcelClick(Sender: TObject);
begin
  OpenDialog1.FileName := '*.tab';
  if not OpenDialog1.Execute then
    exit;

  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  miDo('Open table "%s"', [OpenDialog1.FileName]);
  Tab2ExcelQuicly(miEval('PathToTableName$("%s")', [OpenDialog1.FileName]),
    ChangeFileExt(OpenDialog1.FileName, '.xls'), nil, True);
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnMultiContourClick(Sender: TObject);
begin
  CalcMultiContour;
  Info(ftab + ': Готова к работе..');
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnCode_MOClick(Sender: TObject);
var
  i: Integer;
begin
  Info(' Обновление кода муниципального образования..');
  for i := Low(AnyData) to High(AnyData) do
  begin
    if (AnyData[i].Link = 1) and CText(AnyData[i].field, 'КодНП_МО') and
      CText(AnyData[i + 1].field, 'KodMO') then
    if (st16 or (AnyData[i].gko = 1)) then
    begin// заполняем Код_Муниципального_образования
      miDo('Select %s, %s from %s where %s <> "" into SelTemp noSelect',
        [AnyData[i].Alias, AnyData[i + 1].Alias, ftab, AnyData[i].Alias]);
      miDo('Update SelTemp Set col2=Left$(col1, 2)');
      miDo('close table SelTemp');
      break;
    end;
  end;
  miDo('Commit table %s', [ftab]);
  Info(ftab + ': Готова к работе..');
end;
{-----------------------определение наличия размеров буферных зон-------------}
procedure TForm2.Elipce2RectClick(Sender: TObject);
var
  tab : String;
  j, row : Integer;
begin
  OpenDialog1.FileName := '*.tab';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);

  miDo('Dim pn As Pen');
  miDo('Dim br As Brush');
  tab := miEval('PathToTableName$("%s")', [OpenDialog1.FileName]);
  miDo('Open Table "%s"', [OpenDialog1.FileName]);
  miDo('set coordsys table %s', [tab]);
  row := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
  miDo('Fetch First From %s', [tab]);
  miDo('ob = %s.obj', [tab]);
  if (miEvalInt('ObjectInfo(%s.obj, %d)', [tab, OBJ_INFO_TYPE]) = 2) then//OBJ_TYPE_ELLIPSE
  begin//заменяем элипсы квадратами
    pg.Max := row;
    miDo('pn = ObjectInfo(ob, 2)');
    miDo('br = ObjectInfo(ob, 3)');
    miDo('Set Style Pen pn');
    miDo('Set Style Brush br');
    for j := 1 to row do
    begin
      if j mod 100 = 0 then pg.Position := j;
      miDo('Fetch Rec %d From %s', [j, tab]);
      miDo('ob = %s.obj', [tab]);//превращаем в прямоугольники
      miDo('Create Rect into variable ob (CentroidX(ob)-2, CentroidY(ob)-2)(CentroidX(ob)+2, CentroidY(ob)+2)');
      miDo('update %s set obj=ob where rowid=%d', [tab, j]);
    end;
    miDo('Fetch First From %s', [tab]);
  end else
    Note('Первый объект не элипс!');
  miDo('UnDim pn');
  miDo('UnDim br');
  pg.Max := 0;
  note('Ok');
end;

function TForm2.ExamenBufferSize : Boolean;
var
  fold : TStringList;
  sprpath : String;
  k : Integer;
begin
  Result := False;
  fold := TStringList.Create;
  sprpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'коммуникации\';
  try
    FindFolders(sprpath, fold, false);
    if fold.Count = 0 then
      Exc('Нет подкаталогов в каталоге ' + sprpath);
    for k := 0 to fold.Count - 1 do // идем по частям коммуникаций
      if not FileExists(AddSlesh(fold[k]) + 'Buffer.txt') then
        Exc(Format('Нет файла "Buffer.txt" в каталоге %s', [fold[k]]));
    Result := True;
  finally
    fold.Free;
  end;
end;
//C:\Users\TroitskikhVV\AppData\Local\Temp\MIPRO
{---- объединение частей линейных и точечных объектов по названию каталога---- }
procedure TForm2.CalcMerge;
var
  i, j, k, row: Integer;
  tab, compath, respath, prepared, path, buf: String;
  fold, sl : TStringList;
  open_prepared: Boolean; // признак открытия готовой таблицы
begin
  path := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'коммуникации\';
  compath := ResFolder + 'коммуникации\';
  respath := ResFolder + 'коммуникации\RESULT\';
  prepared := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'PREEPARED\';

  ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [compath])), SW_HIDE, true, 0);
  fold := TStringList.Create;
  sl := TStringList.Create;
  try
    FindFolders(GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'коммуникации\', fold, false);
    // в справочнике д.б. только части коммуникаций
    if fold.Count = 0 then
      Exc('CalcMerge : Нет подкаталогов в каталоге ' + path);

    ForceDirectories(compath);
    ForceDirectories(respath);
    ForceDirectories(prepared);
    SafeCloseTAble('All_CN');

    miDo('Select %s from %s order by %s into temp noSelect', [fcn, ftab, fcn]);
    miDo('Commit Table temp As "%sAll_CN.TAB" TYPE NATIVE Charset "WindowsCyrillic"',
      [respath]);//выбираем из базовой таблицв (examole KN17) все КН
    miDo('Close Table temp');
    miDo('Open Table "%sAll_CN.TAB" Interactive', [respath]);

    miDo('set coordsys table All_CN');
    for k := 0 to fold.Count - 1 do // идем по частям коммуникаций
    begin
      sl.LoadFromFile(AddSlesh(fold[k]) + 'Buffer.txt');
      buf := sl[0];
      sl.Clear;// ищем источники,потребители...
      FindRecursive(AddSlesh(fold[k]), '*.tab', false, sl);
      // в каждой части объедииняем линейные и точечные объекты
      for i := 0 to sl.Count - 1 do
      begin
        path := AddSlesh(ExtractFileName(ExtractFileDir(sl[0])));
        ForceDirectories(prepared + path);//подкаталог с таким же название как и в CF_x
        tab := StringReplace(ExtractFileName(sl[i]), ' ', '_', [rfReplaceAll]);
        open_prepared := false;

        if FileExists(prepared + path + tab) then
        begin// открываем готовую таблицу-буфер
          miDo('Open Table "%s"', [prepared + path + tab]);
          open_prepared := true;//открываем готовые буфера
        end else
          miDo('Open Table "%s"', [sl[i]]);//открываем таблицу

        tab := miEval('PathToTableName$("%s")', [sl[i]]);
        try//только линейные объекты
          if (miEvalInt('ObjectInfo(%s.obj, %d)', [tab, OBJ_INFO_TYPE]) = 4) then
          begin //строим буфера для линейных объектов
            Info(tab + ': построение буферов...');
            row := 1;
            pg.Max := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
            While miEval('EOT(%s)', [tab]) <> 'T' do
            begin
              if row mod 10 = 0 then
                pg.Position := row;
              miDo('Fetch Rec %d From %s', [row, tab]);
              miDo('ob = buffer(%s.obj, 12, %s, "m")', [tab, buf]);
              miDo('update %s set obj=ob where rowid=%d', [tab, row]);
              Inc(row);
            end;
            pg.Max := 0;
          end else
            open_prepared := true;//буфера не строятся, берем элипсы из потребителей

          Info(tab + ': выбор данных..');
          // границы участков частично лежат внутри коммуникаций
          miDo('Select * from All_CN,%s where All_CN.Obj Partly Within %s.Obj into Temp noSelect',
            [tab, tab]);
          //добавляем результаты пересечений в итоговую таблицу
          if i = 0 then
          begin
            Info(tab + ': Создание итоговой таблицы...');
            miDo('Commit Table Temp As "%s%s.tab" TYPE NATIVE Charset "WindowsCyrillic"',
              [compath, ExtractFileName(ExtractFileDir(sl[i]))]);
            miDo('Open Table "%s%s.tab" As Itog',
              [compath, ExtractFileName(ExtractFileDir(sl[i]))]);
          end else
          begin
            Info(' Добавление в итоговую таблицу...');
            miDo('Insert Into Itog Select * From Temp');
          end;
          miDo('Close table Temp');

        finally
          if not open_prepared then//сохраним построенные буфера в каталоге PREEPARED
            miDo('Commit table %s As "%s%s.tab" TYPE NATIVE Charset "WindowsCyrillic"',
              [tab, prepared + path, tab]);
          miDo('Rollback Table %s', [tab]);//если строились буфера,то восстановим
          miDo('Close table %s', [tab]);
        end;
      end;
      Info(tab + ': Сохранение таблицы');//типа водоотведение, Водоснабжение и др.
      miDo('Commit table Itog');
      miDo('Close table Itog');
    end;
  finally
    fold.Free;
    sl.Free;
  end;
end;
{------------------- расчет пересечений по созданным буферам-------------------}
procedure TForm2.CalcIntersect2;
var
  k : Integer;
  compath, respath, prepared, s : String;
  fold : TStringList;
begin
  compath := ResFolder + 'коммуникации\';
  respath := ResFolder + 'коммуникации\RESULT\';
  prepared := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'BUFFERS\';

  ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [compath])), SW_HIDE, true, 0);
  ForceDirectories(compath);
  ForceDirectories(respath);
  SafeCloseTAble('All_CN');
  fold := TStringList.Create;
  try//выбираем из базовой таблицы (например KN17) все КН
    FindFolders(prepared, fold, false);
    miDo('Select %s from %s order by %s into temp noSelect', [fcn, ftab, fcn]);
    miDo('Commit Table temp As "%sAll_CN.TAB" TYPE NATIVE Charset "WindowsCyrillic"',
      [respath]);//сохраняем КН в каталоге ...коммуникации\RESULT
    miDo('Close Table temp');
    cpu := fold.Count;
    for k := 0 to fold.Count - 1 do // идем по частям буферов коммуникаций
    begin
      s := TempPath + IntTostr(k+1) + '\';
      ForceDirectories(s);
      ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%sAll_CN.*" "%s"',
       [respath, s])), SW_HIDE, true, 0);
      ExCreateProcess(format('"%sCommunications.exe" "%s" "%s" "%s" %d',
        [GetAppFolder, s, AddSlesh(fold[k]), compath, handle]), '', SW_MINIMIZE, false);
    end;
    Repeat
      Sleep(1000);
      Application.ProcessMessages;
    Until cpu = 0;
  finally
    fold.Free;
  end;
end;
{--------определение приоритета коммуникаций по каждому виду------------------}
procedure TForm2.CalcFactors;
var
  i, k, row, prior, predprior: Integer;
  compath, sprpath, respath, t, cn, predcn, s: String;
  sl: TStringList;
begin
  compath := ResFolder + 'коммуникации\';
  respath := ResFolder + 'коммуникации\RESULT\';
  sprpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'коммуникации\';

  sl := TStringList.Create;
  try
    FindRecursive(compath, '*.tab', false, sl);
    // объекты,входящие в коммуникации - все части
    if sl.Count = 0 then
      Exc('CalcFactors : Нет результатов расчета');

    SafeCloseTAble('Коммуникации');
    SafeCloseTAble('Градации_коммуникаций');
    Info(' определение приоритета коммуникаций..'); // открываем из справочника
    miDo('Register Table "%sГрадации_коммуникаций.xlsx" TYPE XLS' +
      ' Titles Range "Приоритет_Значения_коммуникаций"' +
      ' Into "%sГрадации_коммуникаций.TAB"', [sprpath, sprpath]);
    miDo('Open table "%sГрадации_коммуникаций.tab"', [sprpath]);

    miDo('Open Table "%sAll_CN.TAB"', [respath]); //если это меню вызывается отдельно

    Info(' Группировка по кадастровым номерам...');
    miDo('Select * from All_CN group by %s into Коммуникации noSelect', [fcn]);
    // выбираем уникальные кадастровые номера группируя их из общей таблицы КН
    miDo('Commit Table Коммуникации As "%sКоммуникации.TAB" TYPE NATIVE Charset'
      + ' "WindowsCyrillic"', [respath]);
    miDo('Close Table Коммуникации');
    miDo('Open Table "%sКоммуникации.TAB"', [respath]);
    miDo('Close table All_CN');
    Info(' Добавление колонок...');
    s := 'Alter Table "Коммуникации" (add ';
    for i := 1 to sl.Count do//добавляем колонки
      s := s + format
        ('IDCGKO%d Char(30),Тип_коммуникации%d Char(32),Примечание%d Char(50),'
        + 'Графика%d Char(20),Приоритет%d Integer,', [i, i, i, i, i]);
    s[Length(s)] := ')';
    miDo(s);

    for i := 0 to sl.Count - 1 do
    begin
      miDo('Open table "%s"', [sl[i]]);//открываем водоотведение, Водоснабжение и др.
      t := miEval('PathToTableName$("%s")', [sl[i]]);
      Info(t + ': Выборка..');
      miDo('Select %s.%s,%s.IDCGKO,%s.Тип_коммуникации,%s.Примечание,%s.Графика,'
        + 'Градации_коммуникаций.Приоритет from %s, Градации_коммуникаций' +
        ' where %s.Примечание=Градации_коммуникаций.Значение' +
        ' group by Градации_коммуникаций.Приоритет,%s.%s' +
        ' order by %s.%s, %s.Графика Desc into Link noSelect', [t, fcn, t, t, t, t, t, t, t, fcn,
        t, fcn, t]);

      miDo('Commit table Link As "%sLink.tab" TYPE NATIVE Charset "WindowsCyrillic"',
        [TempPath]);
      miDo('Close table Link');
      miDo('Close table %s', [t]);
      miDo('Open table "%sLink.tab"', [TempPath]);

      Info(ftab + ': Минимум..');
      pg.Max := miEvalInt('TableInfo(Link, %d)', [TAB_INFO_NROWS]);
      miDo('Fetch First From Link');
      predcn := miEval('Link.%s', [fcn]);
      predprior := miEvalInt('Link.Приоритет');
      row := 2;
      While miEval('EOT(Link)') <> 'T' do
      begin//определение приоритета
        miDo('Fetch Rec %d From Link', [row]);
        cn := miEval('Link.%s', [fcn]);
        prior := miEvalInt('Link.Приоритет');
        if cn = predcn then
        begin
          if prior > predprior then
            miDo('Delete from Link where Rowid = %d', [row])
          else
            miDo('Delete from Link where Rowid = %d', [row - 1]);
          if prior < predprior then
            predprior := prior;
        end
        else
          predprior := prior;
        predcn := cn;
        Inc(row);
        if row mod 100 = 0 then
          pg.Position := row;
      end;
      pg.Max := 0;

      miDo('Select Коммуникации.%s, Коммуникации.IDCGKO%d, Коммуникации.Тип_коммуникации%d,'
        + ' Коммуникации.Примечание%d, Коммуникации.Графика%d, Коммуникации.Приоритет%d from Коммуникации,'
        + ' Link where Коммуникации.%s=Link.%s into Sel noSelect',
        [fcn, i + 1, i + 1, i + 1, i + 1, i + 1, fcn, fcn]);
      Info(t + ': Обновление колоноки IDCGKO...');
      miDo('Add Column "Sel" (IDCGKO%d) From Link Set To IDCGKO Where COL1 = COL1',
        [i + 1]);
      Info(t + ': Обновление колоноки Тип_коммуникации...');
      miDo('Add Column "Sel" (Тип_коммуникации%d) From Link Set To Тип_коммуникации Where COL1 = COL1',
        [i + 1]);
      Info(t + ': Обновление колоноки Примечание...');
      miDo('Add Column "Sel" (Примечание%d) From Link Set To Примечание Where COL1 = COL1',
        [i + 1]);
      Info(t + ': Обновление колоноки Графика...');
      miDo('Add Column "Sel" (Графика%d) From Link Set To Графика Where COL1 = COL1',
        [i + 1]);
      Info(t + ': Обновление колоноки Приоритет...');
      miDo('Add Column "Sel" (Приоритет%d) From Link Set To Приоритет Where COL1 = COL1',
        [i + 1]);
      miDo('Close table Sel');
      miDo('Drop table Link');
    end;
  finally
    sl.Free;
    miDo('Commit table Коммуникации');
    SafeCloseTable('Градации_коммуникаций');
//    if TableExists('All_CN') then
//      miDo('Drop table "All_CN"');
    Info('Готово.');
  end;
end;
{ ---------------Расчет приоритета по графике MapInfo--------------------------}
procedure TForm2.mnGraphicClick(Sender: TObject);
begin
  if (not st16) and (not ExamenBufferSize) then Exit;//ГКО не созданы буфера для коммуникаций
  if not CheckWorkTable('mnGraphicClick') then Exit;

  miDo('Open table "%s"', [ExtractFilePath(worktable) + 'Ценовые_Факторы_Итог.tab']);

//  if st16 then
    CalcMerge;//объединение частей линейных и точечных объектов по названию каталога
//  else
//    CalcIntersect2;
  CalcFactors;
  if Sender <> nil then
    miDo('Browse * From Коммуникации');
end;
{------------------------упаковка таблицы Acces--------------------------------}
// Оператор SELECT  Microsoft Access  tr!!!
// https://docs.microsoft.com/ru-ru/office/client-developer/access/desktop-database-reference/select-statement-microsoft-access-sql
function TForm2.CompactDatabase(SrcName: String): Boolean;
var
  tempname: String;
  v: Variant;
begin
  Info('Упаковка базы...');
  ADOConnection1.Connected := false;
  Result := false;
  try
    try
      v := CreateOleObject('DAO.DBEngine.120');
      tempname := ExtractFilePath(SrcName) + 'temp_file.accdb';
      DeleteFile(tempname);
      v.CompactDatabase(SrcName, tempname);
      if FileExists(tempname) then
      begin
        DeleteFile(SrcName);
        RenameFile(tempname, SrcName);
        Result := true;
      end;
    finally
      v := Unassigned;
    end;
  except
    on E: Exception do
      Note(E.message);
  end;
end;

(*
  MSysIMEXSpec
  DateDelim: определяет разделитель дат. Обычно это / или -. Это может быть только один или нулевой символ.
  DateFourDigitYear: Поле «Да/Нет», указывающее, форматируются ли даты
  с использованием четырех цифр для обозначения лет. Если установлено значение yes,
  такие значения, как 10/10/20, не будут считаться действительными датами.
  DateLeadingZeros: еще одно поле «Да/Нет», указывающее, есть ли в датах начальные нули.
  DateOrder: определяет формат даты, который мы хотим использовать. Действительны следующие коды:
  0: ДМИ
  1: ДИМ
  2: МДГ
  3: МИД
  4: ярд
  5: ГМД
  DecimalPoint: Указывает символ для разделения целой и дробной части десятичного числа.
  Это должен быть только один символ.
  FieldSeparator: разделяет поля в текстовом файле. Для файла CSV это будет запятая.
  Кроме того, это должен быть только один символ.
  FileType: аналогичен набору символов schema.ini. Как упоминалось в предыдущей статье,
  вы можете найти все возможные кодовые страницы, поддерживаемые на вашем компьютере с Windows,
  в разделе реестра Computer\HKEY_CLASSES_ROOT\MIME\Database\Codepage\.
  Он определяет кодировку текстового файла. Вы можете использовать тот же идентификатор кодовой страницы.
  Однако есть одна проблема. Тип файла определяется как целое число.
  Следовательно, кодовая страница 65000 (например, Unicode) будет отображаться как -536.
  Чтобы преобразовать любой идентификатор кодовой страницы, превышающий 32767,
  вы можете использовать формулу CInt("&H" & Hex(x)) и, чтобы отменить операцию, CLng("&H" & Hex(x)).
  Например, преобразование кодовой страницы 65000 пройдет через CInt("&H" & Hex(65000)) и даст -536.
  Чтобы изменить, CLng("&H" & Hex(-536)) чтобы снова получить 65000.
  SpecID: Автонумерация для таблицы. Обратите внимание, что это не первичный ключ таблицы.
  Этот столбец используется с MSysIMEXColumns.SpecID для связи записей.
  Однако между двумя таблицами нет ограничений по внешнему ключу, поэтому могут происходить потерянные удаления.
  SpecName: первичный ключ таблицы. Он однозначно идентифицирует спецификацию.
  Всякий раз, когда вы сохраняете спецификацию, указанное вами имя будет записано в этом поле.
  Когда вы импортируете или экспортируете с помощью мастеров,
  Access может автоматически создать новую спецификацию, даже если вы не использовали Advanced
  и явно не сохранили ее.
  SpecType: может быть либо 1 для обозначения файла с разделителями, либо 2 для
  обозначения файла с фиксированной шириной.
  StartRow: определяет строку в текстовом файле, с которой начинается сканирование для импорта.
  Это может быть любое число, поэтому можно пропустить несколько строк вверху файла.
  Это может быть полезно для текстовых файлов неправильного формата с несоответствующими заголовками.
  TextDelim: одиночный символ, разделяющий текстовое значение.
  Его можно опустить, если текст не разделен. Если вы можете контролировать формат текстовых файлов,
  обычно рекомендуется использовать текстовые разделители.
  Текстовый файл с текстовыми разделителями останется правильно сформированным,
  даже если текст содержит разделители полей (например, запятую) внутри текста.
  TimeDelim: указывает символ-разделитель, используемый для времени.
  Это может быть любой одиночный символ или быть пустым и по умолчанию : (символ двоеточия).

  MSysIMEXColumns
  Attributes/Атрибуты: Неизвестно. В моих тестах я никогда не получал ничего, кроме 0. Если вы знаете,
  для чего это нужно и как его можно использовать, не стесняйтесь оставлять комментарии.
  DataType/DataType: тип данных столбца. Номер должен соответствовать DAO.DataTypeEnum.
  FieldName/FieldName: имя столбца. Обратите внимание, что Access по умолчанию использует FieldN,
  где N — положительное целое число.
  IndexType/IndexType: может быть 0, чтобы указать отсутствие индексов, 1, чтобы указать обычный индекс или 2,
  чтобы указать уникальный индекс.
  SkipColumns/SkipColumns: логическое значение указывает, следует ли пропустить столбец.
  Если пропустить, он будет недоступен для запроса или чтения.
  SpecID/SpecID: соответствует автонумерации MSysIMEXSpecs.SpecID.
  Обратите внимание, что между двумя таблицами нет ограничений внешнего ключа.
  Если вы удалили записи, в таблице MSysIMEXSpecs могут остаться потерянные записи.
  Start/Начало: целое число, указывающее, когда столбец начинается в текстовом файле.
  Это важно для файла с фиксированной шириной.
  Для файлов с разделителями Access по-прежнему будет использовать позицию столбца первой строки,
  в которой находится поле, для определения начала.
  Кроме того, Access будет использовать столбец для обозначения порядкового порядка столбцов.
  Width/Ширина: еще одно целое число, определяющее размер столбца.
  Для файла фиксированной ширины это также важно. Имейте в виду, что Microsoft Access
  будет использовать эту информацию для типов данных переменной длины, таких как текст или двоичный файл,
  для соотв
  етствующего размера столбца даже в файле с разделителями.
*)
{------------------------Расчет приоритета по базе Access----------------------}
procedure TForm2.mnBaseClick(Sender: TObject);
var
  i, ii, j, k: Integer;
  sprpath, respath, compath, prepared, s, ss: String;
  fold: TStringList;
  ds: TADODataset;
  fd: TFieldDef;

  procedure Access2MapInfo(filetxt: String);
  var
    i, j, k: Integer;
    mif, mid: String;
    sl: TStringList;
    sa: StringArray;
    AccApp: Variant;
  begin
    sl := TStringList.Create;
    try
      Info('Создание таблицы MSysIMEXColumns');
      ds.Close;
      ds.CommandText := 'select Top 1 * from [T0]'; // для выбора названий полей
      ds.Open; // для определения названия полей FieldDefs

      Query.sql.Clear;
      Query.sql.Add
        ('update mSysIMEXSpecs set DateDelim=''.'', FieldSeparator='';'',FileType=1251,SpecType=1');
      Query.ExecSQL;

      Query.sql.Clear;
      Query.sql.Add('Delete from MSysIMEXColumns');
      Query.ExecSQL;
      k := 1;
      for j := 0 to ds.FieldDefs.Count - 1 do
      begin
        fd := ds.FieldDefs.Items[j];
        if not(fd.DataType in [ftInteger, ftFloat]) then
          i := fd.Size
        else
          i := 11; // размер для целого числа
        Query.sql.Clear;
        Query.sql.Add
          (format('INSERT INTO MSysIMEXColumns VALUES(0,10,''%s'',0,0,9,%d,%d)',
          [fd.name, k, i]));
        Query.ExecSQL;
        Inc(k, i);
      end;
      ADOConnection1.Connected := false;
      AccApp := CreateOleObject('Access.Application');
      try
        Info('Открытие Коммуникации_свод.accdb...');
        AccApp.OpenCurrentDatabase(format('%sКоммуникации_свод.accdb', [compath]));
        Info('Экпорт Коммуникации_свод.accdb...');
        AccApp.DoCmd.TransferText(2, 'specNote', 'T0', filetxt, false, '', 0);// без заговка
      finally
        AccApp := Unassigned;
      end;

      Info('Создание таблицы выгрузки..');
      sl.LoadFromFile(filetxt);
      for i := 0 to sl.Count - 1 do
      // выбираем наибольший размер по каждому полю
      begin
        sa := SplitStr(sl[i], ';');
        for j := 1 to ds.FieldDefs.Count - 1 do
        begin
          fd := ds.FieldDefs.Items[j];
          if not(fd.DataType in [ftInteger, ftFloat]) then
          begin
            k := Length(sa[j]);
            if k > 254 then
              k := 254; // задаем максимальный размер
            if i = 0 then // задаем размеры по первой строке
              fd.Size := k
            else if fd.Size < k then
              fd.Size := k;
          end;
        end;
        SetLength(sa, 0);
      end;

      DeleteFile(ChangeFileExt(filetxt, '.MID'));
      RenameFile(filetxt, ChangeFileExt(filetxt, '.MID'));
      mif := ChangeFileExt(filetxt, '.MIF');
      mid := ChangeFileExt(filetxt, '.MID');

      k := sl.Count;
      sl.Clear;
      sl.Add('Version   300');
      sl.Add('Charset "WindowsCyrillic"');
      sl.Add('Delimiter ";"');
      sl.Add(format('Columns %d', [ds.FieldDefs.Count]));
      for j := 0 to ds.FieldDefs.Count - 1 do
      begin
        fd := ds.FieldDefs.Items[j];
        if fd.DataType in [ftInteger, ftFloat] then
          sl.Add(fd.name + ' Integer')
        else
          sl.Add(fd.name + format(' CHAR(%d)', [fd.Size]));
      end;
      sl.Add('Data');
      for i := 1 to k do
        sl.Add('None');
      sl.SaveToFile(mif);
      miDo('Import "%s" Type "MIF" Into "%s"  Overwrite',
        [mif, ChangeFileExt(mif, '.TAB')]);
      DeleteFile(mif);
      DeleteFile(mid);
      miDo('Create Index On Выгрузка (cadastralnumber)');
      if miEval('TableInfo(Выгрузка, %d)', [TAB_INFO_MAPPABLE]) = 'T' then
        miDo('drop map Выгрузка');
    finally
      sl.Free;
      Info('Готово.');
    end;
  end;

begin//mnBaseClick работает независимо от таблиц MapInfo
  if ResFolder = '' then
    Exc('Результирующий каталог или таблица не определены');
  sprpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'коммуникации\';
  respath := ResFolder + 'коммуникации\RESULT\';
  compath := ResFolder + 'коммуникации\';
  prepared := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'PREEPARED\';

  if not FileExists(sprpath + 'Градации_коммуникаций.xlsx') then
    Exc('mnBaseClick : файл ' + sprpath +
      'Градации_коммуникаций.xlsx не найден');
  if not FileExists(sprpath + 'Коммуникации_свод.accdb') then
    Exc('mnBaseClick : файл ' + sprpath + 'Коммуникации_свод.accdb не найден');
  Info('Готово.');
  if FileExists(prepared + 'Выгрузка.tab') then
    exit; // если есть готовый файл то выходим
//ConnectionString для ADOConnection1 менять для каждой новой версии CF
  Info('Копирование базы...');
  ADOConnection1.Connected := false;
  ADOConnection1.ConnectionString := format(strConnect1, [ResFolder + 'коммуникации\Коммуникации_свод.accdb']);
  ForceDirectories(compath);
  CopyFile(PWideChar(sprpath + 'Коммуникации_свод.accdb'),
    PWideChar(compath + 'Коммуникации_свод.accdb'), false);
  Info('Определение приоритета по КН...');
  Query.sql.Clear;
  Query.sql.Add
    ('SELECT [Свод коммуникаций].CadastralNumber, [Свод коммуникаций].[Тип коммуникации],'
    + ' Min(Приоритет_Значения_коммуникаций.Приоритет) AS [Min-Приоритет] INTO [Приоритеты по КН]'
    + ' FROM [Свод коммуникаций] INNER JOIN Приоритет_Значения_коммуникаций ON'
    + ' ([Свод коммуникаций].Значение = Приоритет_Значения_коммуникаций.Значение)'
    + ' AND ([Свод коммуникаций].[Тип коммуникации] = Приоритет_Значения_коммуникаций.[Тип коммуникации])'
    + ' GROUP BY [Свод коммуникаций].CadastralNumber, [Свод коммуникаций].[Тип коммуникации]'
    + ' HAVING (((Min(Приоритет_Значения_коммуникаций.Приоритет)) Is Not Null And'
    + ' (Min(Приоритет_Значения_коммуникаций.Приоритет))>0))');
  Query.ExecSQL;

  // Общая длина группирующего предложения Group By не должна превышать 255 символов.
  Info('Выбор коммуникаций по КН...');
  Query.sql.Clear;
  Query.sql.Add
    ('SELECT [Приоритеты по КН].CadastralNumber, Приоритет_Значения_коммуникаций.[Тип коммуникации],'
    + ' Приоритет_Значения_коммуникаций.Значение, Приоритет_Значения_коммуникаций.Источник3,'
    + ' Приоритет_Значения_коммуникаций.Приоритет, Приоритет_Значения_коммуникаций.Значение_уникальное INTO'
    + ' [Коммуникации по КН] FROM [Приоритеты по КН] INNER JOIN Приоритет_Значения_коммуникаций ON'
    + ' ([Приоритеты по КН].[Тип коммуникации] = Приоритет_Значения_коммуникаций.[Тип коммуникации]) AND'
    + ' ([Приоритеты по КН].[Min-Приоритет] = Приоритет_Значения_коммуникаций.Приоритет)');
  Query.ExecSQL;

  Info('Группировка кадастровых номеров...');
  // выбираем уникальные кадастровые номера в новую таблицу
  Query.sql.Clear;
  Query.sql.Add
    ('Select [Свод коммуникаций].cadastralnumber into T0 from [Свод коммуникаций] GROUP BY cadastralnumber');
  Query.ExecSQL;

  Query.sql.Clear;
  Query.sql.Add('Drop table [Приоритеты по КН]');
  Query.ExecSQL;

  Query.sql.Clear;
  Query.sql.Add('Drop table [Свод коммуникаций]');
  Query.ExecSQL;

  Query.sql.Clear;
  Query.sql.Add('Drop table Приоритет_Значения_коммуникаций');
  Query.ExecSQL;

  CompactDatabase(compath + 'Коммуникации_свод.accdb');
  try; // растягиваем поля в одну строку для каждого КН
    Info('Изменение структуры базы...');
    fold := TStringList.Create;
    ds := TADODataset.Create(nil);
    ds.CacheSize := 100;
    ds.CursorLocation := clUseServer;
    ds.Connection := ADOConnection1;
    try
      FindFolders(sprpath, fold, false);
      k := fold.Count; // кол-во коммуникаций включая части
      ds.CommandText := 'select Top 1 * from [Коммуникации по КН]';
      // для выбора названий полей
      ds.Open; // для определения названия полей FieldDefs
      s := '';
      ii := 0;
      for i := 1 to k do
      begin
        ss := ExtractFileName(fold[i - 1]);
        if (Pos('_', ss) = 0) or (Pos('часть1', ss) > 0) then
        begin
          Inc(ii);
          for j := 1 to ds.FieldDefs.Count - 1 do
          // кадастровый номер пропускаем
          begin
            fd := ds.FieldDefs.Items[j];
            if fd.DataType in [ftInteger, ftFloat] then
              s := s + 'COLUMN ' + StringReplace(fd.name, ' ', '_',
                [rfReplaceAll]) + IntToStr(ii) + ' INT,'
            else
              s := s + 'COLUMN ' + StringReplace(fd.name, ' ', '_',
                [rfReplaceAll]) + IntToStr(ii) + format(' VARCHAR(%d),',
                [fd.Size]);
          end;
        end;
      end;
      Query.sql.Clear;
      Query.sql.Add('ALTER TABLE T0 add ' + Copy(s, 1, Length(s) - 1));
      Query.ExecSQL;

      ii := 0;
      for i := 1 to k do
      begin
        s := ExtractFileName(fold[i - 1]);
        if (Pos('_', s) = 0) or (Pos('часть1', s) > 0) then
        begin
          Inc(ii);
          j := Pos('_', s);
          if j > 0 then
            s := Copy(s, 1, j - 1);

          Query.sql.Clear;
          Query.sql.Add
            (format('SELECT * INTO T1 from [Коммуникации по КН] Where [Тип коммуникации]="%s"',
            [s]));
          Query.ExecSQL;
          // обновление колонок
          for j := 1 to ds.FieldDefs.Count - 1 do
          // кадастровый номер пропускаем
          begin
            fd := ds.FieldDefs.Items[j];
            ss := StringReplace(fd.name, ' ', '_', [rfReplaceAll]) +
              IntToStr(ii);
            Info('Обновление ' + ExtractFileName(fold[i - 1]) + ': ' + ss);
            Query.sql.Clear;
            Query.sql.Add
              (format('UPDATE T0 INNER JOIN T1 ON T0.cadastralnumber=T1.cadastralnumber'
              + ' SET T0.[%s]=T1.[%s]', [ss, fd.name]));
            Query.ExecSQL;
          end;
          Query.sql.Clear;
          Query.sql.Add('Drop table T1');
          Query.ExecSQL;
        end;
      end;
      CompactDatabase(compath + 'Коммуникации_свод.accdb');
      Query.sql.Clear;
      Query.sql.Add('Drop table [Коммуникации по КН]');
      Query.ExecSQL;
      Access2MapInfo(prepared + 'Выгрузка.TXT');
    finally
      fold.Free;
      ds.Close;
      ds.Free;
    end;
  except
    on E: Exception do
      Exc('mnBaseClick : Ошибка при изменении структуры базы' + #10 +
        E.message);
  end;

  if Sender <> nil then
    Note('процес завершен.');
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnBuffersClick(Sender: TObject);
var
  i, j, k, row: Integer;
  tab, cat, oldtab, bufpath, buf, gko, s: String;
  fold, sl : TStringList;   t : Cardinal;

  procedure CloseItog;
  var
    j : Integer;
  begin
    miDo('Select * from Itog group by IDCGKO into groupGKO noSelect');
    for j := 1 to miEvalInt('TableInfo(groupGKO, %d)', [TAB_INFO_NROWS]) do
    begin
      miDo('Fetch Rec %d From groupGKO', [j]);
      gko := miEval('groupGKO.IDCGKO');
      Info('Itog: ' + gko);
      miDo('Select * from Itog Where IDCGKO="%s" into selectGKO noSelect', [gko]);
      miDo('Objects Combine DropIndex Auto Data IDCGKO=IDCGKO,Примечание=Примечание,'+
        'Тип_коммуникации=Тип_коммуникации,Графика=Графика');
      miDo('Objects Disaggregate Into Table Itog DropIndex Auto '+
        'Data IDCGKO=IDCGKO,Примечание=Примечание,Тип_коммуникации=Тип_коммуникации,'+
        'Графика=Графика');
    end;
    miDo('Commit table Itog');
    miDo('Pack Table Itog Graphic Data');
    Safeclosetable('Itog');
  end;

begin//mnBuffersClick
  t := GetTickCount;
  bufpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'BUFFERS\';
  s := 'Каталог ' + bufpath + #10 + 'существует заменить ?';
  if DirectoryExists(bufpath) then
    if MessageDlg (s, mtError, mbOKCancel, 0 ) = mrCancel then Exit;
  ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [bufpath])), SW_HIDE, true, 0);//удаление каталога со всем содержимым и подкаталогами

  fold := TStringList.Create;
  sl := TStringList.Create;
  try// в справочнике д.б. только части коммуникаций
    s := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'коммуникации\';
    FindFolders(s, fold, false);
    if fold.Count = 0 then
      Exc('mnBuffersClick : Нет подкаталогов в каталоге ' + s);
    oldtab := '';
    for k := 0 to fold.Count-1 do // идем по частям коммуникаций
    begin
      sl.LoadFromFile(AddSlesh(fold[k]) + 'Buffer.txt');
      buf := sl[0];
      sl.Clear;// ищем таблицы внутри каталога

      FindRecursive(AddSlesh(fold[k]), '*.tab', false, sl);
      for i := 0 to sl.Count - 1 do//Водоотведение_часть1, Водоотведение_часть2
      begin// в каждой части объедииняем линейные объекты
        tab := miEval('PathToTableName$("%s")', [sl[i]]);
        miDo('Open Table "%s"', [sl[i]]);
        row := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);

        if (miEvalInt('ObjectInfo(%s.obj, %d)', [tab, OBJ_INFO_TYPE]) = 4) then
        try//только линейные объекты строим буфера
          if oldtab <> tab then
          begin
            if oldtab <> '' then
              CloseItog;//другой тип коммуникаций, сохраняем предыдущий итог
            cat := ExtractFileName(fold[k]);
            j := Pos('_часть', cat);
            if j > 0 then
              cat := Copy(cat, 1, j-1);
            bufpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'BUFFERS\' + cat + '\';
            ForceDirectories(bufpath);
            Safeclosetable('Itog');//делаем фиктивный выбор для создания таблицы
            miDo('Select IDCGKO, Примечание, Тип_коммуникации, Графика '+
              'from %s where rowid=-1 into selectGKO noSelect', [tab]);
            miDo('Commit table selectGKO As "%s%s.tab"', [bufpath, tab]);
            miDo('Open Table "%s%s.tab" As Itog', [bufpath, tab]);
            oldtab := tab;
          end;

          miDo('Select * from %s group by IDCGKO into groupGKO noSelect', [tab]);//группировка
          for j := 1 to miEvalInt('TableInfo(groupGKO, %d)', [TAB_INFO_NROWS]) do
          begin//Построение буферов для каждого IDCGKO
            miDo('Fetch Rec %d From groupGKO', [j]);
            gko := miEval('groupGKO.IDCGKO');
            Info(tab + ': ' + gko);
            miDo('Select * from %s Where IDCGKO="%s" into selectGKO', [tab, gko]);
            miDo('Create Object As Buffer From selectGKO Width %s Units "m" '+
            'Type Cartesian Resolution 12 Into Table %s Data IDCGKO=IDCGKO,'+
            'Примечание=Примечание,Тип_коммуникации=Тип_коммуникации,'+
            'Графика=Графика noSelect', [buf, tab]);
            miDo('Objects Disaggregate Into Table %s DropIndex Auto Data '+
              'IDCGKO=IDCGKO,Примечание=Примечание,'+
              'Тип_коммуникации=Тип_коммуникации,Графика=Графика', [tab]);
            miDo('Select IDCGKO, Примечание, Тип_коммуникации, Графика from '+
              '%s where rowid>%d into selectGKO noSelect', [tab, row]);
            miDo('Insert into Itog select * from selectGKO');
          end;
        finally
          miDo('Rollback Table %s', [tab]);//восстановим таблицу ???mievalint('tableinfo(%s, %d)', [tab, TAB_INFO_EDITED])
        end;
        miDo('Close table %s', [tab]);
      end;
    end;
    CloseItog;

    oldtab := '';
    for k := 0 to fold.Count-1 do // идем по частям коммуникаций
    begin
      sl.Clear;// ищем таблицы внутри каталога
      FindRecursive(AddSlesh(fold[k]), '*.tab', false, sl);
      for i := 0 to sl.Count - 1 do//Водоотведение_часть1, Водоотведение_часть2
      begin// в каждой части объедииняем потребителей
        tab := miEval('PathToTableName$("%s")', [sl[i]]);
        miDo('Open Table "%s"', [sl[i]]);
        miDo('set coordsys table %s', [tab]);
        row := miEvalInt('TableInfo(%s, %d)', [tab, TAB_INFO_NROWS]);
        miDo('Fetch First From %s', [tab]);
        miDo('ob = %s.obj', [tab]);
        if (miEvalInt('ObjectInfo(%s.obj, %d)', [tab, OBJ_INFO_TYPE]) = 2) then//OBJ_TYPE_ELLIPSE
        try
          if oldtab <> tab then
          begin
            cat := ExtractFileName(fold[k]);
            j := Pos('_часть', cat);
            if j > 0 then
              cat := Copy(cat, 1, j-1);
            bufpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'BUFFERS\' + cat + '\';
            ForceDirectories(bufpath);
            Safeclosetable('Itog');//делаем фиктивный выбор для создания таблицы
            miDo('Select IDCGKO, Примечание, Тип_коммуникации, Графика '+
              'from %s where rowid=-1 into selectGKO noSelect', [tab]);
            miDo('Commit table selectGKO As "%s%s.tab"', [bufpath, tab]);
            miDo('Open Table "%s%s.tab" As Itog', [bufpath, tab]);
            oldtab := tab;
          end;
//          miDo('Select * from %s group by IDCGKO into groupGKO', [tab]);//группировка
//          for j := 1 to miEvalInt('TableInfo(groupGKO, %d)', [TAB_INFO_NROWS]) do
//          begin//объединение
//            miDo('Fetch Rec %d From groupGKO', [j]);
//            gko := miEval('groupGKO.IDCGKO');
//            Info(tab + ': ' + gko);
//            miDo('Select * from %s Where IDCGKO="%s" into selectGKO', [tab, gko]);
//            miDo('Objects Combine Data IDCGKO=IDCGKO,Примечание=Примечание,'+
//              'Тип_коммуникации=Тип_коммуникации,Графика=Графика');
//            miDo('Objects Disaggregate Data IDCGKO=IDCGKO,Примечание=Примечание,'+
//              'Тип_коммуникации=Тип_коммуникации,Графика=Графика Into Table %s', [tab]);
//          end;
//          miDo('Select IDCGKO, Примечание, Тип_коммуникации, Графика from '+
//            '%s where rowid>%d intO selectGKO', [tab, row]);
          miDo('Select IDCGKO, Примечание, Тип_коммуникации, Графика from '+
            '%s into selectGKO', [tab, row]);
          miDo('Insert into Itog select * from selectGKO');
          miDo('Commit table Itog');//ВСЕГДА!! таблица д.б. одна в первом каталоге
          Safeclosetable('Itog');//делаем фиктивный выбор для создания таблицы

        finally
          miDo('Rollback Table %s', [tab]);//восстановим таблицу
        end;
        miDo('Close table %s', [tab]);
      end;
    end;

    Info('Готово.');
    Note(format('процес завершен.%sВремя обработки %d минут', [#10, Round((GetTickCount - t) / 60000)]));
  finally
    fold.Free;
    sl.Free;
  end;
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnCommunicationItogClick(Sender: TObject);
var
  i, j, m, k, fcount, kcount: Integer;
  compath, respath, prepared, s, cn: String;
  stl, sl: TStringList;

  //в CF_ZOUIT_Communication добавим поля и данные из Коммуникации_итог
  procedure UpdateCommunication;
  var
    i: Integer;
    sl: TStringList;
    fields: String;
  begin
    if FileExists(ResFolder + 'CF_ZOUIT_Communication.TAB') then
    begin
      miDo('Open table "%sCF_ZOUIT_Communication.TAB"', [ResFolder]);
      fields := '';
      sl := getFieldsList('Коммуникации_итог');
      try
        for i := 1 to sl.Count - 1 do // число полей в Коммуникации_итог
          fields := fields + sl[i] + ' ' + getFieldsType(sl, i) + ',';
        try //добавляем поля из таблицы Коммуникации_итог
          miDo('Alter table CF_ZOUIT_Communication (add %s)',
            [Copy(fields, 1, Length(fields) - 1)]);
        except
          Info('Ошибка при добавлении полей...');
        end;
        for i := 1 to sl.Count - 1 do//добавляем данные из коммуникаций
          miDo('Add Column CF_ZOUIT_Communication (%s) From Коммуникации_итог Set To %s Where COL2 = COL1',
            [sl[i], sl[i]])
      finally
        sl.Free;
      end;
      miDo('Commit table CF_ZOUIT_Communication');
      miDo('browse * from CF_ZOUIT_Communication');
    end else
      Note('Таблица CF_ZOUIT_Communication.TAB не найдена' + #10 +
        'Вероятно не сделан расчет ценовых факторов');
  end;

begin//mnCommunicationItogClick
  if worktable = '' then
    Exc('mnCommunicationItogClick : Рабочая таблица не открыта.');

  compath := ResFolder + 'коммуникации\';
  respath := ResFolder + 'коммуникации\RESULT\';
  prepared := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'PREEPARED\';

  if not FileExists(respath + 'Коммуникации.TAB') then
    Exc('mnAplayClick : Не найден файл ' + #10 + respath + 'Коммуникации.TAB');
  if not FileExists(prepared + 'Выгрузка.TAB') then
    Exc('mnAplayClick : Не найден файл ' + #10 + prepared + 'Выгрузка.TAB');

  miDo('Open table "%s"', [respath + 'Коммуникации.TAB']);
  miDo('Open Table "%sВыгрузка.TAB"', [prepared]);

  stl := nil;
  sl := TStringList.Create;
  try
    FindRecursive(compath, '*.tab', false, sl);
    if sl.Count = 0 then
      Exc('mnAplayClick : Нет результатов расчета');

    cn := miEval('ColumnInfo("Коммуникации", "col1", %d)', [COL_INFO_NAME]);
    for i := 0 to sl.Count - 1 do // приводим к одному названию
    begin
      s := ExtractFileName(sl[i]);
      k := Pos('_', s);
      if k = 0 then
        sl[i] := AnsiUpperCase(Copy(s, 1, Length(s) - 4))
      else
        sl[i] := AnsiUpperCase(Copy(s, 1, k - 1));
    end;

    kcount := 1; // подсчет количествы коммуникаций исключая части
    for i := 1 to sl.Count - 1 do
      if sl[i] <> sl[i - 1] then
        Inc(kcount); // число коммуникаций

    SafeCloseTAble('Коммуникации_итог');
    s := getColumInfo('Коммуникации');
    i := Pos(',IDCGKO' + IntToStr(kcount + 1), s);
    if i > 0 then // есть части
      s := Copy(s, 1, i - 1) + ')';
    s := format
      ('Create table "Коммуникации_итог" %s file "%sКоммуникации_итог.tab"' +
      ' TYPE NATIVE Charset "WindowsCyrillic"', [s, respath]);
    miDo(s); // созлаем итоговую таблицу для объединения частей

    Info('Вставка кадастровых номеров...');
    miDo('Insert Into Коммуникации_итог (COL1) Select COL1 From Коммуникации');
    miDo('commit table Коммуникации_итог');

    stl := getFieldsList('Коммуникации_итог'); // поля для заполнения
    fcount := 0;
    for i := 1 to stl.Count - 1 do
    begin
      if Pos('IDCGKO2', stl[i]) = 1 then
        break; // число полей 1 коммуникации без кадастрового номера
      Inc(fcount);
    end;

    j := 2; // номер колонки в сводной таблие (по чястям)
    k := 1; // номер колонки в итоговой таблице
    for i := 0 to sl.Count - 1 do
    begin
      if (i = 0) or (sl[i] <> sl[i - 1]) then
      begin
        Info('Выбор полей для вставки...');
        miDo('Select col1, col%d, col%d, col%d, col%d, col%d from Коммуникации into'
          + ' Temp where col%d > 0 NoSelect', [j, j + 1, j + 2, j + 3,
          j + 4, j + 4]);
        for m := 1 to fcount do
        begin
          Info('Вставка поля ' + stl[k + m - 1]);
          miDo('Add Column "Коммуникации_итог" (COL%d) From Temp Set To COL%d Where COL1 = COL1',
            [k + m, m + 1]);
        end;
        Inc(k, fcount);
      end
      else
      begin
        Info('Выбор полей для вставки...');
        miDo('Select col1, col%d, col%d, col%d, col%d, col%d from Коммуникации into'
          + ' Temp where col%d > 0 NoSelect', [j, j + 1, j + 2, j + 3,
          j + 4, j + 4]);
        miDo('select * from Коммуникации_итог where %s=0 into Temp_0 noSelect',
          [stl[k - 1]]);
        for m := 1 to fcount do
        begin
          Info('Вставка поля ' + stl[k + m - fcount - 1]);
          miDo('Add Column "Temp_0" (COL%d) From Temp Set To COL%d Where COL1 = COL1',
            [k + m - fcount, m + 1]);
        end;
      end;
      Inc(j, fcount);
    end;
    miDo('Commit table Коммуникации_итог');

    j := 2; // номер колонки в сводной таблие (по частям)
    k := 1; // номер колонки в итоговой таблице
    for i := 0 to sl.Count - 1 do
    begin
      if (i = 0) or (sl[i] <> sl[i - 1]) then
        Inc(k, fcount)
      else
      begin
        Info('Проверка приоритетов ' + stl[k - 1]);
        s := 'Select';
        for m := 1 to fcount do
          s := s + format(' Коммуникации_итог.COL%d,', [k + m - fcount]);
        for m := 1 to fcount do
          s := s + format(' Коммуникации.COL%d,', [j + m - 1]);
        s := Copy(s, 1, Length(s) - 1) +
          ' from Коммуникации_итог, Коммуникации ' +
          format('where Коммуникации_итог.%s=Коммуникации.%s AND', [cn, cn]) +
          format(' Коммуникации.COL%d > 0 AND', [j + fcount - 1]) +
          format(' Коммуникации.COL%d<Коммуникации_итог.COL%d into Temp',
          [j + fcount - 1, k]);
        miDo(s);
        if miEvalInt('TableInfo(Temp, %d)', [TAB_INFO_NROWS]) > 0 then
          for m := 1 to 5 do
            miDo('Update Temp Set COL%d = COL%d', [m, m + fcount]);
      end;
      Inc(j, fcount);
    end;
  finally
    sl.Free;
    if stl <> nil then
      stl.Free;
  end;

  Info('Сохранение итога коммуникаций...');
  miDo('Commit table Коммуникации_итог');
  miDo('Create Index On Коммуникации_итог (%s)', [cn]);

  for i := 1 to kcount do
  begin
    Info('Обновление данных колонка ' + IntToStr(i) + '...');
    s := format
      ('Select Коммуникации_итог.Тип_коммуникации%d, Коммуникации_итог.Примечание%d,'
      + ' Коммуникации_итог.Графика%d, Коммуникации_итог.Приоритет%d, Выгрузка.Тип_коммуникации%d,'
      + ' Выгрузка.Значение%d, Выгрузка.Источник3%d, Выгрузка.Приоритет%d' +
      format(' from Коммуникации_итог, Выгрузка where Коммуникации_итог.%s=Выгрузка.cadastralnumber',
      [cn]) + ' AND Коммуникации_итог.Приоритет%d > 0 AND Коммуникации_итог.Приоритет%d < Выгрузка.Приоритет%d'
      + ' into Link', [i, i, i, i, i, i, i, i, i, i, i]);
    miDo(s);
    m := miEvalInt('TableInfo(Link, %d)', [TAB_INFO_NROWS]);
    if m > 0 then // практически не работает
    begin
      miDo('drop index Выгрузка (cadastralnumber)');
      s := 'Update Link Set COL5 = COL1, COL6 = COL2, COL7 = COL3, COL8 = COL4';
      miDo(s);
      miDo('Commit table Выгрузка');
    end;
  end;
  miDo('Commit table Выгрузка');
  Info('Обновление структуры Коммуникации_итог ...');
  miDo('Alter Table "Коммуникации_итог" (rename Тип_коммуникации1 Наименование1,'
    + 'Тип_коммуникации2 Наименование2,Тип_коммуникации3 Наименование3,Тип_коммуникации4 Наименование4 '
    + 'modify Наименование1 Char(150),Наименование2 Char(150),Наименование3 Char(150),Наименование4 Char(150))');
  miDo('Update Коммуникации_итог Set Наименование1 = Примечание1+"_"+Наименование1+"_"+Графика1 DropIndex Auto');
  miDo('Update Коммуникации_итог Set Наименование2 = Примечание2+"_"+Наименование2+"_"+Графика2 DropIndex Auto');
  miDo('Update Коммуникации_итог Set Наименование3 = Примечание3+"_"+Наименование3+"_"+Графика3 DropIndex Auto');
  miDo('Update Коммуникации_итог Set Наименование4 = Примечание4+"_"+Наименование4+"_"+Графика4 DropIndex Auto');
  miDo('Commit table Коммуникации_итог');
  miDo('Alter Table "Коммуникации_итог" (drop Примечание1,Графика1,Примечание2,Графика2,Примечание3,Графика3,Примечание4,Графика4)');
  miDo('select * from Коммуникации_итог where Наименование1="__" into temp noSelect');
  miDo('Update temp Set Наименование1 = "" DropIndex Auto');
  miDo('select * from Коммуникации_итог where Наименование2="__" into temp noSelect');
  miDo('Update temp Set Наименование2 = "" DropIndex Auto');
  miDo('select * from Коммуникации_итог where Наименование3="__" into temp noSelect');
  miDo('Update temp Set Наименование3 = "" DropIndex Auto');
  miDo('select * from Коммуникации_итог where Наименование4="__" into temp noSelect');
  miDo('Update temp Set Наименование4 = "" DropIndex Auto');
  miDo('close table temp');
  miDo('Commit table Коммуникации_итог');

  FillEmptyFields('Коммуникации_итог');

  if st16 then
    UpdateCommunication;
  Info('Готово.');
  miDo('Browse * From Коммуникации_итог');
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnComplexCommunicationClick(Sender: TObject);
var
  t: Cardinal;
begin
{$IFDEF DEBUG}SetCurrentDir('D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\TEST\');{$ENDIF}
  if not CheckWorkTable('mnComplexCommunicationClick') then Exit;

  miDo('Open table "%s"', [ExtractFilePath(worktable) + 'Ценовые_Факторы_Итог.tab']);

  t := GetTickCount;
  mnGraphicClick(nil);
  mnBaseClick(nil);
  mnCommunicationItogClick(nil);
  Note(format('процес завершен.%sВремя обработки %d минут',
    [#10, Round((GetTickCount - t) / 60000)]));
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnZOUITClick(Sender: TObject);
var
  zouit, sprpath, respath, prepared, command, path, s, Title, fields: String;
  i, j, k, n, m, r, d: Integer;
  t: Cardinal;
  sl: TStringList;
  res: TList;
  pInfo, predinfo: PzouitObj;
  sa: StringArray;

  procedure SplitZOUIT;
  var
    i: Integer;
    x0, y0, x1, y1, col, row: Double;
  begin
    if FileExists(prepared + zouit + '.tab') then // есть разрезанная таблица
    begin
      s := prepared + zouit;
      CopyTable(s, respath);
      exit;
    end;

    SafeCloseTAble('Setka');
    miDo('Select * from Oblast Where rowId = -1 into Setka NoSelect');
    // для создания таблицы с CS области
    miDo('Commit table Setka As "%sSetka.tab" TYPE NATIVE Charset "WindowsCyrillic"',
      [respath]);
    miDo('Close table Setka');
    miDo('Open table "%sSetka.tab"', [respath]);

    miDo('Map From Setka Position (-11,-11) Width 5 Height 5');
    x0 := miEvalDbl('ObjectGeography(ob, %d)', [OBJ_GEO_MINX]);
    y0 := miEvalDbl('ObjectGeography(ob, %d)', [OBJ_GEO_MINY]);
    x1 := miEvalDbl('ObjectGeography(ob, %d)', [OBJ_GEO_MAXX]);
    y1 := miEvalDbl('ObjectGeography(ob, %d)', [OBJ_GEO_MAXY]);
    row := Trunc(y0 / d) * d;
    pg.Max := Trunc((x1 - x0) / d) * Trunc((y1 - y0) / d);

    Info(zouit + ' : подготовка частей...');
    While row <= y1 do // строятся квадраты
    begin
      col := Trunc(x0 / d) * d;
      While col <= x1 do
      begin
        miDo('Create Rect Into Variable ob (%.6f,%.6f)(%.6f,%.6f)',
          [col, row, col + d, row + d]);
        miDo('Insert into Setka (Obj) Values(ob)');
        col := col + d;
      end;
      row := row + d;
    end;
    miDo('Commit Table Setka');

    miDo('Map From %s Position (-11,-11) Width 5 Height 5', [zouit]);
    miDo('Set Map Window FrontWindow() Layer 1 Editable On');

    miDo('Select * from %s noSelect', [zouit]);
    miDo('Set Target On');
    miDo('Select obj from Setka noSelect');
    command := 'Objects Split Into Target Concurrency All DropIndex Auto Data ';
    // настраиваем набор полей для обрезки сеткой
    for i := 1 to miEvalInt('TableInfo(%s, %d)', [zouit, TAB_INFO_NCOLS]) do
    begin
      s := miEval('ColumnInfo(%s, "COL%d", %d)', [zouit, i, COL_INFO_NAME]);
      command := format('%s %s=%s,', [command, s, s]);
    end;
    SetLength(command, Length(command) - 1);
    Info(zouit + ' : деление на части...');
    miDo(command);

    miDo('Commit Table %s As "%s%s.TAB" TYPE NATIVE Charset "WindowsCyrillic"',
      [zouit, respath, zouit]);//сохраняем разрезанную таблицу в другом каталоге

    miDo('Commit Table %s As "%s%s.TAB" TYPE NATIVE Charset "WindowsCyrillic"',
      [zouit, prepared, zouit]);//сохраняем разрезанную таблицу в другом каталоге

    miDo('RollBack table %s', [zouit]);
    miDo('Drop table Setka');
  end;

  procedure UpdateZOUIT;
  var
    r: Integer;
    field: String[32];
  begin
    path := ResFolder + 'ЗОУИТ\';
    if FileExists(ResFolder + 'CF_ZOUIT_Communication.TAB') then
    begin
      miDo('Open table "%sCF_ZOUIT_Communication.TAB"', [ResFolder]);
      for r := 2 to miEvalInt('TableInfo(Oblast, %d)', [TAB_INFO_NROWS]) do
      begin// идем по числу записей ЗОУИТ
        miDo('Fetch rec %d From Oblast', [r]);//здесь Дополнительные_затраты
        zouit := miEval('Oblast.table');
        if FileExists(path + zouit + '.tab') then
        begin
          miDo('Open table "%s%s.tab"', [path, zouit]);
          field := zouit + '_Перекрытие';
          try
            miDo('Alter table CF_ZOUIT_Communication (add %s Float)', [field]);
          except
            Info('Поля добавлены...');
          end;
          miDo('Add Column "CF_ZOUIT_Communication" (%s) From %s Set To Площадь_перекрытия Where COL2 = COL2',
            [field, zouit]);
          // Выбраем площади перекрытий
          miDo('select * from CF_ZOUIT_Communication where %s>0 into Temp NoSelect',
            [field]);
          // площади перекрытий делим на общую площадь участка
          miDo('Update Temp Set %s = format$(%s*100/area, "#.##") DropIndex Auto',
            [field, field]);
          // могут появиться значения > 100%
          miDo('select * from CF_ZOUIT_Communication where %s>100 into Temp NoSelect',
            [field]);
          miDo('Update Temp Set %s = 100 DropIndex Auto', [field]);
          // переводим в части
          miDo('Update CF_ZOUIT_Communication Set %s = %s/100 DropIndex Auto',
            [field, field]);
          miDo('Close table Temp');
          miDo('Commit table CF_ZOUIT_Communication');
        end;
      end;
      miDo('browse * from CF_ZOUIT_Communication');
    end
    else
      Exc('Таблица CF_ZOUIT_Communication.TAB не найдена' + #10 +
        'Вероятно не сделан расчет ценовых факторов');
  end;

begin//mnZOUITClick
  if not CheckWorkTable('mnZOUITClick') then Exit;

  sprpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'ЗОУИТ\';
  if not FileExists(sprpath + 'Oblast.tab') then
    Exc('Файл ' + sprpath + 'Oblast.tab не найден.');
  miDo('Open table "%s"', [sprpath + 'Oblast.tab']);

  miDo('Open table "%s"', [ExtractFilePath(worktable) + 'Ценовые_Факторы_Итог.tab']);

  respath := ResFolder + 'ЗОУИТ\RESULT\';
  prepared := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'PREEPARED\';
  t := GetTickCount;

  s := format('/c rd /S /Q  "%s"', [respath]);
  ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
  ForceDirectories(respath);

  miDo('Set CoordSys table Oblast');
  miDo('ob = Oblast.obj');//первая запись имеет графику - это границы области
  Title := 'IDN' + #9 + 'Кадастровый_номер' + #9 + 'Площадь_участка' + #9 +
    'Площадь_перекрытия' + #9 + 'Процент_перекрытия';
  fields := 'IDN,' + fcn;
  // число записей ЗОУИТ это Обременение_целое, Дополнительные_затраты или др.
  k := miEvalInt('TableInfo(Oblast, %d)', [TAB_INFO_NROWS]);
  for r := 2 to k do
  begin
    miDo('Fetch rec %d From Oblast', [r]);
    zouit := miEval('Oblast.table');
    if not FileExists(sprpath + zouit + '.tab') then
      Exc('Таблица ' + zouit + ' не найдена.');

    SafeCloseTAble(zouit);
    miDo('Open table "%s.tab"', [sprpath + zouit]);
    d := miEvalInt('Oblast.id'); // шаг сетки
    if d <> 0 then
      SplitZOUIT
    else
      miDo('Commit Table %s As "%s%s.TAB" TYPE NATIVE Charset "WindowsCyrillic"',
        [zouit, respath, zouit]);

    Info('ZOUIT Определение пересечений...');
    j := CalcCPUCount; // максимальное число частей таблицы участков
    m := miEvalInt('TableInfo(%s, %d)', [ftab, TAB_INFO_NROWS]);//размер таблицы участков
    if m < 1000 then // расчитываем все в одной части (было 5000)
    begin
      j := 1;
      m := m + 1;
    end
    else
      m := m div 4;

    row := 0;
    cpu := 0;
    for i := 1 to j do // разбиваем на части по количеству cpu
    begin
      path := respath + IntToStr(i) + '\';
      ForceDirectories(path);
      miDo('Select * from %s where RowId >= %d AND RowID < %d AND Str$(obj) = "Region" into work noSelect',
        [ftab, row, m * i]);

      miDo('Commit table "work" As "%s%s.tab"', [path, ftab]);
      miDo('Close table work');
      s := respath + zouit;
      CopyTable(s, path);
      // определяем процент площаль и перекрытия с каждой зоной
      s := path + zouit + '.TAB';
      ExCreateProcess(format('"%sCalcSplit.exe" "%s" "%s" "%s" "%s" %d 1',// пустой КН
        [GetAppFolder, path + ftab + '.tab', s, fields, '-', Handle]), '',
        SW_MINIMIZE, false);

      Inc(cpu);
      Inc(row, m);
      if i = j - 1 then
        Inc(m, m); // увеличим, чтобы гарантированно захватить конец таблицы
    end;

    Repeat
      Sleep(1000);
      Application.ProcessMessages;
    Until cpu = 0;

    cn.Clear;
    FindRecursive(respath, '*.txt', true, cn); // собираем все текстовые файлы
    sl := TStringList.Create; // для загрузки
    res := TList.Create;
    try
      SafeCloseTable(zouit);
      for i := 0 to cn.Count - 1 do
      begin
        sl.LoadFromFile(cn[i]);
        for j := 0 to sl.Count - 1 do
        begin
          pInfo := AllocMem(SizeOf(pInfo^));
          res.Add(pInfo);
          sa := SplitStr(sl[j], #9);
          Val(sa[0], pInfo.idn, m);
          pInfo.cn := sa[1];
          Val(sa[2], pInfo.area, m);
          Val(sa[3], pInfo.overlap, m);
          Val(sa[4], pInfo.proc, m);
          SetLength(sa, 0);
        end;
      end;

      if res.Count > 0 then
      begin
        res.Sort(@SortInfo1);
        pg.Max := res.Count;
        Info('Объединение кадастровых номеров');
        predinfo := res[0];
        for i := 1 to res.Count - 1 do
        begin
          if i mod 100 = 0 then
            pg.Position := i;
          pInfo := res[i];
          if predinfo.cn = pInfo.cn then
          begin
            predinfo.overlap := predinfo.overlap + pInfo.overlap;
            // накапливаем перекрытия
            if predinfo.idn <> pInfo.idn then // другая часть контура
            begin
              predinfo.area := predinfo.area + pInfo.area;
              predinfo.idn := pInfo.idn;
            end;
            predinfo.proc := -2; // признак перевычисления процента
            pInfo.proc := -1; // признак удаления
          end
          else
            predinfo := pInfo;
        end;

        for i := res.Count - 1 downto 0 do // удаляем дубли
        begin
          pInfo := res[i];
          if pInfo.proc = -1 then
          begin
            Dispose(res.Items[i]);
            res.Delete(i);
          end;
        end;
        pg.Max := 0;

        pg.Max := res.Count;
        Info('Расчет площадей');
        for i := 0 to res.Count - 1 do
        begin
          if i mod 100 = 0 then
            pg.Position := i;
          pInfo := res[i];
          if pInfo.proc = -2 then // измененные значения исправляем
          begin
            if pInfo.overlap > pInfo.area then
              pInfo.overlap := pInfo.area;
            pInfo.proc := pInfo.overlap / pInfo.area * 100;//перевод в проценты
          end;
        end;
        pg.Max := 0; // sum_square

        sl.Clear;
        for i := 0 to res.Count - 1 do
        begin
          pInfo := res[i];
          sl.Add(format('%d%s%s%s%.4f%s%.4f%s%.4f', [pInfo.idn, #9, pInfo.cn,
            #9, pInfo.area, #9, pInfo.overlap, #9, pInfo.proc]));
        end;
        sl.Insert(0, Title);
        path := ResFolder + 'ЗОУИТ\';
        sl.SaveToFile(format('%s%s.txt', [path, zouit]));

        Info('Регистрация временного файла...');
        miDo('Register Table "%s%s.txt" TYPE ASCII Delimiter 9 ' +
          'Titles Charset "WindowsCyrillic" Into "%s%s.TAB"',
          [path, zouit, TempPath, zouit]);
        miDo('Open Table "%s%s.TAB"', [TempPath, zouit]);
        miDo('Commit Table %s As "%s%s.TAB"', [zouit, path, zouit]);
        miDo('Close Table %s', [zouit]);
        miDo('Open Table "%s%s.TAB"', [path, zouit]);

        FillEmptyFields(zouit);
        miDo('browse * from %s', [zouit]);
      end
      else // если нет результатов то создадим пустой файл
      begin
        path := ResFolder + 'ЗОУИТ\';
        miDo('Create Table "%s" (IDN Integer,Кадастровый_номер Char(35),' +
          'Площадь_участка Float,Площадь_перекрытия Float,Процент_перекрытия Float) '
          + 'file "%s%s.tab" TYPE NATIVE Charset "WindowsCyrillic"',
          [zouit, path, zouit]);
        miDo('Open Table "%s%s.TAB"', [path, zouit]);
        miDo('Insert Into %s (IDN,Кадастровый_номер) Select IDN,Кадастровый_номер From %s',
          [zouit, ftab]);
        miDo('Commit Table %s', [zouit]);
      end;

    finally
      for i := 0 to res.Count - 1 do
        Dispose(res.Items[i]);
      res.Free;
      sl.Free;
    end;
  end;

  if st16 then
    UpdateZOUIT;
  miDo('Close table Oblast');

  s := format('/c rd /S /Q  "%s"', [respath]);
  ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога с промежуточными результатами
  Info('Готово.');
  pg.Max := 0;
  if Sender <> nil then
    Note(format('процес завершен.%sВремя обработки %d минут',
      [#10, Round((GetTickCount - t) / 60000)]));
end;

{------------------------------------------------------------------------------}
procedure TForm2.mnSoilClick(Sender: TObject);
var
  i, j, m: Integer;
  sprpath, respath, path, soil, f1, f2, f3, s, fields, Title, dir: String;
  pInfo, predinfo: PzouitObj;
  t: Cardinal;
  sl: TStringList;
  res: TList;
  sa: StringArray;

  procedure ToXLS;
  begin
    SafeCloseTAble('Soil');
    miDo('Open Table "%sЦеновые_Факторы_Итог"', [ResFolder]);

    miDo('Select Ценовые_Факторы_Итог.Кадастровый_номер,Ценовые_Факторы_Итог.CAD_NUM,'+
      'Ценовые_Факторы_Итог.Area,Ценовые_Факторы_Итог.Area,'+
      'Ценовые_Факторы_Итог.Area,Ценовые_Факторы_Итог.Area,Ценовые_Факторы_Итог.число_участков,'+
//      'Агроклиматические_Зоны.Площадь_участка,Агроклиматические_Зоны.Площадь_участка,Ценовые_Факторы_Итог.число_участков,'+
      'Агроклиматические_Зоны.Площадь_перекрытия,Агроклиматические_Зоны.Площадь_перекрытия,'+
      'Ценовые_Факторы_Итог.Код_населенного_пункта,Ценовые_Факторы_Итог.Код_Муниципального_образования,'+
      'Агроклиматические_Зоны.SOIL0_TXT,Агроклиматические_Зоны.Номер_Подзоны,Агроклиматические_Зоны.Почва_и_АП '+
      'from Ценовые_Факторы_Итог,Агроклиматические_Зоны where Ценовые_Факторы_Итог.Кадастровый_номер='+
      'Агроклиматические_Зоны.Кадастровый_номер into Soil noSelect');
    miDo('Commit table Soil as "%sSoil.tab"', [path]);
    miDo('Close table Soil');
    miDo('Open table "%sSoil.tab"', [path]);

    miDo('Alter table Soil (Rename CAD_NUM Кадастровый_квартал,Area Площадь_ЕГРН_кв_м,'+
      'Area_2 Площадь_ЕГРН_га,Area_3 Площадь_по_графике_кв_м,Area_4 Площадь_по_графике_га,'+
//      'Area_2 Площадь_ЕГРН_га,Площадь_участка Площадь_по_графике_кв_м,Площадь_участка_2 Площадь_по_графике_га,'+
      'число_участков Кол_во_контуров,Площадь_перекрытия Площадь_части_КН_кв_м,'+
      'Площадь_перекрытия_2 Площадь_части_КН_га,SOIL0_TXT Почвенная_разновидность,'+
      'Номер_Подзоны Агроклиматическая_подзона)');

    miDo('Update Soil Set Площадь_ЕГРН_кв_м=0,Площадь_ЕГРН_га=0,'+
      'Площадь_по_графике_га=Format$(Площадь_по_графике_кв_м/10000, "#.####"),'+
      'Площадь_части_КН_га=Format$(Площадь_части_КН_кв_м/10000, "#.####")');

    miDo('Commit table Soil');
    Info('Создание файла Excel...');
    if st16 then
      Tab2Excel('Soil', '', nil);
//    miDo('Close table Soil');
  end;

begin//mnSoilClick
  if not CheckWorkTable('mnSoilClick') then Exit;

  miDo('Open table "%s"', [ExtractFilePath(worktable) + 'Ценовые_Факторы_Итог.tab']);

  sprpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) +
    'Агроклиматические_Зоны\';
  respath := ResFolder + 'Агроклиматические_Зоны\RESULT\';
  ForceDirectories(respath);

  cn.Clear;
  FindRecursive(sprpath, '*.tab', false, cn);
  if cn.Count = 0 then
    Exc('Нет таблиц в каталоге ' + sprpath);
  if cn.Count > 1 then
    Exc('В каталоге ' + sprpath + 'должна быть 1 таблица');

  soil := miEval('PathToTableName$("%s")', [cn[0]]);
  miDo('Open table "%s"', [cn[0]]);
  res := TList.Create;
  sl := getFieldsList(soil); // Нужны поля soil0_txt, Номер_подзоны, Почва_и_АП
  f1 := '';  f2 := '';  f3 := '';
  try
    for i := 0 to sl.Count - 1 do
    begin
      if CText('soil0_txt', sl[i]) then
        f1 := sl[i];
      if CText('Номер_подзоны', sl[i]) then
        f2 := sl[i];
      if CText('Почва_и_АП', sl[i]) then
        f3 := sl[i];
    end;
    if f1 = '' then
      Exc('Поле soil0_txt не найдено в таблице ' + soil);
    if f2 = '' then
      Exc('Поле Номер_подзоны не найдено в таблице ' + soil);
    if f3 = '' then
      Exc('Поле Почва_и_АП не найдено в таблице ' + soil);

    fields := f1 + ',' + f2 + ',' + f3;
    Title := 'IDN' + #9 + 'Кадастровый_номер' + #9 + f1 + #9 + f2 + #9 + f3 + #9 +
      'Площадь_участка' + #9 + 'Площадь_перекрытия' + #9 + 'Процент_перекрытия';

    t := GetTickCount;
    Info('Почвы Определение пересечений...');
    j := CalcCPUCount;
    m := miEvalInt('TableInfo(%s, %d)', [ftab, TAB_INFO_NROWS]);//размер таблицы участков
    if m < 1000 then // расчитываем все в одной части (было 5000)
    begin
      j := 1;
      m := m + 1;
    end
    else
      m := m div 4;

    row := 0;
    cpu := 0;
    for i := 1 to j do // разбиваем на части по количеству cpu
    begin
      path := respath + IntToStr(i) + '\';
      ForceDirectories(path);
      miDo('Select * from %s where RowId >= %d AND RowID < %d AND Str$(obj) = "Region" into work noSelect',
        [ftab, row, m * i]);
      miDo('Commit table "work" As "%s%s.tab"', [path, ftab]);
      miDo('Close table work');
      s := Copy(cn[0], 1, Length(cn[0]) - 4);
      CopyTable(s, path);
      // определяем процент площадь и перекрытия с каждой зоной
      s := path + ExtractFileName(cn[0]);
      ExCreateProcess(format('"%sCalcSplit.exe" "%s" "%s" "%s" "%s" %d 2',
        [GetAppFolder, path + ftab + '.tab', s, fields, fcn, Handle]), '',
        SW_MINIMIZE, false);

      Inc(cpu);
      Inc(row, m);
      if i = j - 1 then
        Inc(m, m); // увеличим, чтобы гарантированно захватить конец таблицы
    end;

    Repeat
      Sleep(1000);
      Application.ProcessMessages;
    Until cpu = 0;

    SafeCloseTable('Агроклиматические_Зоны');
    cn.Clear;
    FindRecursive(respath, '*.txt', true, cn);
    for i := 0 to cn.Count - 1 do
    begin
      sl.LoadFromFile(cn[i]);
      for j := 0 to sl.Count - 1 do
      begin
        pInfo := AllocMem(SizeOf(pInfo^));
        res.Add(pInfo);
        sa := SplitStr(sl[j], #9);
        Val(sa[0], pInfo.idn, m);
        pInfo.cn := sa[1];
        pInfo.zone := sa[2];
        Val(sa[3], pInfo.subzone, m);
        pInfo.soil_ap := sa[4];
        Val(sa[5], pInfo.area, m);
        Val(sa[6], pInfo.overlap, m);
        Val(sa[7], pInfo.proc, m);
        SetLength(sa, 0);
      end;
    end;

    if res.Count > 0 then
    begin
      res.Sort(@SortInfo2);
      pg.Max := res.Count;
      Info('Объединение кадастровых номеров');
      predinfo := res[0];
      for i := 1 to res.Count - 1 do
      begin
        if i mod 100 = 0 then
          pg.Position := i;
        pInfo := res[i];
        if (predinfo.cn = pInfo.cn) and (predinfo.zone = pInfo.zone) and
          (predinfo.subzone = pInfo.subzone) then
        begin
          predinfo.area := predinfo.area + pInfo.area; // накапливаем площади
          predinfo.overlap := predinfo.overlap + pInfo.overlap;//накапливаем перекрытия

          predinfo.proc := -2; // признак перевычисления процента
          pInfo.proc := -1; // признак удаления
        end else
          predinfo := pInfo;
      end;

      for i := res.Count - 1 downto 0 do // удаляем дубли
      begin
        pInfo := res[i];
        if pInfo.proc = -1 then
        begin
          Dispose(res.Items[i]);
          res.Delete(i);
        end;
      end;
      pg.Max := 0;

      pg.Max := res.Count;
      Info('Расчет площадей');
      for i := 0 to res.Count - 1 do
      begin
        if i mod 100 = 0 then
          pg.Position := i;
        pInfo := res[i];
        if pInfo.proc = -2 then
        begin
          if pInfo.overlap > pInfo.area then
            pInfo.overlap := pInfo.area;
          pInfo.proc := pInfo.overlap / pInfo.area * 100;//определяется в процентах
        end;
      end;

      for i := 0 to res.Count - 1 do
      begin
        pInfo := res[i];
        pInfo.proc := pInfo.proc / 100; // переводим в части все расчеты
      end;
      pg.Max := 0; // sum_square

      sl.Clear;
      for i := 0 to res.Count - 1 do
      begin
        pInfo := res[i];
        sl.Add(format('%d%s%s%s%s%s%d%s%s%s%.4f%s%.4f%s%.4f',
          [pInfo.idn, #9, pInfo.cn, #9, pInfo.zone, #9, pInfo.subzone, #9,
          pInfo.soil_ap, #9, pInfo.area, #9, pInfo.overlap, #9, pInfo.proc]));
      end;

      sl.Insert(0, Title);
      path := ResFolder + 'Агроклиматические_Зоны\';
      sl.SaveToFile(format('%sАгроклиматические_Зоны.txt', [path]));

      Info('Регистрация временного файла...');
      miDo('Register Table "%sАгроклиматические_Зоны.txt" TYPE ASCII Delimiter 9 '
        + 'Titles Charset "WindowsCyrillic" Into "%sАгроклиматические_Зоны.TAB"',
        [path, TempPath]);
      miDo('Open Table "%sАгроклиматические_Зоны.TAB"', [TempPath]);
      miDo('Commit Table Агроклиматические_Зоны As "%sАгроклиматические_Зоны.TAB"',
        [path]);
      miDo('Close Table Агроклиматические_Зоны');
      miDo('Open Table "%sАгроклиматические_Зоны.TAB"', [path]);
      FillEmptyFields('Агроклиматические_Зоны');
      miDo('browse * from Агроклиматические_Зоны');

      ToXLS;
    end else // создаем пустую таблицу и заполняем все нулями
    begin
      Info('Нет пересечений с Агроклиматическими зонами');
      path := ResFolder + 'Агроклиматические_Зоны\';
      miDo('Create Table "Агроклиматические_Зоны" (IDN Integer,Кадастровый_номер Char(35),'
        + 'SOIL0_TXT Char(75), Номер_Подзоны Integer, Почва_и_АП Char(81),'
        + 'Площадь_участка Float,Площадь_перекрытия Float,Процент_перекрытия Float) '
        + 'file "%sАгроклиматические_Зоны.tab" TYPE NATIVE Charset "WindowsCyrillic"',
        [path]);
      miDo('Open Table "%sАгроклиматические_Зоны.TAB"', [path]);
      miDo('Insert Into Агроклиматические_Зоны (IDN,Кадастровый_номер) Select IDN,Кадастровый_номер From %s',
        [ftab]);
      miDo('Commit Table Агроклиматические_Зоны');
    end;

  //в CF_ZOUIT_Communication добавим поле Почва_и_АП и данные из Агроклиматические_Зоны
  if st16 and FileExists(ResFolder + 'CF_ZOUIT_Communication.TAB') then
  begin
    miDo('Open table "%sCF_ZOUIT_Communication.TAB"', [ResFolder]);
    sl.Clear;
    m := -1;
    sl := getFieldsList('CF_ZOUIT_Communication');
    for i := 1 to sl.Count - 1 do //проверяем наличие поля Почва_и_АП
      if CText(sl[i], 'Почва_и_АП') then m := i;
    if m = -1 then//добавляем поле Почва_и_АП
      miDo('Alter table CF_ZOUIT_Communication (add Почва_и_АП Char(81))');
    miDo('select * from Агроклиматические_Зоны where Процент_перекрытия>=0.5 into temp noSelect');
    //добавляем данные из Агроклиматические_Зоны
    miDo('Add Column "CF_ZOUIT_Communication" (Почва_и_АП )From temp Set To Почва_и_АП Where COL2=COL2');
    miDo('Commit table CF_ZOUIT_Communication');

    miDo('select * from CF_ZOUIT_Communication where Почва_и_АП="" into temp noSelect');
    miDo('Update temp Set Почва_и_АП = "Нет пересечений" DropIndex Auto');
    miDo('Commit table CF_ZOUIT_Communication');
    miDo('Close table temp');
//miDo('run menu command 311');//инвертируем выборку
  end;

  finally
    for i := 0 to res.Count - 1 do
      Dispose(res.Items[i]);
    res.Free;
    sl.Free;
    pg.Max := 0;
  end;
  s := format('/c rd /S /Q  "%s"', [respath]);
  ExShellExecute('cmd.exe', s, SW_HIDE, true, 0);//удалене каталога со всем содержимым и подкаталогами
  Info('Готово.');
  pg.Max := 0;
  if Sender <> nil then
    Note(format('процес завершен.%sВремя обработки %d минут',
      [#10, Round((GetTickCount - t) / 60000)]));
end;
{--{сoздание таблицы Ценовые_Факторы_Итог и копию ее в CF_ZOUIT_Communication--}
//многоконтурные объекты должны быть расчитаны!!!
procedure TForm2.mnItogCFClick(Sender: TObject);
begin
  if not CheckWorkTable('mnItogCFClick') then Exit;

  Info('Создание итога ЦФ...');
  SafeCloseTable('Ценовые_Факторы_Итог');
  // выберем уникальные кадастровые номера, выбирая последнюю запись в многоконтурных объектах
  //одновременно суммируем площади с одинаковыми КН
  miDo('Select idn, Кадастровый_номер, Count(*), max(idn), sum(Area) From %s Group'+
    ' By Кадастровый_номер into Ценовые_Факторы_Итог NoSelect', [ftab]);
  miDo('Commit table Ценовые_Факторы_Итог as "%sЦеновые_Факторы_Итог.tab"',
    [TempPath]);//сохраним во временном каталоге
  miDo('Close table Ценовые_Факторы_Итог');
  miDo('open table "%sЦеновые_Факторы_Итог.tab"', [TempPath]);
  miDo('Select * from %s, Ценовые_Факторы_Итог where %s.IDN=Ценовые_Факторы_Итог._COL4 into Temp NoSelect',
    [ftab, ftab]);
  miDo('Commit Table Temp As "%sЦеновые_Факторы_Итог.TAB" TYPE NATIVE Charset "WindowsCyrillic"',
    [ResFolder]);//сохраним в основном каталоге
  miDo('Close table Ценовые_Факторы_Итог');
  miDo('open table "%sЦеновые_Факторы_Итог.tab"', [ResFolder]);
  miDo('Update Ценовые_Факторы_Итог Set Area = _COL5 DropIndex Auto');
  miDo('Commit Table Ценовые_Факторы_Итог');
  //удалим колонки Count(*), max(idn), sum(Area)
  miDo('Alter Table "Ценовые_Факторы_Итог" (drop IDN_2,Кадастровый_номер_2,_COL3,_COL4,_COL5)');
  miDo('Create Index On Ценовые_Факторы_Итог (Кадастровый_номер)');
  miDo('browse * from Ценовые_Факторы_Итог');
  Info(ftab + ': Готова к работе...');
  // подготовим копию таблицы для добавления ЗОУИТ и коммуникаций
  //в эту таблицу будут добавлены колонки с информацией по коммуникациям и зонам др.
  SafeCloseTable('CF_ZOUIT_Communication');
  if st16 then
    ExShellExecute('cmd', PChar(format('/c copy /b /Y "%s.*" "%s.*"',
      [ResFolder+'Ценовые_Факторы_Итог', ResFolder+'CF_ZOUIT_Communication'])), SW_HIDE, false, 0);
end;
{ ---------совместный расчет ЗОУИТ и почв------------------------------------- }
procedure TForm2.mnComplexOtherClick(Sender: TObject);
var
  sprpath: String;
  t: Cardinal;
begin
  if not CheckWorkTable('mnComplexOtherClick') then Exit;

  miDo('Open table "%s"', [ExtractFilePath(worktable) + 'Ценовые_Факторы_Итог.tab']);

  t := GetTickCount;
  sprpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'ЗОУИТ\';
  if not FileExists(sprpath + 'Oblast.tab') then
    Exc('Файл ' + sprpath + 'Oblast.tab не найден.');

  mnZOUITClick(nil);
  mnSoilClick(nil);
  Note(format('процес завершен.%sВремя обработки %d минут',
    [#10, Round((GetTickCount - t) / 60000)]));
end;
{------------создание загрузочного файла из каталога с расчетными данными------}
procedure TForm2.CreateLoaderFile;
var
  i, r, c, row, col, k, startrow, part : Integer;
  sl, cn, st : TStringList;
  fields, path, s, colname : String;
  sa: StringArray;
  ExcelApp, Workbook, WorkSheet, range, DataParcels{, data} : OleVariant;
begin
  if LoadFields.Count > 0 then
  begin
    if not FileExists(ResFolder + 'CF_ZOUIT_Communication.TAB') then
      Exc('Нет таблицы CF_ZOUIT_Communication.TAB!');

    miDo('Open table "%sCF_ZOUIT_Communication.TAB"', [ResFolder]);
    sl := getFieldsList('CF_ZOUIT_Communication');
    cn := TStringList.Create;
    st := TStringList.Create;//колонки для загрузки
    try
      for i := 0 to LoadFields.Count - 1 do
        if sl.IndexOf(LoadFields.names[i]) > -1 then//сравниваем загруку с полями таблицы
        begin
          if st16 then
            st.Add(LoadFields.names[i])
          else
          begin
            sa := SplitStr(LoadFields.Values[LoadFields.names[i]], ',');
            if sa[0] ='1' then
              st.Add(LoadFields.names[i]);
            SetLength(sa, 0);
          end;
        end;
      if st.Count = 0 then
        Exc('Поля для загрузки не определены!');
      fields := '';
      for i := 0 to st.Count-1 do
        fields := fields + st[i] + ',';

      row := miEvalInt('TableInfo(CF_ZOUIT_Communication, %d)', [TAB_INFO_NROWS]);
      part := 0;//расчитываем кол-во частей исходя из размера 750000 записей
      while row > 0 do
      begin
        Inc(part);
        row := row - 750000;
      end;

      startrow := 1;
      miDo('drop index CF_ZOUIT_Communication (Кадастровый_номер)');
      for k := 1 to part do
      begin
        SafeCloseTAble('Загрузка');
        miDo('Select %s from CF_ZOUIT_Communication into Загрузка where '+
          'rowid>=%d and rowid < %d noSelect',
          [Copy(fields, 1, Length(fields) - 1), startrow, startrow+750000]);
        Inc(startrow, 750000);
        miDo('Commit table Загрузка as "%sЗагрузка.tab"', [TempPath]);
        miDo('Close table Загрузка');
        miDo('Open table "%sЗагрузка.tab"', [TempPath]);

        //проверяем записи без графики
        if (k = 1) and st16 then
        begin
          SafeCloseTable('TableLessObjects');
          if FileExists(AddSlesh(ResFolder) + 'TableLessObjects.tab') then
          begin
            miDo('Open table "%sTableLessObjects.tab"', [ResFolder]);
            While miEval('EOT(TableLessObjects)') <> 'T' do
            begin
              cn.Add(miEval('TableLessObjects.Col1'));//кадастровые номера в список
              miDo('Fetch Next From TableLessObjects');
            end;
            if miEvalInt('TableInfo(TableLessObjects, %d)', [TAB_INFO_NROWS]) = 0 then
              miDo('Drop table TableLessObjects')
            else
              miDo('Close table TableLessObjects');
          end;
        end;

        Info(' Заполнение Excel...');
        row := miEvalInt('TableInfo(Загрузка, %d)', [TAB_INFO_NROWS]);
        col := st.Count;
        ExcelApp := CreateOleObject('Excel.Application');
        try
          ExcelApp.Application.EnableEvents := false;
          Workbook := ExcelApp.Workbooks.Add; // Создаем Книгу (Workbook)
          WorkSheet := Workbook.WorkSheets.item[1];
          WorkSheet.Name := 'Загрузка';
          ///первые 3 строки служебные, первый столбец пустой
          DataParcels := VarArrayCreate([1, row, 1, 1], varVariant);// 1 колонка
          try
            WorkSheet.Cells[1, 1] := '&T0';;
            WorkSheet.Cells[2, 1] := '&H1';;
            WorkSheet.Cells[3, 1] := '&H2';;
           for i := 0 to st.Count - 1 do //заполняем названия столбцов
            begin
              sa := SplitStr(LoadFields.Values[st[i]], ',');
              WorkSheet.Cells[2, i+2] := sa[1];
              WorkSheet.Cells[3, i+2] := sa[2];
              SetLength(sa, 0);
            end;

            pg.Max := st.Count;
            for i := 0 to st.Count-1 do
            begin//вставка колонок MI в Excel
              pg.Position := i+1;
              miDo('Select %s from Загрузка into Temp noSelect', [st[i]]);
              miDo('Export "Temp" Into "%sTemp.txt" Type "ASCII" Overwrite CharSet "WindowsCyrillic"', [TempPath]);
              sl.LoadFromFile(TempPath + 'Temp.txt');
              for c := 0 to sl.Count-1 do
              begin
                if sl[c][1] = '"' then
                  sl[c] := Copy(sl[c], 2, Length(sl[c]));
                if sl[c][Length(sl[c])] = '"' then
                  sl[c] := Copy(sl[c], 1, Length(sl[c])-1);
                if st[i] = 'КодМО' then//дополним код МО литерой "МО" then
                  sl[c] := 'МО' + sl[c];
                DataParcels[c+1, 1] := sl[c];
              end;
              s := ExcelNum2Str(WorkSheet, i+2);//вставка начинается с колонки B
              range := WorkSheet.range[format('%s%d:%s%d', [s, 4, s, row+3])];
              range.value := DataParcels;
            end;
            pg.Max := 0;

          finally
            VarClear(DataParcels);
          end;

          Info(' дополнительные поля...');
          if st16 and (AddFields.Count > 0) then
          begin
            DataParcels := VarArrayCreate([1, row+3+cn.Count, 1, AddFields.Count], varVariant);
            for i := 0 to AddFields.Count - 1 do // заполняем верхние 3 строки
            begin
              sa := SplitStr(AddFields.Values[AddFields.names[i]], ',');
              DataParcels[3, i+1] := AddFields.names[i];//название
              DataParcels[2, i+1] := sa[0];//additional_xxx
              for r := 1 to row+cn.Count do
                DataParcels[r+3, i+1] := sa[1];
              SetLength(sa, 0);
            end;
            range := WorkSheet.range[format('%s1:%s%d', [ExcelNum2Str(WorkSheet, col+2),
              ExcelNum2Str(WorkSheet, col+1+AddFields.Count), row+3+cn.Count])];
            range.value := DataParcels;
            VarClear(DataParcels);
            for i := 0 to cn.Count-1 do //вставим кадстровые номера
              WorkSheet.Cells[row+4+i, 2] := cn[i];
          end;

          Info(' заполнение полей без графики...');
          if st16 and (PatternFields.Count > 0) then
          begin
            for i := 3 to col+1 do
            begin
              colname := String(WorkSheet.Cells[2, i]);
              for c := 0 to PatternFields.Count-1 do
              begin
                if CText(colname, PatternFields.names[c]) then
                  for r := 0 to cn.Count-1 do //вставим значениия по шаблону
                    WorkSheet.Cells[row+4+r, i] := PatternFields.Values[PatternFields.names[c]];
              end;
            end;
          end;

          Info(' вставка шаблонов...');
          if st16 then
          begin
            for i := 0 to cn.Count-1 do //вставим кадстровые номера
              WorkSheet.Cells[row+4+i, 2] := cn[i];
            for i := 3 to col+1 do
            begin
              colname := String(WorkSheet.Cells[2, i]);
              for c := 0 to PatternFields.Count-1 do
              begin
                if CText(colname, PatternFields.names[c]) then
                  for r := 0 to cn.Count-1 do //вставим значениия по шаблону
                    WorkSheet.Cells[row+4+r, i] := PatternFields.Values[PatternFields.names[c]];
              end;
            end;
          end;

          ExcelApp.ActiveSheet.PageSetup.Orientation := 2;//ориентация-альбомная
          ExcelApp.Visible := true;
        finally
          ExcelApp := Unassigned;
          Workbook := Unassigned;
          WorkSheet := Unassigned;
          range := Unassigned;
        end;
      end;

      SafeCloseTable('Загрузка');
      Info(' Готово');
    finally
      sl.Free;
      cn.Free;
      st.Free;
    end;
  end;
//  else if Sender <> nil then
//    Note('Нет списка полей в файле CostFactory.xlsx! + #10 + Невозможно создат файл загрузки.');
end;
{------------------------------------------------------------------------------}
procedure TForm2.FillEmptyFields(table: String);
var
  sl : TStringList;
  i, ob, Size{, j, k, c} : Integer;
begin
  sl := getFieldsList(table);
  Info(table + ' Заполнение пустых полей...');
  pg.Max := sl.Count;
  for i := 0 to sl.Count - 1 do
  begin
    pg.Position := pg.Position + 1;
    ob := Abs(Integer(sl.Objects[i]));
    if (ob > 1000) and (ob < 2000) then // текстовое поле
    begin
      Size := ob - 1000;
      if Size < 23 then
      begin
        miDo('Commit table %s', [table]);
        miDo('Alter table %s (Modify %s Char(23))', [table, sl[i]]);
      end;
      miDo('Select %s from %s Where %s = "" Into Temp NoSelect',
        [sl[i], table, sl[i]]);
      miDo('Update temp Set %s="Отсутствует информация"', [sl[i]]);
      miDo('Close table temp');
    end;
  end;
  pg.Max := 0;
  miDo('Commit table %s', [table]);
  sl.Free;
end;
{------------------------------------------------------------------------------}
function TForm2.CheckSprPath : Boolean;
var
  sprpath : String;
begin
  Result := false;
  sprpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'ЗОУИТ\';
  if not FileExists(sprpath + 'Oblast.tab') then
    Exc('файл ' + sprpath + 'Oblast.tab не найден.');
  sprpath := GetAppFolder + AddSlesh(stBar.Panels[0].Text) + 'коммуникации\';
  if not FileExists(sprpath + 'Градации_коммуникаций.xlsx') then
    Exc('файл ' + sprpath + 'Градации_коммуникаций.xlsx не найден');
  if not FileExists(sprpath + 'Коммуникации_свод.accdb') then
    Exc('файл ' + sprpath + 'Коммуникации_свод.accdb не найден');
  if not FileExists(GetAppFolder + 'Base\Дежурная_карта.tab') then
    Exc(GetAppFolder + 'Base\Дежурная_карта.tab не найдена');
  Result := true;
end;
{------------запуск программы создания файла Для_экспорта_SHP.tab--------------}
procedure TForm2.RunMapBasic(dir : String);
begin
  if FileExists(GetAppFolder + 'Для_SHP.mbx') then
  begin
    Info(' Работает MapBasic...');
    ForceDirectories(dir);
    miDo('Run Application "%s"',[GetAppFolder + 'Для_SHP.mbx']);
    Application.ProcessMessages;
    Info(mi.MBApplications.Item('Для_SHP').Eval(format('Path:%s', [dir])));
    Application.ProcessMessages;
  end;
end;
{ ------------------выбор объектов по кадастровым номерам----------------- }
 //рабочий каталог MapInfo  C:\Users\TroitskikhVV\AppData\Local\Temp\MIPRO
procedure TForm2.mnObjectsFromTextClick(Sender: TObject);
var
  t : Cardinal;
  dir : String;
begin
  if not CheckSprPath then exit;
//{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\TEST'; {$ENDIF} //двинин.txt
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST'; {$ENDIF} //WA.txt
  OpenDialog1.FileName := '*.txt';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
  dir := AddSlesh(OpenDialog1.InitialDir) + AddSlesh(curdata);//исходный текстовый файл + curdata

  t := GetTickCount;
  if not CF_from_CN then
    Exc('Ошибка в расчетах.');

  Compare_13_segment;
  RunMapBasic(dir);

  SetCurrentDir(TempPath);//входим в каталог
  ExShellExecute('cmd.exe', format('/c rd /s/q  "%s"', [TempPath]), SW_HIDE, false, 0);//чистим каталог

  Note(format('процес завершен.%sВремя обработки %d минут',
    [#10, Round((GetTickCount - t) / 60000)]));
end;
{------------------------------------------------------------------------------}
function TForm2.Xls2Txt : Boolean;
var
  ExcelApp, WorkSheet, range, DataParcels : OleVariant;
  row, col, col2, row2, i : Integer;
  sl : TStringList;
  dir : String;
begin
  Info(' Создание таблицы из К.Н.');
  Result := False;
  curdata := CurrentData;
  sl := TStringList.Create;
  ExcelApp := CreateOleObject('Excel.Application');
  try
    ExcelApp.Workbooks.Open(OpenDialog1.FileName);
    WorkSheet := ExcelApp.ActiveWorkbook.WorkSheets.item[1];//активный лист 1

    if not FindCadNum(WorkSheet, row, col, row2, col2) then Exit;

    range := WorkSheet.range[format('%s%d:%s%d',
      [ExcelNum2Str(WorkSheet, col), row, ExcelNum2Str(WorkSheet, col2), row2])];
    DataParcels := range.value;
    if (row = row2) and (col = col2) then//только 1 ячейка
      sl.Add(WorkSheet.Cells[row, col])
    else
    if col2 > col then
      for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
        sl.Add(String(DataParcels[i,1]) + #9 + String(DataParcels[i,2]))
    else
      for i := VarArrayLowBound(DataParcels,1) to VarArrayHighBound(DataParcels,1) do
        sl.Add(String(DataParcels[i,1]));
    dir := ChangeFileExt(OpenDialog1.FileName, '.txt');
    sl.SaveToFile(dir);
    OpenDialog1.FileName := dir;
    OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);
    Result := true;
  finally
    sl.Free;
    VarClear(DataParcels);
    ExcelApp.Workbooks.Close;
    ExcelApp.Quit;
    ExcelApp := Unassigned;
    WorkSheet := Unassigned;
    range := Unassigned;
  end;
end;
procedure TForm2.XML1Click(Sender: TObject);
var
  i, j, n, m, part : Integer;
  sourcepath, s, sub : String;
  sl, xml, res : TStringList;
  mc, mc2 : TMatchCollection;
begin
  sourcepath := GetAppFolder;
  {$IFDEF DEBUG}sourcepath := 'D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\от_Шмелева_КН и стоимость\';{$ENDIF}
  if not GetFolder(sourcepath, Handle) then
    exit;
  sl := TStringList.Create;
  xml := TStringList.Create;
  res := TStringList.Create;
  try
    FindRecursive(sourcepath, '*.xml', false, sl);
    pg.Max := sl.Count;
    part := 0;
    for i := 0 to sl.Count-1 do
    begin
      pg.Position := i;
      xml.LoadFromFile(sl[i]);
      s := xml.Text;
      mc := reg.Matches(s, '\d\d:\d\d:\d{1,7}:\d{1,}', [roMultiLine]);
      if mc.Count > 0 then
      begin
        Inc(part, mc.Count);
        mc2 := reg.Matches(s, '<CadastralCost Value="', [roMultiLine]);
        if mc2.Count = mc.Count then
        begin
          for j := 0 to mc.Count-1 do
          begin
            sub := Copy(s, mc2.item[j].Index + mc2.item[j].Length, 20);
            m := Pos('"', sub);
            sub := Copy(sub, 1, m-1);
            res.Add(format('%s;%s', [mc.Item[j].Value, sub]));
          end;
        end else
          res.Add(format('mc.Count %d;mc2.Count %d', [mc.Count, mc2.Count]));
        if part >= 500000 then
        begin
          res.SaveToFile(sourcepath + 'KN_Cost' + InttoStr(i) + '.csv');
          res.Clear;
          part := 0;
        end;
      end;
    end;
    res.SaveToFile(sourcepath + 'KN_Cost' + InttoStr(i) + '.csv');
  finally
    pg.Max := 0;
    sl.Free;
    xml.Free;
    res.Free;
  end;
end;
{------------------------------------------------------------------------------}
procedure TForm2.ZIP_mifmid_tabClick(Sender: TObject);
var
  destpath, s : String;
  sl, mif : TStringList;
  i, j, col, k, firstcol : Integer;
begin
  destpath := GetAppFolder;
  if not GetFolder(destpath, Handle) then exit;
  miDo('Close All');
  sl := TStringList.Create;
  mif := TStringList.Create;
  miDo('Set ProgressBars Off');
  try
    Info('Рапаковка архивов...');
    SetCurrentDir(destpath);//обязательно
    ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [destpath+'MifMid'])), SW_HIDE, true, 0);//удаление каталога со всем содержимым и подкаталогами
    //распаковываем все зипы в новый каталог (MifMid короткое имя каталога, создается автоматически)
    ExCreateProcess(format('"%s" e * -o"MifMid"',
      [ServiceMap_Unit.ServiceMapForm.CatZip.Text]), '', SW_SHOWNORMAL, true);
//    "C:\Program Files\7-Zip\7z.exe" e * -o"D:\temp\1\MifMid"
//    FindRecursive(destpath, '*.zip', false, sl);//ищем zip
//    Info('Распакрвка архивов...');
//    pg.Max := sl.Count;
//    for i := 0 to sl.Count-1 do
//    begin
//      if (not FileExists(ChangeFileExt(sl[i], '.mif'))) or (not FileExists(ChangeFileExt(sl[i], '.mid'))) then
//        ExCreateProcess(format('"%s" x "%s" -y -o"%s"',//распаковываем все зипы
//          [ServiceMap_Unit.ServiceMapForm.CatZip.Text, sl[i], destpath]), '', SW_HIDE, true);
//      if i mod 1000 = 0 then
//        pg.Position := i;
//
//    end;
//    pg.Max := 0;

    sl.Clear;
    FindRecursive(destpath + 'MifMid\', '*.mif', false, sl);//ищем mif
    Info('Создание сводной таблицы...');
    pg.Max := sl.Count;
    for i := 0 to sl.Count-1 do
    begin
      if (i > 0) and (i mod 1000 = 0) then
      begin
        pg.Position := i;
        miDo('Commit table Itog');
      end;
      mif.LoadFromFile(sl[i]);
      for j := 1 to mif.count-1 do
        if Pos('CoordSys NonEarth', mif[j]) > 0 then
          mif[j] := 'CoordSys Earth Projection 8, 9999, 3, 23.57, -140.95, -79.80, 0.0, -0.35, -0.79, -0.22, 0, 7, 60.05, 0, 1, 1500000, -5911057.63'
        else
        if Pos('Columns', mif[j]) > 0 then
        begin
          col := StrToInt(Copy(mif[j], 9, 100));
          mif.SaveToFile(sl[i]);
          Break;
        end;
      if i = 0 then
      begin
        miDo('Import "%s" Type "MIF" Into "%sITOG.TAB" Overwrite', [sl[i], destpath + 'MifMid\']);
        firstcol := col;
      end else
      begin
        miDo('Import "%s" Type "MIF" Into "Temp" Overwrite', [sl[i]]);
        if col > firstcol then
          col := firstcol;
        s := 'COL1';
        for k := 2 to col  do
          s := s + Format(', COL%d', [k]);
        miDo('Insert Into ITOG (' + s +') Select ' + s + ' From Temp');
        miDo('Drop table Temp');
      end;
    end;
    pg.Max := 0;
    if TableExists('Itog') then
      miDo('Commit table Itog');
    if sl.Count = 0 then
      Note('Нет MIF файлов!')
    else
      miDo('Browse * From Itog');
  finally
    sl.Free;
    mif.Free;
    pg.Max := 0;
  end;
end;
{------------------------------------------------------------------------------}
procedure TForm2.mnObjectsFromXlsClick(Sender: TObject);
var
  dir : String;
  t : Cardinal;
begin
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP\test\TEST_здания_зу'; {$ENDIF}
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\TEST'; {$ENDIF}
{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\TEST\Двинин_16_05_2025'; {$ENDIF}
  if not CheckSprPath then exit;
  OpenDialog1.FileName := '*.xlsx';
  if not OpenDialog1.Execute then
    exit;
  OpenDialog1.InitialDir := ExtractFilePath(OpenDialog1.FileName);

  if not Xls2Txt then Exit;

  dir := AddSlesh(OpenDialog1.InitialDir) + AddSlesh(curdata);//исходный текстовый файл + curdata
  t := GetTickCount;
  if not CF_from_CN then
    Exc('Ошибка в расчетах.');

  Compare_13_segment;
  RunMapBasic(dir);

  SetCurrentDir(TempPath);//входим в каталог, чтобы сохранить его не удаляя
  ExShellExecute('cmd.exe', format('/c rd /s/q  "%s"', [TempPath]), SW_HIDE, false, 0);//чистим каталог

  Note(format('процес завершен.%sВремя обработки %d минут',
    [#10, Round((GetTickCount - t) / 60000)]));
end;
{------------------------------------------------------------------------------}
procedure TForm2.Compare_13_segment;
var
  row, k, i, j, n, m  : Integer;
  ar1, ar2, overlap: Double;
  s : String;
  sl : TStringList;
begin
//{$IFDEF DEBUG} OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP\VER_7\Win32\Debug\Ценовые факторы_4'; {$ENDIF}
  if not FileExists(GetAppFolder + 'Base\графика13_сегмент.tab') then
    Exc(GetAppFolder + 'Base\графика13_сегмент.tab не найдена');

  miDo('Open table "%s"', [GetAppFolder + 'Base\графика13_сегмент.tab']);
  sl := TStringList.Create;
  try
    s := fcn;
    k := miEvalInt('TableInfo(графика13_сегмент, %d)', [TAB_INFO_NCOLS]);
    for j := 1 to k do
      s := s + #9 + miEval('ColumnInfo("графика13_сегмент", "COL%d", %d)', [j, COL_INFO_NAME]);
    sl.Add(s);
    miDo('select * from %s where Str$(obj)="Region" into Regions_ noSelect', [ftab]);
    row := miEvalInt('TableInfo(Regions_, %d)', [TAB_INFO_NROWS]);
    Info('Анализ графика13_сегмент');
    pg.Max := row;
    for i := 1 to row do
    begin
      pg.Position := i;
      miDo('Fetch Rec %d from Regions_', [i]);
      miDo('ob1 = Regions_.obj');
      ar1 := miEvalDbl('CartesianArea(ob1, "sq m")');
      miDo('Select * from графика13_сегмент Where obj Partly Within ob1 into TempMO NoSelect');//границы объекта A частично лежат внутри границ объекта B

      m := miEvalInt('TableInfo(TempMO, %d)', [TAB_INFO_NROWS]);
      for n := 1 to m do
      begin
        miDo('Fetch Rec %d from TempMO', [n]);
        miDo('ob2 = TempMO.obj');
        ar2 := miEvalDbl('CartesianArea(ob2, "sq m")');//площадь старого участка "графика13_сегмент"
        if (ar1 >= ar2) or (Abs(ar2 - ar1) < 0.000001) then
        begin
          miDo('ob2 = Overlap(ob1, ob2)');
          overlap := miEvalDbl('CartesianArea(ob2, "sq m")');
          if Round(overlap / ar2 * 100) >= 95 then
          begin
            s := miEval('Regions_.%s', [fcn]);
            for j := 1 to k do
              s := s + #9 + miEval('TempMO.COL%d', [j]);
            sl.Add(s);
          end;
        end;
      end;
      miDo('Close table TempMO');
    end;
    List2Excel(sl, #9, '', nil);
  finally
    miDo('Close table Regions_');
    pg.Max := 0;
    Info('Готово');
    sl.Free;
  end;
end;


end.
