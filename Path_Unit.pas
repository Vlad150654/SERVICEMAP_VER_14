unit Path_Unit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  System.IniFiles,
  Mapbasic_INT, Vcl.ExtCtrls,
  Common, cfCommon;

type
  TPath_Form = class(TForm)
    btCancel: TBitBtn;
    btOk: TBitBtn;
    btXLS: TBitBtn;
    btDBF: TBitBtn;
    btXML: TBitBtn;
    btPDF: TBitBtn;
    Edit1: TEdit;
    Edit3: TEdit;
    Edit2: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    cbDBF: TCheckBox;
    procedure btXLSClick(Sender: TObject);
    procedure btDBFClick(Sender: TObject);
    procedure btXMLClick(Sender: TObject);
    procedure btPDFClick(Sender: TObject);
    procedure btOkClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    s : String;
    function CheckIskl : Boolean;
    function CheckFields : Boolean;
  public
    { Public declarations }
    path, respath, fieldsdbf, tabdbf, xls, source : String;
    radius, decimal : Integer;
    procent : Double;
  end;

var
  Path_Form: TPath_Form;

implementation

uses ServiceMap_Unit;

{$R *.dfm}

procedure TPath_Form.btPDFClick(Sender: TObject);
begin
  path := 'W:\ЦГКО\IT\_ОБМЕН ДАННЫМИ_\Штань\PDF_2\foto2';//tr!!!
  path := 'W:\ЦГКО\IT\_ОБМЕН ДАННЫМИ_\Штань\PDF';//tr!!!
  if GetFolder(path, Handle) then
    Edit4.Text := path;
//  Winapi.Windows.SetFocus(Handle);
end;

//из исправленной таблицы (от Штаня) выбираем поля,описанные в секции [FIELDS_DBF]
function TPath_Form.CheckIskl : Boolean;
var
  i, j : Integer;
  sl, log, fields : TStringList;
  ini : TMemIniFile;
begin
  Result := false;
  SafeCloseTable('TempTab');
  sl := TStringList.Create;
  log := TStringList.Create;
  fields := GetFieldsList(tabdbf);
  fields.Sort;
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
    ini.ReadSectionValues('FIELDS_DBF', sl);//набор полей файла DBF
    fieldsdbf := 'Select ' + sl[0];//для копирование исключений
    for i := 1 to sl.Count-1 do
    begin
      if not fields.Find(sl[i], j) then
        log.Add(sl[i] + ' не найден в таблице ' + tabdbf)
      else
        fieldsdbf := fieldsdbf + format(',%s', [sl[i]]);
    end;
    fieldsdbf := fieldsdbf + ' from ' + tabdbf + ' Into tempTab';//строка-запрос

    Result := log.Count = 0;
    s := ChangeFileExt(ServiceMapForm.OpenDialog1.FileName, '.log');
    log.SaveToFile(s);
    if log.Count > 0 then
      ExShellExecute('Notepad.exe', s, SW_NORMAL, false, 0)
    else
      DeleteFile(s);
  finally
    ini.Free;
    sl.Free;
    log.Free;
    fields.Free;
  end;
end;

function TPath_Form.CheckFields : Boolean;
Const//обязательно нужны для работы с аналогами
  fnames : Array [0..13] of String = ('ID','Площадь участка', 'Комментарий',
    'pdf', 'mapLng', 'mapLat', 'Адрес', 'Адрес из парсинга', 'Название предложения',
    'Цена', 'Площадь округленная=1', 'x_41=2', 'y_41=3', 'coord_wgs=4');
var
  i, k : Integer;
  log, fields : TStringList;
  ini : TMemIniFile;
begin
  Result := false;
  log := TStringList.Create;
  fields := TStringList.Create;//должный набор полей исходного файла
  radius := 0;
  ini := TMemIniFile.Create(GetAppFolder + 'ServiceMap.ini');
  try
    ini.ReadSectionValues('Interface', fields);
    for i := 0 to fields.Count-1 do
    begin
      if CText(fields.Names[i], 'Radius_Roads') then
        Val(fields.ValueFromIndex[i], radius, k) else
      if CText(fields.Names[i], 'StringCompare_Decimal') then
        Val(fields.ValueFromIndex[i], decimal, k);
      if CText(fields.Names[i], 'StringCompare_Procent') then
        Val(fields.ValueFromIndex[i], procent, k);
    end;
    fields.Clear;

    ini.ReadSectionValues('Analysis', fields);
    for i := fields.Count-1 downto 0 do
      if Pos('//', fields[i]) = 1 then
        fields.Delete((i));
    for i := Low(fnames) to High(fnames) do
      if fields.IndexOf(fnames[i]) = -1 then
        log.Add('Не найдена колонка ' + fnames[i]);

    if radius = 0 then radius := 30;
    if decimal = 0 then decimal := 3;
    if procent = 0 then procent := 0.95;
    Result := log.Count = 0;
    s := ChangeFileExt(ServiceMapForm.OpenDialog1.FileName, '.log');
    log.SaveToFile(s);
    if log.Count > 0 then
      ExShellExecute('Notepad.exe', s, SW_NORMAL, false, 0)
    else
      DeleteFile(s);
  finally
    ini.Free;
    fields.Free;
    log.Free;
  end;
end;

procedure TPath_Form.FormCreate(Sender: TObject);
begin
{$IFDEF DEBUG} ServiceMapForm.OpenDialog1.InitialDir := 'D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\Аналоги_test4'; {$ENDIF}
  if ServiceMapForm.OpenDialog1.InitialDir = '' then
    ServiceMapForm.OpenDialog1.InitialDir := GetCurrentDir;
  Note('Нет сервера W:\ЦГКО...');
end;

procedure TPath_Form.btOkClick(Sender: TObject);
begin
  if not FileExists(Edit1.Text) then
    Exc('Нет файла для расчета аналогов');

  if not cbDBF.Checked then//полный цикл
  begin
    if not FileExists(Edit3.Text) then
      Exc('Нет файла XML');
    if not DirectoryExists(Edit4.Text) then
      Exc('Нет каталога PDF');
    if (Length(Edit5.Text) > 1) and (not (Edit5.Text[1] in ['0'..'9'])) then
      Exc('Ошибка в задании даты выгрузки');

    respath := GetAppFolder + 'Аналоги\' + AddSlesh(CurrentData);
    ForceDirectories(respath);
    source := Edit1.Text;//tr!!!
    xls := ExtractFileName(Edit1.Text);
    xls := respath + ChangeFileExt(xls, '_Результат.xlsx');
    if FileExists(xls) then
      if not DeleteFile(xls) then
        Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [xls, #10]));
    ExShellExecute('cmd', PChar(format('/C Copy /b /Y "%s" "%s"', [Edit1.Text, xls])), SW_HIDE, true, 0);

    s := ExtractFilePath(xls) + 'Z' + Edit5.Text + 'graf.dbf';
    if FileExists(s) then
      if not DeleteFile(s) then
        Exc(format('Файл "%s" открыт.%s Закройте файл и повторите попытку.', [s, #10]));

    if not CheckFields then
      exit;
  end else//берем готовый
  begin
    if not FileExists(Edit2.Text) then
      Exc('Нет файла с анализом данных');
    xls := Edit1.Text;
    if Pos('_Результат.xlsx', xls) = 0 then
      Exc('Необходимо выбрать файл результата вычислений xxx_Результат.xlsx!');
    respath := ExtractFilePath(xls);

    s := ChangeFileExt(xls, '_2.xlsx');
    if FileExists(s) then
      if not DeleteFile(s) then
        Exc(format('Файл "%s" открыт в Excel.%s Закройте файл и повторите попытку.', [s, #10]));

    s := ChangeFileExt(Edit2.Text, '.tab');//z241101ispr.dbf
    tabdbf := miEval('PathToTableName$("%s")', [s]);//таблица искдючений от Штаня
//    SafeCloseTable(tabdbf);
    miDo('Close All');
    miDo('Register Table "%s"  TYPE DBF Charset "WindowsCyrillic" Into "%s"', [Edit2.Text, s]);
    miDo('Open Table "%s"', [s]);
    if not CheckIskl then
      exit;
  end;

  ModalResult := mrOk;
end;

procedure TPath_Form.btXLSClick(Sender: TObject);
begin
//D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\Аналоги_test4\241101.xlsx
  ServiceMapForm.OpenDialog1.FileName := '*.xlsx';
  if ServiceMapForm.OpenDialog1.Execute then
    Edit1.Text := ServiceMapForm.OpenDialog1.FileName;
  ServiceMapForm.OpenDialog1.InitialDir := ExtractFilePath(Edit1.Text);
end;

procedure TPath_Form.btDBFClick(Sender: TObject);
begin
  ServiceMapForm.OpenDialog1.FileName := '*.dbf';
  if ServiceMapForm.OpenDialog1.Execute then
    Edit2.Text := ServiceMapForm.OpenDialog1.FileName;
  ServiceMapForm.OpenDialog1.InitialDir := ExtractFilePath(Edit2.Text);
end;

procedure TPath_Form.btXMLClick(Sender: TObject);
begin
//D:\Tokyo\SERVICEMAP_VER_12\Win32\Debug\Аналоги_test4\241001.xml
  ServiceMapForm.OpenDialog1.FileName := '*.xml';
  if ServiceMapForm.OpenDialog1.Execute then
    Edit3.Text := ServiceMapForm.OpenDialog1.FileName;
  ServiceMapForm.OpenDialog1.InitialDir := ExtractFilePath(Edit3.Text);
end;

end.
