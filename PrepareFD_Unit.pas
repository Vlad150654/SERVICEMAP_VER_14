unit PrepareFD_Unit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls,
  System.Zip,
  System.RegularExpressions,
  System.StrUtils,
  Common, Vcl.Buttons;

type
  TPrepareForm = class(TForm)
    edPackets: TEdit;
    edSource: TEdit;
    btZIP: TBitBtn;
    btPacket: TBitBtn;
    btOk: TBitBtn;
    btCancel: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btOkClick(Sender: TObject);
    procedure btZIPClick(Sender: TObject);
    procedure btPacketClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    SourcPath, PacketPath: String;
    TempPath : String;//C:\Users\TroitskikhVV\AppData\Local\Temp\MIPRO
    sl : TStringList;
    Location : Boolean;
  end;

var
  PrepareForm: TPrepareForm;

implementation

uses ServiceMap_Unit;

{$R *.dfm}

procedure TPrepareForm.btOkClick(Sender: TObject);
var
  xml : TStringList;
  i, j, n, m : Integer;
  zipname, s, ss, addr : String;
  mc : TMatchCollection;
begin
  if not DirectoryExists(PacketPath) then
    Exc('Нет каталога с пакетами.');

  if not DirectoryExists(SourcPath) then
    Exc('Нет каталога с архивами "ZIP".');

  if edPackets.Text = edSource.Text then
    Exc('ОЙ! Каталоги одинаковые, наверное ошибка');

  Hide;

  zipname := ServiceMap_Unit.ServiceMapForm.CatZip.Text;
  ForceDirectories(TempPath + '1\');
  ForceDirectories(TempPath + '2\');

  sl.Clear;
  FindRecursive(edSource.Text, '*.zip', false, sl);
  for i := 0 to sl.Count-1 do
    ExCreateProcess(format('"%s" x "%s" -y -o"%s"',//распаковываем все зипы в каталог 1
      [zipname, sl[i], TempPath + '1\']), '', SW_HIDE, true);

  if Location then//корректировка узла Location
  begin
    xml := TStringList.Create;
    sl.Clear;
    FindRecursive(TempPath + '1\', 'FD_*.xml', false, sl);
    for i := 0 to sl.Count-1 do
    begin
      xml.LoadFromFile(sl[i], TEncoding.UTF8);
      s := '';
      for j := 0 to xml.Count-1 do
        s := s + Trim(xml[j]);//все в одну строку
      s := StringReplace(s, '</Location>', '</Location>' + lf, [rfReplaceAll]);
      mc := reg.Matches(s, '<Location>.+</Location>', [roMultiLine]);
      for j := 0 to mc.Count-1 do
      begin
        n := Pos('<Note>', mc.Item[j].Value);
        m := PosEx('</Note>', mc.Item[j].Value);
        if (m > 0) and (m > n) then
        begin
          addr := Copy(mc.Item[j].Value, n+6, m-n-6);
          ss := '<Location>' + lf + '<Region>66</Region>' + lf +
            '<Note>' + addr + '</Note>'+lf + '</Location>' + lf;
          s := StringReplace(s, mc.Item[j].Value, ss, []);
        end;
      end;
      s := StringReplace(s, lf, '', [rfReplaceAll]);
///////////
//      xml.Text := s;
//      xml.SaveToFile(sl[i]);
///////////
      s := StringReplace(s, '<Real_Estates>', '<Real_Estates>' + lf, []);
      s := StringReplace(s, '</Real_Estate>', '</Real_Estate>' + lf, [rfReplaceAll]);
      s := StringReplace(s, '</Real_Estates>', '</Real_Estates>' + lf, []);
      xml.Text := s;
      xml.SaveToFile(sl[i]);
    end;
    xml.Free;
  end;

  ExCreateProcess(format('%s a -tzip -mx5 -r0 %s %s',//создаем FD.zip в каталоге 2
    [zipname, TempPath + '2\FD.zip', TempPath + '1\FD_*.xml']), '', SW_HIDE, true);

  s := format('/c rd /S /Q  "%s"', [TempPath + '1\']);
  ExShellExecute('cmd.exe', s, SW_HIDE, false, 0);//удалене каталога 1 со всем содержимым и подкаталогами

  ExShellExecute('cmd', format('/C Copy /b /Y "%s" "%s"',//копируем в каталог с пакетами
    [TempPath + '2\FD.zip', edPackets.Text]), SW_HIDE, true, 0);

  s := format('/c rd /S /Q  "%s"', [TempPath + '2\']);
  ExShellExecute('cmd.exe', s, SW_HIDE, false, 0);//удалене каталога 2 со всем содержимым и подкаталогами

  ModalResult := mrOk;
end;

//W:\ЦГКО\IT\Outbox\Статья 16\2026\2026_02_12_TEST\SOKS_a_2026-02-06_3
//D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\SOKS_a_2026-02-06_3
procedure TPrepareForm.btPacketClick(Sender: TObject);
var
  i : Integer;
begin
  if not GetFolder(PacketPath, Handle) then exit;
  sl.Clear;
  FindFolders(PacketPath, sl, false);
  for i := sl.Count-1 downto 0 do
    if not CText(ExtractFileName(sl[i]), 'Пакеты') then
      sl.Delete(i);//в каталоге должена остаться папка "Пакеты"
  if sl.Count = 0 then
    Note('В выбранном каталоге не найдены пакеты.')
  else
    edPackets.Text := ExtractFilePath(sl[0]);
end;

//W:\ЦГКО\IT\Outbox\Статья 16\2026\2026_02_12_TEST\ОКС_ИТОГ
//D:\Tokyo\SERVICEMAP_VER_14\Win32\Debug\TEST\тест_правка_FD\ОКС_ИТОГ
procedure TPrepareForm.btZIPClick(Sender: TObject);
begin
  if not GetFolder(SourcPath, Handle) then exit;
  sl.Clear;
  FindRecursive(SourcPath, '*.zip', false, sl);
  if sl.Count = 0 then
    Note('В выбранном каталоге нет файлов "ZIP".')
  else
    edSource.Text := ExtractFilePath(sl[0]);
end;

procedure TPrepareForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  sl.Free;
end;

procedure TPrepareForm.FormCreate(Sender: TObject);
begin
  sl := TStringList.Create;
end;

end.
