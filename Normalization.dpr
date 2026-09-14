program Normalization;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  System.Classes,
  Winapi.Windows,
  Vcl.Dialogs,
  MapBasic_INT in '..\MapBasic_INT.pas',
  Common in '..\Common.pas';

var
  zipfile, path : String;

  mid_mif, xml, work : String;
  sl : TStringList;
  i : Integer;
  wnd : HWND;

function GetConsoleWindow: HWND; stdcall; external kernel32;

procedure myMoveFile(namefile, destpath : String; para : Boolean);
var
  j, k : Integer;
  s, s2, ss : String;
begin
  j := 1;
  s := ExtractFileName(namefile);
  if not FileExists(destpath + s) then
  begin//перемещение MIF сохраняет дату
    ExShellExecute('cmd', PChar(format('/c Copy /A "%s" "%s"', [namefile, destpath])), SW_HIDE, true, 0);
    if para then
    begin//перемещение MID меняет дату, поэтому делаем копирование
      k := LastDelimiter('.', namefile);
      ExShellExecute('cmd', PChar(format('/c Copy /A "%s" "%s"',
        [Copy(namefile, 1, k) + 'mid', destpath])), SW_HIDE, true, 0);
    end;
  end else

  Repeat
    ss := StringReplace(s, '.', format('_копия%d.', [j]), [rfIgnoreCase]);
    if not FileExists(destpath + ss) then
    begin
      s2 := ExtractFilePath(namefile) + ss;
      ExShellExecute('cmd', PChar(format('/c Copy /A "%s" "%s"', [namefile, s2])), SW_HIDE, true, 0);
      ExShellExecute('cmd', PChar(format('/c Copy /A "%s" "%s"', [s2, destpath])), SW_HIDE, true, 0);
      if para then
      begin
        ExShellExecute('cmd', PChar(format('/c Copy /A "%s" "%s"',
          [ChangeFileExt(namefile, '.mid'), ChangeFileExt(s2, '.mid')])), SW_HIDE, true, 0);
        ExShellExecute('cmd', PChar(format('/c Copy /A "%s" "%s"',
          [ChangeFileExt(s2, '.mid'), destpath])), SW_HIDE, true, 0);
      end;
      break;
    end;
    Inc(j);
  Until False;
end;

procedure UnPack(sourcepath, destpath : String);
var
  i : Integer;
  s, ss : String;
  sl, dest : TStringList;
begin
  sl := TStringList.Create;
  try
    FindRecursive(sourcepath, '*.7z', false, sl);//ишем 7z
    FindRecursive(sourcepath, '*.zip', false, sl);//добавляем zip
    for i := 0 to sl.Count-1 do//извлекаем все архивы
      ExCreateProcess(format('"%s" x "%s" -y -o"%s"', [zipfile, sl[i], destpath]), '', SW_HIDE, true);
    //в каталоге назначения проверяем все папки
    dest := TStringList.Create;
    try
      FindFolders(destpath, dest, true);
      for i := 0 to dest.Count-1 do//ищем архивные файлы
      begin
        sl.Clear;
        FindRecursive(AddSlesh(dest[i]), '*.7z', false, sl);//ишем 7z
        FindRecursive(AddSlesh(dest[i]), '*.zip', false, sl);//добавляем zip
        if sl.Count > 0 then
          UnPack(AddSlesh(dest[i]), AddSlesh(dest[i]));
      end;
    finally
      dest.Free;
    end;

  finally
    sl.Free;
  end;
end;

begin
  path := ParamStr(1);
  zipfile := ParamStr(2);

  work := Copy(path, 1, Length(path)-1) + '_' +  CurrentData + '\';//отсекаем слеш
  ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [work])), SW_HIDE, true, 0);//удаление каталога со всем содержимым и подкаталогами
  wnd := GetConsoleWindow;
  SetWindowText(wnd, 'каталог: ' + work);

  mid_mif := work + 'MID_MIF\';
  ForceDirectories(mid_mif);
  xml := work + 'XML\';
  ForceDirectories(xml);

  UnPack(path, work);

  sl := TStringList.Create;
  try
    FindRecursive(work, '*.xml', true, sl);
    Writeln('Копирование файлов XML...');
    for i := 0 to sl.Count-1 do
      myMoveFile(sl[i], xml, false);

    sl.Clear;
    FindRecursive(work, '*.mif', true, sl);
    Writeln('Копирование файлов MIF/MID...');
    for i := 0 to sl.Count-1 do
      myMoveFile(sl[i], mid_mif, true);

    Writeln('Чистка каталогов...');
    ExShellExecute('cmd', PChar(format('/C Move "%s*.*" "%s"', [xml, work])), SW_HIDE, true, 0);

    sl.Clear;
    FindFolders(work, sl, false);
    for i := 0 to sl.Count-1 do//проходим по всем каталогам
      if not CText(AddSlesh(sl[i]), mid_mif) then//этот каталог не удаляем
        ExShellExecute('cmd', PChar(format('/C rd /S /Q "%s"', [sl[i]])), SW_HIDE, true, 0);//удаление каталога
  finally
    sl.Free;
  end;

end.
