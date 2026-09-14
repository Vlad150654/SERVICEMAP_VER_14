program Communications;


{$APPTYPE CONSOLE}

{$R *.res}

uses
  Winapi.Windows,
  System.Classes,
  System.SysUtils,
  ComObj,
  Variants,
  Winapi.Messages,
  Common in '..\Common.pas',
  MapBasic_INT in '..\MapBasic_INT.pas';

var
  wnd, mwnd : HWND;
  s, tab : String;
  sl : TStringList;
  i : Integer;

  aCopyData : TCopyDataStruct;
  CDS : TCopyDataStruct;

function GetConsoleWindow: HWND; stdcall; external kernel32;
{---------------------определение пересечения----------------------------------}
begin
  if ParamStr(1) = '' then
    Note('Это файл запускается из ценовых факторов')
  else
  try
    mwnd := GetConsoleWindow;
    s := ExtractFileName(DellSlesh(ParamStr(2)));
    SetWindowText(mwnd, s);
    Val(ParamStr(4), wnd, i);
    mi := CreateOleObject('MapInfo.Application');
    AppWnd := Mi.Eval('SystemInfo(9)');
    Set_Priority(AppWnd, HIGH_PRIORITY_CLASS);
    miDo('Set ProgressBars Off');
    sl := TStringList.Create;

    miDo('Open table "%s"', [ParamStr(1) + 'All_CN.tab']);//путь к каталогу коммуникаций
    miDo('set coordsys table All_CN');
    FindRecursive(ParamStr(2), '*.tab', false, sl);
    for i := 0 to sl.Count - 1 do//выборка по каждой части(линии или потребители)
    try
      miDo('Open Table "%s"', [sl[i]]);
      tab := miEval('PathToTableName$("%s")', [sl[i]]);
      writeln(tab);
      //границы объекта A частично лежат внутри границ объекта B
      // границы участков частично лежат внутри коммуникаций
      miDo('Select * from All_CN,%s where All_CN.Obj Partly Within %s.Obj into Temp noSelect',
        [tab, tab]);
      if i = 0 then//добавляем результаты пересечений в итоговую таблицу
      begin
        miDo('Commit Table Temp As "%s%s.tab" TYPE NATIVE Charset "WindowsCyrillic"',
          [ParamStr(3), ExtractFileName(ExtractFileDir(sl[i]))]);//водоотведение, Водоснабжение и др.
        miDo('Open Table "%s%s.tab" As Itog',
          [ParamStr(3), ExtractFileName(ExtractFileDir(sl[i]))]);
      end else
        miDo('Insert Into Itog Select * From Temp');
      miDo('Close table Temp');
      miDo('Close table %s', [tab]);
    except
      on e : Exception do note(e.Message);
    end;
    miDo('Commit table Itog');
    miDo('Close table Itog');
  finally
    sl.Free;
    miDo('Close All');
    miDo('End MapInfo');
    mi := Unassigned;
    s := 'RESULT';
    CDS.dwData := 0;
    CDS.cbData := (length(s) + 1) * sizeof(s);
    CDS.lpData := PChar(s);
    SendMessage(wnd, WM_COPYDATA, wnd, Integer(@CDS));
  end;
end.
