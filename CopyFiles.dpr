program CopyFiles;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.Classes,
  Winapi.ShellAPI,
  Winapi.Windows,
//  Vcl.Dialogs,
  System.SysUtils;

var
  SEI : TShellExecuteInfo;
  sl, xml : TStringList;
  newname, id, url, filename, path, respath, s : String;
  i, j, n, k : Integer;
  mwnd : hwnd;

function GetConsoleWindow: HWND; stdcall; external kernel32;

begin
  mwnd := GetConsoleWindow;
  try
    if (ParamStr(1) = '') OR (ParamStr(2) = '') OR (ParamStr(3) = '') OR (ParamStr(4) = '') then Exit;
    sl := TStringList.Create;
    sl.LoadFromFile(ParamStr(1));
    xml := TStringList.Create;
    xml.LoadFromFile(ParamStr(2), TEncoding.UTF8);
    path := ParamStr(3);
    respath := ParamStr(4);
    for i  := 0 to sl.Count-1 do
    begin
      j := Pos('~', sl[i]);
      id := Copy(sl[i], 1, j-1);
      url := Copy(sl[i], j+1, Length(sl[i]));
//      SetWindowText(mwnd, id);
      sl[i] := '';
      for j := 0 to xml.Count-2 do
      begin
        if Pos(url, xml[j]) > 0 then
        begin
          filename := xml[j+1];
          n := Pos('<FileName>', filename);
          k := Pos('</FileName>', filename);
          if( n > 0) and (k > 0) then
          begin
            filename := Copy(filename, n+10, k-n-10);
            s := path + ExtractFileName(filename);
            if FileExists(s) then
            begin
              newname := respath + id + '.pdf';
              sl[i] := newname;
SetWindowText(mwnd, newname);

              FillChar(SEI, SizeOf(SEI), 0);
              SEI.cbSize := SizeOf(SEI);
              SEI.fMask := SEE_MASK_NOCLOSEPROCESS;
              {$IFDEF UNICODE}
              SEI.fMask := SEI.fMask or SEE_MASK_UNICODE;
              {$ENDIF}
              SEI.Wnd := 0;
              SEI.lpFile := PChar('cmd');
              SEI.lpParameters := PChar(PChar(format('/C Copy /b /Y "%s" "%s"' ,[s, newname])));
              SEI.nShow := SW_HIDE;
              ShellExecuteEx(@SEI);
              CloseHandle(SEI.hProcess);
            end else
              sl[i] := 'Скан отсутствует';
          end;
          break;
        end;
      end;
    end;
  finally
    sl.SaveToFile(ExtractFilePath(respath) + 'files.txt');
    sl.Free;
  end;
end.
