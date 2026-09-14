library ServiceMap;

uses
  System.SysUtils,
  System.Classes,
  Windows,
  Forms,
  Dialogs,
  MapBasic_INT in '..\MapBasic_INT.pas',
  GeoTransform in '..\GeoTransform.pas',
  Common in '..\Common.pas',
  dbTables in '..\dbTables.pas',
  cfCommon in 'cfCommon.pas',
  CostFactors_Unit in 'CostFactors_Unit.pas' {Form2},
  Letters_Unit in 'Letters_Unit.pas' {LettersForm},
  ManualParsing_Unit in 'ManualParsing_Unit.pas' {Manual},
  Path_Unit in 'Path_Unit.pas' {Path_Form},
  PrepareFD_Unit in 'PrepareFD_Unit.pas' {PrepareForm},
  ServiceMap_Unit in 'ServiceMap_Unit.pas' {ServiceMapForm};

{$R *.res}
procedure OpenServiceMap(MiWnd : HWND); stdcall;
begin
  mbName := 'ServiceMap.mbx';//модуль MBX
  FormatSettings.DecimalSeparator := '.';
  FormatSettings.ThousandSeparator := ' ';
  Forms.Application.UpdateFormatSettings := False;
  Mi := Variant(IDispatch(MiWnd));
  AppWnd := Mi.Eval('SystemInfo(9)');
  MIversion := StrToInt(Mi.Version);
  ServiceMapForm := TServiceMapForm.Create(nil);
  SetWindowLong(ServiceMapForm.Handle, GWL_HWNDPARENT, AppWnd);
  ServiceMapForm.Show;
  Form2 := TForm2.Create(nil);
  LettersForm := TLettersForm.Create(nil);
  SetWindowLong(Form2.Handle, GWL_HWNDPARENT, AppWnd);
  SetWindowLong(LettersForm.Handle, GWL_HWNDPARENT, AppWnd);
end;

procedure CloseServiceMap; stdcall;
begin
  Form2.Free;
  LettersForm.Free;
  ServiceMapForm.Free;
end;

exports
  OpenServiceMap,
  CloseServiceMap;

begin
{$IFDEF DEBUG} ReportMemoryLeaksOnShutdown := True; {$ENDIF}
end.
