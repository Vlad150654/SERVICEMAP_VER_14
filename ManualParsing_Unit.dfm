object Manual: TManual
  Left = 480
  Top = 263
  Caption = #1056#1091#1095#1085#1086#1081' '#1088#1072#1079#1073#1086#1088
  ClientHeight = 198
  ClientWidth = 582
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnActivate = FormActivate
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object sg: TStringGrid
    Left = 0
    Top = 0
    Width = 582
    Height = 173
    Align = alClient
    DefaultColWidth = 96
    FixedCols = 0
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    ParentFont = False
    TabOrder = 0
    OnDblClick = sgDblClick
    OnDrawCell = sgDrawCell
  end
  object BitBtn1: TBitBtn
    Left = 0
    Top = 173
    Width = 582
    Height = 25
    Align = alBottom
    Caption = 
      #1055#1088#1086#1074#1077#1088#1080#1090#1100' '#1074#1074#1077#1076#1077#1085#1085#1099#1077'  '#1080' '#1087#1088#1086#1076#1086#1083#1078#1080#1090#1100' '#1088#1072#1073#1086#1090#1091' ('#1076#1074#1086#1081#1085#1086#1081' '#1082#1083#1080#1082' '#1076#1083#1103' '#1091#1089#1090#1072#1085 +
      #1086#1074#1082#1080'/'#1089#1085#1103#1090#1080#1103' '#1074#1099#1076#1077#1083#1077#1085#1080#1103')'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 1
    OnClick = BitBtn1Click
  end
end
