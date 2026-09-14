object ServiceMapForm: TServiceMapForm
  Left = 0
  Top = 0
  BorderStyle = bsToolWindow
  Caption = #1044#1077#1078#1091#1088#1085#1072#1103' '#1082#1072#1088#1090#1072
  ClientHeight = 128
  ClientWidth = 529
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 424
    Top = 1
    Width = 92
    Height = 13
    Caption = #1050#1072#1090#1077#1075#1086#1088#1080#1080' '#1079#1077#1084#1077#1083#1100
  end
  object Panel4: TPanel
    Left = 0
    Top = 106
    Width = 529
    Height = 22
    Align = alBottom
    TabOrder = 0
    object stBar: TStatusBar
      Left = 1
      Top = 1
      Width = 392
      Height = 20
      Align = alLeft
      Panels = <
        item
          Width = 90
        end
        item
          Width = 50
        end>
    end
    object pg: TProgressBar
      Left = 393
      Top = 1
      Width = 135
      Height = 20
      Align = alClient
      TabOrder = 1
    end
  end
  object cbProgramMB: TComboBox
    Left = 287
    Top = 14
    Width = 123
    Height = 21
    Style = csDropDownList
    TabOrder = 1
  end
  object CatZip: TLabeledEdit
    Left = 86
    Top = 84
    Width = 538
    Height = 21
    EditLabel.Width = 75
    EditLabel.Height = 13
    EditLabel.Caption = #1050#1072#1090#1072#1083#1086#1075' 7-Zip  '
    LabelPosition = lpLeft
    LabelSpacing = 6
    TabOrder = 2
    Text = 'C:\Program Files\7-Zip\7z.exe'
  end
  object cbCategories: TComboBox
    Left = 424
    Top = 14
    Width = 97
    Height = 21
    Style = csDropDownList
    TabOrder = 3
  end
  object cbMapBasic: TCheckBox
    Left = 8
    Top = 11
    Width = 255
    Height = 17
    Caption = #1047#1072#1087#1091#1089#1082#1072#1090#1100' '#1087#1088#1086#1075#1088#1072#1084#1084#1091' MapBasic '#1076#1083#1103' '#1075#1088#1072#1092#1080#1082#1080
    Checked = True
    State = cbChecked
    TabOrder = 4
  end
  object cbRunUpdate: TCheckBox
    Left = 8
    Top = 34
    Width = 273
    Height = 17
    Caption = #1047#1072#1087#1091#1089#1082#1072#1090#1100' '#1086#1073#1085#1086#1074#1083#1077#1085#1080#1077' '#1087#1086#1089#1083#1077' '#1089#1086#1079#1076#1072#1085#1080#1103' '#1075#1088#1072#1092#1080#1082#1080
    TabOrder = 5
  end
  object cbPacked: TCheckBox
    Left = 8
    Top = 58
    Width = 273
    Height = 17
    Caption = #1059#1087#1072#1082#1086#1074#1072#1090#1100' '#1044#1077#1078#1091#1088#1085#1091#1102' '#1082#1072#1088#1090#1091' '#1087#1086#1089#1083#1077' '#1086#1073#1085#1086#1074#1083#1077#1085#1080#1103
    TabOrder = 6
  end
  object MainMenu1: TMainMenu
    AutoHotkeys = maManual
    Left = 296
    Top = 32
    object mnLetters: TMenuItem
      Caption = #1055#1080#1089#1100#1084#1072
      object mnPrepareLetters: TMenuItem
        Caption = #1055#1077#1088#1074#1080#1095#1085#1072#1103' '#1086#1073#1088#1072#1073#1086#1090#1082#1072' '#1087#1080#1089#1077#1084
        OnClick = mnPrepareLettersClick
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object mnExamen: TMenuItem
        Caption = #1053#1086#1088#1084#1072#1083#1080#1079#1086#1074#1072#1090#1100' '#1055#1080#1089#1100#1084#1072
        OnClick = mnExamenClick
      end
      object mnUnion: TMenuItem
        Caption = #1057#1083#1080#1103#1085#1080#1077' XML'
        OnClick = mnUnionClick
      end
      object mnCompareXML: TMenuItem
        Caption = #1057#1088#1072#1074#1085#1077#1085#1080#1077' XML '#1080' '#1089#1086#1079#1076#1072#1085#1080#1077' '#1087#1072#1082#1077#1090#1086#1074
        OnClick = mnCompareXMLClick
      end
    end
    object mnGraphics: TMenuItem
      Caption = #1043#1088#1072#1092#1080#1082#1072
      object mnCreateTables: TMenuItem
        Caption = #1057#1086#1079#1076#1072#1085#1080#1077' '#1090#1072#1073#1083#1080#1094
        OnClick = mnCreateTablesClick
      end
      object mnUpdateBase: TMenuItem
        Caption = #1054#1073#1085#1086#1074#1083#1077#1085#1080#1077' '#1073#1072#1079#1099
        OnClick = mnUpdateBaseClick
      end
    end
    object mnCalcCF: TMenuItem
      Caption = #1062#1077#1085#1086#1074#1099#1077' '#1092#1072#1082#1090#1086#1088#1099
      object mnSt16: TMenuItem
        Caption = #1057#1090#1072#1090#1100#1103' 16'
        OnClick = mnSt16Click
      end
      object mnGKO: TMenuItem
        Caption = #1043#1050#1054
        OnClick = mnGKOClick
      end
    end
    object mnAnalogs: TMenuItem
      Caption = #1040#1085#1072#1083#1086#1075#1080
      object mnCalc: TMenuItem
        Caption = #1056#1072#1089#1095#1077#1090
        OnClick = mnCalcClick
      end
      object mnCalcGroup: TMenuItem
        Caption = #1043#1088#1091#1087#1087#1086#1074#1086#1081' '#1088#1072#1089#1095#1077#1090
        OnClick = mnCalcGroupClick
      end
      object mnCheck: TMenuItem
        Caption = #1055#1088#1086#1074#1077#1088#1082#1072
        OnClick = mnCheckClick
      end
      object mnSelectUserFields: TMenuItem
        Caption = #1042#1099#1073#1086#1088' '#1076#1072#1085#1085#1099#1093' '#1076#1083#1103' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1103
        OnClick = mnSelectUserFieldsClick
      end
      object mnImportXML: TMenuItem
        Caption = 'TopnLab '#1048#1084#1087#1086#1088#1090' '#1080#1079' XML'
        OnClick = mnImportXMLClick
      end
    end
    object mnCorrect: TMenuItem
      Caption = #1055#1088#1072#1074#1082#1072
      object CorrectXML: TMenuItem
        Caption = #1055#1088#1072#1074#1082#1072' XML'
        OnClick = CorrectXMLClick
      end
      object mnFD: TMenuItem
        Caption = #1055#1088#1072#1074#1082#1072' FD'
        object mnFullCorrectFD: TMenuItem
          Caption = #1055#1086#1083#1085#1072#1103' '#1087#1088#1072#1074#1082#1072' FD'
          OnClick = mnFullCorrectFDClick
        end
        object mnCorrectFDPlus: TMenuItem
          Caption = ' '#1055#1086#1083#1085#1072#1103' '#1087#1088#1072#1074#1082#1072' FD+'#1040#1076#1088#1077#1089
          OnClick = mnCorrectFDPlusClick
        end
        object Telda2НСПД: TMenuItem
          Caption = #1058#1077#1083#1076#1072' '#1074' '#1092#1086#1088#1084#1072#1090' '#1076#1083#1103' '#1079#1072#1075#1088#1091#1079#1082#1080' '#1074' '#1053#1057#1055#1044
          OnClick = Telda2НСПДClick
        end
      end
      object N2: TMenuItem
        Caption = #1053#1086#1074#1072#1103' '#1092#1086#1088#1084#1072' '#1072#1082#1090#1086#1074
        OnClick = N2Click
      end
    end
    object mnAny: TMenuItem
      Caption = #1056#1072#1079#1085#1086#1077
      object ogr1: TMenuItem
        Caption = 'GDAL/OGR'
        OnClick = ogr1Click
      end
      object mnDKSEMANTIC: TMenuItem
        Caption = 'DK_SEMANTIC'
        OnClick = mnDKSEMANTICClick
      end
      object mnTransactions: TMenuItem
        Caption = #1057#1076#1077#1083#1082#1080
        OnClick = mnTransactionsClick
      end
    end
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties="dBASE IV";' +
      'Data Source=Provider=MSDASQL.1;Persist Security Info=False;Data ' +
      'Source=dBASE Files;Initial Catalog="D:\Tokyo\TRACK\'#1054#1090'_'#1064#1090#1072#1085#1103'\'#1087#1088#1086#1084 +
      #1099#1096#1083#1077#1085#1085#1086#1077' '#1079#1072#1091#1088#1072#1083#1100#1077'\towns";User ID=Admin;Password=""'
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 216
  end
  object OpenDialog1: TOpenDialog
    OnClose = OpenDialog1Close
    FileName = 'D:\Tokyo\d7zip-master\7z.dll'
    Left = 368
    Top = 40
  end
end
