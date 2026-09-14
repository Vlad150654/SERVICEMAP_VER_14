object Form2: TForm2
  Left = 590
  Top = 620
  BorderStyle = bsToolWindow
  Caption = #1056#1072#1089#1095#1077#1090' '#1094#1077#1085#1086#1074#1099#1093' '#1092#1072#1082#1090#1086#1088#1086#1074
  ClientHeight = 75
  ClientWidth = 392
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poDesigned
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel4: TPanel
    Left = 0
    Top = 53
    Width = 392
    Height = 22
    Align = alBottom
    TabOrder = 0
    object stBar: TStatusBar
      Left = 1
      Top = 1
      Width = 328
      Height = 20
      Align = alLeft
      Panels = <
        item
          Width = 80
        end
        item
          Width = 50
        end>
    end
    object pg: TProgressBar
      Left = 329
      Top = 1
      Width = 62
      Height = 20
      Align = alClient
      TabOrder = 1
    end
  end
  object Procent: TLabeledEdit
    Left = 352
    Top = 26
    Width = 33
    Height = 21
    EditLabel.Width = 108
    EditLabel.Height = 13
    EditLabel.Caption = #1055#1088#1086#1094#1077#1085#1090' '#1087#1077#1088#1077#1082#1088#1099#1090#1080#1103
    LabelPosition = lpLeft
    LabelSpacing = 10
    TabOrder = 1
    Text = '50'
  end
  object cbFindCN: TCheckBox
    Left = 8
    Top = 4
    Width = 145
    Height = 17
    Caption = #1047#1072#1087#1088#1086#1089' '#1050#1053' '#1074' '#1056#1086#1089#1088#1077#1077#1089#1090#1088#1077
    TabOrder = 2
  end
  object TimeOut: TLabeledEdit
    Left = 347
    Top = 2
    Width = 38
    Height = 21
    EditLabel.Width = 144
    EditLabel.Height = 13
    EditLabel.Caption = #1054#1078#1080#1076#1072#1085#1080#1077' '#1084#1077#1078#1076#1091' '#1079#1072#1087#1088#1086#1089#1072#1084#1080
    LabelPosition = lpLeft
    LabelSpacing = 10
    TabOrder = 3
    Text = '100'
  end
  object cbExamenOGR: TCheckBox
    Left = 8
    Top = 27
    Width = 169
    Height = 17
    Caption = #1055#1088#1086#1074#1077#1088#1082#1072' '#1075#1088#1072#1092#1080#1082#1080' '#1087#1086' OGR'
    Checked = True
    State = cbChecked
    TabOrder = 4
  end
  object OpenDialog1: TOpenDialog
    OnClose = OpenDialog1Close
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 168
  end
  object MainMenu1: TMainMenu
    AutoHotkeys = maManual
    Left = 136
    object mnCostFactory: TMenuItem
      Caption = #1060#1072#1082#1090#1086#1088#1099
    end
    object mnTable: TMenuItem
      Caption = #1058#1072#1073#1083#1080#1094#1072
      object mnCreateTableFromTab: TMenuItem
        Caption = #1057#1086#1079#1076#1072#1090#1100' '#1088#1072#1073#1086#1095#1091#1102' '#1090#1072#1073#1083#1080#1094#1091' '#1080#1079' '#1092#1072#1083#1072' "*.tab"'
        OnClick = mnCreateTableFromTabClick
      end
      object mnOpenWorkTable: TMenuItem
        Caption = #1054#1090#1082#1088#1099#1090#1100' '#1088#1072#1073#1086#1095#1091#1102' '#1090#1072#1073#1083#1080#1094#1091
        OnClick = mnOpenWorkTableClick
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object mnApplaCosFactory: TMenuItem
        Caption = #1054#1073#1085#1086#1074#1080#1090#1100' CostFactory'
        OnClick = mnApplaCosFactoryClick
      end
    end
    object mnCalc: TMenuItem
      Caption = #1056#1072#1089#1095#1077#1090
      object mnCF: TMenuItem
        Caption = #1062#1077#1085#1086#1074#1099#1077' '#1092#1072#1082#1090#1086#1088#1099
        OnClick = mnCFClick
      end
      object mnCommunications: TMenuItem
        Caption = #1050#1086#1084#1084#1091#1085#1080#1082#1072#1094#1080#1080
        object mnGraphic: TMenuItem
          Caption = #1056#1072#1089#1095#1077#1090' '#1087#1088#1080#1086#1088#1080#1090#1077#1090#1072' '#1087#1086' '#1075#1088#1072#1092#1080#1082#1077' MapInfo'
          OnClick = mnGraphicClick
        end
        object mnBase: TMenuItem
          Caption = #1056#1072#1089#1095#1077#1090' '#1087#1088#1080#1086#1088#1080#1090#1077#1090#1072' '#1087#1086' '#1073#1072#1079#1077' Access'
          OnClick = mnBaseClick
        end
        object mnCommunicationItog: TMenuItem
          Caption = #1057#1086#1079#1076#1072#1090#1100' '#1080#1090#1086#1075#1086#1074#1091#1102' '#1090#1072#1073#1083#1080#1094#1091' '#1082#1086#1084#1084#1091#1085#1080#1082#1072#1094#1080#1081
          OnClick = mnCommunicationItogClick
        end
        object N7: TMenuItem
          Caption = '-'
        end
        object mnComplexCommunication: TMenuItem
          Caption = #1050#1086#1084#1087#1083#1077#1082#1089' '#1089' '#1086#1090#1082#1088#1099#1090#1080#1077#1084' '#1088#1072#1073#1086#1095#1077#1081' '#1090#1072#1073#1083#1080#1094#1099
          OnClick = mnComplexCommunicationClick
        end
      end
      object mnOthers: TMenuItem
        Caption = #1044#1088#1091#1075#1080#1077' '#1092#1072#1082#1090#1086#1088#1099
        object mnZOUIT: TMenuItem
          Caption = #1047#1054#1059#1048#1058
          OnClick = mnZOUITClick
        end
        object mnSoil: TMenuItem
          Caption = #1055#1086#1095#1074#1099'/'#1040#1075#1088#1086#1082#1083#1080#1084#1072#1090#1080#1095#1077#1089#1082#1084#1077' '#1079#1086#1085#1099
          OnClick = mnSoilClick
        end
        object N9: TMenuItem
          Caption = '-'
        end
        object mnComplexOther: TMenuItem
          Caption = #1050#1086#1084#1087#1083#1077#1082#1089' '#1089' '#1086#1090#1082#1088#1099#1090#1080#1077#1084' '#1090#1072#1073#1083#1080#1094#1099' '#1088#1072#1073#1086#1095#1077#1081' '#1090#1072#1073#1083#1080#1094#1099
          OnClick = mnComplexOtherClick
        end
      end
      object mnObjectsFromText: TMenuItem
        Caption = #1054#1073#1098#1077#1082#1090#1099' '#1080#1079' '#1044#1077#1078#1091#1088#1085#1086#1081' '#1082#1072#1088#1090#1099' '#1087#1086' '#1050#1053' (*.txt)'
        OnClick = mnObjectsFromTextClick
      end
      object mnObjectsFromXls: TMenuItem
        Caption = #1054#1073#1098#1077#1082#1090#1099' '#1080#1079' '#1044#1077#1078#1091#1088#1085#1086#1081' '#1082#1072#1088#1090#1099' '#1087#1086' '#1050#1053' (*.xls)'
        OnClick = mnObjectsFromXlsClick
      end
    end
    object mnUtils: TMenuItem
      Caption = #1059#1090#1080#1083#1080#1090#1099
      object mnBuffers: TMenuItem
        Caption = #1041#1091#1092#1077#1088#1072' '#1076#1083#1103' '#1082#1086#1084#1084#1091#1085#1080#1082#1072#1094#1080#1081
        OnClick = mnBuffersClick
      end
      object SpedTest1: TMenuItem
        Caption = 'Pulkovo_CK42'
        OnClick = SpedTest1Click
      end
      object mnTerZone: TMenuItem
        Caption = #1058#1077#1088#1047#1086#1085#1099
        OnClick = mnTerZoneClick
      end
      object mnMi2Excel: TMenuItem
        Caption = #1057#1086#1079#1076#1072#1090#1100' '#1092#1072#1081#1083' Excel '#1080#1079' '#1090#1072#1073#1083#1080#1094#1099' Mi'
        OnClick = mnMi2ExcelClick
      end
      object ZIP_mifmid_tab: TMenuItem
        Caption = #1056#1072#1089#1087#1072#1082#1086#1074#1072#1090#1100' ZIP '#1089#1086#1073#1088#1072#1090#1100' '#1074' '#1090#1072#1073#1083#1080#1094#1091
        OnClick = ZIP_mifmid_tabClick
      end
      object Elipce2Rect: TMenuItem
        Caption = #1047#1072#1084#1077#1085#1080#1090#1100' '#1101#1083#1080#1087#1089#1099' '#1082#1074#1072#1076#1088#1072#1090#1072#1084#1080
        OnClick = Elipce2RectClick
      end
      object XML1: TMenuItem
        Caption = #1042#1099#1090#1072#1097#1080#1090#1100' '#1050#1053' '#1080' '#1089#1090#1086#1080#1084#1086#1089#1090#1100' '#1080#1079' XML'
        OnClick = XML1Click
      end
      object mnQueryRR: TMenuItem
        Caption = #1047#1072#1087#1088#1086#1089#1099' '#1074' '#1056#1086#1089#1056#1077#1077#1089#1090#1088
        OnClick = mnQueryRRClick
      end
      object Access1: TMenuItem
        Caption = 'Access'
        OnClick = Access1Click
      end
      object MultiCircuitObject1: TMenuItem
        Caption = 'Multi Circuit Object'
        OnClick = MultiCircuitObject1Click
      end
    end
  end
  object Query: TADOQuery
    Parameters = <>
    Left = 488
    Top = 24
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Tokyo\SERVICEMA' +
      'P\Win32\Debug\'#1062#1077#1085#1086#1074#1099#1077' '#1092#1072#1082#1090#1086#1088#1099'_3\'#1082#1086#1084#1084#1091#1085#1080#1082#1072#1094#1080#1080'\'#1050#1086#1084#1084#1091#1085#1080#1082#1072#1094#1080#1080'_'#1089#1074#1086#1076'.a' +
      'ccdb;Persist Security Info=False'
    Provider = 'Microsoft.ACE.OLEDB.12.0'
    Left = 232
    Top = 16
  end
end
