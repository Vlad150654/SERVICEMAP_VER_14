object LettersForm: TLettersForm
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1055#1077#1088#1074#1080#1095#1085#1072#1103' '#1086#1073#1088#1072#1073#1086#1090#1082#1072' '#1087#1080#1089#1077#1084
  ClientHeight = 106
  ClientWidth = 356
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object btCalc: TButton
    Left = 126
    Top = 56
    Width = 98
    Height = 27
    Caption = #1056#1072#1089#1095#1077#1090
    TabOrder = 0
    OnClick = btCalcClick
  end
  object stBar: TStatusBar
    Left = 0
    Top = 87
    Width = 356
    Height = 19
    Panels = <
      item
        Width = 50
      end>
  end
  object leNumberInputinfo: TLabeledEdit
    Left = 8
    Top = 21
    Width = 96
    Height = 21
    EditLabel.Width = 69
    EditLabel.Height = 13
    EditLabel.Caption = #1053#1086#1084#1077#1088' '#1087#1080#1089#1100#1084#1072
    TabOrder = 2
  end
  object leDateInputinfo: TLabeledEdit
    Left = 252
    Top = 21
    Width = 96
    Height = 21
    EditLabel.Width = 64
    EditLabel.Height = 13
    EditLabel.Caption = #1044#1072#1090#1072' '#1087#1080#1089#1100#1084#1072
    TabOrder = 3
  end
  object btPacket: TButton
    Left = 8
    Top = 56
    Width = 96
    Height = 27
    Caption = #1050#1072#1090#1072#1083#1086#1075
    TabOrder = 4
    OnClick = btPacketClick
  end
  object btToBase: TButton
    Left = 252
    Top = 56
    Width = 96
    Height = 27
    Caption = #1055#1088#1080#1085#1103#1090#1100' '#1087#1080#1089#1100#1084#1086
    TabOrder = 5
    OnClick = btToBaseClick
  end
  object CreateCSV: TBitBtn
    Left = 8
    Top = 96
    Width = 96
    Height = 25
    Caption = #1057#1086#1079#1076#1072#1090#1100' CSV'
    TabOrder = 6
    Visible = False
    OnClick = CreateCSVClick
  end
  object LabeledEdit1: TLabeledEdit
    Left = 126
    Top = 100
    Width = 98
    Height = 21
    EditLabel.Width = 74
    EditLabel.Height = 13
    EditLabel.Caption = 'foundationdate'
    TabOrder = 7
    Visible = False
  end
  object OpenDialog1: TOpenDialog
    Left = 72
    Top = 8
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=s;Persist Security Info=True;User I' +
      'D=zisadmin;Initial Catalog=btizu_21st_2022;Data Source=TESTSQL;U' +
      'se Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;' +
      'Workstation ID=G001;Use Encryption for Data=False;Tag with colum' +
      'n collation when possible=False'
    LoginPrompt = False
    Mode = cmReadWrite
    Provider = 'SQLOLEDB.1'
    Left = 144
    Top = 8
  end
  object Query: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 208
    Top = 8
  end
end
