object Form1: TForm1
  Left = 258
  Top = 107
  Width = 580
  Height = 480
  Caption = #1063#1090#1077#1085#1080#1077' Excel'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 315
    Top = 20
    Width = 102
    Height = 16
    Caption = #1054#1073#1097#1072#1103' '#1089#1091#1084#1084#1072':'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Button1: TButton
    Left = 20
    Top = 20
    Width = 75
    Height = 25
    Caption = #1054#1090#1082#1088#1099#1090#1100
    TabOrder = 0
    OnClick = Button1Click
  end
  object BitBtn1: TBitBtn
    Left = 160
    Top = 20
    Width = 75
    Height = 25
    Caption = #1042#1099#1093#1086#1076
    TabOrder = 1
    OnClick = BitBtn1Click
  end
  object GroupBox1: TGroupBox
    Left = 20
    Top = 55
    Width = 290
    Height = 380
    Caption = #1060#1072#1084#1080#1083#1080#1080
    TabOrder = 2
    object ListBox1: TListBox
      Left = 10
      Top = 20
      Width = 260
      Height = 350
      ItemHeight = 13
      TabOrder = 0
    end
  end
  object GroupBox2: TGroupBox
    Left = 330
    Top = 55
    Width = 225
    Height = 385
    Caption = #1057#1091#1084#1084#1099
    TabOrder = 3
    object ListBox2: TListBox
      Left = 10
      Top = 20
      Width = 200
      Height = 350
      ItemHeight = 13
      TabOrder = 0
    end
  end
  object Edit1: TEdit
    Left = 430
    Top = 20
    Width = 121
    Height = 24
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 4
    Text = #1054#1073#1097#1072#1103' '#1089#1091#1084#1084#1072
  end
  object OpenDialog1: TOpenDialog
    Left = 112
    Top = 16
  end
end
