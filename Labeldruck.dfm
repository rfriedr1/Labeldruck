object Form7: TForm7
  Left = 0
  Top = 0
  Caption = 'Labeldruck'
  ClientHeight = 447
  ClientWidth = 305
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 87
    Top = 306
    Width = 78
    Height = 13
    Caption = 'Anzahl Etiketten'
  end
  object edEinsenderNummer: TLabeledEdit
    Left = 129
    Top = 39
    Width = 48
    Height = 21
    EditLabel.Width = 89
    EditLabel.Height = 13
    EditLabel.BiDiMode = bdLeftToRight
    EditLabel.Caption = 'Einsendernummer:'
    EditLabel.ParentBiDiMode = False
    LabelPosition = lpLeft
    TabOrder = 0
    Text = '0'
  end
  object btnQuery: TButton
    Left = 115
    Top = 74
    Width = 75
    Height = 25
    Caption = 'Abfragen'
    Default = True
    TabOrder = 1
    OnClick = btnQueryClick
  end
  object Memo1: TMemo
    Left = 40
    Top = 120
    Width = 225
    Height = 143
    ReadOnly = True
    TabOrder = 2
  end
  object btnPrint: TButton
    Left = 40
    Top = 346
    Width = 75
    Height = 25
    Caption = '&Drucken'
    Enabled = False
    TabOrder = 3
    OnClick = btnPrintClick
  end
  object edAnzahlKopien: TJvSpinEdit
    Left = 41
    Top = 303
    Width = 40
    Height = 21
    ButtonKind = bkClassic
    EditorEnabled = False
    Increment = 2.000000000000000000
    MaxValue = 10.000000000000000000
    MinValue = 2.000000000000000000
    Value = 2.000000000000000000
    TabOrder = 4
  end
  object btnNeueAbfrage: TButton
    Left = 178
    Top = 346
    Width = 87
    Height = 25
    Caption = '&Neue Abfrage'
    TabOrder = 5
    OnClick = btnNeueAbfrageClick
  end
  object cbBarcode: TCheckBox
    Left = 40
    Top = 280
    Width = 97
    Height = 17
    Caption = 'Barcode'
    TabOrder = 6
  end
  object Button1: TButton
    Left = 115
    Top = 392
    Width = 75
    Height = 25
    Caption = 'S&chlie'#223'en'
    TabOrder = 7
    OnClick = Button1Click
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\user\mygroup\c14d' +
      'b\c14.mdb;Persist Security Info=False;'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 136
  end
  object qryDb: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 168
  end
  object DataSource1: TDataSource
    DataSet = qryDb
    Left = 200
  end
end
