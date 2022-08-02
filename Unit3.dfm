object Form3: TForm3
  Left = 0
  Top = 0
  Caption = 'panel'
  ClientHeight = 729
  ClientWidth = 1350
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1366
    Height = 768
    TabOrder = 0
    object Button1: TButton
      Left = 16
      Top = 696
      Width = 75
      Height = 25
      Caption = 'Button1'
      TabOrder = 0
      OnClick = Button1Click
    end
    object StringGrid1: TStringGrid
      Left = 0
      Top = 0
      Width = 1366
      Height = 690
      TabOrder = 1
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 112
    Top = 696
  end
end
