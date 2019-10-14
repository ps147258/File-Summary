object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'File Summary'
  ClientHeight = 393
  ClientWidth = 713
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    713
    393)
  PixelsPerInch = 96
  TextHeight = 13
  object JvFilenameEdit1: TJvFilenameEdit
    Left = 8
    Top = 8
    Width = 697
    Height = 24
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    Anchors = [akLeft, akTop, akRight]
    ParentFont = False
    TabOrder = 0
    Text = ''
    OnChange = JvFilenameEdit1Change
  end
  object ListView1: TListView
    Left = 8
    Top = 38
    Width = 697
    Height = 347
    Anchors = [akLeft, akTop, akRight, akBottom]
    Columns = <
      item
        Caption = #21517#31281
        Width = 120
      end
      item
        Caption = #20540
        Width = 250
      end
      item
        Caption = 'GUID'
        Width = 250
      end
      item
        Caption = 'PID'
      end>
    DoubleBuffered = True
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = #32048#26126#39636
    Font.Style = []
    FlatScrollBars = True
    GridLines = True
    GroupView = True
    RowSelect = True
    ParentDoubleBuffered = False
    ParentFont = False
    TabOrder = 1
    ViewStyle = vsReport
  end
end
