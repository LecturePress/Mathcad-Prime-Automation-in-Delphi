object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'Mathcad COM'
  ClientHeight = 244
  ClientWidth = 571
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object BtnStartMathcad: TButton
    Left = 16
    Top = 16
    Width = 169
    Height = 25
    Caption = 'Start Mathcad Prime'
    TabOrder = 0
    OnClick = BtnStartMathcadClick
  end
  object BtnCloseMathcad: TButton
    Left = 16
    Top = 47
    Width = 169
    Height = 25
    Caption = 'Close Mathcad Prime'
    TabOrder = 1
    OnClick = BtnCloseMathcadClick
  end
  object BtnShowMathcad: TButton
    Left = 16
    Top = 171
    Width = 169
    Height = 25
    Caption = 'Show Mathcad Prime'
    TabOrder = 2
    OnClick = BtnShowMathcadClick
  end
  object BtnHideMathcad: TButton
    Left = 16
    Top = 202
    Width = 169
    Height = 25
    Caption = 'Hide Mathcad Prime'
    TabOrder = 3
    OnClick = BtnHideMathcadClick
  end
  object BtnGetActiveWorksheet: TButton
    Left = 216
    Top = 16
    Width = 169
    Height = 25
    Caption = 'Get Active Worksheet'
    TabOrder = 4
    OnClick = BtnGetActiveWorksheetClick
  end
  object BtnOpenWorksheet: TButton
    Left = 216
    Top = 47
    Width = 169
    Height = 25
    Caption = 'Open Worksheet'
    TabOrder = 5
    OnClick = BtnOpenWorksheetClick
  end
  object BtnOpenNewWorksheet: TButton
    Left = 216
    Top = 78
    Width = 169
    Height = 25
    Caption = 'Open New Worksheet'
    TabOrder = 6
    OnClick = BtnOpenNewWorksheetClick
  end
  object BtnSaveWorksheet: TButton
    Left = 216
    Top = 109
    Width = 169
    Height = 25
    Caption = 'Save Worksheet'
    TabOrder = 7
    OnClick = BtnSaveWorksheetClick
  end
  object BtnSaveAsWorskheet: TButton
    Left = 216
    Top = 140
    Width = 169
    Height = 25
    Caption = 'Save As Worskheet'
    TabOrder = 8
    OnClick = BtnSaveAsWorskheetClick
  end
  object BtnSaveAsMctx: TButton
    Left = 216
    Top = 171
    Width = 169
    Height = 25
    Caption = 'Save As MCTX'
    TabOrder = 9
    OnClick = BtnSaveAsMctxClick
  end
  object BtnCloseWorksheet: TButton
    Left = 216
    Top = 202
    Width = 169
    Height = 25
    Caption = 'Close Worksheet'
    TabOrder = 10
    OnClick = BtnCloseWorksheetClick
  end
  object RGEvents: TRadioGroup
    Left = 416
    Top = 151
    Width = 137
    Height = 76
    Caption = 'Mathcad Prime Events'
    ItemIndex = 1
    Items.Strings = (
      'Enable Events'
      'Disable Events')
    TabOrder = 11
    OnClick = RGEventsClick
  end
end
