object FormMain: TFormMain
  Left = 464
  Top = 122
  Width = 900
  Height = 539
  Caption = 'analyzer (version 1.0.0)'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object pnl_bottom: TPanel
    Left = 0
    Top = 467
    Width = 884
    Height = 33
    Align = alBottom
    TabOrder = 0
  end
  object pnl_right: TPanel
    Left = 446
    Top = 0
    Width = 438
    Height = 467
    Align = alRight
    Anchors = [akLeft, akTop, akBottom]
    TabOrder = 1
    object pnl_top_right: TPanel
      Left = 1
      Top = 1
      Width = 436
      Height = 89
      Align = alTop
      TabOrder = 0
    end
    object mmo_error_text: TMemo
      Left = 1
      Top = 90
      Width = 436
      Height = 376
      Align = alClient
      ReadOnly = True
      TabOrder = 1
    end
  end
  object pnl_left: TPanel
    Left = 0
    Top = 0
    Width = 446
    Height = 467
    Align = alClient
    Anchors = [akTop, akBottom]
    TabOrder = 2
    object pnl_top_left: TPanel
      Left = 1
      Top = 1
      Width = 444
      Height = 64
      Align = alTop
      TabOrder = 0
      object btn_add_files: TButton
        Left = 72
        Top = 8
        Width = 113
        Height = 41
        Caption = #1042#1099#1073#1088#1072#1090#1100' '#1092#1072#1081#1083#1099
        TabOrder = 0
        OnClick = btn_add_filesClick
      end
      object btn_run_analyze: TButton
        Left = 248
        Top = 7
        Width = 113
        Height = 41
        Caption = #1040#1085#1072#1083#1080#1079#1080#1088#1086#1074#1072#1090#1100
        Enabled = False
        TabOrder = 1
        OnClick = btn_run_analyzeClick
      end
    end
    object mmo_list_files: TMemo
      Left = 1
      Top = 65
      Width = 444
      Height = 401
      Align = alClient
      ReadOnly = True
      TabOrder = 1
    end
  end
  object dlg_add_files: TOpenDialog
    Left = 471
    Top = 25
  end
end
