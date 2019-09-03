object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = #1047#1072#1075#1088#1091#1079#1082#1072' csv '#1092#1072#1081#1083#1086#1074
  ClientHeight = 412
  ClientWidth = 699
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object StringGrid1: TStringGrid
    Left = 297
    Top = 0
    Width = 402
    Height = 303
    Align = alClient
    BevelKind = bkSoft
    ColCount = 1
    Ctl3D = False
    DefaultRowHeight = 15
    FixedCols = 0
    RowCount = 1
    FixedRows = 0
    ParentCtl3D = False
    TabOrder = 1
  end
  object Panel1: TPanel
    Left = 0
    Top = 303
    Width = 699
    Height = 109
    Align = alBottom
    Ctl3D = False
    ParentCtl3D = False
    TabOrder = 2
    object Label1: TLabel
      Left = 303
      Top = 6
      Width = 27
      Height = 13
      Caption = 'SPID:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGrayText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 259
      Top = 73
      Width = 27
      Height = 13
      Caption = #1041#1072#1079#1072':'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGrayText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 259
      Top = 59
      Width = 41
      Height = 13
      Caption = #1057#1077#1088#1074#1077#1088':'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGrayText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object ok_lb: TLabel
      Left = 23
      Top = 61
      Width = 24
      Height = 23
      Caption = 'OK'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGrayText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object false_lb: TLabel
      Left = 53
      Top = 61
      Width = 52
      Height = 23
      Caption = 'FALSE'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGrayText
      Font.Height = -19
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object count_lb: TLabel
      Left = 259
      Top = 89
      Width = 7
      Height = 14
      Caption = '0'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      Visible = False
    end
    object Load_All_Btn: TButton
      Left = 6
      Top = 30
      Width = 115
      Height = 25
      Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100
      TabOrder = 0
      OnClick = Load_BtnClick
    end
    object tablename_ed: TEdit
      Left = 2
      Top = 3
      Width = 295
      Height = 19
      TabOrder = 1
    end
    object spid_ed: TEdit
      Left = 333
      Top = 3
      Width = 41
      Height = 19
      TabOrder = 2
      Text = '1'
    end
    object Log_memo: TMemo
      Left = 380
      Top = 1
      Width = 318
      Height = 107
      Align = alRight
      Anchors = [akLeft, akTop, akRight, akBottom]
      TabOrder = 3
    end
    object ServerSet_Btn: TButton
      Left = 259
      Top = 28
      Width = 115
      Height = 25
      Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1072' '#1073#1072#1079#1099
      TabOrder = 4
      OnClick = ServerSet_BtnClick
    end
    object RadioGroup1: TRadioGroup
      Left = 127
      Top = 25
      Width = 98
      Height = 64
      TabOrder = 5
    end
    object RB_sel: TRadioButton
      Left = 140
      Top = 40
      Width = 77
      Height = 17
      Caption = #1074#1099#1073#1088#1072#1085#1085#1099#1077
      Checked = True
      TabOrder = 6
      TabStop = True
    end
    object RB_all: TRadioButton
      Left = 140
      Top = 63
      Width = 77
      Height = 17
      Caption = #1074#1089#1077
      TabOrder = 7
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 297
    Height = 303
    Align = alLeft
    Ctl3D = False
    ParentCtl3D = False
    TabOrder = 0
    object FileListBox: TFileListBox
      Left = 1
      Top = 25
      Width = 295
      Height = 277
      Align = alClient
      ItemHeight = 13
      Mask = '*.csv'
      MultiSelect = True
      TabOrder = 0
      OnClick = FileListBoxClick
    end
    object Panel3: TPanel
      Left = 1
      Top = 1
      Width = 295
      Height = 24
      Align = alTop
      TabOrder = 1
      object Patch_Ed: TEdit
        Left = 1
        Top = 1
        Width = 223
        Height = 22
        Align = alClient
        Ctl3D = False
        ParentCtl3D = False
        TabOrder = 0
        Text = 'c:\'
        ExplicitHeight = 19
      end
      object BitBtn1: TBitBtn
        Left = 224
        Top = 1
        Width = 70
        Height = 22
        Align = alRight
        Caption = #1087#1086#1080#1089#1082
        TabOrder = 1
        OnClick = BitBtn1Click
        Kind = bkRetry
      end
    end
  end
  object ADOConn: TADOConnection
    ConnectionString = 
      'Provider=SQLOLEDB.1;Password=123456;Persist Security Info=True;U' +
      'ser ID=dca;Initial Catalog=IFRS09_SecIssue_test;Data Source=finp' +
      'erftest'
    KeepConnection = False
    Provider = 'SQLOLEDB.1'
    AfterConnect = ADOConnAfterConnect
    BeforeDisconnect = ADOConnBeforeDisconnect
    Left = 304
    Top = 24
  end
  object Qdelete: TADOQuery
    Connection = ADOConn
    Parameters = <>
    Left = 304
    Top = 56
  end
  object Qinsert: TADOQuery
    Connection = ADOConn
    ParamCheck = False
    Parameters = <>
    Left = 304
    Top = 88
  end
  object QSchema: TADOQuery
    Connection = ADOConn
    Parameters = <>
    Left = 304
    Top = 120
  end
  object QChkIdentityField: TADOQuery
    Connection = ADOConn
    Parameters = <>
    Left = 304
    Top = 152
  end
end
