object BD_: TBD_
  Left = 272
  Top = 149
  Width = 696
  Height = 554
  Caption = 'Base de donn'#233'es'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    688
    520)
  PixelsPerInch = 96
  TextHeight = 13
  object SpeedButton1: TSpeedButton
    Left = 112
    Top = 8
    Width = 169
    Height = 22
    Caption = 'Initialisation'
    OnClick = SpeedButton1Click
  end
  object SpeedButton6: TSpeedButton
    Left = 376
    Top = 8
    Width = 153
    Height = 22
    Caption = 'D'#233'conexion'
    Visible = False
    OnClick = SpeedButton6Click
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 40
    Width = 673
    Height = 233
    Anchors = [akLeft, akTop, akRight, akBottom]
    Caption = 'EXCEL'
    TabOrder = 0
    Visible = False
    DesignSize = (
      673
      233)
    object Label1: TLabel
      Left = 272
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Tables'
    end
    object Label2: TLabel
      Left = 480
      Top = 16
      Width = 38
      Height = 13
      Caption = 'Champs'
    end
    object SpeedButton2: TSpeedButton
      Left = 320
      Top = 8
      Width = 23
      Height = 22
      Caption = 'O'
      OnClick = SpeedButton2Click
    end
    object SpeedButton3: TSpeedButton
      Left = 360
      Top = 8
      Width = 23
      Height = 22
      Caption = 'F'
      OnClick = SpeedButton3Click
    end
    object Label5: TLabel
      Left = 72
      Top = 16
      Width = 12
      Height = 13
      Caption = 'N'#176
    end
    object DBGrid1: TDBGrid
      Left = 8
      Top = 64
      Width = 657
      Height = 161
      Anchors = [akLeft, akTop, akRight, akBottom]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
    object ComboBox1: TComboBox
      Left = 272
      Top = 32
      Width = 145
      Height = 21
      DropDownCount = 100
      ItemHeight = 13
      TabOrder = 1
      Text = 'Feuil1$'
    end
    object ComboBox2: TComboBox
      Left = 448
      Top = 32
      Width = 145
      Height = 21
      DropDownCount = 100
      ItemHeight = 13
      TabOrder = 2
      Text = 'ComboBox2'
    end
    object DBNavigator1: TDBNavigator
      Left = 8
      Top = 32
      Width = 240
      Height = 25
      TabOrder = 3
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 280
    Width = 673
    Height = 217
    Anchors = [akLeft, akRight, akBottom]
    Caption = 'ACCESS'
    TabOrder = 1
    Visible = False
    DesignSize = (
      673
      217)
    object Label3: TLabel
      Left = 264
      Top = 16
      Width = 32
      Height = 13
      Caption = 'Tables'
    end
    object Label4: TLabel
      Left = 480
      Top = 16
      Width = 38
      Height = 13
      Caption = 'Champs'
    end
    object SpeedButton4: TSpeedButton
      Left = 312
      Top = 8
      Width = 23
      Height = 22
      Caption = 'O'
      OnClick = SpeedButton4Click
    end
    object SpeedButton5: TSpeedButton
      Left = 352
      Top = 8
      Width = 23
      Height = 22
      Caption = 'F'
      OnClick = SpeedButton5Click
    end
    object Label6: TLabel
      Left = 72
      Top = 8
      Width = 12
      Height = 13
      Caption = 'N'#176
    end
    object DBGrid2: TDBGrid
      Left = 8
      Top = 56
      Width = 657
      Height = 144
      Anchors = [akLeft, akTop, akRight, akBottom]
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
    object DBNavigator2: TDBNavigator
      Left = 8
      Top = 24
      Width = 240
      Height = 25
      TabOrder = 1
    end
    object ComboBox3: TComboBox
      Left = 264
      Top = 32
      Width = 145
      Height = 21
      DropDownCount = 100
      ItemHeight = 13
      TabOrder = 2
      Text = 'T_ACCESS'
    end
    object ComboBox4: TComboBox
      Left = 448
      Top = 32
      Width = 145
      Height = 21
      DropDownCount = 100
      ItemHeight = 13
      TabOrder = 3
      Text = 'ComboBox4'
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 501
    Width = 688
    Height = 19
    Panels = <>
  end
  object ADOConnection1: TADOConnection
    AfterConnect = ADOConnection1AfterConnect
    AfterDisconnect = ADOConnection1AfterDisconnect
    Left = 104
    Top = 120
  end
  object ADOConnection2: TADOConnection
    AfterConnect = ADOConnection2AfterConnect
    AfterDisconnect = ADOConnection2AfterDisconnect
    Left = 32
    Top = 368
  end
  object ADOQuery1: TADOQuery
    AfterOpen = ADOQuery1AfterOpen
    AfterClose = ADOQuery1AfterClose
    Parameters = <>
    Left = 136
    Top = 120
  end
  object ADOQuery2: TADOQuery
    AfterOpen = ADOQuery2AfterOpen
    AfterClose = ADOQuery2AfterClose
    Parameters = <>
    Left = 64
    Top = 368
  end
  object DataSource1: TDataSource
    Left = 168
    Top = 120
  end
  object DataSource2: TDataSource
    Left = 104
    Top = 368
  end
end
