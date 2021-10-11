object F_Transfert: TF_Transfert
  Left = 228
  Top = 131
  Width = 941
  Height = 574
  Caption = 'Transfer tables EXCEL <--> ACCESS'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  DesignSize = (
    933
    522)
  PixelsPerInch = 120
  TextHeight = 16
  object GroupBox1: TGroupBox
    Left = 10
    Top = 10
    Width = 1130
    Height = 119
    Anchors = [akLeft, akTop, akRight]
    Caption = ' Source EXCEL    et Destination  ACCESS '
    TabOrder = 0
    DesignSize = (
      1130
      119)
    object SpeedButton1: TSpeedButton
      Left = 79
      Top = 20
      Width = 28
      Height = 27
      Caption = '../'
      OnClick = SpeedButton1Click
    end
    object SpeedButton2: TSpeedButton
      Left = 79
      Top = 79
      Width = 28
      Height = 27
      Caption = '../'
      OnClick = SpeedButton2Click
    end
    object Label1: TLabel
      Left = 10
      Top = 20
      Width = 43
      Height = 16
      Caption = 'Source'
    end
    object Label2: TLabel
      Left = 10
      Top = 79
      Width = 67
      Height = 16
      Caption = 'Destination'
    end
    object Destination_ACCESS: TEdit
      Left = 118
      Top = 79
      Width = 1002
      Height = 24
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 0
      Text = 'Destination_ACCESS'
    end
    object Source_EXCEL: TEdit
      Left = 118
      Top = 20
      Width = 1002
      Height = 24
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 1
      Text = 'Source_EXCEL'
    end
  end
  object GroupBox2: TGroupBox
    Left = 10
    Top = 226
    Width = 1137
    Height = 386
    Anchors = [akLeft, akTop, akRight, akBottom]
    Caption = ' Transfer '
    TabOrder = 1
    DesignSize = (
      1137
      386)
    object Label5: TLabel
      Left = 69
      Top = 10
      Width = 99
      Height = 16
      Caption = 'Commande SQL'
    end
    object Label3: TLabel
      Left = 89
      Top = 118
      Width = 228
      Height = 25
      Caption = 'INITIALISER LES BASES'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -23
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object SpeedButton12: TSpeedButton
      Left = 384
      Top = 118
      Width = 41
      Height = 27
      Caption = 'Init'
      OnClick = SpeedButton12Click
    end
    object SQL_cd: TEdit
      Left = 10
      Top = 30
      Width = 1117
      Height = 24
      Anchors = [akLeft, akTop, akRight]
      Enabled = False
      TabOrder = 0
      Text = 
        'SELECT * INTO copie_accessr IN "D:\Projets 2010\Programmes\Mouve' +
        'ment du personnel\BD_Equipements_DMPH.mdb" FROM [Feuil1$]'
      OnDblClick = SQL_cdDblClick
    end
    object ComboBox3: TComboBox
      Left = 30
      Top = 30
      Width = 788
      Height = 24
      ItemHeight = 16
      TabOrder = 1
      Text = 'ComboBox3'
      Visible = False
      OnChange = ComboBox3Change
      Items.Strings = (
        
          'SELECT * INTO copie_access IN "D:\Projets 2010\Programmes\Mouvem' +
          'ent du personnel\BD_Equipements_DMPH.mdb" FROM [Feuil1$]'
        
          'SELECT nom,ville INTO Nouvelle_table IN "D:\Projets 2010\Program' +
          'mes\Mouvement du personnel\Transfert_EXCEL_vers_ACCES\Dossier_AC' +
          'CESS.mdb" FROM [Feuil1$]')
    end
    object GroupBox5: TGroupBox
      Left = 423
      Top = 59
      Width = 425
      Height = 218
      Caption = ' Transfert vers    EXCEL '
      TabOrder = 2
      Visible = False
      OnMouseMove = GroupBox5MouseMove
      object Label10: TLabel
        Left = 62
        Top = 30
        Width = 96
        Height = 16
        Caption = 'Table  ACCESS'
      end
      object Label9: TLabel
        Left = 9
        Top = 59
        Width = 151
        Height = 16
        Caption = 'Table Destination EXCEL'
      end
      object SpeedButton7: TSpeedButton
        Left = 118
        Top = 79
        Width = 267
        Height = 27
        Caption = 'Transfert    ACCESS   vers     EXCEL '
        OnClick = SpeedButton7Click
        OnMouseMove = SpeedButton7MouseMove
      end
      object SpeedButton11: TSpeedButton
        Left = 118
        Top = 148
        Width = 267
        Height = 27
        Caption = 'Insertion table ACCESS vers table  EXCEL'
        OnClick = SpeedButton11Click
        OnMouseMove = SpeedButton11MouseMove
      end
      object SpeedButton8: TSpeedButton
        Left = 118
        Top = 112
        Width = 267
        Height = 27
        Caption = 'Copie de EXCEL vers  EXCEL'
        OnClick = SpeedButton8Click
        OnMouseMove = SpeedButton8MouseMove
      end
      object SpeedButton16: TSpeedButton
        Left = 118
        Top = 187
        Width = 267
        Height = 27
        Caption = 'Insertion table   EXCEL vers table  EXCEL'
        OnClick = SpeedButton16Click
        OnMouseMove = SpeedButton16MouseMove
      end
      object Edit13: TEdit
        Left = 177
        Top = 20
        Width = 149
        Height = 24
        TabOrder = 0
        Text = 'T_ACCESS'
      end
      object edit6: TEdit
        Left = 177
        Top = 49
        Width = 149
        Height = 24
        TabOrder = 1
        Text = 'Nouvelle_Feuille1'
      end
    end
    object GroupBox6: TGroupBox
      Left = 10
      Top = 59
      Width = 385
      Height = 218
      Caption = ' Transfert    ACCESS '
      TabOrder = 3
      Visible = False
      OnMouseMove = GroupBox6MouseMove
      object SpeedButton9: TSpeedButton
        Left = 20
        Top = 92
        Width = 296
        Height = 27
        Caption = 'copie    ACCESS    vers     ACCESS'
        OnClick = SpeedButton9Click
        OnMouseMove = SpeedButton9MouseMove
      end
      object SpeedButton10: TSpeedButton
        Left = 20
        Top = 132
        Width = 296
        Height = 27
        Caption = 'Insertion d'#39'une table EXCEL   vers   ACCESS'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clFuchsia
        Font.Height = -15
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        OnClick = SpeedButton10Click
        OnMouseMove = SpeedButton10MouseMove
      end
      object SpeedButton3: TSpeedButton
        Left = 20
        Top = 59
        Width = 296
        Height = 27
        Caption = 'Transfert   EXCEL   vers   ACCESS'
        Visible = False
        OnClick = SpeedButton3Click
        OnMouseMove = SpeedButton3MouseMove
      end
      object SpeedButton6: TSpeedButton
        Left = 197
        Top = 20
        Width = 100
        Height = 27
        Caption = 'Validation'
        OnClick = SpeedButton6Click
        OnMouseMove = SpeedButton6MouseMove
      end
      object SpeedButton15: TSpeedButton
        Left = 10
        Top = 187
        Width = 326
        Height = 27
        Caption = 'Insertion d'#39'une table ACCESS   vers   ACCESS'
        OnClick = SpeedButton15Click
        OnMouseMove = SpeedButton15MouseMove
      end
      object Nouvelle_table: TEdit
        Left = 30
        Top = 20
        Width = 148
        Height = 24
        TabOrder = 0
        Text = 'Nouvelle_table'
        OnKeyUp = Nouvelle_tableKeyUp
      end
      object CheckBox2: TCheckBox
        Left = 138
        Top = 163
        Width = 119
        Height = 20
        Hint = 'SQL dans le RichEdit'
        Caption = 'sql memo'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = CheckBox2Click
      end
    end
    object RichEdit1: TRichEdit
      Left = 0
      Top = 276
      Width = 1134
      Height = 109
      Anchors = [akLeft, akTop, akRight, akBottom]
      Lines.Strings = (
        
          'INSERT INTO [Nouvelle_table] IN "D:\Projets 2010\Programmes\Mouv' +
          'ement du personnel\Transfert_EXCEL_vers_ACCES\Dossier_ACCESS.mdb' +
          '" SELECT * FROM  [T_ACCESS]'
        
          'INSERT INTO [Nouvelle_table] IN "D:\Projets 2010\Programmes\Mouv' +
          'ement du personnel\Transfert_EXCEL_vers_ACCES\Dossier_ACCESS.mdb' +
          '" SELECT nom,ville FROM  [Feuil1$]')
      TabOrder = 4
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 503
    Width = 933
    Height = 19
    Panels = <>
  end
  object GroupBox3: TGroupBox
    Left = 49
    Top = 138
    Width = 257
    Height = 90
    Caption = ' Liste des tables ACCESS '
    Enabled = False
    TabOrder = 3
    object Label4: TLabel
      Left = 54
      Top = 79
      Width = 3
      Height = 16
    end
    object SpeedButton4: TSpeedButton
      Left = 199
      Top = 23
      Width = 29
      Height = 27
      Caption = 'Init'
      Visible = False
      OnClick = SpeedButton4Click
    end
    object SpeedButton5: TSpeedButton
      Left = 20
      Top = 59
      Width = 99
      Height = 27
      Caption = 'Efface la table'
      Visible = False
      OnClick = SpeedButton5Click
    end
    object ComboBox2: TComboBox
      Left = 10
      Top = 25
      Width = 178
      Height = 24
      DropDownCount = 100
      ItemHeight = 16
      TabOrder = 0
      Text = 'ComboBox2'
    end
  end
  object GroupBox4: TGroupBox
    Left = 482
    Top = 136
    Width = 268
    Height = 90
    Caption = ' Liste des tables  EXCEL '
    Enabled = False
    TabOrder = 4
    object SpeedButton13: TSpeedButton
      Left = 226
      Top = 30
      Width = 29
      Height = 27
      Caption = 'Init'
      OnClick = SpeedButton13Click
    end
    object SpeedButton14: TSpeedButton
      Left = 20
      Top = 59
      Width = 109
      Height = 27
      Caption = 'Efface la table'
      OnClick = SpeedButton14Click
    end
    object ComboBox1: TComboBox
      Left = 20
      Top = 30
      Width = 178
      Height = 24
      DropDownCount = 100
      ItemHeight = 16
      TabOrder = 0
      Text = 'Feuil1$'
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 16
    Top = 40
  end
  object MainMenu1: TMainMenu
    Left = 48
    Top = 40
    object Fini1: TMenuItem
      Caption = 'Fini'
      OnClick = Fini1Click
    end
    object Rpertoire1: TMenuItem
      Caption = 'Repertoire'
      OnClick = Rpertoire1Click
    end
    object BD1: TMenuItem
      Caption = 'BD'
      object Voir1: TMenuItem
        Caption = 'Voir'
        OnClick = Voir1Click
      end
    end
  end
  object Timer1: TTimer
    Interval = 100
    OnTimer = Timer1Timer
    Left = 328
    Top = 128
  end
end
