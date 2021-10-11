object frmMaitre: TfrmMaitre
  Left = 226
  Top = 127
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'ADSI Group user rename'
  ClientHeight = 562
  ClientWidth = 793
  Color = clBtnFace
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object Gauge1: TGauge
    Left = 8
    Top = 224
    Width = 777
    Height = 25
    Progress = 0
  end
  object gbRecherche: TGroupBox
    Left = 8
    Top = 8
    Width = 465
    Height = 209
    Caption = ' Search '
    TabOrder = 0
    object Label1: TLabel
      Left = 49
      Top = 31
      Width = 23
      Height = 16
      Caption = 'OU:'
    end
    object Label2: TLabel
      Left = 10
      Top = 135
      Width = 73
      Height = 16
      Caption = 'Search filter:'
    end
    object lblProgression: TLabel
      Left = 324
      Top = 241
      Width = 3
      Height = 16
    end
    object edtOU: TEdit
      Left = 90
      Top = 29
      Width = 357
      Height = 24
      TabOrder = 0
      Text = 'LDAP://DC=XXXX,DC=com'
    end
    object gbTypeObjet: TGroupBox
      Left = 90
      Top = 57
      Width = 359
      Height = 56
      Caption = ' Type Objet a chercher '
      TabOrder = 1
      object cbxUser: TCheckBox
        Left = 58
        Top = 21
        Width = 127
        Height = 22
        Caption = 'Utilisateur'
        TabOrder = 0
      end
      object cbxGroupe: TCheckBox
        Left = 258
        Top = 20
        Width = 71
        Height = 23
        Caption = 'Group'
        TabOrder = 1
      end
    end
    object edtFiltre: TEdit
      Left = 90
      Top = 132
      Width = 357
      Height = 24
      TabOrder = 2
      Text = '*'
    end
    object cmdRecherche: TButton
      Left = 89
      Top = 166
      Width = 360
      Height = 27
      Caption = 'Search'
      TabOrder = 3
      OnClick = cmdRechercheClick
    end
  end
  object GroupBox1: TGroupBox
    Left = 478
    Top = 8
    Width = 307
    Height = 209
    Caption = ' Rename '
    TabOrder = 1
    object Label3: TLabel
      Left = 47
      Top = 21
      Width = 36
      Height = 16
      Caption = 'Mask:'
    end
    object Label4: TLabel
      Left = 14
      Top = 99
      Width = 70
      Height = 16
      Caption = 'Remplacer:'
    end
    object Label5: TLabel
      Left = 59
      Top = 141
      Width = 24
      Height = 16
      Caption = 'Par:'
    end
    object Label6: TLabel
      Left = 42
      Top = 63
      Width = 39
      Height = 16
      Caption = 'Indice:'
    end
    object edtMasque: TEdit
      Left = 94
      Top = 21
      Width = 203
      Height = 24
      TabOrder = 0
      Text = '*'
      OnChange = edtRechercheChange
    end
    object edtRecherche: TEdit
      Left = 94
      Top = 97
      Width = 203
      Height = 24
      TabOrder = 1
      OnChange = edtRechercheChange
    end
    object edtRemplacerPar: TEdit
      Left = 94
      Top = 138
      Width = 203
      Height = 24
      TabOrder = 2
      OnChange = edtRechercheChange
    end
    object SpinEdit1: TSpinEdit
      Left = 94
      Top = 55
      Width = 64
      Height = 26
      MaxValue = 0
      MinValue = 0
      TabOrder = 3
      Value = 0
      OnChange = edtRechercheChange
    end
  end
  object strGrid: TStringGrid
    Left = 8
    Top = 256
    Width = 777
    Height = 265
    ColCount = 4
    DefaultRowHeight = 17
    FixedCols = 0
    TabOrder = 2
  end
  object cmdSupprimer: TButton
    Left = 8
    Top = 528
    Width = 121
    Height = 25
    Caption = 'Delete'
    TabOrder = 3
    OnClick = cmdSupprimerClick
  end
  object cmdRenommer: TButton
    Left = 680
    Top = 528
    Width = 105
    Height = 25
    Caption = 'Rename'
    TabOrder = 4
    OnClick = cmdRenommerClick
  end
  object cmdSavLog: TButton
    Left = 575
    Top = 528
    Width = 98
    Height = 25
    Caption = 'Save Log'
    TabOrder = 5
    OnClick = cmdSavLogClick
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = 'fblog'
    Filter = 'fblog | *.fblog'
    Left = 16
    Top = 368
  end
end
